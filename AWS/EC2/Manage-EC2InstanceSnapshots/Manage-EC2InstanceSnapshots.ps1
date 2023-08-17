<#
.SYNOPSIS
    Manage Amazon EC2 instance snapshots by creating new snapshots or restoring instances from existing snapshots.

.DESCRIPTION
    This script provides functionalities to create snapshots for EC2 instances' volumes and restore instances from those snapshots.
    You can either create new snapshots for specified instances or restore instances from existing snapshots.

.PARAMETER CreateSnapshots
    A switch to indicate if the action is to create new snapshots. This switch is mutually exclusive with RestoreSnapshots.

.PARAMETER RestoreSnapshots
    A switch to indicate if the action is to restore instances from existing snapshots. This switch is mutually exclusive with CreateSnapshots.

.PARAMETER InstanceIds
    An array of EC2 instance IDs that you want to create snapshots for or restore from snapshots.

.PARAMETER VolumeSnapshotPairTagKey
    A tag key to associate with volume and snapshot pairs. Defaults to "VolumeSnapshotPair" if not provided.

.EXAMPLE
    .\Manage-EC2InstanceSnapshots.ps1 -CreateSnapshots -InstanceIds i-12345678, i-23456789

    This example creates new snapshots for the volumes of instances with the specified IDs.

.EXAMPLE
    .\Manage-EC2InstanceSnapshots.ps1 -RestoreSnapshots -InstanceIds i-12345678, i-23456789 -VolumeSnapshotPairTagKey "CustomTag"

    This example restores the instances with the specified IDs from the snapshots associated with the custom tag key.

.INPUTS
    None.

.OUTPUTS
    String.
    - "NotAllVolumesHaveTagOrSnapshots" if the restoration process fails due to missing tags or snapshots.
    - "Error" if an error occurs during the restoration process.

.NOTES
    The script includes two main functions:
    - Restore-EC2InstanceFromSnapshots: Restores EC2 instances from existing snapshots.
    - New-EC2InstanceVolumeSnapshots: Creates new snapshots for EC2 instance volumes.

    Ensure that the necessary AWS cmdlets are installed and available, and that you have the required permissions to perform the actions in your AWS environment.
#>
param (
    [Parameter(Mandatory=$true, ParameterSetName="CreateSnapshots")]
    [switch]$CreateSnapshots,

    [Parameter(Mandatory=$true, ParameterSetName="RestoreSnapshots")]
    [switch]$RestoreSnapshots,

    [Parameter(Mandatory=$true)]
    [string[]]$InstanceIds,

    [string]$VolumeSnapshotPairTagKey = "VolumeSnapshotPair"
)

function Restore-EC2InstanceFromSnapshots {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$InstanceIds,

        [Parameter(Mandatory=$true)]
        [string]$VolumeSnapshotPairTagKey
    )

    begin {
        Write-Verbose "Starting Restore-EC2InstanceFromSnapshots for InstanceId: $InstanceId and VolumeSnapshotPairTagKey: $VolumeSnapshotPairTagKey"
    }

    process {
        foreach ($InstanceId in $InstanceIds) {
            try {
                # Get volumes attached to the instance
                $volumes = Get-EC2Volume -Filter @(@{Name="attachment.instance-id"; Values=$InstanceId}) -ErrorAction Stop

                # Get all snapshots with the specified tag key
                $snapshots = Get-EC2Snapshot -Filter @(@{Name="tag-key"; Values=$VolumeSnapshotPairTagKey}) -ErrorAction Stop

                # Check if all volumes have the VolumeSnapshotPairTagKey and corresponding snapshots exist
                $allVolumesHaveTagAndSnapshots = $true
                foreach ($volume in $volumes) {
                    $pairId = ($volume.Tags | Where-Object { $_.Key -eq $VolumeSnapshotPairTagKey }).Value
                    if ($null -eq $pairId -or -not ($snapshots | Where-Object { ($_.Tags | Where-Object { $_.Key -eq $VolumeSnapshotPairTagKey }).Value -eq $pairId })) {
                        $allVolumesHaveTagAndSnapshots = $false
                        break
                    }
                }

                if ($allVolumesHaveTagAndSnapshots) {
                    # Stop the EC2 instance
                    Stop-EC2Instance -InstanceId $InstanceId

                    # Wait for the instance to stop
                    while ((Get-EC2InstanceStatus -IncludeAllInstance $true -InstanceId $instanceId).InstanceState.Name.Value -ne "stopped") {
                        Write-Host "Waiting for our instance [$instanceId] to reach the state of [stopped]..." -ForegroundColor Blue
                        Start-Sleep -Seconds 5
                    }
                    Write-Host "[$instanceId] is in the [stopped] state" -ForegroundColor Blue

                    # Detach existing volumes and create/attach new volumes from snapshots
                    foreach ($volume in $volumes) {
                        $volumeId = $volume.VolumeId
                        $deviceName = $volume.Attachments.Device

                        # Get the corresponding snapshot for the current volume
                        $pairId = ($volume.Tags | Where-Object { $_.Key -eq $VolumeSnapshotPairTagKey }).Value
                        $snapshot = $snapshots | Where-Object { ($_.Tags | Where-Object { $_.Key -eq $VolumeSnapshotPairTagKey }).Value -eq $pairId }

                        # Detach the volume
                        Dismount-EC2Volume -VolumeId $volumeId -Force

                        # Create a new volume from the snapshot
                        $newVolume = New-EC2Volume -SnapshotId $snapshot.SnapshotId -AvailabilityZone $volume.AvailabilityZone -Size $volume.Size

                        # Attach the new volume to the instance
                        while ((Get-EC2Volume -VolumeId $newVolume.VolumeId).State -ne "available") {
                            Start-Sleep -Seconds 5
                        }
                        Add-EC2Volume -InstanceId $InstanceId -VolumeId $newVolume.VolumeId -Device $deviceName

                        # Copy the tags from the snapshot to the new volume
                        $tags = $snapshot.Tags
                        if ($null -ne $tags) {
                            $tagSpecification = New-Object Amazon.EC2.Model.TagSpecification
                            $tagSpecification.ResourceType = "volume"
                            $tagSpecification.Tags.AddRange($tags)
                            New-EC2Tag -Resources $newVolume.VolumeId -Tags $tagSpecification.Tags
                        }

                        # Delete the detached volume
                        Remove-EC2Volume -VolumeId $volumeId -Force
                    }

                    # Set the Delete on Termination option on all attached volumes
                    $deviceids = (Get-EC2InstanceAttribute -InstanceId $InstanceId -Attribute blockDeviceMapping | Select -ExpandProperty BlockDeviceMappings).DeviceName
                    foreach($deviceid in $deviceids)
                    {
                        Edit-EC2InstanceAttribute -InstanceId $InstanceId -BlockDeviceMapping @{DeviceName=$deviceid;Ebs=@{DeleteOnTermination=$true}};
                    }

                    # Start the EC2 instance
                    Start-EC2Instance -InstanceId $InstanceId

                } else {
                    Write-Verbose "Not all volumes have the required tag or corresponding snapshots."
                    return "NotAllVolumesHaveTagOrSnapshots"
                }
            }
            catch {
                Write-Error "An error occurred during Restore-EC2InstanceFromSnapshots: $_"
                return "Error"
            }
        }
    }
    end {
        Write-Verbose "Finished Restore-EC2InstanceFromSnapshots for InstanceId: $InstanceId and VolumeSnapshotPairTagKey: $VolumeSnapshotPairTagKey"
    }
}

function New-EC2InstanceVolumeSnapshots {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$InstanceIds,

        [Parameter(Mandatory=$true)]
        [string]$VolumeSnapshotPairTagKey
    )

    begin {
        Write-Verbose "Starting New-EC2InstanceVolumeSnapshots for InstanceIds: $($InstanceIds -join ', ') and VolumeSnapshotPairTagKey: $VolumeSnapshotPairTagKey"
    }

    process {
        foreach ($InstanceId in $InstanceIds) {
            try {
                # Get volumes attached to the instance
                $volumes = Get-EC2Volume -Filter @(@{Name="attachment.instance-id"; Values=$InstanceId}) -ErrorAction Stop

                # Create snapshots and assign the VolumeSnapshotPairTagKey with a unique value
                foreach ($volume in $volumes) {
                    # Generate a unique ID for the VolumeSnapshotPairTagKey
                    $uniquePairId = [Guid]::NewGuid().ToString()

                    # Create a snapshot
                    $snapshot = New-EC2Snapshot -VolumeId $volume.VolumeId -Description "Snapshot for instance $InstanceId, volume $($volume.VolumeId)" -ErrorAction Stop

                    # Add the VolumeSnapshotPairTagKey to the snapshot
                    New-EC2Tag -Resources $snapshot.SnapshotId -Tags @{Key=$VolumeSnapshotPairTagKey; Value=$uniquePairId} -ErrorAction Stop

                    # Add the VolumeSnapshotPairTagKey to the volume
                    New-EC2Tag -Resources $volume.VolumeId -Tags @{Key=$VolumeSnapshotPairTagKey; Value=$uniquePairId} -ErrorAction Stop
                }
            }
            catch {
                Write-Error "An error occurred during New-EC2InstanceVolumeSnapshots for InstanceId: $InstanceId : $_"
                continue
            }
        }
    }

    end {
        Write-Verbose "Finished New-EC2InstanceVolumeSnapshots for InstanceIds: $($InstanceIds -join ', ') and VolumeSnapshotPairTagKey: $VolumeSnapshotPairTagKey"
    }
}


if ($CreateSnapshots) {
    New-EC2InstanceVolumeSnapshots -InstanceIds $InstanceIds -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey
} elseif ($RestoreSnapshots) {
    $result = Restore-EC2InstanceFromSnapshots -InstanceIds $InstanceIds -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey
    if ($result -eq "NotAllVolumesHaveTagOrSnapshots") {
        $createSnapshotsResponse = Read-Host -Prompt "Not all volumes have the required tag or corresponding snapshots. Do you want to create them? (Yes/No)"
        $acceptedAffirmativeValues = @("Y", "Yes")
        if ($createSnapshotsResponse.ToUpper() -in $acceptedAffirmativeValues) {
            New-EC2InstanceVolumeSnapshots -InstanceIds $InstanceIds -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey
        }
        else {
            Write-Host "Exiting..."
        }
    }
}
