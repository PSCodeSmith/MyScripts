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
    An array of EC2 instance IDs for which you want to create snapshots or from which you want to restore instances.

.PARAMETER VolumeSnapshotPairTagKey
    A tag key to associate with volume and snapshot pairs. Defaults to "VolumeSnapshotPair" if not provided.

.EXAMPLE
    .\Manage-EC2InstanceSnapshots.ps1 -CreateSnapshots -InstanceIds i-12345678, i-23456789

    Creates new snapshots for the volumes of the specified instances.

.EXAMPLE
    .\Manage-EC2InstanceSnapshots.ps1 -RestoreSnapshots -InstanceIds i-12345678, i-23456789 -VolumeSnapshotPairTagKey "CustomTag"

    Restores the specified instances from snapshots associated with the given custom tag key.

.INPUTS
    None.

.OUTPUTS
    String.
    - "NotAllVolumesHaveTagOrSnapshots" if the restoration process fails due to missing tags or snapshots.
    - "Error" if an error occurs during the restoration process.

.NOTES
    This script relies on AWS Tools for PowerShell cmdlets. Ensure you have the necessary AWS credentials and permissions.
    Functions:
    - Restore-EC2InstanceFromSnapshots: Restores EC2 instances from existing snapshots.
    - New-EC2InstanceVolumeSnapshots: Creates new snapshots for EC2 instance volumes.
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

function Wait-UntilInstanceStopped {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InstanceId
    )

    Write-Verbose "Waiting for instance [$InstanceId] to reach the 'stopped' state."
    while ((Get-EC2InstanceStatus -InstanceId $InstanceId -IncludeAllInstance $true).InstanceState.Name.Value -ne "stopped") {
        Start-Sleep -Seconds 5
    }
    Write-Verbose "Instance [$InstanceId] is now 'stopped'."
}

function Get-AssociatedSnapshot {
    param (
        [Parameter(Mandatory=$true)]
        $Volume,
        [Parameter(Mandatory=$true)]
        $Snapshots,
        [Parameter(Mandatory=$true)]
        [string]$VolumeSnapshotPairTagKey
    )
    # Identify the unique PairId from the volume tags
    $pairId = ($Volume.Tags | Where-Object { $_.Key -eq $VolumeSnapshotPairTagKey }).Value
    if ($null -eq $pairId) {
        return $null
    }
    return $Snapshots | Where-Object {
        ($_.Tags | Where-Object { $_.Key -eq $VolumeSnapshotPairTagKey }).Value -eq $pairId
    }
}

function AllVolumesHaveSnapshots {
    param (
        [Parameter(Mandatory=$true)]
        [System.Collections.Generic.List[Amazon.EC2.Model.Volume]]$Volumes,
        [Parameter(Mandatory=$true)]
        [System.Collections.Generic.List[Amazon.EC2.Model.Snapshot]]$Snapshots,
        [Parameter(Mandatory=$true)]
        [string]$VolumeSnapshotPairTagKey
    )

    foreach ($volume in $Volumes) {
        $snapshot = Get-AssociatedSnapshot -Volume $volume -Snapshots $Snapshots -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey
        if ($null -eq $snapshot) {
            return $false
        }
    }
    return $true
}

function Restore-EC2InstanceFromSnapshots {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string[]]$InstanceIds,

        [Parameter(Mandatory=$true)]
        [string]$VolumeSnapshotPairTagKey
    )

    begin {
        Write-Verbose "Starting restore process for InstanceIds: $($InstanceIds -join ', ') with tag key: $VolumeSnapshotPairTagKey"
    }

    process {
        foreach ($InstanceId in $InstanceIds) {
            Write-Verbose "Processing instance [$InstanceId]."
            try {
                $volumes = Get-EC2Volume -Filter @{Name="attachment.instance-id"; Values=$InstanceId} -ErrorAction Stop
                $snapshots = Get-EC2Snapshot -Filter @{Name="tag-key"; Values=$VolumeSnapshotPairTagKey} -ErrorAction Stop

                if (-not (AllVolumesHaveSnapshots -Volumes $volumes -Snapshots $snapshots -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey)) {
                    Write-Verbose "Not all volumes for [$InstanceId] have the required tag or corresponding snapshots."
                    return "NotAllVolumesHaveTagOrSnapshots"
                }

                # Stop the EC2 instance before modification
                Stop-EC2Instance -InstanceId $InstanceId -ErrorAction Stop
                Wait-UntilInstanceStopped -InstanceId $InstanceId

                # Replace volumes with those created from snapshots
                foreach ($volume in $volumes) {
                    $deviceName = $volume.Attachments.Device
                    $snapshot = Get-AssociatedSnapshot -Volume $volume -Snapshots $snapshots -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey

                    if ($null -eq $snapshot) {
                        # This case should not occur due to earlier checks
                        Write-Error "No corresponding snapshot found for volume [$($volume.VolumeId)] of instance [$InstanceId]."
                        return "Error"
                    }

                    # Detach existing volume
                    Dismount-EC2Volume -VolumeId $volume.VolumeId -Force -ErrorAction Stop

                    # Create a new volume from the snapshot
                    $newVolume = New-EC2Volume -SnapshotId $snapshot.SnapshotId -AvailabilityZone $volume.AvailabilityZone -Size $volume.Size -ErrorAction Stop

                    # Wait until the new volume is available
                    while ((Get-EC2Volume -VolumeId $newVolume.VolumeId).State -ne "available") {
                        Write-Verbose "Waiting for new volume [$($newVolume.VolumeId)] to become available."
                        Start-Sleep -Seconds 5
                    }

                    # Attach the new volume
                    Add-EC2Volume -InstanceId $InstanceId -VolumeId $newVolume.VolumeId -Device $deviceName -ErrorAction Stop

                    # Copy tags from snapshot to the new volume
                    if ($snapshot.Tags) {
                        $tagSpecification = $snapshot.Tags | ForEach-Object {
                            New-EC2Tag -Resources $newVolume.VolumeId -Tags $_ -ErrorAction Stop
                        }
                    }

                    # Delete the old, detached volume
                    Remove-EC2Volume -VolumeId $volume.VolumeId -Force -ErrorAction Stop
                }

                # Set DeleteOnTermination = True for attached volumes
                $deviceMappings = (Get-EC2InstanceAttribute -InstanceId $InstanceId -Attribute blockDeviceMapping | 
                                   Select-Object -ExpandProperty BlockDeviceMappings)
                foreach ($mapping in $deviceMappings) {
                    Edit-EC2InstanceAttribute -InstanceId $InstanceId -BlockDeviceMapping @{
                        DeviceName = $mapping.DeviceName
                        Ebs = @{ DeleteOnTermination = $true }
                    } -ErrorAction Stop
                }

                # Start the instance back up
                Start-EC2Instance -InstanceId $InstanceId -ErrorAction Stop
            }
            catch {
                Write-Error "An error occurred while restoring instance [$InstanceId]: $($_.Exception.Message)"
                return "Error"
            }
        }
    }

    end {
        Write-Verbose "Finished restoring instances: $($InstanceIds -join ', ') from snapshots."
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
        Write-Verbose "Starting snapshot creation for instances: $($InstanceIds -join ', ') with tag key: $VolumeSnapshotPairTagKey"
    }

    process {
        foreach ($InstanceId in $InstanceIds) {
            Write-Verbose "Creating snapshots for instance [$InstanceId]."
            try {
                $volumes = Get-EC2Volume -Filter @{Name="attachment.instance-id"; Values=$InstanceId} -ErrorAction Stop

                foreach ($volume in $volumes) {
                    $uniquePairId = [Guid]::NewGuid().ToString()
                    $snapshotDescription = "Snapshot for instance $InstanceId, volume $($volume.VolumeId)"

                    # Create the snapshot
                    $snapshot = New-EC2Snapshot -VolumeId $volume.VolumeId -Description $snapshotDescription -ErrorAction Stop

                    # Tag both snapshot and volume with the pair ID
                    New-EC2Tag -Resources $snapshot.SnapshotId -Tags @{Key=$VolumeSnapshotPairTagKey; Value=$uniquePairId} -ErrorAction Stop
                    New-EC2Tag -Resources $volume.VolumeId -Tags @{Key=$VolumeSnapshotPairTagKey; Value=$uniquePairId} -ErrorAction Stop
                }
            }
            catch {
                Write-Error "An error occurred while creating snapshots for instance [$InstanceId]: $($_.Exception.Message)"
                continue
            }
        }
    }

    end {
        Write-Verbose "Finished creating volume snapshots for instances: $($InstanceIds -join ', ')."
    }
}

# Main execution logic
if ($CreateSnapshots) {
    New-EC2InstanceVolumeSnapshots -InstanceIds $InstanceIds -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey
} elseif ($RestoreSnapshots) {
    $result = Restore-EC2InstanceFromSnapshots -InstanceIds $InstanceIds -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey
    if ($result -eq "NotAllVolumesHaveTagOrSnapshots") {
        $createSnapshotsResponse = Read-Host -Prompt "Not all volumes have the required tags/snapshots. Do you want to create them now? (Yes/No)"
        $affirmativeValues = @("Y", "YES")
        if ($createSnapshotsResponse.ToUpper() -in $affirmativeValues) {
            New-EC2InstanceVolumeSnapshots -InstanceIds $InstanceIds -VolumeSnapshotPairTagKey $VolumeSnapshotPairTagKey
        } else {
            Write-Host "Exiting without creating snapshots."
        }
    }
}