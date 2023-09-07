
# Splunk Queries for Domain Administrators

This is a collection of useful Splunk queries that can help Domain Admins in monitoring and maintaining the security and operations of their domain.

## Table of Contents

1. [Failed Logon Attempts](#failed-logon-attempts)
2. [Successful Administrative Logons](#successful-administrative-logons)
3. [Account Lockouts](#account-lockouts)
4. [Monitoring USB Device Usage](#monitoring-usb-device-usage)
5. [Monitoring Group Policy Changes](#monitoring-group-policy-changes)
6. [Newly Created Accounts](#newly-created-accounts)
7. [Service Creation Events](#service-creation-events)
8. [PowerShell Commands Execution](#powershell-commands-execution)
9. [Permission Changes to Important Shares or Files](#permission-changes-to-important-shares-or-files)
10. [Monitoring DNS Queries](#monitoring-dns-queries)
11. [Privileged Group Modifications](#privileged-group-modifications)

---

### Failed Logon Attempts

Monitor failed logon attempts to catch suspicious activities.

```splunk
index="wineventlog" sourcetype="WinEventLog:Security" EventCode=4625 earliest=-1d
| stats count by src_ip, Account_Name
| sort - count
```

---

### Successful Administrative Logons

Track successful logons to privileged accounts.

```splunk
index="wineventlog" sourcetype="WinEventLog:Security" EventCode=4624 Logon_Type=10 Account_Name!="*$" earliest=-1d
| stats count by Account_Name, ComputerName
```

---

### Account Lockouts

Identify accounts that are getting locked out frequently.

```splunk
index="wineventlog" sourcetype="WinEventLog:Security" EventCode=4740 earliest=-1d
| stats count by Account_Name, ComputerName
| sort - count
```

---

### Monitoring USB Device Usage

Keep an eye on USB device connections to domain computers.

```splunk
index="wineventlog" sourcetype="WinEventLog:System" EventCode=6416 earliest=-1d
| stats count by host, EventCode, DeviceName, DeviceManufacturer
```

---

### Monitoring Group Policy Changes

Track changes made to Group Policies.

```splunk
index="wineventlog" sourcetype="WinEventLog:Security" EventCode=5136 earliest=-7d
| stats count by ObjectName, host
```

---

### Newly Created Accounts

Identify newly created user accounts in the domain.

```splunk
index="wineventlog" sourcetype="WinEventLog:Security" EventCode=4720 earliest=-7d
| stats count by Account_Name, ComputerName
```

---

### Service Creation Events

Watch for the creation of new services, which could be a sign of compromise.

```splunk
index="wineventlog" sourcetype="WinEventLog:Security" EventCode=4697 earliest=-7d
| stats count by ObjectName, host
```

---

### PowerShell Commands Execution

Monitor PowerShell command execution for potential malicious activities.

```splunk
index="powershell" EventCode=4104 earliest=-1d
| stats count by ScriptBlockText, ComputerName
```

---

### Permission Changes to Important Shares or Files

Track changes to permissions on critical shares or files.

```splunk
index="wineventlog" sourcetype="WinEventLog:Security" EventCode=4670 earliest=-7d
| stats count by ObjectName, host
```

---

### Monitoring DNS Queries

Identify potentially harmful or excessive DNS queries.

```splunk
index="dns" sourcetype="dns:query" earliest=-1h
| stats count by query, src_ip
| sort - count
```

---

### Privileged Group Modifications

Monitor privileged group modifications for maintaining security.

```splunk
index="wineventlog" sourcetype="WinEventLog:Security" 
(EventCode IN (4728, 4729, 4756, 4757)) host=*DC-* earliest=-10m
| eval Group_Name=mvindex(split(Group_Name, " "), -1)
| search Group_Name IN ("Administrators", "Schema Admins", "Enterprise Admins", "Domain Admins", "Server Admins", "Account Operators", "Server Operators", "Backup Operators", "Cryptographic Operators", "Cert Publishers", "DHCP Administrators", "DnsAdmins", "Print Operators", "Replicator", "Group Policy Creator Owners", "Pre-Windows 2000 Compatible Access", "Protected Users", "Exchange Windows Permissions", "Exchange Trusted Subsystem")
| eval Time=strftime(_time, "%m/%d %H:%M %Z"), 
      EventCode=case(EventCode==4728, "4728 [+] GG Sec", 
                     EventCode==4729, "4729 [-] GG Sec", 
                     EventCode==4756, "4756 [+] UG Sec", 
                     EventCode==4757, "4757 [-] UG Sec")
| rex field=_raw "Subject:\s+[^
]+\s+Account Name:\s+(?<ActionBy>.*)"
| rex field=_raw "Member:\s+\w+\s\w+:\s+(?<Member>.*)"
| rex field=_raw "Group:\s+[^
]+\s+Group Name:\s+(?<Group>.*)"
| table Time, Group, ActionBy, Member, host, EventCode
| rename Group AS "Group Name", 
          ActionBy AS "Action By", 
          Member AS "Action To", 
          host AS "Computer", 
          EventCode AS "Event Code"
| sort - _time
```

