' Description: Uses disk quotas to return information about disk space usage per-user for drive C on the local computer.


Set colDiskQuotas = CreateObject("Microsoft.DiskQuota.1")
colDiskQuotas.Initialize "C:\", True
 
For Each objUser in colDiskQuotas
    Wscript.Echo "Logon name: " & objUser.LogonName
    Wscript.Echo "Quota used: " & objUser.QuotaUsed
Next

