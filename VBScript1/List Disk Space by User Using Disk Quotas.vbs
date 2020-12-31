' Description: Uses disk quotas to report the amount of disk space being used by each individual user on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colQuotas = objWMIService.ExecQuery("Select * from Win32_DiskQuota")

For Each objQuota in colQuotas
    Wscript.Echo "Volume: "& objQuota.QuotaVolume
    Wscript.Echo "User: "& objQuota.User      
    Wscript.Echo "Disk Space Used: "& objQuota.DiskSpaceUsed
Next

