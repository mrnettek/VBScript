strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

strDrive = "Win32_LogicalDisk.DeviceID=" & chr(34) & "C:" & chr(34)

Set colQuotas = objWMIService.ExecQuery _
    ("Select * from Win32_DiskQuota Where QuotaVolume = '" & strDrive & "'")

For Each objQuota in colQuotas
    Wscript.Echo "User: "& objQuota.User      
    Wscript.Echo "Disk Space Used: "& objQuota.DiskSpaceUsed
Next
  


