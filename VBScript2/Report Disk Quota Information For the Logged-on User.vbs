strComputer = "."

Set objNetwork = CreateObject("Wscript.Network")
strUser = objNetwork.UserName
strDomain = objNetwork.UserDomain

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colDisks = objWMIService.ExecQuery _
    ("Select * from Win32_QuotaSetting Where State = 1")

For Each objDisk in colDisks
    strDrive = objDisk.VolumePath
    strDrive = Replace(strDrive, "\", "")

    Set objQuota = objWMIService.Get _
        ("Win32_DiskQuota.QuotaVolume='Win32_LogicalDisk.DeviceID=" & chr(34) & strDrive & chr(34) & "'," & _
            "User='Win32_Account.Domain=" & chr(34) & strDomain & chr(34) & _
                ",Name=" & chr(34) & strUser & chr(34) & "'")

    Wscript.Echo "Drive: " & objDisk.VolumePath
    Wscript.Echo "Disk Space Used: " & Int(objQuota.DiskSpaceUsed / 1048576) & " megabytes"
    Wscript.Echo "Quota Limit: " & Int(objQuota.Limit  / 1048576) & " megabytes"
    Wscript.Echo "Disk Space Remaining: " & Int((objQuota.Limit - objQuota.DiskSpaceUsed) / 1048576) & _
        " megabytes"
    Wscript.Echo
Next
  


