' Description: Deletes a disk quota entry for the user fabrikam\bob.


strDrive = "C:"
strDomain = "fabrikam"
strUser = "bob"

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objQuota = objWMIService.Get _
    ("Win32_DiskQuota.QuotaVolume='Win32_LogicalDisk.DeviceID=""" " & _
        " & strDrive & """',User='Win32_Account.Domain=""" " & _
            " & strDomain & """,Name=""" & strUser & """'")

objQuota.Delete_

