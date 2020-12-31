' Description: Modifies the disk quota limit for the user fabrikam\bob.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objAccount = objWMIService.Get _
    ("Win32_Account.Domain='fabrikam',Name='bob'")
Set objDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='C:'")
Set objQuota = objWMIService.Get _
    ("Win32_DiskQuota.QuotaVolume= " & _
      "'Win32_LogicalDisk.DeviceID=""C:""'," & _ 
         "User='Win32_Account.Domain=""fabrikam"",Name=""bob""'")

objQuota.Limit = 11111111
objQuota.Put_

