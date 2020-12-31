' Description: Ensures that Autochk.exe will not run against drive C the next time the computer reboots, even if the "dirty bit" has been set on drive C.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objDisk = objWMIService.Get("Win32_LogicalDisk")

errReturn = objDisk.ExcludeFromAutoChk(Array("C:"))

