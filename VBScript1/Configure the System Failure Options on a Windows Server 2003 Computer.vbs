Const SMALL_MEMORY_DUMP = 3

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colRecoveryOptions = objWMIService.ExecQuery _
    ("Select * From Win32_OSRecoveryConfiguration")

For Each objOption in colRecoveryOptions 
    objOption.DebugInfoType = SMALL_MEMORY_DUMP
    objOption.AutoReboot = FALSE
    objOption.SendAdminAlert = FALSE
    objOption.Put_
Next
  


