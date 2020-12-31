' Description: Changes the warning threshold and default quota limits for drive C.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colQuotaSettings = objWMIService.ExecQuery _
    ("Select * from Win32_QuotaSetting")

For Each objQuotaSetting in colQuotaSettings
    objQuotaSetting.DefaultLimit = 10000000
    objQuotaSetting.DefaultWarningLimit = 9000000
    objQuotaSetting.Put_
Next

