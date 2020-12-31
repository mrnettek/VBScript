On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_WindowsProductActivation",,48)
For Each objItem in colItems
    Wscript.Echo "ActivationRequired: " & objItem.ActivationRequired
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "IsNotificationOn: " & objItem.IsNotificationOn
    Wscript.Echo "ProductID: " & objItem.ProductID
    Wscript.Echo "RemainingEvaluationPeriod: " & objItem.RemainingEvaluationPeriod
    Wscript.Echo "RemainingGracePeriod: " & objItem.RemainingGracePeriod
    Wscript.Echo "ServerName: " & objItem.ServerName
    Wscript.Echo "SettingID: " & objItem.SettingID
Next

