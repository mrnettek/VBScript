On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_QuotaSetting",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DefaultLimit: " & objItem.DefaultLimit
    Wscript.Echo "DefaultWarningLimit: " & objItem.DefaultWarningLimit
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "ExceededNotification: " & objItem.ExceededNotification
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "VolumePath: " & objItem.VolumePath
    Wscript.Echo "WarningExceededNotification: " & objItem.WarningExceededNotification
Next

