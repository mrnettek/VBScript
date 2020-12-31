On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSClientSetting",,48)
For Each objItem in colItems
    Wscript.Echo "AudioMapping: " & objItem.AudioMapping
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ClipboardMapping: " & objItem.ClipboardMapping
    Wscript.Echo "ColorDepth: " & objItem.ColorDepth
    Wscript.Echo "ColorDepthPolicy: " & objItem.ColorDepthPolicy
    Wscript.Echo "COMPortMapping: " & objItem.COMPortMapping
    Wscript.Echo "ConnectClientDrivesAtLogon: " & objItem.ConnectClientDrivesAtLogon
    Wscript.Echo "ConnectionPolicy: " & objItem.ConnectionPolicy
    Wscript.Echo "ConnectPrinterAtLogon: " & objItem.ConnectPrinterAtLogon
    Wscript.Echo "DefaultToClientPrinter: " & objItem.DefaultToClientPrinter
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DriveMapping: " & objItem.DriveMapping
    Wscript.Echo "LPTPortMapping: " & objItem.LPTPortMapping
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "TerminalName: " & objItem.TerminalName
    Wscript.Echo "WindowsPrinterMapping: " & objItem.WindowsPrinterMapping
Next

