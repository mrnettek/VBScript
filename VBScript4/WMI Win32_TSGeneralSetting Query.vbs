On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSGeneralSetting",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Comment: " & objItem.Comment
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "MinEncryptionLevel: " & objItem.MinEncryptionLevel
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "TerminalName: " & objItem.TerminalName
    Wscript.Echo "TerminalProtocol: " & objItem.TerminalProtocol
    Wscript.Echo "Transport: " & objItem.Transport
    Wscript.Echo "WindowsAuthentication: " & objItem.WindowsAuthentication
Next

