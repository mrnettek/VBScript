On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TSLogonSetting",,48)
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "ClientLogonInfoPolicy: " & objItem.ClientLogonInfoPolicy
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "PromptForPassword: " & objItem.PromptForPassword
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "TerminalName: " & objItem.TerminalName
    Wscript.Echo "UserName: " & objItem.UserName
Next

