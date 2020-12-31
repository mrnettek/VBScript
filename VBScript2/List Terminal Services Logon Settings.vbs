' Description: Returns information about the Terminal Service logon policies configured on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSLogonSetting")

For Each objItem in colItems
    Wscript.Echo "Client logon information policy: " & _
        objItem.ClientLogonInfoPolicy
    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "Prompt for password: " & objItem.PromptForPassword
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Terminal name: " & objItem.TerminalName
    Wscript.Echo "User name: " & objItem.UserName
    Wscript.Echo
Next

