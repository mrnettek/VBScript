' Description: Returns general information about how Terminal Services has been configured on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_TSGeneralSetting")

For Each objItem in colItems
    Wscript.Echo "Comment: " & objItem.Comment
    Wscript.Echo "Minimum encryption level: " & objItem.MinEncryptionLevel
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Terminal name: " & objItem.TerminalName
    Wscript.Echo "Terminal protocol: " & objItem.TerminalProtocol
    Wscript.Echo "Transport: " & objItem.Transport
    Wscript.Echo "Windows authentication: " & objItem.WindowsAuthentication
    Wscript.Echo
Next

