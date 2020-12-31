' Description: Retrieve a user's Terminal Services Profile settings.


Set objUser = GetObject("LDAP://cn=youngrob,ou=r&d,dc=fabrikam,dc=com") 
 
WScript.Echo objUser.Name & " Terminal Services Profile Settings"
WScript.Echo "--------------------------------------------------"
 
WScript.Echo "Allow Logon: " & objUser.AllowLogon
WScript.Echo "Terminal Services Home Directory: " & _
    objUser.TerminalServicesHomeDirectory
WScript.Echo "Terminal Services Home Drive: " & _
    objUser.TerminalServicesHomeDrive
WScript.Echo "Terminal Services Profile Path: " & _
    objUser.TerminalServicesProfilePath
 
WScript.Echo "Enable Remote Control: " & objUser.EnableRemoteControl
 
WScript.Echo "Broken Connection Action: " & objUser.BrokenConnectionAction
WScript.Echo "Max Connection Time: " & objUser.MaxConnectionTime
WScript.Echo "Max Disconnection Time: " & objUser.MaxDisconnectionTime
WScript.Echo "Max Idle Time: " & objUser.MaxIdleTime
WScript.Echo "Reconnection Action: " & objUser.ReconnectionAction
 
WScript.Echo "Connect Client Drives At Logon: " & _
    objUser.ConnectClientDrivesAtLogon
WScript.Echo "Connect Client Printers At Logon: " & _
    objUser.ConnectClientPrintersAtLogon
WScript.Echo "Default To Main Printer: " & _
    objUser.DefaultToMainPrinter
WScript.Echo "Terminal Services Initial Program: " & _
    objUser.TerminalServicesInitialProgram
WScript.Echo "Terminal Services Work Directory: " & _
    objUser.TerminalServicesWorkDirectory

