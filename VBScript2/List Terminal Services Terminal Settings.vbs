' Description: Lists the Terminal Service configuration settings on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_TerminalServiceSetting")

For Each objItem in colItems
    Wscript.Echo "Active Desktop: " & objItem.ActiveDesktop
    Wscript.Echo "Allow TS connections: " & objItem.AllowTSConnections
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Delete temporary folders: " & objItem.DeleteTempFolders
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Direct connect license servers: " & _
        objItem.DirectConnectLicenseServers
    Wscript.Echo "Disable forcible logoff: " & objItem.DisableForcibleLogoff
    Wscript.Echo "Help: " & objItem.Help
    Wscript.Echo "Home directory: " & objItem.HomeDirectory
    Wscript.Echo "Licensing description: " & objItem.LicensingDescription
    Wscript.Echo "Licensing name: " & objItem.LicensingName
    Wscript.Echo "Licensing type: " & objItem.LicensingType
    Wscript.Echo "Logons: " & objItem.Logons
    Wscript.Echo "Profile path: " & objItem.ProfilePath
    Wscript.Echo "Server name: " & objItem.ServerName
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Single session: " & objItem.SingleSession
    Wscript.Echo "Terminal Server mode: " & objItem.TerminalServerMode
    Wscript.Echo "Time zone redirection: " & objItem.TimeZoneRedirection
    Wscript.Echo "User permission: " & objItem.UserPermission
    Wscript.Echo "Use temporary folders: " & objItem.UseTempFolders
    Wscript.Echo
Next

