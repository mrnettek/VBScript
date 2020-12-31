On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_TerminalServiceSetting",,48)
For Each objItem in colItems
    Wscript.Echo "ActiveDesktop: " & objItem.ActiveDesktop
    Wscript.Echo "AllowTSConnections: " & objItem.AllowTSConnections
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DeleteTempFolders: " & objItem.DeleteTempFolders
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DirectConnectLicenseServers: " & objItem.DirectConnectLicenseServers
    Wscript.Echo "HomeDirectory: " & objItem.HomeDirectory
    Wscript.Echo "LicensingDescription: " & objItem.LicensingDescription
    Wscript.Echo "LicensingName: " & objItem.LicensingName
    Wscript.Echo "LicensingType: " & objItem.LicensingType
    Wscript.Echo "Logons: " & objItem.Logons
    Wscript.Echo "ProfilePath: " & objItem.ProfilePath
    Wscript.Echo "ServerName: " & objItem.ServerName
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "SingleSession: " & objItem.SingleSession
    Wscript.Echo "TerminalServerMode: " & objItem.TerminalServerMode
    Wscript.Echo "UserPermission: " & objItem.UserPermission
    Wscript.Echo "UseTempFolders: " & objItem.UseTempFolders
Next

