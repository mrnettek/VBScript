' Description: Lists all authorized applications for the Windows Firewall standard profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy

Set objProfile = objPolicy.GetProfileByType(1)
Set colApplications = objProfile.AuthorizedApplications

For Each objApplication in colApplications
    Wscript.Echo "Authorized application: " & objApplication.Name
    Wscript.Echo "Application enabled: " & objApplication.Enabled
    Wscript.Echo "Application IP version: " & objApplication.IPVersion
    Wscript.Echo "Application process image file name: " & _
        objApplication.ProcessImageFileName
    Wscript.Echo "Application remote addresses: " & _
        objApplication.RemoteAddresses
    Wscript.Echo "Application scope: " & objApplication.Scope
    Wscript.Echo
Next

