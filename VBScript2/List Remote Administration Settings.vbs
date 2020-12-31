' Description: Lists remote administration settings for the Windows Firewall current profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set objAdminSettings = objPolicy.RemoteAdminSettings
Wscript.Echo "Remote administration settings enabled: " & _
    objAdminSettings.Enabled
Wscript.Echo "Remote administration addresses: " & _
    objAdminSettings.RemoteAddresses
Wscript.Echo "Remote administration scope: " & objAdminSettings.Scope
Wscript.Echo "Remote administration IP version: " & objAdminSettings.IPVersion

