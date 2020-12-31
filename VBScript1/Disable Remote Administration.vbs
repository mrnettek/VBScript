' Description: Disable Windows Firewall remote administration.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set objAdminSettings = objPolicy.RemoteAdminSettings
objAdminSettings.Enabled = FALSE

