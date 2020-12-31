' Description: Demonstration script that modifies a Windows Firewall ICMP setting for the current profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

Set objICMPSettings = objPolicy.ICMPSettings
objICMPSettings.AllowRedirect = TRUE

