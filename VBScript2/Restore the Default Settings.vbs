' Description: Restore the Windows Firewall default settings.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
objFirewall.RestoreDefaults()

