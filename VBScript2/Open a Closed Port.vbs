' Description: Opens closed port 9999 for the Windows Firewall current profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
Set colPorts = objPolicy.GloballyOpenPorts

Set objPort = colPorts.Item(9999,6)
objPort.Enabled = TRUE

