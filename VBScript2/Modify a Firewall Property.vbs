' Description: Demonstration script that modifies Windows Firewall properties for the current profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile

objPolicy.ExceptionsNotAllowed = TRUE
objPolicy.NotificationsDisabled = TRUE
objPolicy.UnicastResponsestoMulticastBroadcastDisabled = TRUE

