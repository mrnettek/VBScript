' Description: Demonstration script that connects to and returns information about the Windows Firewall standard profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy
Set objProfile = objPolicy.GetProfileByType(1)

Wscript.Echo "Firewall enabled: " & objProfile.FirewallEnabled
Wscript.Echo "Exceptions not allowed: " & objProfile.ExceptionsNotAllowed
Wscript.Echo "Notifications disabled: " & objProfile.NotificationsDisabled
Wscript.Echo "Unicast responses to multicast broadcast disabled: " & -
    objProfile.UnicastResponsestoMulticastBroadcastDisabled

