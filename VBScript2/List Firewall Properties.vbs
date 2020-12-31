' Description: Lists Windows Firewall properties for the current profile.


Set objFirewall = CreateObject("HNetCfg.FwMgr")
Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
Wscript.Echo "Current profile type: " & objFirewall.CurrentProfileType

Wscript.Echo "Firewall enabled: " & objPolicy.FirewallEnabled
Wscript.Echo "Exceptions not allowed: " & objPolicy.ExceptionsNotAllowed
Wscript.Echo "Notifications disabled: " & objPolicy.NotificationsDisabled
Wscript.Echo "Unicast responses to multicast broadcast disabled: " & _
    objPolicy.UnicastResponsestoMulticastBroadcastDisabled

