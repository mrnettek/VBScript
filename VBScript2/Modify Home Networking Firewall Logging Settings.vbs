' Description: Disables the logging of dropped packets with Internet Connection Firewall. To enable logging of dropped packets, simply set the value of the property LogDroppedPackets to True.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery _
    ("Select * from HNet_FirewallLoggingSettings")

For Each objItem in colItems
    objItem.LogDroppedPackets = False
    objItem.Put_
Next

