' Description: Enumerates the ICMP settings for Internet Connection Firewall.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery _
    ("Select * from HNet_ConnectionICMPSetting")

For Each objItem in colItems
    Wscript.Echo "Connection: " & objItem.Connection
    Wscript.Echo "ICMP Settings: " & objItem.ICMPSettings
Next

