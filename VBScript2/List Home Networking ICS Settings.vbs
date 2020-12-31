' Description: Reports whether the Internet Connection Firewall settings have been configured to allow DHCP and DNS.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\microsoft\homenet")

Set colItems = objWMIService.ExecQuery("Select * from HNet_ICSSettings")

For Each objItem in colItems
    Wscript.Echo "DHCP Enabled: " & objItem.DHCPEnabled
    Wscript.Echo "DNS Enabled: " & objItem.DNSEnabled
    Wscript.Echo "ID: " & objItem.ID
Next

