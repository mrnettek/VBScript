' Description: Lists configuration information for all the network adapters installed on a computer.


On Error Resume Next

strComputer = "."Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_NetworkAdapterConfiguration")

For Each objItem in colItems
    Wscript.Echo "ARP Always Source Route: " & objItem.ArpAlwaysSourceRoute
    Wscript.Echo "ARP Use EtherSNAP: " & objItem.ArpUseEtherSNAP
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Database Path: " & objItem.DatabasePath
    Wscript.Echo "Dead GW Detection Enabled: " & objItem.DeadGWDetectEnabled
    Wscript.Echo "Default IP Gateway: " & objItem.DefaultIPGateway
    Wscript.Echo "Default TOS: " & objItem.DefaultTOS
    Wscript.Echo "Default TTL: " & objItem.DefaultTTL
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DHCP Enabled: " & objItem.DHCPEnabled
    Wscript.Echo "DHCP Lease Expires: " & objItem.DHCPLeaseExpires
    Wscript.Echo "DHCP Lease Obtained: " & objItem.DHCPLeaseObtained
    Wscript.Echo "DHCP Server: " & objItem.DHCPServer
    Wscript.Echo "DNS Domain: " & objItem.DNSDomain
    Wscript.Echo "DNS Domain Suffix Search Order: " & _
        objItem.DNSDomainSuffixSearchOrder
    Wscript.Echo "DNS Enabled For WINS Resolution: " & _
        objItem.DNSEnabledForWINSResolution
    Wscript.Echo "DNS Host Name: " & objItem.DNSHostName
    Wscript.Echo "DNS Server Search Order: " & objItem.DNSServerSearchOrder
    Wscript.Echo "Domain DNS Registration Enabled: " & _
        objItem.DomainDNSRegistrationEnabled
    Wscript.Echo "Forward Buffer Memory: " & objItem.ForwardBufferMemory
    Wscript.Echo "Full DNS Registration Enabled: " & _
        objItem.FullDNSRegistrationEnabled
    Wscript.Echo "Gateway Cost Metric: " & objItem.GatewayCostMetric
    Wscript.Echo "IGMP Level: " & objItem.IGMPLevel
    Wscript.Echo "Index: " & objItem.Index
    Wscript.Echo "IP Address: " & objItem.IPAddress
    Wscript.Echo "IP Connection Metric: " & objItem.IPConnectionMetric
    Wscript.Echo "IP Enabled: " & objItem.IPEnabled
    Wscript.Echo "IP Filter Security Enabled: " & _
        objItem.IPFilterSecurityEnabled
    Wscript.Echo "IP Port Security Enabled: " & objItem.IPPortSecurityEnabled
    Wscript.Echo "IPSec Permit IP Protocols: " & objItem.IPSecPermitIPProtocols
    Wscript.Echo "IPSec Permit TCP Ports: " & objItem.IPSecPermitTCPPorts
    Wscript.Echo "IPSec Permit UDP Ports: " & objItem.IPSecPermitUDPPorts
    Wscript.Echo "IP Subnet: " & objItem.IPSubnet
    Wscript.Echo "IP Use Zero Broadcast: " & objItem.IPUseZeroBroadcast
    Wscript.Echo "IPX Address: " & objItem.IPXAddress
    Wscript.Echo "IPX Enabled: " & objItem.IPXEnabled
    Wscript.Echo "IPX Frame Type: " & objItem.IPXFrameType
    Wscript.Echo "IPX Media Type: " & objItem.IPXMediaType
    Wscript.Echo "IPX Network Number: " & objItem.IPXNetworkNumber
    Wscript.Echo "IPX Virtual Net Number: " & objItem.IPXVirtualNetNumber
    Wscript.Echo "Keep Alive Interval: " & objItem.KeepAliveInterval
    Wscript.Echo "Keep Alive Time: " & objItem.KeepAliveTime
    Wscript.Echo "MAC Address: " & objItem.MACAddress
    Wscript.Echo "MTU: " & objItem.MTU
    Wscript.Echo "Number of Forward Packets: " & objItem.NumForwardPackets
    Wscript.Echo "PMTUBH Detect Enabled: " & objItem.PMTUBHDetectEnabled
    Wscript.Echo "PMTU Discovery Enabled: " & objItem.PMTUDiscoveryEnabled
    Wscript.Echo "Service Name: " & objItem.ServiceName
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "TCPIP Netbios Options: " & objItem.TcpipNetbiosOptions
    Wscript.Echo "TCP Maximum Connect Retransmissions: " & _
        objItem.TcpMaxConnectRetransmissions
    Wscript.Echo "TCP Maximum Data Retransmissions: " & _
        objItem.TcpMaxDataRetransmissions
    Wscript.Echo "TCP NumC onnections: " & objItem.TcpNumConnections
    Wscript.Echo "TCP Use RFC1122 Urgent Pointer: " & _
        objItem.TcpUseRFC1122UrgentPointer
    Wscript.Echo "TCP Window Size: " & objItem.TcpWindowSize
    Wscript.Echo "WINS Enable LMHosts Lookup: " & _
        objItem.WINSEnableLMHostsLookup
    Wscript.Echo "WINS Host Lookup File: " & objItem.WINSHostLookupFile
    Wscript.Echo "WINS Primary Server: " & objItem.WINSPrimaryServer
    Wscript.Echo "WINS Scope ID: " & objItem.WINSScopeID
    Wscript.Echo "WINS Secondary Server: " & objItem.WINSSecondaryServer
Next

