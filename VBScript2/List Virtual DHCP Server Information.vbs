' Description: Lists Virtual DHCP Server information for a network named Internal Network.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objNetwork = objVS.FindVirtualNetwork("Internal Network")

Set objDHCPServer = objNetwork.DHCPVirtualNetworkServer
Wscript.Echo "Default gateway address: " & objDHCPServer.DefaultGatewayAddress
Wscript.Echo "DNS servers: " & objDHCPServer.DNSServers
Wscript.Echo "Ending IP address: " & objDHCPServer.EndingIPAddress
Wscript.Echo "Is enabled: " & objDHCPServer.IsEnabled
Wscript.Echo "Lease rebinding time: " & objDHCPServer.LeaseRebindingTime
Wscript.Echo "Lease renewal time: " & objDHCPServer.LeaseRenewalTime
Wscript.Echo "Lease time: " & objDHCPServer.LeaseTime
Wscript.Echo "Network: " & objDHCPServer.Network
Wscript.Echo "Network mask: " & objDHCPServer.NetworkMask
Wscript.Echo "Server IP address: " & objDHCPServer.ServerIPAddress
Wscript.Echo "Starting IP address: " & objDHCPServer.StartingIPAddress
Wscript.Echo "WINS Server: " & objDHCPServer.WINSServers

