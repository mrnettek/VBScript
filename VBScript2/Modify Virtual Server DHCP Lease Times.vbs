' Description: Modifies the DHCP lease times for a virtual network named Internal Network.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objNetwork = objVS.FindVirtualNetwork("Internal Network")

Set objDHCPServer = objNetwork.DHCPVirtualNetworkServer
errReturn = objDHCPServer.ConfigureDHCPLeaseTimes(129630,64830,97230)

