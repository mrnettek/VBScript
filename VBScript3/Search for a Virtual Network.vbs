' Description: Locates a virtual network named Scripted Network.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objNetwork = objVS.FindVirtualNetwork("Scripted Network")

