' Description: Modifies the Notes property for a virtual network named Scripted Network.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objNetwork = objVS.FindVirtualNetwork("Scripted Network")

objNetwork.Notes = "This note was added via a script."

