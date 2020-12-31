' Description: Registers a virtual network named Scripted Network.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
errReturn = objVS.RegisterVirtualNetwork _
    ("Scripted Network","C:\Virtual Machines")

