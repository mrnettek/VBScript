' Description: Create a virtual network named Scripted Network.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
errReturn = objVS.CreateVirtualNetwork _
    ("Scripted Network","C:\Virtual Machines")

