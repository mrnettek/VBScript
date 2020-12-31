' Description: Configures the IPX network number and frame type for a network adapter. In this case, the frame type is set to AUTO (value of 255). Because the frame type is set to AUTO, the network number must be set to 0. Note that both the network number and the frame types must be passed as arrays, even if there is only a single element (for example, one frame type).


On Error Resume Next
 
Const AUTO = 255
Const NETWORK_NUMBER = 0
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colNetCards = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objNetCard in colNetCards
    arrNetworkNumber = Array(NETWORK_NUMBER)
    arrFrameTypes = Array(AUTO)
    objNetCard.SetIPXFrameTypeNetworkPairs arrNetworkNumber, arrFrameTypes
Next

