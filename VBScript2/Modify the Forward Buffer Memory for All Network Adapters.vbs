' Description: Configures the forward buffer memory for all network adapters on a computer. Values should be supplied in multiples of 256.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetForwardBufferMemory(74240)

