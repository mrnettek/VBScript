' Description: Enables WINS on all the network adapters installed in a computer.


On Error Resume Next
 
Const DNS_ENABLED_FOR_WINS_RESOLUTION = True
Const USE_LMHOST_FILE = False

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
errResult = objNetworkSettings.EnableWINS(DNS_ENABLED_FOR_WINS_RESOLUTION,_
     USE_LMHOST_FILE)

