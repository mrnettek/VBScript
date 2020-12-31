' Description: Configures a computer to use the RFC 1122 specification for urgent data rather than the mode used by Berkeley Software Design- (BSD) derived systems. To use the BSD mode instead, pass the value False to the SetTcpRFC1122UrgentPointer method.


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
objNetworkSettings.SetTcpUseRFC1122UrgentPointer(True)

