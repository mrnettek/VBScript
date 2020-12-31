' Description: Configure a computer to use two DNS suffixes -- hr.fabrikam.com and research.fabrikam.com -- when performing DNS searches. Note that even if a computer uses only a single DNS suffix, this value must still be passed to the SetDNSSuffixSearchOrder method as an array (in that case, an array with a single element).


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
arrDNSSuffixes = Array("hr.fabrikam.com", "research.fabrikam.com")
objNetworkSettings.SetDNSSuffixSearchOrder(arrDNSSuffixes)

