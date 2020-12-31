' Description: Configures dynamic DNS registration for a network adapter.


On Error Resume Next
 
Const FULL_DNS_REGISTRATION = True
Const DOMAIN_DNS_REGISTRATION = False
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colNetCards = objWMIService.ExecQuery _
    ("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objNetCard in colNetCards
    objNetCard.SetDynamicDNSRegistration FULL_DNS_REGISTRATION, _
        DOMAIN_DNS_REGISTRATION
Next

