' Description: Creates an ATM address to name (ATMA) record on a DNS server.


strDNSServer = "atl-dc-03.fabrikam.com"
strContainer = "fabrikam.com"
strOwner = "atm.fabrikam.com"
intRecordClass = 1
intTTL = 600 
intFormat = 1
strATMAddress = "47.0079.00010200000000000000.00a03e000002.00"
 
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\MicrosoftDNS")

Set objItem = objWMIService.Get("MicrosoftDNS_ATMAType")
errResult = objItem.CreateInstanceFromPropertyData _
    (strDNSServer, strContainer, strOwner, intRecordClass, intTTL, _
        intFormat, strATMAddress)

