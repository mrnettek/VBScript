' Description: Displays the global SMTP advanced delivery metabase property values for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpServiceSetting") 

For Each objItem in colItems
    Wscript.Echo "Enable Reverse DNS Lookup: " & _
        objItem.EnableReverseDnsLookup
    Wscript.Echo "Fully Qualified Domain Name: " & _
        objItem.FullyQualifiedDomainName
    Wscript.Echo "Hop Count: " & objItem.HopCount
    Wscript.Echo "Masquerade Domain: " & objItem.MasqueradeDomain
    Wscript.Echo "Smart Host: " & objItem.SmartHost
    Wscript.Echo "Smart Host Type: " & objItem.SmartHostType
Next

