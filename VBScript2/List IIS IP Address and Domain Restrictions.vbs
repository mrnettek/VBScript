' Description: Lists IIS IP address and domain restrictions.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsIPSecuritySetting")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    For Each strDeny in objItem.DomainDeny
        Wscript.Echo "Domain Deny: " & strDeny
    Next
    For Each strGrant in objItem.DomainGrant
        Wscript.Echo "Domain Grant: " & strGrant
    Next
    Wscript.Echo "Grant By Default: " & objItem.GrantByDefault
    For Each strDeny in objItem.IPDeny
        Wscript.Echo "IP Deny: " & strDeny
    Next
    For Each strGrant in objItem.IPGrant
        Wscript.Echo "IP Grant: " & strGrant
    Next
Next

