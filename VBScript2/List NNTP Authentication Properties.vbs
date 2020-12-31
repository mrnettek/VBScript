' Description: Displays global NNTP service authentication properties.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpServiceSetting") 

For Each objItem in colItems
    Wscript.Echo "Access SSL Map Certificate: " & _
        objItem.AccessSSLMapCert
    Wscript.Echo "Access SSL Negotiate Certifcate: " & _
        objItem.AccessSSLNegotiateCert
    Wscript.Echo "Access SSL Require Certificate: " & _
        objItem.AccessSSLRequireCert
    Wscript.Echo "Authentication Anonymous: " & objItem.AuthAnonymous
    Wscript.Echo "Authentication Basic: " & objItem.AuthBasic
    Wscript.Echo "Authentication NTLM: " & objItem.AuthNTLM
Next

