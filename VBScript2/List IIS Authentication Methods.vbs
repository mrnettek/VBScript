' Description: Lists IIS authentication methods.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    Wscript.Echo "Anonymous User Name: " & objItem.AnonymousUserName
    Wscript.Echo "Anonymous User Password: " & _
        objItem.AnonymousUserPass
    Wscript.Echo "Authentication Anonymous: " & objItem.AuthAnonymous
    Wscript.Echo "Authentication Basic: " & objItem.AuthBasic
    Wscript.Echo "Authentication MD5: " & objItem.AuthMD5
    Wscript.Echo "Authentication NTLM: " & objItem.AuthNTLM
    Wscript.Echo "Authentication Passport: " & objItem.AuthPassport
    Wscript.Echo "Default Logon Domain: " & objItem.DefaultLogonDomain
    Wscript.Echo "Realm: " & objItem.Realm
Next

