' Description: Lists the SMTP LDAP routing properties on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpRoutingSourceSetting")

For Each objItem in colItems
    Wscript.Echo "SMTP DS Account: " & objItem.SmtpDsAccount
    Wscript.Echo "SMTP DS Bind Type: " & objItem.SmtpDsBindType
    Wscript.Echo "SMTP DS Domain: " & objItem.SmtpDsDomain
    Wscript.Echo "SMTP DS Host: " & objItem.SmtpDsHost
    Wscript.Echo "SMTP DS Naming Context: " & _
        objItem.SmtpDsNamingContext
    Wscript.Echo "SMTP DS Password: " & objItem.SmtpDsPassword
    Wscript.Echo "SMTP DS Schema Type: " & objItem.SmtpDsSchemaType
Next

