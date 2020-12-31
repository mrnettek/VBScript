' Description: Lists all the SMTP routing source setting properties on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsSmtpRoutingSourceSetting")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "SMTP DS Account: " & objItem.SmtpDsAccount
    Wscript.Echo "SMTP DS Bind Type: " & objItem.SmtpDsBindType
    Wscript.Echo "SMTP DS Data Directory: " & _
        objItem.SmtpDsDataDirectory
    Wscript.Echo "SMTP DS Default Mail Root: " & _
        objItem.SmtpDsDefaultMailRoot
    Wscript.Echo "SMTP DS Domain: " & objItem.SmtpDsDomain
    Wscript.Echo "SMTP DS Flags: " & objItem.SmtpDsFlags
    Wscript.Echo "SMTP DS Host: " & objItem.SmtpDsHost
    Wscript.Echo "SMTP DS Naming Context: " & _
        objItem.SmtpDsNamingContext
    Wscript.Echo "SMTP DS Password: " & objItem.SmtpDsPassword
    Wscript.Echo "SMTP DS Port: " & objItem.SmtpDsPort
    Wscript.Echo "SMTP DS Schema Type: " & objItem.SmtpDsSchemaType
    Wscript.Echo "SMTP DS Use Catalog: " & objItem.SmtpDsUseCat
    Wscript.Echo "SMTP Routing Table Type: " & _
        objItem.SmtpRoutingTableType
Next

