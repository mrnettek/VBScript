' Description: Displays basic SMTP server configuration information for a computer running IIS 6.0.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsSmtpServer")

For Each objItem in colItems
    For Each strRoute in objItem.DomainRouting
        Wscript.Echo "Domain Routing: " & strRoute
    Next
    Wscript.Echo "Local Domains:" & objItem.LocalDomains
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Postmaster Email: " & objItem.PostmasterEmail
    Wscript.Echo "Postmaster Name: " & objItem.PostmasterName
    For Each strSource in objItem.RoutingSources
        Wscript.Echo "Routing Sources: " & strSource
    Next
    Wscript.Echo "Server State: " & objItem.ServerState
    Wscript.Echo "SMTP Service Version: " & objItem.SmtpServiceVersion
    Wscript.Echo "Status: " & objItem.Status
Next

