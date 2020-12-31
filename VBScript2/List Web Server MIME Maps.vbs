' Description: Lists all Web server MIME map extensions and types.


strComputer = "."
 
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsWebServerSetting")
 
For Each objItem in colItems
    For i = 0 to Ubound(objItem.MimeMap)
        Wscript.Echo "Extension: " & objItem.MimeMap(i).Extension
        Wscript.Echo "MIME Type: " & objItem.MimeMap(i).MimeType
        Wscript.Echo
    Next
Next

