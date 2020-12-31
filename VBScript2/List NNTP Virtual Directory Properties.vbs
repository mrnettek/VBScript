' Description: Returns basic configuration information for all the NNTP virtual directories on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsNntpVirtualDirSetting")

For Each objItem in colItems
    Wscript.Echo "Content Indexed: " & objItem.ContentIndexed
    Wscript.Echo "Don't Log: " & objItem.DontLog
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path: " & objItem.Path
Next

