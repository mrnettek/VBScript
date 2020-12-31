' Description: Displays global document property values for IIS Web sites.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")
For Each objItem in colItems
    Wscript.Echo "Default Doc: " & objItem.DefaultDoc
    Wscript.Echo "Default Doc Footer: " & objItem.DefaultDocFooter
    Wscript.Echo "Enable Default Doc: " & objItem.EnableDefaultDoc
    Wscript.Echo "Enable Doc Footer: " & objItem.EnableDocFooter
Next

