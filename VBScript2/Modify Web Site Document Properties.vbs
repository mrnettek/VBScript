' Description: Demonstration script that modifies default Web site document metabase properties on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsWebServiceSetting")

For Each objItem in colItems
    strDocs = objItem.DefaultDoc
    objItem.DefaultDoc = strDocs & ",index.htm"
    objItem.DefaultDocFooter = "FILE:c:\config\footer.htm"
    objItem.EnableDefaultDoc = True
    objItem.EnableDocFooter = True
    objItem.Put_
Next

