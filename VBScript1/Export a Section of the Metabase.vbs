' Description: Exports a section of the metabase using the password er$3qld9o.


Const EXPORT_CHILDREN = 0

strComputer = "."

Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsComputer")

For Each objItem in colItems
    objItem.Export "er$3qld9o", "C:\backups\export.xml", _
        "/lm/logging/custom logging", EXPORT_CHILDREN
Next

