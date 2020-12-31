' Description: Deletes the BITS_Update.dll extension file record.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From IIsWebService")

For Each objItem in colItems
    objItem.DeleteExtensionFileRecord _
        "C:\WINDOWS\system32\bits_update.dll"
Next

