On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from CIM_SupportAccess",,48)
For Each objItem in colItems
    Wscript.Echo "CommunicationInfo: " & objItem.CommunicationInfo
    Wscript.Echo "CommunicationMode: " & objItem.CommunicationMode
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Locale: " & objItem.Locale
    Wscript.Echo "SupportAccessId: " & objItem.SupportAccessId
Next

