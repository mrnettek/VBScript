' Description: Lists IIS access control entries.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery("Select * from IIsACE")
 
For Each objItem in colItems
    Wscript.Echo "Access Mask: " & objItem.AccessMask
    Wscript.Echo "Ace Flags: " & objItem.AceFlags
    Wscript.Echo "Ace Type: " & objItem.AceType
    Wscript.Echo "Flags: " & objItem.Flags
    Wscript.Echo "Inherited Object Type: " & _
        objItem.InheritedObjectType
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Object Type: " & objItem.ObjectType
    Wscript.Echo "Trustee: " & objItem.Trustee
Next

