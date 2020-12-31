' Description: Lists Virtual Server access rights.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objSecurity = objVS.Security
Set colAccessRights = objSecurity.AccessRights

For Each objAccessRight in colAccessRights
    Wscript.Echo "Change permissions: " & objAccessRight.ChangePermissions
    Wscript.Echo "Delete access: " & objAccessRight.DeleteAccess
    Wscript.Echo "Execute access: " & objAccessRight.ExecuteAccess
    Wscript.Echo "Flags: " & objAccessRight.Flags
    Wscript.Echo "Name: " & objAccessRight.Name
    Wscript.Echo "Read access: " & objAccessRight.ReadAccess
    Wscript.Echo "Read permissions: " & objAccessRight.ReadPermissions
    Wscript.Echo "SID: " & objAccessRight.SID
    Wscript.Echo "Special access: " & objAccessRight.SpecialAccess
    Wscript.Echo "Type: " & objAccessRight.Type
    Wscript.Echo "Write access: " & objAccessRight.WriteAccess
    Wscript.Echo
Next

