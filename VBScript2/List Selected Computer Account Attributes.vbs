' Description: Demonstration script that retrieves the location and description attributes for a computer account in Active Directory.


On Error Resume Next

Set objComputer = GetObject _
    ("LDAP://CN=atl-dc-01,CN=Computers,DC=fabrikam,DC=com")

objProperty = objComputer.Get("Location")
If IsNull(objProperty) Then
    Wscript.Echo "The location has not been set."
Else
    Wscript.Echo "Location: " & objProperty
    objProperty = Null
End If

objProperty = objComputer.Get("Description")
If IsNull(objProperty) Then
    Wscript.Echo "The description has not been set."
Else
    Wscript.Echo "Description: " & objProperty
    objProperty = Null
End If

