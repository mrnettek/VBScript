' Description: Determines the Active Directory class type for the organizational-person object.


strClassName = "cn=organizational-person"
 
Set objSchemaClass = GetObject _
    ("LDAP://" & strClassName & _
        ",cn=schema,cn=configuration,dc=fabrikam,dc=com")
 
intClassCategory = objSchemaClass.Get("objectClassCategory")

Select Case intClassCategory
    Case 0
        strCategory = "88"
    Case 1
        strCategory = "structural"
    Case 2
        strCategory = "abstract"
    Case 3
        strCategory = "auxiliary"
End Select

Wscript.Echo strClassName & " is categorized as " & strCategory & "."

