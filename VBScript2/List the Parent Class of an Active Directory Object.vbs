' Description: Determines the parent class of the Computer object within Active Directory.


strClassName = "cn=computer"
 
Set objSchemaClass = GetObject _
    ("LDAP://" & strClassName & _
        ",cn=schema,cn=configuration,dc=fabrikam,dc=com")
 
strSubClassOf = objSchemaClass.Get("subClassOf")
WScript.Echo "The " & strClassName & _
    " class is a child of the " & strSubClassOf & " class."

