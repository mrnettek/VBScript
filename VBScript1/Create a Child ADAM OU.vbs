' Description: Creates a new ADAM OU within the Management2 OU.


On Error Resume Next

Set objParentOU = GetObject _
    ("LDAP://localhost:389/ou=Management2,ou=Management,dc=fabrikam,dc=com")

Set objChildOU = objParentOU.Create("organizationalUnit", "ou=Level3")
objChildOU.SetInfo

