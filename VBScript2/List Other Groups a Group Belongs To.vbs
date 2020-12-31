' Description: Returns a list of all the groups that the Active Directory security group Scientists is a member of.


On Error Resume Next
 
Set objGroup = GetObject _
    ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com")
objGroup.GetInfo
 
arrMembersOf = objGroup.GetEx("memberOf")
 
WScript.Echo "MembersOf:"
For Each strMemberOf in arrMembersOf
    WScript.Echo strMemberOf
Next

