' Description: Removes the Jackson user account from all Active Directory security groups.


On Error Resume Next

Const ADS_PROPERTY_DELETE = 4
Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
 
Set objUser = GetObject _
    ("LDAP://cn=Jackson,ou=Management,dc=NA,dc=fabrikam,dc=com") 
arrMemberOf = objUser.GetEx("memberOf")
 
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "This account is not a member of any security groups."
    WScript.Quit
End If
 
For Each Group in arrMemberOf
    Set objGroup = GetObject("LDAP://" & Group) 
    objGroup.PutEx ADS_PROPERTY_DELETE, _
        "member", Array("cn=Jackson,ou=Management,dc=NA,dc=fabrikam,dc=com")
    objGroup.SetInfo
Next

