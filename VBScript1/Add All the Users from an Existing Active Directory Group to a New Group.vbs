Const ADS_GROUP_TYPE_GLOBAL_GROUP = &H2

Set objOU = GetObject("LDAP://OU=Finance, dc=fabrikam, dc=com")
Set objOldGroup = GetObject("LDAP://CN=Finance Managers, ou=Finance, dc=fabrikam, dc=com")

Set objNewGroup = objOU.Create("Group", "Finance Department")
objNewGroup.sAMAccountName = "financedept"
objNewGroup.groupType = ADS_GROUP_TYPE_GLOBAL_GROUP
objNewGroup.Set Info

For Each objUser in objOldGroup.Member
    objNewGroup.Add "LDAP://" & objUser
Next
  


