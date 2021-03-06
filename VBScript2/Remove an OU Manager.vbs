' Description: Removes the manager entry for the Active Directory OU named Sales. When this group is run, the OU will no longer have an assigned manager.


Const ADS_PROPERTY_CLEAR = 1 
 
Set objContainer = GetObject _
  ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")

objContainer.PutEx ADS_PROPERTY_CLEAR, "managedBy", 0
objContainer.SetInfo

