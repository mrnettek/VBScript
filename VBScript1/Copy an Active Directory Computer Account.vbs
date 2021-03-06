' Description: Retrieves the attributes of an existing computer object and copies the attributes to a new computer object created by the script.


Set objCompt = _
    GetObject("LDAP://cn=Computers,dc=NA,dc=fabrikam,dc=com")
Set objComptCopy = objCompt.Create("computer", "cn=SEA-SQL-01")
objComptCopy.Put "sAMAccountName", "sea-sql-01"
objComptCopy.SetInfo
 
Set objComptTemplate = GetObject _
    ("LDAP://cn=SEA-PM-01,cn=Computers,dc=NA,dc=fabrikam,dc=com")
arrAttributes = Array("description", "location")
 
For Each strAttrib in arrAttributes
    strValue = objComptTemplate.Get(strAttrib)
    objComptCopy.Put strAttrib, strValue
Next
 
objComptCopy.SetInfo

