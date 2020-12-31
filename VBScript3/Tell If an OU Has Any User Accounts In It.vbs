On Error Resume Next

Const ADS_SCOPE_ONELEVEL = 1

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_ONELEVEL 

objCommand.CommandText = _
    "SELECT Name FROM 'LDAP://OU=finance,dc=fabrikam,dc=com' WHERE objectCategory='user'"  
Set objRecordSet = objCommand.Execute

Wscript.Echo "Number of user accounts: " &  objRecordSet.RecordCount
  


