On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")

objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.CommandText = _
    "SELECT ADsPath FROM 'LDAP://dc=fabrikam,dc=com' WHERE objectCategory='user' " & _
        "AND samAccountName = 'kenmyer'" 

Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    strADsPath = objRecordSet.Fields("ADsPath").Value
    Set objGroup = GetObject("LDAP://cn=Finance Managers,ou=Finance,dc=fabrikam,dc=com")
    objGroup.Add(strADsPath)
    objRecordSet.MoveNext
Loop
  


