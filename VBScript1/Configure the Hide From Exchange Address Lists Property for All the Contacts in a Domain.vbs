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
    "SELECT ADsPath FROM 'LDAP://dc=fabrikam,dc=com' WHERE objectClass='contact'"  
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    strContactPath = objRecordSet.Fields("ADsPath").Value
    Set objContact = GetObject(strContactPath)
    objContact.MSExchHideFromAddressLists = TRUE
    objContact.SetInfo
    objRecordSet.MoveNext
Loop
  


