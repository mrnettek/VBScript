On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2
Const ADS_PROPERTY_CLEAR = 1

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = _
    "SELECT AdsPath FROM 'LDAP://dc=fabrikam,dc=com' WHERE objectCategory='user'"
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    Set objUser = GetObject(objRecordSet.Fields("AdsPath").Value)
    objUser.PutEx ADS_PROPERTY_CLEAR, "scriptPath", 0
    objUser.SetInfo
    objRecordSet.MoveNext
Loop
  

