On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = "SELECT AdsPath FROM 'LDAP://dc=fabrikam,dc=com' WHERE objectCategory='user'"  

Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    strUserPath = objRecordSet.Fields("AdsPath").Value
    Set objUser = GetObject(strUserPath)
    strTelephone = objUser.telephoneNumber
    objUser.Description = strTelephone
    objUser.SetInfo
    objRecordSet.MoveNext
Loop
  


