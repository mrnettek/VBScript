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
    "SELECT Name FROM 'LDAP://cn=Schema,cn=Configuration,dc=fabrikam,dc=com' WHERE " & _
        "objectClass='attributeSchema' AND isMemberOfPartialAttributeSet=TRUE"  
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields("Name").Value
    objRecordSet.MoveNext
Loop
  


