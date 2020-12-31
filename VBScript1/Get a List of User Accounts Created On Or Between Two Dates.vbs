On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

dtmCreationDate1 = "20070701000000.0Z"
dtmCreationDate2 = "20070731000000.0Z"

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = _
    "SELECT Name, whenCreated FROM 'LDAP://dc=fabrikam,dc=com' WHERE objectClass='user' "  & _
        "AND whenCreated>='" & dtmCreationDate1 & "' AND whenCreated<='" & dtmCreationDate2 & "'" 
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields("Name").Value, objRecordSet.Fields("whenCreated").Value
    objRecordSet.MoveNext
Loop
  


