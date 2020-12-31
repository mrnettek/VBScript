On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
    "Select Name From 'LDAP://DC=fabrikam,DC=com' Where objectClass='computer'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    strComputer = objRecordSet.Fields("Name").Value

    Set objUser = GetObject("WinNT://" & strComputer & "/Administrator")
    objUser.SetPassword "x%tY7iu8%4f"

    objRecordSet.MoveNext
Loop
  


