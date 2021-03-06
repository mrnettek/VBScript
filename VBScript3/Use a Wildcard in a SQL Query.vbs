' Description: Demonstration script that uses a wildcard search to return a list of all Active Directory groups whose common name begins with the letters HR.


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
    "SELECT cn FROM 'LDAP://dc=fabrikam,dc=com' WHERE " _
        & "objectCategory='group' AND cn = 'HR*' "  
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields("cn").Value
    objRecordSet.MoveNext
Loop

