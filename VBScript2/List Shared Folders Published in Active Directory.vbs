' Description: Lists all the shared folders that have been published in Active Directory.


Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = "Select Name, unCName, ManagedBy from " _
    & "'LDAP://DC=Fabrikam,DC=com' where objectClass='volume'"
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    Wscript.Echo "Share Name: " & objRecordSet.Fields("Name").Value
    Wscript.Echo "UNC Name: " & objRecordSet.Fields("uNCName").Value
    Wscript.Echo "Managed By: " & objRecordSet.Fields("ManagedBy").Value
    objRecordSet.MoveNext
Loop

