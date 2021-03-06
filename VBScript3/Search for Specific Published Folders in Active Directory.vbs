' Description: Searches Active Directory for any shared folders that have the keyword "finance."


On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCOmmand.ActiveConnection = objConnection
objCommand.CommandText = "Select Name, unCName, ManagedBy from "
    & "'LDAP://DC=Reskit,DC=com'" _
        & " where objectClass='volume' and Keywords = 'finance*'"
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    Wscript.Echo "Share Name: " & objRecordSet.Fields("Name").Value
    Wscript.Echo "UNC Name: " & objRecordSet.Fields("uNCName").Value
    Wscript.Echo "Managed By: " & objRecordSet.Fields("ManagedBy").Value
    objRecordSet.MoveNext
Loop

