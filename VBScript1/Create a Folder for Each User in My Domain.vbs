On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = _
    "SELECT givenName, sn FROM 'LDAP://DC=fabrikam,DC=com' WHERE objectCategory='user'" 

Set objRecordSet = objCommand.Execute

Set objFSO = CreateObject("Scripting.FileSystemObject")

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    strInitial = Left(objRecordSet.Fields("givenName").Value, 1)
    strFolderName = "C:\Public\" & strInitial & objRecordSet.Fields("sn").Value
    Set objFolder = objFSO.CreateFolder(strFolderName)
    objRecordSet.MoveNext
Loop
  


