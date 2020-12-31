Const ADS_SCOPE_SUBTREE = 2

strComputer = InputBox("Please enter the computer name:", "Delete Computer Account")

If strComputer = "" Then
    Wscript.Quit
End If

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.CommandText = "Select ADsPath From " & _
    "'LDAP://DC=fabrikam,DC=com' Where objectClass='computer'" & _
        " and Name = '" & strComputer & "'"  
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    Set objComputer = GetObject(objRecordSet.Fields("ADsPath").Value)
    objComputer.DeleteObject (0)
    objRecordSet.MoveNext
Loop
  


