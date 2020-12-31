On Error Resume Next

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000

objCommand.CommandText = _
    "<LDAP://dc=fabrikam,dc=com>;" & _
    "(&(objectCategory=User)(userAccountControl:1.2.840.113556.1.4.803:=2));" & _
            "Name,homeDirectory;Subtree"
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    Wscript.Echo objRecordSet.Fields("Name").Value
    Wscript.Echo objRecordSet.Fields("homeDirectory").Value
    Wscript.Echo 
    objRecordSet.MoveNext
Loop
  


