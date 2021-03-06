' Description: Determines which Active Directory attributes are indexed and which attributes are in the global catalog.


Const IS_INDEXED = 1
 
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"
 
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
 
objCommand.Properties("Sort On") = "isMemberOfPartialAttributeSet" 
 
objCommand.CommandText = _
    "<LDAP://CN=Schema,CN=Configuration,DC=fabrikam,DC=com>;" & _
        "(objectClass=attributeSchema);" & _
            "lDAPDisplayName, isMemberOfPartialAttributeSet,searchFlags;onelevel"
 
Set objRecordSet = objCommand.Execute
 
Do Until objRecordSet.EOF
    WScript.Echo objRecordset.Fields("lDAPDisplayName") 
    If objRecordset.Fields("isMemberOfPartialAttributeSet")Then
        WScript.Echo "In the global catalog."
    Else
        WScript.Echo "Not in the global catlog."
    End If
 
    If IS_INDEXED AND objRecordset.Fields("searchFlags") Then
        WScript.Echo "Is indexed."
    Else
        WScript.Echo "Is not indexed."
    End If
    Wscript.Echo VbCrLf
    objRecordSet.MoveNext
Loop
 
objConnection.Close

