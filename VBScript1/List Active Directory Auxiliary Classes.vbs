' Description: Returns a list of all the Active Directory auxiliary classes directly applied to the User class.


On Error Resume Next

strClassName = "cn=user"
 
Set objSchemaClass = GetObject _
    ("LDAP://" & strClassName & _
        ",cn=schema,cn=configuration,dc=fabrikam,dc=com")
 
arrSystemAuxiliaryClass = _
objSchemaClass.GetEx("systemAuxiliaryClass")
 
If isEmpty(arrSystemAuxiliaryClass) Then
    WScript.Echo "There are no auxiliary classes" & _
        " applied directly to this class."
    Else
        WScript.StdOut.Write "Auxiliary classes: "
    For Each strAuxiliaryClass in arrSystemAuxiliaryClass
        WScript.StdOut.Write strAuxiliaryClass & " | "
    Next
    WScript.Echo
End If

