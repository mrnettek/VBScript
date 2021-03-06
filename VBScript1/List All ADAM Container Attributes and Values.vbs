' Description: Lists all the attributes and values for the ADAM container named Users.


On Error Resume Next

Set objUser = GetObject("LDAP://localhost:389/cn=Users,dc=fabrikam,dc=com")
Set objUserProperties = GetObject("LDAP://localhost:389/schema/container")

For Each strAttribute in objUserProperties.MandatoryProperties
    strValues = objUser.GetEx(strAttribute)
    For Each strItem in strValues
        Wscript.Echo strAttribute & " -- " & strItem
    Next
Next

For Each strAttribute in objUserProperties.OptionalProperties
    strValues = objUser.GetEx(strAttribute)
    If Err = 0 Then
        For Each strItem in strValues
            Wscript.Echo strAttribute & " -- " & strItem
        Next
    Else
        Wscript.Echo strAttribute & " --  No value set"
        Err.Clear
    End If
Next

