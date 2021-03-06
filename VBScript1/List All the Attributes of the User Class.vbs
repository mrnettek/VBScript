' Description: Returns a list of mandatory and optional attributes for the User class in Active Directory.


Set objUserClass = GetObject("LDAP://schema/user")
Set objSchemaClass = GetObject(objUserClass.Parent)
 
i = 0
WScript.Echo "Mandatory attributes:"
For Each strAttribute in objUserClass.MandatoryProperties
    i= i + 1
    WScript.Echo i & vbTab & strAttribute
    Set objAttribute = objSchemaClass.GetObject("Property",  strAttribute)
    WScript.Echo " (Syntax: " & objAttribute.Syntax & ")"
    If objAttribute.MultiValued Then
        WScript.Echo " Multivalued"
    Else
        WScript.Echo " Single-valued"
    End If
Next
 
WScript.Echo VbCrLf & "Optional attributes:"
For Each strAttribute in objUserClass.OptionalProperties
    i=i + 1
    WScript.Echo i & vbTab & strAttribute
    Set objAttribute = objSchemaClass.GetObject("Property",  strAttribute)
    WScript.Echo " [Syntax: " & objAttribute.Syntax & "]"
    If objAttribute.MultiValued Then
        WScript.Echo " Multivalued"
    Else
        WScript.Echo " Single-valued"
    End If
Next

