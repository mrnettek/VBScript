' Description: Returns a list of mandatory and optional attributes of the group class (as stored in the Active Directory schema).


Set objGroupClass = GetObject("LDAP://schema/group")
Set objSchemaClass = GetObject(objGroupClass.Parent)
 
i = 0
WScript.Echo "Mandatory attributes:"
For Each strAttribute in objGroupClass.MandatoryProperties
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
For Each strAttribute in objGroupClass.OptionalProperties
    i= i + 1
    Wscript.Echo i & vbTab & strAttribute
    Set objAttribute = objSchemaClass.GetObject("Property",  strAttribute)
    Wscript.Echo " [Syntax: " & objAttribute.Syntax & "]"
    If objAttribute.MultiValued Then
        WScript.Echo " Multivalued"
    Else
        WScript.Echo " Single-valued"
    End If
Next

