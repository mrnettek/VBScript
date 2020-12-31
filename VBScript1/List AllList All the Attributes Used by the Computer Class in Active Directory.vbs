Set objSchema = GetObject("LDAP://schema/computer")
 
Wscript.Echo "Mandatory attributes"

For Each strAttribute in objSchema.MandatoryProperties
    Wscript.Echo strAttribute
Next
 
Wscript.Echo ""

Wscript.Echo "Optional attributes"

For Each strAttribute in objSchema.OptionalProperties
    Wscript.Echo strAttribute
Next
  


