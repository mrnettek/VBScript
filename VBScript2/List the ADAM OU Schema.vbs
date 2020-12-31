' Description: Lists all the mandatory and optional attributes of the ADAM organizationalUnit object.


On Error Resume Next

Set objUserProperties = GetObject("LDAP://localhost:389/schema/organizationalUnit")
WScript.Echo "Mandatory (Must-Contain) attributes"
For Each strAttribute in objUserProperties.MandatoryProperties
    WScript.Echo strAttribute
Next

Wscript.Echo

WScript.Echo VbCrLf & "Optional (May-Contain) attributes"
For Each strAttribute in objUserProperties.OptionalProperties
    WScript.Echo strAttribute
Next

