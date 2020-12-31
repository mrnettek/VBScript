' Description: Displays mandatory and optional attributes (and their values) for a local user account named kenmyer on a computer named atl-win2k-01.


On Error Resume Next
 
strComputer = "atl-win2k-01"
Set objUser = GetObject("WinNT://" & strComputer & "/kenmyer ")
Set objClass = GetObject(objUser.Schema)
 
WScript.Echo "Mandatory properties for " & objUser.Name & ":"
For Each property in objClass.MandatoryProperties
    WScript.Echo property, objUser.Get(property)
Next
 
WScript.Echo "Optional properties for " & objUser.Name & ":"
For Each property in objClass.OptionalProperties
    WScript.Echo property, objUser.Get(property)
Next

