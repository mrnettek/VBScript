strDomain = "fabrikam"

Set colGroups = GetObject("WinNT://" & strDomain & "")
colGroups.Filter = Array("group")

For Each objGroup In colGroups
    Wscript.Echo objGroup.Name 
Next
  


