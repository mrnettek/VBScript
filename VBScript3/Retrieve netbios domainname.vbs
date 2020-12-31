Dim oSystemInfo : Set oSystemInfo = CreateObject("ADSystemInfo") 
WScript.Echo "Domain shortname: " & oSystemInfo.DomainShortName

