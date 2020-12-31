Dim oComputerSystem 	: Set oComputerSystem 	= GetObject("winmgmts://.").InstancesOf("Win32_ComputerSystem")
Dim bMember 		: bMember 		= false


For Each oEntry in oComputerSystem
	If oEntry.DomainRole = 1 then
		WScript.Echo "Computer is Member of a Domain"
		bMember = true
		Exit For
	End if
Next

If (bMember = false) then
	WScript.Echo "Computer is not Member of a Domain"
End if
