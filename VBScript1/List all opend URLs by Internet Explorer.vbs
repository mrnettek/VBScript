Dim oShell 		: Set oShell = CreateObject("Shell.Application")
Dim oIEApp 		: Set oIEApp = CreateObject("InternetExplorer.Application") 
Dim oShellWindows 	: Set oShellWindows = oShell.windows

For Each oIEApp in oShellWindows
	If Right(LCase(oIEApp.Fullname),12) = "iexplore.exe" then
		If Len(oIEApp.LocationURL) > 0 then
			msgbox oIEApp.LocationURL
		End if
	End if
Next
