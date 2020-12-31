
msgbox CreateGUID()

' --------------------------- create GUID
Private Function CreateGUID()
Dim sReturn
Dim oTypeLib : Set oTypeLib = CreateObject("Scriptlet.TypeLib")
	sReturn = cStr(oTypeLib.Guid)
	Set oTypeLib = Nothing

	sReturn = Replace(sReturn, "{", "", 1, -1, 1)
	sReturn = Replace(sReturn, "}", "", 1, -1, 1)
	sReturn = Replace(sReturn, Chr(0), "", 1, -1, 1)

	CreateGUID = "{" & Left(Trim(sReturn),Len(Trim(sReturn))-1) & "}"
End Function

