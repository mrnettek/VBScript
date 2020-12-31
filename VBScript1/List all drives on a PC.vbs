Dim ofso : Set ofso = CreateObject("Scripting.FileSystemObject")
Dim sList
Dim oDrives : Set oDrives = ofso.drives

For Each oDrive in oDrives
	sList = sList & oDrive & vbCr 
next

msgbox sList


