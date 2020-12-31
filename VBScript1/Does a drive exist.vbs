Dim ofso : Set ofso = CreateObject("Scripting.FileSystemObject")

sDrive = "C:"

If ofso.driveexists(sDrive) then 
	msgbox("Drive " & sDrive & " exist!")
Else
	msgbox("Drive " & sDrive & " not exist!")
End if
