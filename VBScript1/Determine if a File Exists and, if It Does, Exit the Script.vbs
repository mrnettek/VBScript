Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists("C:\Scripts\Test.txt") Then
    Wscript.Quit
Else
    Wscript.Echo "The file does not exist."
End If
  


