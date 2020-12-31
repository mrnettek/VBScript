Const ForReading = 1
blnFound = False

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)
strContents = objFile.ReadAll
objFile.Close

If InStr(strContents, "Windows 2000") Then
    blnFound = True
End If

If InStr(strContents, "Windows XP") Then
    blnFound = True
End If

If blnFound Then
    Wscript.Echo "Either Windows 2000 or Windows XP appears in this file."
Else
    Wscript.Echo "Neither Windows 2000 nor Windows XP appears in this file."
End If
  


