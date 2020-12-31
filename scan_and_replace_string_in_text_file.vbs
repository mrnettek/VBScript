' MrNetTek
' eddiejackson.net/blog
' 10/14/2019
' free for public use 
' free to claim as your own

'Option Explicit

On error resume next
 
dim objFSO, strFolder, strFilePath, tmpFile, strLineInput, Settings
 
Const ForReading = 1
 
Const ForWriting = 2
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
strFolder = "C:\ProgramData\ABC_Program\"
 
strFilePath = strFolder & "TheFile.ini"
 
Set Settings = objFSO.OpenTextFile(strFilePath, ForReading, True)
 
Set tmpFile = objFSO.OpenTextFile(strFilePath & ".tmp", ForWriting, True)
 
Do While Not Settings.AtEndofStream
 
strLineInput = Settings.ReadLine
 
If InStr(strLineInput, "AutoUpdate=") Then
 
strLineInput = "AutoUpdate=0"
 
End If
 
tmpFile.WriteLine strLineInput
 
Loop
 
Settings.Close
 
tmpFile.Close
 
objFSO.DeleteFile(strFilePath)
 
objFSO.MoveFile strFilePath&".tmp", strFilePath