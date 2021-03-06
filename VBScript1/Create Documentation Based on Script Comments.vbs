' Description: Demonstrates the use of the FileSystemObject as a way to copy comments from a script to a separate text file. Requires comments to have been marked using '*.


Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objScriptFile = objFSO.OpenTextFile("c:\scripts\Service_Monitor.vbs", _
    ForReading)
Set objCommentFile = objFSO.OpenTextFile("c:\scripts\Comments.txt", _ 
    ForWriting, TRUE)

Do While objScriptFile.AtEndOfStream <> TRUE
    strCurrentLine = objScriptFile.ReadLine
    intIsComment = Instr(1,strCurrentLine,"'*")
    If intIsComment > 0 Then
        objCommentFile.Write strCurrentLine & VbCrLf
    End If
Loop

objScriptFile.Close
objCommentFile.Close

