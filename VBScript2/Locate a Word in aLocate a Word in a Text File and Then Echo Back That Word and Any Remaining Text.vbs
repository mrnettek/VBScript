Const ForReading = 1

Set obJFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

strContents = objFile.ReadAll
objFile.Close

arrWords = Split(strContents, " ")
intStart = 0

For Each strWord in arrWords
    If InStr(strWord, "telescope") Then
        intStart = intStart + 1
    End If

    If intStart > 0 Then
        strFinalText = strFinalText & strWord & " "
    End If
Next
    
Wscript.Echo strFinalText
  


