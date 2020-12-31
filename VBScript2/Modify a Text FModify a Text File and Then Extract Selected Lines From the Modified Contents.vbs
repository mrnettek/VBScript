Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

strContents = objFile.ReadAll
objFile.Close

strContents = Replace(strContents, "~", "~" & vbCrLf)

arrLines =Split(strContents, vbCrLf)

For Each strLine in arrLines
    If Left(strLine, 2) = "N9" or Left(strLine, 2) = "B4" Then
        strNewFile = strNewFile & strLine & vbCrLf
    End If
Next

Set objFile = objFSO.CreateTextFile("C:\Scripts\Test2.txt")
objFile.Write strNewFile
objFile.Close
  


