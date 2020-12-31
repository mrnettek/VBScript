Const ForReading = 1

x = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine

    If Left(strLine, 4) = "****" Then
        x = x + 1
    End If

    If x = 2 Then
        Exit Do
    End If
   
    If Left(strLine, 4) <> "****" And x = 1 Then
        strText = strText & strLine & vbCrLf
    End If
Loop

Wscript.Echo strText
  


