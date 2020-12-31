Const ForReading = 1

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Pattern = "^[1-9]...GRP"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

Do Until objFile.AtEndOfStream
    strSearchString = objFile.ReadLine
    Set colMatches = objRegEx.Execute(strSearchString)  
    If colMatches.Count > 0 Then
        For Each strMatch in colMatches   
            Wscript.Echo strSearchString 
        Next
    End If
Loop

objFile.Close
  


