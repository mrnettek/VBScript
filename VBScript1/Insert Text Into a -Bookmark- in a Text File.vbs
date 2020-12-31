Const ForReading = 1
Const ForWriting = 2

Set objNetwork = CreateObject("Wscript.Network")
strUser = objNetwork.UserName

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForReading)

strText = objFile.ReadAll
objFile.Close

strText = Replace(strText, "[BOOKMARK #1]", Date)
strText = Replace(strText, "[BOOKMARK #2]", strUser)

Set objFile = objFSO.OpenTextFile("C:\Scripts\Test.txt", ForWriting)
objFile.WriteLine strText
objFile.Close
  


