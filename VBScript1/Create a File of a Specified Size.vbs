Const ForWriting = 2

intBytes = InputBox("Enter the size of the file, in bytes:", "File Size")

intBytes = intBytes / 2

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile _
    ("Testfile.txt", ForWriting, True)

For i = 1 to intBytes
    objFile.Write "."
Next

objFile.Close
  


