Set objFSO = CreateObject("Scripting.FileSystemObject")

intStartingFolder = InputBox("Please enter the starting number:")
intEndingFolder = InputBox("Please enter the ending number:")

For i = intStartingFolder to intEndingFolder
    strNumber = i
    Do While Len(strNumber) < 5
        strNumber = "0" & strNumber
    Loop

    strFolder = "C:\Scripts\2007-" & strNumber
    Set objFolder = objFSO.CreateFolder(strFolder)
Next
  


