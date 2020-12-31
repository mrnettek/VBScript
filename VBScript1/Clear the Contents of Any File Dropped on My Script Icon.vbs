Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each strArgument in Wscript.Arguments
    Set objFile = objFSO.OpenTextFile(strArgument, ForWriting)
    objFile.Write ""
    objFile.Close
Next
  


