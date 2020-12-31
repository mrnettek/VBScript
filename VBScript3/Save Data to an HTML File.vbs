' Description: Sample HTML function that saves the data found in a SPAN named DataArea to a text file named test.htm.


Sub RunScript
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateTextFile("test.htm")
    Set objFile = objFSO.OpenTextFile("test.htm", 2)
    objFile.WriteLine DataArea.InnerHTML
    objFile.Close
End Sub

