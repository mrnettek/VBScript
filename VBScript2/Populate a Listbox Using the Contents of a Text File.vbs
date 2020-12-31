' Description: Sample HTML function that opens a text file and adds the contents to a listbox each time a Web page or HTA is loaded.


Sub Window_Onload
    ForReading = 1
    strNewFile = "computers.txt"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile _
        (strNewFile, ForReading)
    Do Until objFile.AtEndOfStream
        strLine = objFile.ReadLine
        Set objOption = Document.createElement("OPTION")
        objOption.Text = strLine
        objOption.Value = strLine
        AvailableComputers.Add(objOption)
    Loop
    objFile.Close
End Sub

