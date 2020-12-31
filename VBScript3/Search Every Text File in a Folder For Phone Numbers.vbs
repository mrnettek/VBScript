Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Scripts'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFileList
    strName = objFile.FileName 
    strName = Replace(strName, "_", " ")
    Wscript.Echo strName

    Set objFile = objFSO.OpenTextFile(objFile.Name, ForReading)
    strSearchString = objFile.ReadAll
    objFile.Close

    Set objRegEx = CreateObject("VBScript.RegExp")

    objRegEx.Global = True   
    objRegEx.Pattern = "\d{3}-\d{3}-\d{4} \([a-zA-Z]*\)"

    Set colMatches = objRegEx.Execute(strSearchString)  

    If colMatches.Count > 0 Then
       For Each strMatch in colMatches   
           Wscript.Echo strMatch.Value 
       Next
    End If

    Wscript.Echo
Next
  


