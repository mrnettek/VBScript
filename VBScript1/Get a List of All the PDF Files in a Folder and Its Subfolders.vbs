strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

strFolderName = "C:\Test"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.CreateTextFile("C:\Scripts\Test.txt")

Set colSubfolders = objWMIService.ExecQuery _
    ("Associators of {Win32_Directory.Name='" & strFolderName & "'} " _
        & "Where AssocClass = Win32_Subdirectory " _
            & "ResultRole = PartComponent")

Set colFiles = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='" & strFolderName & "'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile in colFiles
    If objFile.Extension = "pdf" Then
        objTextFile.WriteLine objFile.FileName 
    End If
Next

For Each objFolder in colSubfolders
    GetSubFolders strFolderName
Next

Sub GetSubFolders(strFolderName)

    Set colSubfolders2 = objWMIService.ExecQuery _
        ("Associators of {Win32_Directory.Name='" & strFolderName & "'} " _
            & "Where AssocClass = Win32_Subdirectory " _
                & "ResultRole = PartComponent")

    For Each objFolder2 in colSubfolders2
        strFolderName = objFolder2.Name

    Set colFiles = objWMIService.ExecQuery _
        ("ASSOCIATORS OF {Win32_Directory.Name='" & strFolderName & "'} Where " _
            & "ResultClass = CIM_DataFile")

    For Each objFile in colFiles
        If objFile.Extension = "pdf" Then
            objTextFile.WriteLine objFile.FileName 
        End If
    Next

        GetSubFolders strFolderName
    Next
End Sub
  


