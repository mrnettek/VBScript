On Error Resume Next

Dim arrFolders()
intSize = 0

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

strFolderName = "c:\scripts"

GetSubFolders strFolderName

Sub GetSubFolders(strFolderName)
    Set colSubfolders = objWMIService.ExecQuery _
        ("Associators of {Win32_Directory.Name='" & strFolderName & "'} " _
            & "Where AssocClass = Win32_Subdirectory " _
                & "ResultRole = PartComponent")

    For Each objFolder in colSubfolders
        strFolderName = objFolder.Name
        ReDim Preserve arrFolders(intSize)
        arrFolders(intSize) = strFolderName
        intSize = intSize + 1
        GetSubFolders strFolderName
    Next
End Sub

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each strFolder in arrFolders
    strFolderName = strFolder 
    strNewFolder = Replace(strFolderName, "c:\scripts", "c:\test")
    Set objFolder = objFSO.CreateFolder(strNewFolder)
Next
  


