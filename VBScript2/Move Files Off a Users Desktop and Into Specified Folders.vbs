Const DESKTOP = &H10&

Set objShell = CreateObject("Shell.Application")

Set objFolder = objShell.Namespace(DESKTOP)
Set objFolderItem = objFolder.Self
strPath = objFolderItem.Path

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='" & strpath & "'} Where " _
        & "ResultClass = CIM_DataFile")

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each objFile in colFiles
    strDesktopFolder = strPath & "\" & objFile.Extension
    If objFSO.FolderExists(strDesktopFolder) Then
        strTarget = strDesktopFolder & "\" 
        objFSO.MoveFile objFile.Name, strTarget
    Else
        Set objFolder = objFSO.CreateFolder(strDesktopFolder)
        strTarget = strDesktopFolder & "\" 
        objFSO.MoveFile objFile.Name, strTarget
    End If
Next
  


