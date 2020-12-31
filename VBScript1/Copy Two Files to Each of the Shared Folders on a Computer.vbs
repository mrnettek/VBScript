Const OverwriteExisting = TRUE

strComputer = "atl-fs-01"

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * From Win32_Share Where Type = 0")

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each objItem in colItems
    strFolderName = objItem.Name
    If InStr(strFolderName, "$") = 0 Then
        strPath = "\\" & strComputer & "\" & strFolderName & "\"
        objFSO.CopyFile "C:\Scripts\Test.txt", strPath, OverwriteExisting
        objFSO.CopyFile "C:\Scripts\Test2.txt", strPath, OverwriteExisting
    End If
Next
  


