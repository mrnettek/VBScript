On Error Resume Next

Dim arrFolders()
intSize = 0

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

strFolderName = "c:\scripts"

Set colSubfolders = objWMIService.ExecQuery _
    ("Associators of {Win32_Directory.Name='" & strFolderName & "'} " _
        & "Where AssocClass = Win32_Subdirectory " _
            & "ResultRole = PartComponent")

ReDim Preserve arrFolders(intSize)
arrFolders(intSize) = strFolderName
intSize = intSize + 1

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
        ReDim Preserve arrFolders(intSize)
        arrFolders(intSize) = strFolderName
        intSize = intSize + 1
        GetSubFolders strFolderName
    Next
End Sub

For Each strFolder in arrFolders
    Wscript.Echo strFolder
Next
  


