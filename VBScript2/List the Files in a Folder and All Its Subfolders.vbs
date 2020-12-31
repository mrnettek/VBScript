strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

strFolderName = "c:\scripts"

Set colSubfolders = objWMIService.ExecQuery _
    ("Associators of {Win32_Directory.Name='" & strFolderName & "'} " _
        & "Where AssocClass = Win32_Subdirectory " _
            & "ResultRole = PartComponent")

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
        Wscript.Echo objFolder2.Name
        GetSubFolders strFolderName
    Next
End Sub
  


