strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colSubfolders = objWMIService.ExecQuery _
    ("Associators of {Win32_Directory.Name='C:\Absentee Reports'} " _
        & "Where AssocClass = Win32_Subdirectory " _
            & "ResultRole = PartComponent")

For Each objFolder in colSubfolders
    Set colFiles = objWMIService.ExecQuery _
        ("ASSOCIATORS OF {Win32_Directory.Name='" & objFolder.Name & "'} Where " _
            & "ResultClass = CIM_DataFile")

    If colFiles.Count => 4 Then 
        Select Case colFiles.Count
            Case 4
                Wscript.Echo objFolder.Name & " has 4 files in it."
            Case 5
                Wscript.Echo objFolder.Name & " has 5 files in it."
            Case 6
                Wscript.Echo objFolder.Name & " has 6 files in it."
            Case 7
                Wscript.Echo objFolder.Name & " has 7 files in it."
            Case 8
                Wscript.Echo objFolder.Name & " has 8 files in it."
        End Select
    End If
Next
  


