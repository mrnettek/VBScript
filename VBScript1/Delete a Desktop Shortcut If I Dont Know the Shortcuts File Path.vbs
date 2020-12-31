strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_ShortcutFile Where FileName = 'Digital Voice Editor 2'")

For Each objItem in colItems
    If Instr(objItem.Name, "desktop") Then
        strPath = objItem.Name
        strPath = Replace(strPath, "\", "\\")
        Set colFiles = objWMIService.ExecQuery _
            ("Select * From CIM_Datafile Where Name = '" & strpath & "'")
        For Each objFile in colFiles
            objFile.Delete
        Next
    End If
Next
  


