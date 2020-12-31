strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Process Where Name = 'Notepad.exe'")

If colItems.Count < 2 Then
    Wscript.Quit
End If

dtmTarget = Now

For Each objItem in colItems
    dtmDateHolder = objItem.CreationDate
    
    dtmDateHolder = CDate(Mid(dtmDateHolder, 5, 2) & "/" & _
        Mid(dtmDateHolder, 7, 2) & "/" & Left(dtmDateHolder, 4) _
            & " " & Mid (dtmDateHolder, 9, 2) & ":" & _
                Mid(dtmDateHolder, 11, 2) & ":" & Mid(dtmDateHolder, 13, 2))

    If dtmDateHolder < dtmTarget Then
         intProcessID = objItem.ProcessID
         dtmTarget = dtmDateHolder
    End If
Next

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Process Where Name = 'Notepad.exe' " & _
        "AND ProcessID <> " & intProcessID)

For Each objItem in colItems
    objItem.Terminate
Next
  


