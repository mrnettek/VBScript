dtmDate = Date - 7

strDay = Day(dtmDate)

If Len(strDay) < 2 Then
    strDay = "0" & strDay
End If

strMonth = Month(dtmDate)

If Len(strMonth) < 2 Then
    strMonth = "0" & strMonth
End If

strYear = Year(dtmDate)

strTargetDate = strYear & strMonth & strDay

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set FileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Scripts'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In FileList
    strDate = Left(objFile.CreationDate, 8)
    If strDate < strTargetDate Then
        If objFile.Extension = "bak" Then
            objFile.Delete
        End If
    End If
Next
  


