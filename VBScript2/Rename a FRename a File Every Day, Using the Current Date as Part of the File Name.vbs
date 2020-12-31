strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

strMonth = Month(Date - 1)

If Len(strMonth) = 1 Then
    strMonth = "0" & strMonth
End If

strDay = Day(Date - 1)

If Len(strDay) = 1 Then
    strDay = "0" & strDay
End If

strYear = Year(Date - 1)

strFileName = "C:\\Test\\BackupFile-" & strMonth & strDay & strYear  & ".txt"

Set colFiles = objWMIService.ExecQuery _
    ("Select * From CIM_DataFile Where Name = '" & strFileName & "'")

For Each objFile in colFiles

    strMonth = Month(Date)

    If Len(strMonth) = 1 Then
        strMonth = "0" & strMonth
    End If

    strDay = Day(Date)

    If Len(strDay) = 1 Then
        strDay = "0" & strDay
    End If

    strYear = Year(Date)

    strNewFileName = "C:\\Test\\BackupFile-" & strMonth & strDay & strYear  & ".txt"

    errResult = objFile.Rename(strNewFileName)
Next
  


