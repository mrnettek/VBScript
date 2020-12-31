Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)

i = 2

Do Until x = 1
    If objWorksheet.Cells(i,1) = "" Then
        Exit Do
    End If

    intColor = objWorksheet.Cells(i,2).Interior.ColorIndex

    Select Case intColor
        Case 3 strStatus = "Behind schedule"
        Case 4 strStatus = "Project complete"
        Case 6 strStatus = "Project on schedule"
        Case Else strStatus = "No information"
    End Select

    Wscript.Echo objWorksheet.Cells(i,1).Value & " -- " & strStatus
    i = i + 1
Loop
  


