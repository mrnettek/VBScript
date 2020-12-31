Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
Set objWorksheet = objWorkbook.Worksheets(1)

Set objRange = objWorksheet.UsedRange

For Each objCell in objRange
    If IsNumeric(objCell.Value) Then
        If objCell.Value > 999 Then
            objCell.Value = 999
        End If
    End If
Next
  


