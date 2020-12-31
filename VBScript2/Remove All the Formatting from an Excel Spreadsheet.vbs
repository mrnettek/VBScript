Const xlHAlignCenter = -4108

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

For i = 1 to 14
    objExcel.Cells(i, 1).Value = i
    objExcel.Cells(i, 2).Interior.ColorIndex = i
Next

Set objRange = objWorksheet.UsedRange
objRange.HorizontalAlignment = xlHAlignCenter 
objRange.Font.Bold = True

Wscript.Sleep 5000

objRange.ClearFormats
  


