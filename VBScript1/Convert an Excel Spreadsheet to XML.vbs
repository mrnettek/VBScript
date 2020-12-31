Const xlXMLSpreadsheet = 46

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

k = 1

For i = 1 to 5
    For j = 1 to 3
        objWorksheet.Cells(i,j) = k
        k = k + 1
    Next
Next

objWorkbook.SaveAs "C:\Scripts\Test.xml", xlXMLSpreadsheet
  


