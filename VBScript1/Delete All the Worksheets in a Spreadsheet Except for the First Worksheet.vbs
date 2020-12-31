Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open("C:\Scripts\Test.xls")
i = objWorkbook.Worksheets.Count

Do Until i = 1 
  objWorkbook.Worksheets(i).Delete
  i = i - 1
Loop
  


