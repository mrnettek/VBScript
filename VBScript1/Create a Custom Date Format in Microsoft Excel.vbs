Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

objExcel.Workbooks.Add

objExcel.Cells(1,1).Value = "01/01/2006"
objExcel.Cells(1,1).NumberFormat = "yyyymmdd"
  


