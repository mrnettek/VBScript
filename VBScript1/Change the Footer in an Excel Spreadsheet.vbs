Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.PageSetup.LeftFooter = "Left footer"
objWorksheet.PageSetup.CenterFooter = "Center footer"
objWorksheet.PageSetup.RightFooter = "Right footer"
  


