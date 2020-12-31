Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

objWorksheet.Cells(1,1) = "Operating System"
objWorksheet.Cells(2,1) = "Windows Server 2003"
objWorksheet.Cells(3,1) = "Windows XP"
objWorksheet.Cells(4,1) = "Windows 2000"
objWorksheet.Cells(5,1) = "Windows NT 4.0"
objWorksheet.Cells(6,1) = "Other"

objWorksheet.Cells(1,2) = "Number of Computers"
objWorksheet.Cells(2,2) = 145
objWorksheet.Cells(3,2) = 987
objWorksheet.Cells(4,2) = 611
objWorksheet.Cells(5,2) = 41
objWorksheet.Cells(6,2) = 56

Set objRange = objWorksheet.UsedRange
objRange.Select

Set colCharts = objExcel.Charts
colCharts.Add()

Set objChart = colCharts(1)
objChart.Activate

objChart.Export "C:\Scripts\Test.jpg", "JPG"
  


