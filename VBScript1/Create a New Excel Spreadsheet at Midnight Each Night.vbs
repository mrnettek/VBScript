dtmDate = Date

strMonth = Month(Date)
strDay = Day(Date)
strYear = Right(Year(Date),2)

strFileName = "C:\Scripts\" & strMonth & "-" & strDay & "-" & strYear & ".xls"

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
objWorkbook.SaveAs(strFileName)

objExcel.Quit
  


