Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)

objExcel.Cells(1, 1).Value = "A"
objExcel.Cells(2, 1).Value = "B"
objExcel.Cells(3, 1).Value = "C"
objExcel.Cells(4, 1).Value = "D"
objExcel.Cells(5, 1).Value = "E"
objExcel.Cells(6, 1).Value = "F"
objExcel.Cells(7, 1).Value = "G"
objExcel.Cells(8, 1).Value = "H"

objExcel.Cells(1, 1).Interior.ColorIndex = 7
objExcel.Cells(2, 1).Interior.ColorIndex = 8
objExcel.Cells(3, 1).Interior.ColorIndex = 9
objExcel.Cells(4, 1).Interior.ColorIndex = 10
objExcel.Cells(5, 1).Interior.ColorIndex = 7
objExcel.Cells(6, 1).Interior.ColorIndex = 7
objExcel.Cells(7, 1).Interior.ColorIndex = 8
objExcel.Cells(8, 1).Interior.ColorIndex = 10

i = 1

Do Until objExcel.Cells(i, 1).Value = ""
    intColor = objExcel.Cells(i, 1).Interior.ColorIndex

    Select Case intColor
        Case 7 intSum = intSum + 5
        Case 8 intSum = intSum + 10
        Case 9 intSum = intSum + 15
        Case 10 intSum = intSum + 20
    End Select

    i = i + 1
Loop

Wscript.Echo intSum
  


