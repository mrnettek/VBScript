' Description: Demonstration script that displays the various colors -- and their related color index -- available when programmatically controlling Microsoft Excel.


Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True
objExcel.Workbooks.Add

For i = 1 to 56
    objExcel.Cells(i, 1).Value = i
    objExcel.Cells(i, 1).Interior.ColorIndex = i
Next

