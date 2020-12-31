arrValues = Array(1,5,7,9,11,13,15,17)

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

i = 1

For Each intValue in arrValues
    objExcel.Cells(i, 1).Value = intValue
    i = i + 1
Next

objExcel.Cells(i, 1).Formula = "=SUM(A1:A" & i - 1 & ")"
  


