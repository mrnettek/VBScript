Const wdAlignParagraphRight = 2 
Const NUMBER_OF_ROWS = 1
Const NUMBER_OF_COLUMNS = 3

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()

Set objRange = objDoc.Range()
objDoc.Tables.Add objRange, NUMBER_OF_ROWS, NUMBER_OF_COLUMNS
Set objTable = objDoc.Tables(1)

x=2

objTable.Cell(1, 1).Range.Text = "Process Name"
objTable.Cell(1, 2).Range.text = "Process ID"
objTable.Cell(1, 3).Range.text = "Handle Count"

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Process")
For Each objItem in colItems
    If x > 1 Then
        objTable.Rows.Add()
    End If
    objTable.Cell(x, 1).Range.Text = objItem.Name
    objTable.Cell(x, 2).Range.text = objItem.ProcessID
    objTable.Cell(x, 3).Range.text = objItem.HandleCount
    x = x + 1
Next

objDoc.Tables(1).Columns(3).Select
Set objSelection = objWord.Selection

objSelection.ParagraphFormat.Alignment = wdAlignParagraphRight
  


