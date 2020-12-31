Const NUMBER_OF_ROWS = 1
Const NUMBER_OF_COLUMNS = 2

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()

Set objRange = objDoc.Range()
objDoc.Tables.Add objRange, NUMBER_OF_ROWS, NUMBER_OF_COLUMNS
Set objTable = objDoc.Tables(1)

objTable.Cell(1, 1).Range.Text = "Date"
objTable.Cell(1, 2).Range.Text = "Notes"

dtmDate = #4/1/2006#
dtmMonth = Month(dtmDate)
i = 2

Do While True
    objTable.Rows.Add()
    objTable.Cell(i, 1).Range.Text = dtmDate
    dtmDate = dtmDate + 1
    If Month(dtmDate) <> dtmMonth Then
        Exit Do
    End If
    i = i +1
Loop
  


