Const END_OF_STORY = 6

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.TypeText "Table 1"
objSelection.TypeParagraph()

Set objRange = objSelection.Range
objDoc.Tables.Add objRange, 1, 2
Set objTable = objDoc.Tables(1)

objTable.Cell(1, 1).Range.Text = "This is cell 1."
objTable.Cell(1, 2).Range.Text = "This is cell 2."
objSelection.EndKey END_OF_STORY
objSelection.TypeParagraph()

objSelection.TypeText "Table 2"
objSelection.TypeParagraph()

Set objRange = objSelection.Range
objDoc.Tables.Add objRange, 1, 2
Set objTable = objDoc.Tables(2)

objTable.Cell(1, 1).Range.Text = "This is cell 1."
objTable.Cell(1, 2).Range.Text = "This is cell 2."

objSelection.EndKey END_OF_STORY
objSelection.TypeParagraph()
  


