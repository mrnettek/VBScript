Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.TypeText "This is the first paragraph."
objSelection.TypeParagraph()
objSelection.TypeText "This is the second paragraph."
objSelection.TypeParagraph()
objSelection.TypeText "This is the third paragraph."
objSelection.TypeParagraph()

Set objRange = objDoc.Range()
objRange.Font.Name = "Arial"
objRange.Font.Size = 10
  


