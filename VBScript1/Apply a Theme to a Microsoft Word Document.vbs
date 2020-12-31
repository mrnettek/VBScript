Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
objSelection.TypeText "Here is a bulleted list."
objSelection.TypeParagraph()
objSelection.TypeParagraph()

Set objRange = objDoc.Paragraphs(1).Range
objRange.Style = "Normal"

Set objRange = objDoc.Paragraphs(3).Range
objRange.ListFormat.ApplyBulletDefault

objSelection.TypeText "Item 1"
objSelection.TypeParagraph()
objSelection.TypeText "Item 2"
objSelection.TypeParagraph()
objSelection.TypeText "Item 3"
objSelection.TypeParagraph()

Set objRange = objDoc.Paragraphs(6).Range
objRange.Style = "Normal"

objSelection.TypeParagraph()
objSelection.TypeText "No longer in a bulleted list."

objDoc.ApplyTheme "Balance"
  


