Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.TypeText "This paragraph is followed by a paragraph return."
objSelection.TypeParagraph()

objSelection.TypeText "This paragraph is followed by a line break." & Chr(11)

objSelection.TypeText "This paragraph is also followed by a line break." & Chr(11)

objSelection.TypeText "This paragraph is followed by a paragraph return."
objSelection.TypeParagraph()
  


