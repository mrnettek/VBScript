Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.InsertSymbol 74, "Wingdings"
  


