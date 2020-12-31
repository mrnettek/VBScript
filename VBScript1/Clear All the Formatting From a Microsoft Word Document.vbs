Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")

objDoc.Select

Set objSelection = objWord.Selection
objSelection.ClearFormatting
  


