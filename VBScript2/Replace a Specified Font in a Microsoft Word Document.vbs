Const wdReplaceAll = 2

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")
Set objSelection = objWord.Selection

objSelection.Find.Font.Name = "Gigi"
objSelection.Find.Forward = TRUE

objSelection.Find.Replacement.Font.Name = "Arial"

objSelection.Find.Execute "", ,False, , , , , , , ,wdReplaceAll
  


