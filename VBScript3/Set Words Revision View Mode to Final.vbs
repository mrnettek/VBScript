Const wdRevisionsViewFinal = 0

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open("c:\scripts\test.doc")

Set objView = objWord.ActiveDocument.ActiveWindow.View

objView.RevisionsView = wdRevisionsViewFinal 
objView.ShowRevisionsAndComments = False 

objWord.Visible = True
  


