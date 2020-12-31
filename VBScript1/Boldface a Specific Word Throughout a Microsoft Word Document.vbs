Const wdReplaceAll  = 2

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")
Set objSelection = objWord.Selection

objSelection.Find.Text = "Fabrikam"
objSelection.Find.Forward = TRUE
objSelection.Find.MatchWholeWord = TRUE

objSelection.Find.Replacement.Font.Bold = True

objSelection.Find.Execute ,,,,,,,,,,wdReplaceAll
  


