Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")
Set objSelection = objWord.Selection

objSelection.Find.Forward = True
objSelection.Find.Format = True
objSelection.Find.Font.Bold = True

Do While True
    objSelection.Find.Execute
    If objSelection.Find.Found Then
        objSelection.Text = "<b>" & objSelection.Text & "</b>"
    Else
        Exit Do
    End If
Loop
  


