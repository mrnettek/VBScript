Const wdColorRed = 255

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")
Set objSelection = objWord.Selection

objSelection.Find.Forward = True
objSelection.Find.Format = True
objSelection.Find.Font.Color = wdColorRed

Do While True
    objSelection.Find.Execute
    If objSelection.Find.Found Then
        Wscript.Echo objSelection.Text
    Else
        Exit Do
    End If
Loop
  


