Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")
Set objSelection = objWord.Selection

objSelection.Find.Forward = True
objSelection.Find.MatchWildcards = True
objSelection.Find.Text = "\[\[*\]\]"

Do While True
    objSelection.Find.Execute
    If objSelection.Find.Found Then
        strWord = objSelection.Text
        strWord = Replace(strWord, "[[", "")
        strWord = Replace(strWord, "]]", "")
        Wscript.Echo strWord
    Else
        Exit Do
    End If
Loop
  


