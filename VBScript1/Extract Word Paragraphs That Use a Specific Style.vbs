Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")

Set objSelection = objWord.Selection
Set colParagraphs = objDoc.Paragraphs

For Each objParagraph in colParagraphs
    If objParagraph.Style <> "Heading 1" Then
        objParagraph.Range.Select
        objSelection.Cut
    End If    
Next

objDoc.SaveAs("C:\Scripts\Headings.doc")
objWord.Quit
  


