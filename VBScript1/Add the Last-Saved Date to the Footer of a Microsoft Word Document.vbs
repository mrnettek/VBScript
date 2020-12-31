Const wdFieldSaveDate = 22
Const wdAlignParagraphCenter = 1

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()

Set objRange = objDoc.Sections(1).Footers(1).Range
objDoc.Fields.Add objRange, wdFieldSaveDate

objRange.ParagraphFormat.Alignment = wdAlignParagraphCenter

objDoc.SaveAs("C:\Scripts\Test.doc")
  


