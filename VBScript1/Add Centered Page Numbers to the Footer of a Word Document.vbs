Const wdAlignPageNumberCenter = 1

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()

objDoc.Sections(1).Footers(1).PageNumbers.Add(wdAlignPageNumberCenter)
  


