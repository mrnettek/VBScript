Const wdOrientLandscape = 1

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()

objDoc.PageSetup.Orientation = wdOrientLandscape
  


