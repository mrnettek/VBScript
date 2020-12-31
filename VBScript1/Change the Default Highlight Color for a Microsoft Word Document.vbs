Const wdBrightGreen = 4

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()

Set objOptions = objWord.Options
objOptions.DefaultHighlightColorIndex = wdBrightGreen

objWord.Visible = True
  


