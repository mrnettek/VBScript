Const wdFormatDocument = 0

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open("c:\scripts\test.rtf")
objDoc.SaveAs "C:\Scripts\test.doc", wdFormatDocument

objWord.Quit
  


