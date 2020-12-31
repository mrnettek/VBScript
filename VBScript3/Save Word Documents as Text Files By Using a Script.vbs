Const wdFormatText = 2

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open("c:\scripts\mylog.doc")
objDoc.SaveAs "c:\scripts\mylog.txt", wdFormatText

objWord.Quit
  


