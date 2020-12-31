' Description: Demonstration script that opens and prints and existing Microsoft Word document.


Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open("c:\scripts\inventory.doc")

objDoc.PrintOut()
objWord.Quit

