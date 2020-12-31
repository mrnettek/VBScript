' Description: Demonstration script that creates and displays a new Microsoft Word document.


Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()

