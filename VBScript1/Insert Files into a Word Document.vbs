Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.TypeText "Hardware Inventory"
objSelection.TypeParagraph()
objSelection.InsertFile("C:\Scripts\Hardware.txt")

objSelection.TypeParagraph()

objSelection.TypeText "Software Inventory"
objSelection.TypeParagraph()
objSelection.InsertFile("C:\Scripts\Software.txt")
  


