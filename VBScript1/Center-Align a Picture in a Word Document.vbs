Const wdShapeCenter = -999995 
Const wdWrapSquare = 0

Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Add()

Set objShape = objDoc.Shapes
objShape.AddPicture("C:\Scripts\Welder-small.jpg")

Set objShapeRange = objDoc.Shapes.Range(1)
objShapeRange.Left = wdShapeCenter

objShapeRange.WrapFormat.Type = wdWrapSquare
  


