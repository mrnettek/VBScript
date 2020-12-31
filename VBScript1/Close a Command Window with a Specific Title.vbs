Set objWord = CreateObject("Word.Application")
Set colTasks = objWord.Tasks

If colTasks.Exists("My Window") Then
    colTasks("My Window").Close
End If

objWord.Quit
  


