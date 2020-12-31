Set objWord = CreateObject("Word.Application")
objWord.Visible = TRUE
Set objDoc = objWord.Documents.Add()
Set objRange = objDoc.Range()
Set objLink = objDoc.Hyperlinks.Add _
    (objRange, " http://www.microsoft.com/technet/scriptcenter ", , , "Script Center")
  


