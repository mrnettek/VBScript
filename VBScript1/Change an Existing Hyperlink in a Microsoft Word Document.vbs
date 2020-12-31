Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("c:\scripts\test.doc")

Set colHyperlinks = objDoc.Hyperlinks

For Each objHyperlink in colHyperlinks
    If objHyperlink.Address = "http://www.microsoft.com/" Then                                
        objHyperlink.Address = "http://www.microsoft.com/technet/scriptcenter/"
    End If
Next
  


