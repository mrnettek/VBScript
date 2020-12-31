On Error Resume Next

Set objExplorer = CreateObject("InternetExplorer.Application")

objExplorer.Navigate "http://www.microsoft.com/technet/scriptcenter"   
objExplorer.Visible = 1

Wscript.Sleep 5000

Set objDoc = objExplorer.Document

Do While True
    Wscript.Sleep 30000
    objDoc.Location.Reload(True)
    If Err <> 0 Then
        Wscript.Quit
    End If
 Loop
  


