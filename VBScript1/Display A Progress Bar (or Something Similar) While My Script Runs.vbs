On Error Resume Next

Set objExplorer = CreateObject _
    ("InternetExplorer.Application")

objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width = 400
objExplorer.Height = 200 
objExplorer.Visible = 1             

objExplorer.Document.Title = "Logon script in progress"
objExplorer.Document.Body.InnerHTML = "Your logon script is being processed. " _
    & "This might take several minutes to complete."

Wscript.Sleep 10000

objExplorer.Document.Body.InnerHTML = "Your logon script is now complete."

Wscript.Sleep 5000
objExplorer.Quit
  


