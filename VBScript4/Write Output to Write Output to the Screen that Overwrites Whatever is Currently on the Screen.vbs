Set objExplorer = CreateObject("InternetExplorer.Application")
objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width = 400
objExplorer.Height = 200 
objExplorer.Left = 0
objExplorer.Top = 0

Do While (objExplorer.Busy)
    Wscript.Sleep 200
 Loop    

objExplorer.Document.Title = "Process Information"   
objExplorer.Visible = 1  
        
objExplorer.Document.Body.InnerHTML = "Retrieving process information." 

Wscript.Sleep 2000

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Process")
For Each objItem in colItems
    objExplorer.Document.Body.InnerHTML = objItem.Name
    Wscript.Sleep 500
Next

objExplorer.Document.Body.InnerHTML = "Process information retrieved."
Wscript.Sleep 3000
objExplorer.Quit
  


