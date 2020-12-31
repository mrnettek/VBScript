strComputer = "."

Set objExplorer = WScript.CreateObject("InternetExplorer.Application")
objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Visible = 1

Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
  ("SELECT * FROM Win32_Service")
 
For Each objItem in colItems
    strHTML = objItem.DisplayName  & " = " & objItem.State & "<BR>"
    objExplorer.Document.Body.InnerHTML = strHTML
Next
  


