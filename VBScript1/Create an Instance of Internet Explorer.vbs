' Description: Demonstration script that creates an instance of Internet Explorer, opened to a blank page.


Set objExplorer = WScript.CreateObject("InternetExplorer.Application")

objExplorer.Navigate "about:blank"   
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width=300
objExplorer.Height = 150 
objExplorer.Left = 0
objExplorer.Top = 0
objExplorer.Visible = 1

