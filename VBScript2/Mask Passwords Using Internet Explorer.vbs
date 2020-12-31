' Description: Demonstration script that creates an instance of Internet Explorer, and retrieves a password typed into a password-style text box. Requires a Web page named password.htm with the appropriate text box.


Set objExplorer = WScript.CreateObject _
    ("InternetExplorer.Application", "IE_")

objExplorer.Navigate "file:///c:\scripts\password.htm"   
objExplorer.Visible = 1             
objExplorer.ToolBar = 0
objExplorer.StatusBar = 0
objExplorer.Width=400
objExplorer.Height = 250 
objExplorer.Left = 0
objExplorer.Top = 0

Do While (objExplorer.Document.Body.All.OKClicked.Value = "")
    Wscript.Sleep 250                 
Loop 

strPassword = objExplorer.Document.Body.All.PasswordBox.Value
objExplorer.Quit

Wscript.Sleep 250
Wscript.Echo strPassword

