Set objPassword = CreateObject("ScriptPW.Password") 
WScript.StdOut.Write "Please enter your password:" 

strPassword = objPassword.GetPassword() 
Wscript.Echo
Wscript.Echo "Your password is: " & strPassword
  


