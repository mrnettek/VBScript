' Description: Demonstration script that uses ScriptPW.dll to mask passwords entered at the command line.


Set objPassword = CreateObject("ScriptPW.Password") 
WScript.StdOut.Write "Please enter your password:" 

strPassword = objPassword.GetPassword() 
Wscript.Echo
Wscript.Echo "Your password is: " & strPassword

