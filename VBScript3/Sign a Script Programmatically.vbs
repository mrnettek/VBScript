' Description: Uses the Scripting Runtime Signer object to digitally sign a script. Requires a valid digital certificate.


Set objSigner = WScript.CreateObject("Scripting.Signer")
objSigner.SignFile "C:\Scripts\CreateUsers.vbs", "IT Department"

