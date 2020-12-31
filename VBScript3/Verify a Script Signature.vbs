' Description: Verifies that an individual script has been digitally signed.


blnShowGUI = False

Set objSigner = WScript.CreateObject("Scripting.Signer")
blnIsSigned = objSigner.VerifyFile("C:\Scripts\CreateUser.vbs", blnShowGUI)

If blnIsSigned Then
    WScript.Echo "Script has been signed."
Else
   WScript.Echo " Script has not been signed."
End If

