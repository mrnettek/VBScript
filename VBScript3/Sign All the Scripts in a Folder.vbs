' Description: Uses the Scripting Runtime Signer object to digitally sign all the scripts in a folder. Requires a valid digital certificate.


Set objSigner = WScript.CreateObject("Scripting.Signer")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("c:\scripts")
Set colListOfFiles = objFolder.Files

For Each objFile in colListOfFiles
    objSigner.SignFile objFile.Path, "IT Department"
Next

