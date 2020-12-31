Set objFSO = CreateObject("Scripting.FileSystemObject")

For i = 1 to 5
    Wscript.Echo i
    Wscript.Sleep 1000
Next

strScript = Wscript.ScriptFullName
objFSO.DeleteFile(strScript)
  


