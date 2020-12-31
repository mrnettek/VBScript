strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFolders = objWMIService.ExecQuery _
    ("SELECT * FROM Win32_Directory WHERE FileName LIKE 'December%'")

For Each objFolder in colFolders
    Wscript.Echo objFolder.Name
Next
  


