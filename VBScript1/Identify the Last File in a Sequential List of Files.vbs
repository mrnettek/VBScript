strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\PDFs'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFileList
    strFile = objFile.FileName
Next

Wscript.Echo strFile
  


