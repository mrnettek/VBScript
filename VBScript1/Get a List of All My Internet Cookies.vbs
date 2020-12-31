Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
strValueName = "Cookies"
objRegistry.GetStringValue HKEY_CURRENT_USER,strKeyPath,strValueName,strFolder

Set colFileList = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='" & strFolder & "'} Where " _
        & "ResultClass = CIM_DataFile")

For Each objFile In colFileList
    Wscript.Echo objFile.FileName
Next
  


