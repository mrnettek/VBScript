Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
Set objRegistry = GetObject("winmgmts:\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName = "Test Value"
objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

If IsNull(strValue) Then
    Wscript.Echo "The registry key does not exist."
Else
    Wscript.Echo "The registry key exists."
End If
  


