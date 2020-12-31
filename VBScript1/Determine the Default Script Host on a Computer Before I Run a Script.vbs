Const HKEY_CLASSES_ROOT = &H80000000

strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "VBSFile\Shell\Open\Command"
objRegistry.GetExpandedStringValue HKEY_CLASSES_ROOT,strKeyPath,vbNullString,strValue

strValue = LCase(strValue)

If InStr(strValue, "wscript.exe") then
    Wscript.Echo "WScript"
Else
    Wscript.Echo "CScript"
End If
  


