Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
 
Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\5.0\Cache\Content"
strValueName = "CacheLimit"
dwValue = 358400
objRegistry.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Cache\Content"
strValueName = "CacheLimit"
dwValue = 358400
objRegistry.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue
  


