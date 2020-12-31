HKEY_CURRENT_USER = &H80000001

strComputer = "."

Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "Console"
strValueName = "HistoryBufferSize"

objRegistry.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, strValueName, dwValue

Wscript.Echo dwValue
  


