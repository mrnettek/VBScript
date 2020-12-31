Const HKEY_CURRENT_USER = &H80000001

strComputer = "."

Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
 
strKeyPath = "Control Panel\Screen Saver.Marquee"
strValueName = "Text"
strStringValue = FormatDateTime(Date, vbLongDate)

objRegistry.SetStringValue HKEY_CURRENT_USER, strKeyPath, strValueName, strStringValue
  


