Const ADS_NAME_INITTYPE_GC = 3
Const ADS_NAME_TYPE_NT4 = 3
Const ADS_NAME_TYPE_GUID = 7

strUserName = "fabrikam\kenmyer"

Set objTranslator = CreateObject("NameTranslate")

objTranslator.Init ADS_NAME_INITTYPE_GC, ""
objTranslator.Set ADS_NAME_TYPE_NT4, strUserName

strUserGUID = objTranslator.Get(ADS_NAME_TYPE_GUID)

Wscript.Echo strUserGUID
  


