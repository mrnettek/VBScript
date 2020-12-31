Set objRegEx = CreateObject("VBScript.RegExp")

objRegEx.Global = True   
objRegEx.Pattern = "[^A-Za-z]"

strSearchString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890abcdefghijklmnopqrstuvwxyz!@#$%&*(_+"

strSearchString = objRegEx.Replace(strSearchString, "")

Wscript.Echo strSearchString
  


