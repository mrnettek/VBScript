strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & _
    "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_Summary")

For Each strIESetting in colIESettings
    Wscript.Echo "Version: " & strIESetting.Version
    Wscript.Echo "Product ID: " & strIESetting.ProductID
    Wscript.Echo "Cipher strength: " & strIESetting.CipherStrength
Next
  


