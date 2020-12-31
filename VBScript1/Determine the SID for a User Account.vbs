strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set objAccount = objWMIService.Get _
    ("Win32_UserAccount.Name='kenmyer',Domain='fabrikam'")
Wscript.Echo objAccount.SID
  


