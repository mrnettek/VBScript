strComputer = "atl-ws-01"
Set objUser = GetObject("WinNT://" & strComputer & "/Guest")

Wscript.Echo "Guest account disabled: " & objUser.AccountDisabled
  


