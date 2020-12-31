strComputer = "atl-ws-01"

Set objAdmins = GetObject("WinNT://" & strComputer & "/Administrators")
Set objGroup = GetObject("WinNT://fabrikam/accounting")

objAdmins.Add(objGroup.ADsPath)
  


