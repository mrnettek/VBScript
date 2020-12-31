strComputer = "atl-ws-01"
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators")
Set objUser = GetObject("WinNT://" & strComputer & "/kenmyer")
objGroup.Add(objUser.ADsPath)
  


