' Description: Adds a user (kenmyer) to the local Administrators group on a computer named atl-ws-01.


strComputer = "atl-ws-01"
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")

Set objUser = GetObject("WinNT://" & strComputer & "/kenmyer,user")
objGroup.Add(objUser.ADsPath)

