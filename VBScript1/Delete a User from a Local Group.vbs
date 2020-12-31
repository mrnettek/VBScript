' Description: Removes kenmyer from the local Administrators group on a computer named atl-ws-01.


strComputer = "atl-ws-01"
Set objGroup = GetObject("WinNT://" & strComputer & "/Adminstrators,group")
Set objUser = GetObject("WinNT://" & strComputer & "/kenmyer,user")
 
objGroup.Remove(objUser.ADsPath)

