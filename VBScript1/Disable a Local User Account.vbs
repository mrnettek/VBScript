' Description: Disables the local Guest account on a computer named atl-ws-01.


strComputer = "atl-ws-01"
Set objUser = GetObject("WinNT://" & strComputer & "/Guest")

objUser.AccountDisabled = True
objUser.SetInfo

