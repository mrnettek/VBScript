' Description: Deletes the local user account Admin2 from a computer named atl-ws-01.


strComputer = "atl-ws-01"
strUser = "Admin2"

Set objComputer = GetObject("WinNT://" & strComputer & "")
objComputer.Delete "user", strUser

