' Description: Deletes a local group named FinanceUsers from a computer named atl-ws-01.


strComputer = "atl-ws-01"
Set objComputer = GetObject("WinNT://" & strComputer & "")
objComputer.Delete "group", "FinanceUsers"

