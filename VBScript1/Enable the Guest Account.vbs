' Description: Uses the Shell object to enable the Guest account on the local computer.


Set objComputer = CreateObject("Shell.LocalMachine")
objComputer.EnableGuest(1)

