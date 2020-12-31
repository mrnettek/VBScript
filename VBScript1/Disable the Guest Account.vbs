' Description: Uses the Shell object to disable the Guest account on the local computer.


Set objComputer = CreateObject("Shell.LocalMachine")
objComputer.DisableGuest(0)

