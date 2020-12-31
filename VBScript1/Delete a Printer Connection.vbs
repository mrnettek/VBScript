' Description: Removes a printer connection to a network printer. Script must be run on the local computer.


Set objNetwork = WScript.CreateObject("WScript.Network")
objNetwork.RemovePrinterConnection "\\PrintServer\xerox3006"

