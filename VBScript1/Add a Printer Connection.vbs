' Description: Adds a printer connection to a network printer. Script must be run on the local computer.


Set WshNetwork = CreateObject("WScript.Network")

WshNetwork.AddWindowsPrinterConnection "\\PrintServer1\Xerox300"
WshNetwork.SetDefaultPrinter "\\PrintServer1\Xerox300"

