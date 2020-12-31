' Description: Demonstration script that counts the number of key-item pairs in a Script Runtime Dictionary. Script must be run on the local computer.


Set objDictionary = CreateObject("Scripting.Dictionary")

objDictionary.Add "Printer 1", "Printing"   
objDictionary.Add "Printer 2", "Offline"
objDictionary.Add "Printer 3", "Printing"
Wscript.Echo objDictionary.Count

