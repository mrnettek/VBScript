' Description: Demonstration script that deletes a specific key-item pair from a Script Runtime Dictionary. Script must be run on the local computer.


Set objDictionary = CreateObject("Scripting.Dictionary")

objDictionary.Add "Printer 1", "Printing"   
objDictionary.Add "Printer 2", "Offline"
objDictionary.Add "Printer 3", "Printing"

colKeys = objDictionary.Keys

Wscript.Echo "First run: " 
For Each strKey in colKeys
    Wscript.Echo strKey
Next
 
objDictionary.Remove("Printer 2")
colKeys = objDictionary.Keys

Wscript.Echo VbCrLf & "Second run: " 
For Each strKey in colKeys
    Wscript.Echo strKey
Next

