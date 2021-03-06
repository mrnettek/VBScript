' Description: Demonstration script that verifies the existence of a particular key within a Script Runtime Dictionary. Script must be run on the local computer.


Set objDictionary = CreateObject("Scripting.Dictionary")

objDictionary.Add "Printer 1", "Printing"   
objDictionary.Add "Printer 2", "Offline"
objDictionary.Add "Printer 3", "Printing"

If objDictionary.Exists("Printer 4") Then
    Wscript.Echo "Printer 4 is in the Dictionary."
Else
    Wscript.Echo "Printer 4 is not in the Dictionary."
End If

