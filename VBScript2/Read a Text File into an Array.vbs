' Description: Demonstration script that uses the VBScript Split command to read a line from a commas-separated values file, and then place the individual items in that line into an array.


Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("c:\scripts\servers and services.txt", ForReading)

Do Until objTextFile.AtEndOfStream
    strNextLine = objTextFile.Readline
    arrServiceList = Split(strNextLine , ",")
    Wscript.Echo "Server name: " & arrServiceList(0)
    For i = 1 to Ubound(arrServiceList)
        Wscript.Echo "Service: " & arrServiceList(i)
    Next
Loop

