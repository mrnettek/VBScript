' Description: Demonstration script that opens a hypothetical text file consisting of server names, then retrieves service information from each on the servers in the file.


Const ForReading = 1

Set objDictionary = CreateObject("Scripting.Dictionary")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("c:\scripts\servers.txt", ForReading)
i = 0

Do Until objTextFile.AtEndOfStream 
    strNextLine = objTextFile.Readline
    objDictionary.Add i, strNextLine
    i = i + 1
Loop

For Each objItem in objDictionary
    Set colServices = GetObject("winmgmts://" & _
        objDictionary.Item(objItem) _
            & "").ExecQuery("Select * from Win32_Service")
    Wscript.Echo colServices.Count
Next

