' Description: Demonstration script that uses the FileSystemObject to read a text file character-by-character, and individually echo those characters to the screen. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.OpenTextFile("C:\FSO\New Text Document.txt", 1)
Do Until objFile.AtEndOfStream
    strCharacters = objFile.Read(1)
    Wscript.Echo strCharacters
Loop

