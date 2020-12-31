' Description: Demonstration script that uses the FileSystemObject to return available disk space on a specific disk drive. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objDrive = objFSO.GetDrive("C:")
Wscript.Echo "Available space: " & objDrive.AvailableSpace

