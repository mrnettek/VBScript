' Description: Demonstration script that uses the FileSystemObject to return the properties of all the disk drives installed on a computer. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives

For Each objDrive in colDrives
    Wscript.Echo "Available space: " & objDrive.AvailableSpace
    Wscript.Echo "Drive letter: " & objDrive.DriveLetter
    Wscript.Echo "Drive type: " & objDrive.DriveType
    Wscript.Echo "File system: " & objDrive.FileSystem
    Wscript.Echo "Free space: " & objDrive.FreeSpace
    Wscript.Echo "Is ready: " & objDrive.IsReady
    Wscript.Echo "Path: " & objDrive.Path
    Wscript.Echo "Root folder: " & objDrive.RootFolder
    Wscript.Echo "Serial number: " & objDrive.SerialNumber
    Wscript.Echo "Share name: " & objDrive.ShareName
    Wscript.Echo "Total size: " & objDrive.TotalSize
    Wscript.Echo "Volume name: " & objDrive.VolumeName
Next

