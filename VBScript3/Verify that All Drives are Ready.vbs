' Description: Demonstration script that uses the FileSystemObject to ensure that all drives are ready (i.e., media is inserted) before echoing the drive letter and available disk space. Script must be run on the local computer.


Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives

For Each objDrive in colDrives
    If objDrive.IsReady = True Then
        Wscript.Echo "Drive letter: " & objDrive.DriveLetter
        Wscript.Echo "Free space: " & objDrive.FreeSpace
    Else
        Wscript.Echo "Drive letter: " & objDrive.DriveLetter
    End If
Next

