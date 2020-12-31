Set objNetwork = CreateObject("Wscript.Network")

Set colDrives = objNetwork.EnumNetworkDrives

For i = 0 to colDrives.Count-1 Step 2
    If colDrives.Item(i + 1) = "\\server1\share" Then
        strDriveLetter = colDrives.Item(i)
        objNetwork.RemoveNetworkDrive strDriveLetter
        objNetwork.MapNetworkDrive strDriveLetter, "\\server2\share"
    End If
Next
  


