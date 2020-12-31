Set objNetwork = CreateObject("Wscript.Network")
strComputer = objNetwork.ComputerName
strUser = objNetwork.UserName

Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators")
For Each objUser in objGroup.Members
    If objUser.Name = strUser Then
        Wscript.Echo strUser & " is a local administrator."
    End If
Next
  


