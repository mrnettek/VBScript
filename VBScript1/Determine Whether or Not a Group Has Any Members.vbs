strComputer = "atl-ws-01"

Set objGroup = GetObject("WinNT://" & strComputer & "/Remote Desktop Users")
i = 0
For Each objUser in objGroup.Members
    i = i + 1
Next

Wscript.Echo "Number of users in group: " & i
  


