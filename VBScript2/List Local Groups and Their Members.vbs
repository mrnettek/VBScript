' Description: Returns a list of local groups (and their members) found on a computer named atl-win2k-01.


strComputer = "atl-win2k-01"
Set colGroups = GetObject("WinNT://" & strComputer & "")
colGroups.Filter = Array("group")

For Each objGroup In colGroups
    Wscript.Echo objGroup.Name 
    For Each objUser in objGroup.Members
        Wscript.Echo vbTab & objUser.Name
    Next
Next

