' Description: Returns a list of processes and all the services currently running in each process.


Set objIdDictionary = CreateObject("Scripting.Dictionary")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where State <> 'Stopped'")

For Each objService in colServices
    If objIdDictionary.Exists(objService.ProcessID) Then
    Else
        objIdDictionary.Add objService.ProcessID, objService.ProcessID
    End If
Next

colProcessIDs = objIdDictionary.Items

For i = 0 to objIdDictionary.Count - 1
    Set colServices = objWMIService.ExecQuery _
        ("Select * from Win32_Service Where ProcessID = '" & _
            colProcessIDs(i) & "'")
    Wscript.Echo "Process ID: " & colProcessIDs(i)
    For Each objService in colServices
        Wscript.Echo VbTab & objService.DisplayName 
    Next
Next

