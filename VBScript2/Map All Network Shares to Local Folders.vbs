' Description: Uses a WMI Associators of query to return the local path of a network share named Scripts.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colShares = objWMIService.ExecQuery("Select * From Win32_Share")

For Each objShare in colShares
    Set colAssociations = objWMIService.ExecQuery _
        ("Associators of {Win32_Share.Name='" & objShare.Name & "'} " _
            & " Where AssocClass=Win32_ShareToDirectory")
    For Each objFolder in colAssociations
        Wscript.Echo objShare.Name & vbTab & objFolder.Name
    Next
Next

