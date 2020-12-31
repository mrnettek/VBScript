' Description: Uses a WMI Associators of query to return the local path of all the network shares on a computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colShares = objWMIService.ExecQuery _
    ("Associators of {Win32_Share.Name='Scripts'} Where " _
        & "AssocClass=Win32_ShareToDirectory")

For Each objFolder in colShares
    Wscript.Echo objFolder.Name
Next

