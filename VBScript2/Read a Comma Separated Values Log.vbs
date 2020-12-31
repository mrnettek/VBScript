' Description: Extracts the information in the DHCP Server log to individual fields and records.


Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("C:\Windows\System32\DHCP\" _
    & "DhcpSrvLog-Mon.log", ForReading)

Do While objTextFile.AtEndOfStream <> True
    If inStr(objtextFile.Readline, ",") Then
        arrDHCPRecord = split(objTextFile.Readline, ",")
        Wscript.Echo "Event ID: " & arrDHCPRecord(0)
        Wscript.Echo "Date: " & arrDHCPRecord(1)
        Wscript.Echo "Time: " & arrDHCPRecord(2)
        Wscript.Echo "Description: " & arrDHCPRecord(3)
        Wscript.Echo "IP Address: " & arrDHCPRecord(4)
        Wscript.Echo "Host Name: " & arrDHCPRecord(5)
        Wscript.Echo "MAC Address: " & arrDHCPRecord(6)
    Else
        objTextFile.Skipline
    End If
    i = i + 1
Loop

