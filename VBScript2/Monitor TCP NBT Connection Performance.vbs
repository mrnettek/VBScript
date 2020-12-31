' Description: Uses cooked performance counters to monitor the rates at which bytes are sent and received over the NetBIOS over TCP/IP (NBT) connection between the local computer and a remote computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_TCP_NBTConnection").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Bytes Received Per Second: " & _
            objItem.BytesReceivedPersec
        Wscript.Echo "Bytes Sent Per Second: " & objItem.BytesSentPersec
        Wscript.Echo "Bytes Total Per Second: " & objItem.BytesTotalPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Name: " & objItem.Name
        objRefresher.Refresh
    Next
Next

