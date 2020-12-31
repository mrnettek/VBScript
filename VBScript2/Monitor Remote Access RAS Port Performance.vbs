' Description: Uses cooked performance counters to monitor individual Remote Access Service ports of the RAS device on the computer.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService," & _
    "Win32_PerfFormattedData_PerfProc_RemoteAccess_RASPort").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Alignment Errors: " & objItem.AlignmentErrors
        Wscript.Echo "Buffer Overrun Errors: " & objItem.BufferOverrunErrors
        Wscript.Echo "Bytes Received: " & objItem.BytesReceived
        Wscript.Echo "Bytes Received Per Second: " & _
            objItem.BytesReceivedPerSec
        Wscript.Echo "Bytes Transmitted: " & objItem.BytesTransmitted
        Wscript.Echo "Bytes Transmitted Per Second: " & _
            objItem.BytesTransmittedPerSec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "CRC Errors: " & objItem.CRCErrors
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Frames Received: " & objItem.FramesReceived
        Wscript.Echo "Frames Received Per Second: " & _
            objItem.FramesReceivedPerSec
        Wscript.Echo "Frames Transmitted: " & objItem.FramesTransmitted
        Wscript.Echo "Frames Transmitted Per Second: " & _
            objItem.FramesTransmittedPerSec
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Percent Compression In: " & objItem.PercentCompressionIn
        Wscript.Echo "Percent Compression Out: " & _
            objItem.PercentCompressionOut
        Wscript.Echo "Serial Overrun Errors: " & objItem.SerialOverrunErrors
        Wscript.Echo "Timeout Errors: " & objItem.TimeoutErrors
        Wscript.Echo "Total Errors: " & objItem.TotalErrors
        Wscript.Echo "Total Errors Per Second: " & objItem.TotalErrorsPerSec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

