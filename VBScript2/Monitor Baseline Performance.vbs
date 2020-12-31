' Description: Uses cooked performance counters and the SWbemRefresher object to monitor three performance counters on a computer, and then save that data to a text file.


Const ForAppending = 8

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
  & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.Swbemrefresher")
Set objMemory = objRefresher.AddEnum _
  (objWMIService, "Win32_PerfFormattedData_PerfOS_Memory").objectSet
Set objDiskSpace = objRefresher.AddEnum _
  (objWMIService, "Win32_PerfFormattedData_PerfDisk_LogicalDisk").objectSet
Set objQueueLength = objRefresher.AddEnum _
  (objWMIService, "Win32_PerfFormattedData_PerfNet_ServerWorkQueues").objectSet
objRefresher.Refresh

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile _
  ("c:\scripts\performance.csv", ForAppending, True)

For i = 1 to 10
  For Each intAvailableBytes in objMemory
      objLogFile.Write(intAvailableBytes.AvailableMBytes) & "," 
  Next
  For each intQueueLength in objDiskSpace
      objLogFile.Write(intQueueLength.CurrentDiskQueueLength) & "," 
  Next
  For each intServerQueueLength in objQueueLength
      objLogFile.Write(intServerQueueLength.QueueLength) & ","
  Next
  objLogFile.Write VbCrLf
  Wscript.Sleep 10000
  objRefresher.Refresh
Next
objLogFile.Close

