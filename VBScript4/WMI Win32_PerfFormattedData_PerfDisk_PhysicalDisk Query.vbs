On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfDisk_PhysicalDisk",,48)
For Each objItem in colItems
    Wscript.Echo "AvgDiskBytesPerRead: " & objItem.AvgDiskBytesPerRead
    Wscript.Echo "AvgDiskBytesPerTransfer: " & objItem.AvgDiskBytesPerTransfer
    Wscript.Echo "AvgDiskBytesPerWrite: " & objItem.AvgDiskBytesPerWrite
    Wscript.Echo "AvgDiskQueueLength: " & objItem.AvgDiskQueueLength
    Wscript.Echo "AvgDiskReadQueueLength: " & objItem.AvgDiskReadQueueLength
    Wscript.Echo "AvgDisksecPerRead: " & objItem.AvgDisksecPerRead
    Wscript.Echo "AvgDisksecPerTransfer: " & objItem.AvgDisksecPerTransfer
    Wscript.Echo "AvgDisksecPerWrite: " & objItem.AvgDisksecPerWrite
    Wscript.Echo "AvgDiskWriteQueueLength: " & objItem.AvgDiskWriteQueueLength
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CurrentDiskQueueLength: " & objItem.CurrentDiskQueueLength
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DiskBytesPersec: " & objItem.DiskBytesPersec
    Wscript.Echo "DiskReadBytesPersec: " & objItem.DiskReadBytesPersec
    Wscript.Echo "DiskReadsPersec: " & objItem.DiskReadsPersec
    Wscript.Echo "DiskTransfersPersec: " & objItem.DiskTransfersPersec
    Wscript.Echo "DiskWriteBytesPersec: " & objItem.DiskWriteBytesPersec
    Wscript.Echo "DiskWritesPersec: " & objItem.DiskWritesPersec
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PercentDiskReadTime: " & objItem.PercentDiskReadTime
    Wscript.Echo "PercentDiskTime: " & objItem.PercentDiskTime
    Wscript.Echo "PercentDiskWriteTime: " & objItem.PercentDiskWriteTime
    Wscript.Echo "PercentIdleTime: " & objItem.PercentIdleTime
    Wscript.Echo "SplitIOPerSec: " & objItem.SplitIOPerSec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

