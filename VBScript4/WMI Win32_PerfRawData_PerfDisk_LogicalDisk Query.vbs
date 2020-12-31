On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfDisk_LogicalDisk",,48)
For Each objItem in colItems
    Wscript.Echo "AvgDiskBytesPerRead: " & objItem.AvgDiskBytesPerRead
    Wscript.Echo "AvgDiskBytesPerRead_Base: " & objItem.AvgDiskBytesPerRead_Base
    Wscript.Echo "AvgDiskBytesPerTransfer: " & objItem.AvgDiskBytesPerTransfer
    Wscript.Echo "AvgDiskBytesPerTransfer_Base: " & objItem.AvgDiskBytesPerTransfer_Base
    Wscript.Echo "AvgDiskBytesPerWrite: " & objItem.AvgDiskBytesPerWrite
    Wscript.Echo "AvgDiskBytesPerWrite_Base: " & objItem.AvgDiskBytesPerWrite_Base
    Wscript.Echo "AvgDiskQueueLength: " & objItem.AvgDiskQueueLength
    Wscript.Echo "AvgDiskReadQueueLength: " & objItem.AvgDiskReadQueueLength
    Wscript.Echo "AvgDisksecPerRead: " & objItem.AvgDisksecPerRead
    Wscript.Echo "AvgDisksecPerRead_Base: " & objItem.AvgDisksecPerRead_Base
    Wscript.Echo "AvgDisksecPerTransfer: " & objItem.AvgDisksecPerTransfer
    Wscript.Echo "AvgDisksecPerTransfer_Base: " & objItem.AvgDisksecPerTransfer_Base
    Wscript.Echo "AvgDisksecPerWrite: " & objItem.AvgDisksecPerWrite
    Wscript.Echo "AvgDisksecPerWrite_Base: " & objItem.AvgDisksecPerWrite_Base
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
    Wscript.Echo "FreeMegabytes: " & objItem.FreeMegabytes
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PercentDiskReadTime: " & objItem.PercentDiskReadTime
    Wscript.Echo "PercentDiskReadTime_Base: " & objItem.PercentDiskReadTime_Base
    Wscript.Echo "PercentDiskTime: " & objItem.PercentDiskTime
    Wscript.Echo "PercentDiskTime_Base: " & objItem.PercentDiskTime_Base
    Wscript.Echo "PercentDiskWriteTime: " & objItem.PercentDiskWriteTime
    Wscript.Echo "PercentDiskWriteTime_Base: " & objItem.PercentDiskWriteTime_Base
    Wscript.Echo "PercentFreeSpace: " & objItem.PercentFreeSpace
    Wscript.Echo "PercentFreeSpace_Base: " & objItem.PercentFreeSpace_Base
    Wscript.Echo "PercentIdleTime: " & objItem.PercentIdleTime
    Wscript.Echo "PercentIdleTime_Base: " & objItem.PercentIdleTime_Base
    Wscript.Echo "SplitIOPerSec: " & objItem.SplitIOPerSec
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
Next

