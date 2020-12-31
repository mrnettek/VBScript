On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\WMI")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM TraceLogger", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AgeLimit: " & objItem.AgeLimit
      WScript.Echo "BufferSize: " & objItem.BufferSize
      WScript.Echo "BuffersWritten: " & objItem.BuffersWritten
      strEnableFlags = Join(objItem.EnableFlags, ",")
         WScript.Echo "EnableFlags: " & strEnableFlags
      WScript.Echo "EventsLost: " & objItem.EventsLost
      WScript.Echo "FlushTimer: " & objItem.FlushTimer
      WScript.Echo "FreeBuffers: " & objItem.FreeBuffers
      strGuid = Join(objItem.Guid, ",")
         WScript.Echo "Guid: " & strGuid
      strLevel = Join(objItem.Level, ",")
         WScript.Echo "Level: " & strLevel
      WScript.Echo "LogBuffersLost: " & objItem.LogBuffersLost
      WScript.Echo "LogFileMode: " & objItem.LogFileMode
      WScript.Echo "LogFileName: " & objItem.LogFileName
      WScript.Echo "LoggerId: " & objItem.LoggerId
      WScript.Echo "LoggerThreadId: " & objItem.LoggerThreadId
      WScript.Echo "MaximumBuffers: " & objItem.MaximumBuffers
      WScript.Echo "MaximumFileSize: " & objItem.MaximumFileSize
      WScript.Echo "MinimumBuffers: " & objItem.MinimumBuffers
      WScript.Echo "Name: " & objItem.Name
      WScript.Echo "NumberOfBuffers: " & objItem.NumberOfBuffers
      WScript.Echo "RealTimeBuffersLost: " & objItem.RealTimeBuffersLost
      WScript.Echo
   Next
Next

