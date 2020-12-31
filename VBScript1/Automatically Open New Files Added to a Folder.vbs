Set objShell = CreateObject("Wscript.Shell")

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("SELECT * FROM __InstanceCreationEvent WITHIN 10 WHERE " _
        & "Targetinstance ISA 'CIM_DirectoryContainsFile' and " _
            & "TargetInstance.GroupComponent= " _
                & "'Win32_Directory.Name=""c:\\\\scripts""'")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
    strNewFile = objLatestEvent.TargetInstance.PartComponent
    arrNewFile = Split(strNewFile, "=")
    strFileName = arrNewFile(1)
    strFileName = Replace(strFileName, "\\", "\")
    strFileName = Replace(strFileName, Chr(34), "")
    objShell.Run("notepad.exe " & strFileName)
Loop
  


