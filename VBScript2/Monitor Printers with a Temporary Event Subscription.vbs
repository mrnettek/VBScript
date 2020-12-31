' Description: Uses a temporary event consumer to issues alerts any time a printer changes status.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService. _ 
    ExecNotificationQuery("Select * from __instancemodificationevent " _ 
        & "within 30 where TargetInstance isa 'Win32_Printer'")
i = 0

Do While i = 0
    Set objPrinter = colPrinters.NextEvent
    If objPrinter.TargetInstance.PrinterStatus <> _ 
        objPrinter.PreviousInstance.PrinterStatus Then
        Select Case objPrinter.TargetInstance.PrinterStatus
            Case 1 strCurrentState = "Other" 
            Case 2 strCurrentState = "Unknown" 
            Case 3 strCurrentState = "Idle" 
            Case 4 strCurrentState = "Printing" 
            Case 5 strCurrentState = "Warming Up" 
        End Select
        Select Case objPrinter.PreviousInstance.PrinterStatus
            Case 1 strPreviousState = "Other" 
            Case 2 strPreviousState = "Unknown" 
            Case 3 strPreviousState = "Idle" 
            Case 4 strPreviousState = "Printing" 
            Case 5 strPreviousState = "Warming Up" 
        End Select
        Wscript.Echo objPrinter.TargetInstance.Name _ 
            &  " is " & strCurrentState _
                & ". The printer previously was " & strPreviousState & "."
    End If
Loop

