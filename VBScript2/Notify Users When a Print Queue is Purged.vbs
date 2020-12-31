' Description: Uses Msg.exe to send a network alert to any users who had documents in a print queue about to be purged. After sending the alerts, the script purges the print queue.


Set WshShell = Wscript.CreateObject("Wscript.Shell")
Set objDictionary = CreateObject("Scripting.Dictionary")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colInstalledPrintJobs =  objWMIService.ExecQuery _
    ("Select * from Win32_PrintJob")

For Each objPrintJob in colInstalledPrintJobs
    strPrinterName = Split(objPrintJob.Name,",",-1,1)
    If objDictionary.Exists(objPrintJob.Notify) Then
    Else
        objDictionary.Add objPrintJob.Notify, strPrinterName(0)
    End If
Next

arrKeys = objDictionary.Keys
arrItems = objDictionary.Items

For i = 0 to objDictionary.Count - 1
    Message = "The documents you were printing on printer "
    Message = Message & arrItems(i)
    Message = Message & " had to be deleted from the print queue. "
    Message = Message & "You will need to reprint these documents."
    CommandString = "%comspec% /c msg " & arrKeys(i) & " " & Chr(34) _
        & Message & Chr(34)
    WshShell.Run CommandString, 0, True
Next

Set colInstalledPrinters = objWMIService.ExecQuery _
    ("Select * from Win32_Printer")
For Each objPrinter in colInstalledPrinters
    objPrinter.CancelAllJobs()
Next

