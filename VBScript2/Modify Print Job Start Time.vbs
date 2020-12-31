' Description: Uses ADSI to change the start time for all print jobs larger than 400K to 2:00 AM.


Set objPrinter = GetObject("WinNT://atl-dc-02/ArtDepartmentPrinter,printqueue")

For Each objPrintQueue in objPrinter.PrintJobs
    If objPrintQueue.Size > 400000 Then
        objPrintQueue.Put "StartTime" , TimeValue("2:00:00 AM")
        objPrintQueue.SetInfo
    End If
Next

