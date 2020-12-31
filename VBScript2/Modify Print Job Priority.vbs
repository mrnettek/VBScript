' Description: Uses ADSI to change the priority of current print jobs based on the size of those print jobs.


Set objPrinter = GetObject _
    ("WinNT://atl-dc-02/ArtDepartmentPrinter, printqueue")

For Each objPrintJob in objPrinter.PrintJobs
    If objPrintJob.Size > 400000 Then
        objPrintJob.Put "Priority" , 2
        objPrintJob.SetInfo
    Else
        objPrintJob.Put "Priority" , 3
        objPrintJob.SetInfo
    End If
Next

