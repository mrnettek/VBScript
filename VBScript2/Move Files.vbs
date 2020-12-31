' Description: Moves all the Windows Media (.wma) files to the folder C:\Media Archive.


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService. _
    ExecQuery("Select * from CIM_DataFile where Extension = 'wma'")

For Each objFile in colFiles
    strCopy = "C:\Media Archive\" & objFile.FileName _
        & "." & objFile.Extension
    objFile.Copy(strCopy)
    objFile.Delete
Next

