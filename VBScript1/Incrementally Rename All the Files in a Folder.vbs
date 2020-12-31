strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Data'} Where " _
        & "ResultClass = CIM_DataFile")

i = 1

For Each objFile in colFiles
    strNewName = "C:\Data\ABC_f" & i & ".inv"
    errResult = objFile.Rename(strNewName)
    i = i + 1
Next
  


