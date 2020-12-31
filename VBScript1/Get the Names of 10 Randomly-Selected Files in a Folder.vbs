strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colFiles = objWMIService.ExecQuery _
    ("ASSOCIATORS OF {Win32_Directory.Name='C:\Scripts'} Where " _
        & "ResultClass = CIM_DataFile")

intFiles = colFiles.Count

Set objDictionary = CreateObject("Scripting.Dictionary")

intHighNumber = intFiles
intLowNumber = 1

For i = 1 to 10
    x = 0
    Do Until x = 1
        Randomize
        intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
        If objDictionary.Exists(intNumber) Then
            x = 0
        Else
            objDictionary.Add intNumber, intNumber
            x = 1
        End If
    Loop
Next

i = 1

For Each objFile in colFiles
    If objDictionary.Exists(i) Then
        Wscript.Echo objFile.Name
    End If
    i = i + 1
Next
  


