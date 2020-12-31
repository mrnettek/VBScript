Const ForWriting = 2

Set objNetwork = CreateObject("Wscript.Network")

strName = objNetwork.UserName
strDomain = objNetwork.UserDomain
strUser = strDomain & "\" & strName

strText = strUser & vbCrLf

strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colPrinters = objWMIService.ExecQuery _
    ("Select * From Win32_Printer Where Local = FALSE")

For Each objPrinter in colPrinters
    strText = strText & objPrinter.Name & vbCrLf
Next

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.CreateTextFile _
    ("C:\Scripts\Printers.txt", ForWriting, True)

objFile.Write strText

objFile.Close
  


