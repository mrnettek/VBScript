strPath = "\\atl-fs-01\public"

strPath = Replace(strPath, "\\", "")

arrPath = Split(strPath, "\")

strComputer = arrPath(0)
strShare = arrPath(1)

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_Share Where Name = '" & strShare & "'")

For Each objItem in colItems
    Wscript.Echo objItem.Path
Next
  


