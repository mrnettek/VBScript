Const RECYCLE_BIN = &Ha&
Const FILE_SIZE = 3

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(RECYCLE_BIN)

Set colItems = objFolder.Items

For Each objItem in colItems
    strSize = objFolder.GetDetailsOf(objItem, FILE_SIZE)
    arrSize = Split(strSize, " ")
    intSize = intSize + CLng(arrSize(0))
Next

Wscript.Echo intSize & " KB"
  


