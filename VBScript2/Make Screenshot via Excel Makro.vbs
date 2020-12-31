'On Error Resume Next

Dim ExcelApp 		: Set ExcelApp		= CreateObject("Excel.Application")
ExcelApp.Visible 	= False
Dim ExcelAppWBk 	: Set ExcelAppWBk 	= ExcelApp.Workbooks.Add
Dim ExcelAppMod 	: Set ExcelAppMod 	= ExcelAppWBk.VBProject.VBComponents.Add(1)
Dim oWshShell 		: Set oWshShell 	= CreateObject("WScript.Shell")

ExcelAppMod.CodeModule.AddFromString 	"Private Declare Sub keybd_event Lib " & Chr(34) & "User32" & Chr(34) & " (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)" & _
					VBCrLf & "Public Function GetScreenshot" &_
					VBCrLf & "keybd_event " & Chr(38) & "H2C, 1, 0, 0" &_
					VBCrLf & "DoEvents" &_
					VBCrLf & "End Function"

ExcelApp.Run "GetScreenshot()"
ExcelAppWBk.Close False
ExcelApp.Quit

oWshShell.Run "mspaint",1,False
wscript.sleep 1000
oWshShell.Sendkeys "^(v)"

