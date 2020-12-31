On Error Resume Next

Dim ExcelApp 		: Set ExcelApp		= CreateObject("Excel.Application")
ExcelApp.Visible 	= False
Dim ExcelAppWBk 	: Set ExcelAppWBk 	= ExcelApp.Workbooks.Add
Dim ExcelAppMod 	: Set ExcelAppMod 	= ExcelAppWBk.VBProject.VBComponents.Add(1)
Dim ofso 		: Set ofso 		= CreateObject("Scripting.FileSystemObject")
Dim oWshShell 		: Set oWshShell 	= CreateObject("WScript.Shell")

ExcelAppMod.CodeModule.AddFromString 	"Private Declare Function SendMessage Lib " & Chr(34) & "user32" & Chr(34) & " Alias " & Chr(34) & "SendMessageA" & Chr(34) & " (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long" & _
					VBCrLf & "Public Const HWND_BROADCAST = " & Chr(38) & "HFFFF" & Chr(38) &_
					VBCrLf & "Public Const SC_MONITORPOWER = " & Chr(38) & "HF170" & Chr(38) &_
					VBCrLf & "Public Const MONITOR_OFF = 2" & Chr(38) &_
					VBCrLf & "Public Const WM_SYSCOMMAND = " & Chr(38) & "H112" &_
					VBCrLf & "Public Function MonitorOff()" & _
					VBCrLf & "SendMessage HWND_BROADCAST, WM_SYSCOMMAND, SC_MONITORPOWER, MONITOR_OFF" & _
					VBCrLf & "End Function"


Dim oFile : Set oFile = ofso.OpenTextFile("killexcel.vbs",2,True)

	oFile.WriteLine "wscript.sleep 4000"
	oFile.WriteLine "Call Taskkill(" & Chr(34) & "excel.exe" & Chr(34) & ")"
	oFile.WriteLine "Private Function Taskkill(strProcessName)"
	oFile.WriteLine "Dim oWMIService : Set oWMIService = GetObject(" & Chr(34) & "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2" & Chr(34) & ")"
	oFile.WriteLine "Dim colProcessList : Set colProcessList = oWMIService.ExecQuery(" & Chr(34) & "Select * from Win32_Process Where Name = '" & Chr(34) & " & strProcessName & " & Chr(34) & "'" & Chr(34) & ")"
	oFile.WriteLine "For Each oProcess in colProcessList"
	oFile.WriteLine "oProcess.Terminate()"
	oFile.WriteLine "Next"
	oFile.WriteLine "End Function"

	oFile.close

	oWshShell.Run "killexcel.vbs",0,False
	wscript.sleep 1500
	ofso.DeleteFile "killexcel.vbs",True

ExcelApp.Run "MonitorOff"
