On Error Resume Next
Dim iDesktopStyle : iDesktopStyle = 99

Call GetArguments(ArgArray)


If IsArray(ArgArray) then

	For Each ArrayElement In ArgArray
		iDesktopStyle = ArrayElement
	Next

End if


	Select Case iDesktopStyle

		Case 0,1,2,3
			Call Main()

		Case Else
			WScript.Echo "Possible Arguments for the Script are 0,1,2 or 3" & VBCrLF & VBCrLF & "Call the Script like:" & VBCrLF & Chr(34) & "Set Desktop Style via Excel Makro.vbs 2" & Chr(34) & VBCrLF & VBCrLF & "Arguments Detail:" &_
			VBCrLF & "0 = Style Normal Icon" & VBCrLF & "1 = Style Report" & VBCrLF & "2 = Style Small Icon" & VBCrLF & "3 = Style List"
	End Select

' --------
Sub Main()

Dim ExcelApp 		: Set ExcelApp		= CreateObject("Excel.Application")
ExcelApp.Visible 	= False
Dim ExcelAppWBk 	: Set ExcelAppWBk 	= ExcelApp.Workbooks.Add
Dim ExcelAppMod 	: Set ExcelAppMod 	= ExcelAppWBk.VBProject.VBComponents.Add(1)
Dim ofso 		: Set ofso 		= CreateObject("Scripting.FileSystemObject")
Dim oWshShell 		: Set oWshShell 	= CreateObject("WScript.Shell")

ExcelAppMod.CodeModule.AddFromString 	"Private Declare Function SendMessage Lib " & Chr(34) & "user32" & Chr(34) & " Alias " & Chr(34) & "SendMessageA" & Chr(34) & " (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long" & _
					VBCrLf & "Private Declare Function GetWindowLong Lib " & Chr(34) & "user32" & Chr(34) & " Alias " & Chr(34) & "GetWindowLongA" & Chr(34) & " ( ByVal hWnd As Long, ByVal nIndex As Long) As Long" & _
					VBCrLf & "Private Declare Function FindWindow Lib " & Chr(34) & "user32" & Chr(34) & " Alias " & Chr(34) & "FindWindowA" & Chr(34) & " ( ByVal lpClassName As String, ByVal lpWindowName As String) As Long" & _
					VBCrLf & "Private Declare Function GetWindow Lib " & Chr(34) & "user32" & Chr(34) & " ( ByVal hWnd As Long, ByVal wCmd As Long) As Long" & _
					VBCrLf & "Public Const VM_ICON = " & Chr(38) & "H0" &_
					VBCrLf & "Public Const VM_REPORT = " & Chr(38) & "H1" &_
					VBCrLf & "Public Const VM_SMALLICON = " & Chr(38) & "H2" &_
					VBCrLf & "Public Const VM_LIST = " & Chr(38) & "H3" &_
					VBCrLf & "Private Const GW_CHILD = 5" &_
					VBCrLf & "Private Const GWL_STYLE = (-16)" &_
					VBCrLf & "Private Const LVS_TYPEMASK = " & Chr(38) & "H3" &_
					VBCrLf & "Private Const WM_STYLECHANGED = " & Chr(38) & "H7D" &_
					VBCrLf & "Private Type StyleBits" &_
					VBCrLf & "dwOld As Long" &_
					VBCrLf & "dwNew As Long" &_
					VBCrLf & "End Type" &_
					VBCrLf & "Public Function SetDesktopStyle2(ByVal Flag As Long)" &_
					VBCrLf & "Dim hWnd As Long" &_
					VBCrLf & "Dim sb As StyleBits" &_
					VBCrLf & "hWnd = FindWindow(" & Chr(34) & "Progman" & Chr(34) & ", " & Chr(34) & "Program Manager" & Chr(34) & ")" &_
					VBCrLf & "hWnd = GetWindow(hWnd, GW_CHILD)" &_
					VBCrLf & "hWnd = GetWindow(hWnd, GW_CHILD)" &_
					VBCrLf & "With sb" &_
					VBCrLf & ".dwOld = GetWindowLong(hWnd, GWL_STYLE)" &_
					VBCrLf & ".dwNew = .dwOld" &_
					VBCrLf & ".dwNew = .dwNew And Not LVS_TYPEMASK" &_
					VBCrLf & ".dwNew = .dwNew Or Flag" &_
					VBCrLf & "End With" &_
					VBCrLf & " SendMessage hWnd, WM_STYLECHANGED, GWL_STYLE, sb" &_
					VBCrLf & "End Function" &_
					VBCrLf & "Public Function SetDesktopStyle(iStyle)" &_
					VBCrLf & "Select Case iStyle" &_
					VBCrLf & "Case 0" &_
					VBCrLf & "Call SetDesktopStyle2(VM_ICON)" &_
					VBCrLf & "Case 1" &_
					VBCrLf & "Call SetDesktopStyle2(VM_REPORT)" &_
					VBCrLf & "Case 2" &_
					VBCrLf & "Call SetDesktopStyle2(VM_SMALLICON)" &_
					VBCrLf & "Case 3" &_
					VBCrLf & "Call SetDesktopStyle2(VM_LIST)" &_
					VBCrLf & "End Select" &_
					VBCrLf & "End Function"

ExcelApp.Run "SetDesktopStyle(" & iDesktopStyle & ")"
ExcelAppWBk.Close False
ExcelApp.Quit
WScript.Quit

End Sub

' ----------------------------------------
Private Function GetArguments(SourceArray)

Dim iCount : iCount = 0

	If wscript.arguments.count > 0 then

		ReDim ArgArray(wscript.arguments.count -1)

		For Each Argument in wscript.arguments

			ArgArray(iCount) = Argument
			iCount = iCount +1
		Next


	iCount = Null
	GetArguments = ArgArray
		

	End if

End Function
