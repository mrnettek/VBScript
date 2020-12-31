On Error Resume Next

Dim sPointer, arrPointer, xPos, yPos, lClick


Call GetArguments(ArgArray)


If IsArray(ArgArray) then

	For Each ArrayElement In ArgArray
		sPointer = ArrayElement
	Next

End if

	If Instr(1, sPointer, ",", 1) > 0 then
		arrPointer = Split(sPointer, ",", -1, 1)

		If Ubound(arrPointer) = 2 then

			If IsNumeric(arrPointer(0)) and IsNumeric(arrPointer(1)) and IsNumeric(arrPointer(2)) then


				If arrPointer(2) = "0" or arrPointer(2) = "1" then

					xPos = CLng(arrPointer(0))
					yPos = CLng(arrPointer(1))
					lClick = arrPointer(2)
					Call Main()

				Else
					Wscript.Echo "Wrong Parameters!" & VbCrLf & VbCrLf & "Call the Script like this" & VbCrLf & VbCrLf & "Set Mouse Cursor Position.vbs xPos,yPos,LeftClick (0=False,1=True)" & VbCrLf & VbCrLf & "Example:" & VbCrLf & "Set Mouse Cursor Position.vbs 300,300,0"

				End if				

			Else
				Wscript.Echo "Wrong Parameters!" & VbCrLf & VbCrLf & "Call the Script like this" & VbCrLf & VbCrLf & "Set Mouse Cursor Position.vbs xPos,yPos,LeftClick (0=False,1=True)" & VbCrLf & VbCrLf & "Example:" & VbCrLf & "Set Mouse Cursor Position.vbs 300,300,0"

			End if

		Else
			Wscript.Echo "Wrong Parameters!" & VbCrLf & VbCrLf & "Call the Script like this" & VbCrLf & VbCrLf & "Set Mouse Cursor Position.vbs xPos,yPos,LeftClick (0=False,1=True)" & VbCrLf & VbCrLf & "Example:" & VbCrLf & "Set Mouse Cursor Position.vbs 300,300,0"

		End if

	Else

		Wscript.Echo "Wrong Parameters!" & VbCrLf & VbCrLf & "Call the Script like this" & VbCrLf & VbCrLf & "Set Mouse Cursor Position.vbs xPos,yPos,LeftClick (0=False,1=True)" & VbCrLf & VbCrLf & "Example:" & VbCrLf & "Set Mouse Cursor Position.vbs 300,300,0"

	End if



' --------
Sub Main()

Dim ExcelApp 		: Set ExcelApp		= CreateObject("Excel.Application")
ExcelApp.Visible 	= False
Dim ExcelAppWBk 	: Set ExcelAppWBk 	= ExcelApp.Workbooks.Add
Dim ExcelAppMod 	: Set ExcelAppMod 	= ExcelAppWBk.VBProject.VBComponents.Add(1)
Dim ofso 		: Set ofso 		= CreateObject("Scripting.FileSystemObject")
Dim oWshShell 		: Set oWshShell 	= CreateObject("WScript.Shell")

ExcelAppMod.CodeModule.AddFromString 	"Private Declare Function SetCursorPos Lib " & Chr(34) & "user32" & Chr(34) & " (ByVal x As Long, ByVal y As Long) As Long" & _
					VBCrLf & "Private Declare Sub mouse_event Lib " & Chr(34) & "user32" & Chr(34) & " (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)" & _
					VBCrLf & "Private Const MOUSEEVENTF_LEFTDOWN =" & Chr(38) & "H2" &_
					VBCrLf & "Private Const MOUSEEVENTF_LEFTUP =" & Chr(38) & "H4" &_
					VBCrLf & "Public Function SetCursorPosition(xPos as Long, yPos as Long, lClick)" &_
					VBCrLf & "SetCursorPos xPos, yPos" &_
					VBCrLf & "If lClick = 1 then" &_
					VBCrLf & "mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0" &_
					VBCrLf & "mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0" &_
					VBCrLf & "End if" &_
					VBCrLf & "End Function"

ExcelApp.Run "SetCursorPosition(" & xPos & "," & yPos & "," & lClick & ")"
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

