On Error Resume Next

Dim ExcelApp 		: Set ExcelApp		= CreateObject("Excel.Application")
ExcelApp.Visible 	= False
Dim ExcelAppWBk 	: Set ExcelAppWBk 	= ExcelApp.Workbooks.Add
Dim ExcelAppMod 	: Set ExcelAppMod 	= ExcelAppWBk.VBProject.VBComponents.Add(1)
Dim ExcelAppFrm 	: Set ExcelAppFrm 	= ExcelAppWBk.VBProject.VBComponents.Add(3)
ExcelAppFrm.Name = "frmMain"

ExcelAppFrm.CodeModule.AddFromString 	"Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & _
					VBCrLf & "Cancel=True" &_
					VBCrLf & "End Sub"

ExcelAppMod.CodeModule.AddFromString 	"Sub ShowForm()" &_
					VBCrLf & "frmMain.Caption = " & Chr(34) & "Prevent Logoff Window" & Chr(34) &_
					VBCrLf & "frmMain.Width = 200" &_
					VBCrLf & "frmMain.Height = 60" &_
					VBCrLf & "Set LabelX = frmMain.Controls.Add(" & Chr(34) & "Forms.Label.1" & Chr(34) & ", " & Chr(34) & "LabelX" & Chr(34) & ")" &_
					VBCrLf & "With LabelX" &_
					VBCrLf & ".Left = 15" &_
					VBCrLf & ".Top = 10" &_
					VBCrLf & ".Width = 200" &_
					VBCrLf & ".Caption =" & Chr(34) & "to close this window kill the excel.exe process" & Chr(34) &_
					VBCrLf & "End With" &_
					VBCrLf & "frmMain.Show" &_
					VBCrLf & "End Sub"


ExcelApp.Run "ShowForm()"
ExcelAppWBk.Close False
ExcelApp.Quit

