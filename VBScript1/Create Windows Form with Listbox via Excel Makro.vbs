On Error Resume Next

Dim ExcelApp 		: Set ExcelApp		= CreateObject("Excel.Application")
ExcelApp.Visible 	= False
Dim ExcelAppWBk 	: Set ExcelAppWBk 	= ExcelApp.Workbooks.Add
Dim ExcelAppMod 	: Set ExcelAppMod 	= ExcelAppWBk.VBProject.VBComponents.Add(1)
Dim ExcelAppCls 	: Set ExcelAppCls 	= ExcelAppWBk.VBProject.VBComponents.Add(2)
ExcelAppCls.Name = "ClassEvents"

ExcelAppCls.CodeModule.AddFromString 	"Public WithEvents CloseCommandButton As MSForms.CommandButton" &_
					VBCrLf & "Private Sub CloseCommandButton_Click()" &_
					VBCrLf & "Unload UserForm1" &_
					VBCrLf & "End Sub"

ExcelAppMod.CodeModule.AddFromString 	"Sub CreateForm()" &_
					VBCrLf & "Dim frm As Object" &_
					VBCrLf & "Set frm = ThisWorkbook.VBProject.VBComponents.Add(3)" &_
					VBCrLf & "frm.Name = " & Chr(34) & "UserForm1" & Chr(34) &_
					VBCrLf & "ShowForm" &_
					VBCrLf & "End Sub" &_
					VBCrLf & "Sub ShowForm()" &_
					VBCrLf & "Set BtnClose = UserForm1.Controls.Add(" & Chr(34) & "Forms.CommandButton.1" & Chr(34) & ", " & Chr(34) & "BtnClose" & Chr(34) & ")" &_
					VBCrLf & "With BtnClose" &_
					VBCrLf & ".Caption =" & Chr(34) & "Close" & Chr(34) &_
					VBCrLf & ".Left = 183" &_
					VBCrLf & ".Top = 200" &_
					VBCrLf & ".Height = 20" &_
					VBCrLf & ".Width = 40" &_
					VBCrLf & ".Font.Bold = True" &_
					VBCrLf & ".Font.Size = 8" &_
					VBCrLf & "End With" &_
					VBCrLf & "Set ListX = UserForm1.Controls.Add(" & Chr(34) & "Forms.ListBox.1" & Chr(34) & ", " & Chr(34) & "ListX" & Chr(34) & ")" &_
					VBCrLf & "With ListX" &_
					VBCrLf & ".Left = 22" &_
					VBCrLf & ".Top = 20" &_
					VBCrLf & ".Width = 200" &_
					VBCrLf & ".Height = 170" &_
					VBCrLf & "End With" &_
					VBCrLf & "sFile = Dir$(Environ$(" & Chr(34) & "windir" & Chr(34) & ") & " & Chr(34) & "\*.*" & Chr(34) & ")" &_
					VBCrLf & "Do While Len(sFile) <> 0" &_
					VBCrLf & "ListX.AddItem sFile, 0" &_
					VBCrLf & "sFile = Dir$" &_
					VBCrLf & "Loop" &_
					VBCrLf & "Set LabelX = UserForm1.Controls.Add(" & Chr(34) & "Forms.Label.1" & Chr(34) & ", " & Chr(34) & "LabelX" & Chr(34) & ")" &_
					VBCrLf & "With LabelX" &_
					VBCrLf & ".Left = 22" &_
					VBCrLf & ".Top = 10" &_
					VBCrLf & ".Width = 200" &_
					VBCrLf & ".Caption = " & Chr(34) & "Files in " & Chr(34) & " & Environ$(" & Chr(34) & "windir" & Chr(34) & ")" &_
					VBCrLf & "End With" &_
					VBCrLf & "Set CloseCommandButton = New ClassEvents" &_
					VBCrLf & "Set CloseCommandButton.CloseCommandButton = BtnClose" &_
					VBCrLf & "UserForm1.Caption = " & Chr(34) & "My Window" & Chr(34) &_
					VBCrLf & "UserForm1.Width = 250" &_
					VBCrLf & "UserForm1.Height = 250" &_
					VBCrLf & "UserForm1.Show" &_
					VBCrLf & "End Sub"

ExcelApp.Run "CreateForm"
ExcelAppWBk.Close False
ExcelApp.Quit

