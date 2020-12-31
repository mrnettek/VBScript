Dim App : Set App = new cApp

If App.PrevInstance() = True then
	msgbox "There is more then one Instance of this Script"
Else
	msgbox "No other Instance running!"
End if


Class cApp

	Private Sub Class_Initialize()

	End Sub
          
	Private Sub Class_Terminate()

	End Sub

	' ----------------------------
	Public Function PrevInstance()

		Dim strComputer		: strComputer		= "."
		Dim oWMIService 	: Set oWMIService 	= GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
		Dim colItems 		: Set colItems 		= oWMIService.ExecQuery("Select * From Win32_Process")
		Dim sScriptCommandLine	: sScriptCommandLine	= wscript.ScriptFullName

		Dim iCount : iCount = 0

		For Each oItem in colItems
			If InStr(oItem.CommandLine, sScriptCommandLine) > 0 and InStr(UCase(oItem.CommandLine), "CSCRIPT.EXE") > 0 or InStr(UCase(oItem.CommandLine), "WSCRIPT.EXE") > 0 Then
				iCount = iCount +1
			End If
		Next

		If iCount > 1 then
			PrevInstance = True
		Else
			PrevInstance = False
		End if

	End Function

End Class
