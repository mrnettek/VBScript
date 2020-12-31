Dim App : Set App = new cApp		'### Must be the first line in the Script

wscript.echo App.StartTime()
wscript.echo App.WorkDir()
wscript.echo App.AppDir()
wscript.echo App.ScriptName()
wscript.echo App.ProcessID()
wscript.echo App.ProcessorArchitecture()
wscript.echo App.ScriptHost()
If App.PrevInstance() = True then
	wscript.echo "There is more then one Instance of this Script"
Else
	wscript.echo "No other Instance of this Script running!"
End if


Class cApp


	Public Property Get WorkDir()
      		WorkDir = GetWorkPath()
	End Property

	Public Property Get AppDir()
      		AppDir = GetScriptPath()
	End Property

	Public Property Get PrevInstance()
      		PrevInstance = GetPrevInstance()
	End Property

	Public Property Get ScriptName()
      		ScriptName = wscript.ScriptName
	End Property

	Public Property Get ProcessID()
      		ProcessID = GetProcessID()
	End Property

	Public Property Get ScriptHost()
      		ScriptHost = GetScriptHost()
	End Property

	Public Property Get ProcessorArchitecture()
      		ProcessorArchitecture = GetEnvironment("PROCESSOR_ARCHITECTURE")
	End Property

	Private dStartTime
	Public Property Get StartTime()
		StartTime = dStartTime
	End Property

	Private Property Set StartTime(dTime)
		dStartTime = dTime
	End Property


	' ------------------------------
	Private Function GetScriptHost()

		Dim ofso		: Set ofso	 	= CreateObject("Scripting.FileSystemObject")
		GetScriptHost = ofso.GetFile(WScript.FullName).Name

	End Function

	' ------------------------------
	Private Function GetEnvironment(sEnvironment)

		Dim oWshShell		: Set oWshShell 	= WScript.CreateObject("WSCript.shell")
		GetEnvironment = oWshShell.ExpandEnvironmentStrings("%" & sEnvironment & "%")

	End Function

	' ------------------------------
	Private Function GetScriptPath()

		Dim ofso		: Set ofso	 	= CreateObject("Scripting.FileSystemObject")
		Dim sScriptPath 	: sScriptPath 		= ofso.GetFile(wscript.ScriptFullName).ParentFolder
		Set ofso = Nothing
		GetScriptPath = sScriptPath

	End Function

	' ----------------------------
	Private Function GetWorkPath()

		Dim oWshShell		: Set oWshShell 	= WScript.CreateObject("WSCript.shell")
		Dim sWorkPath 		: sWorkPath 		= oWshShell.CurrentDirectory
		Set oWshShell = Nothing
		GetWorkPath = sWorkPath

	End Function

	' --------------------------------
	Private Function GetPrevInstance()

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

		Set oWMIService = Nothing
		Set colItems = Nothing

		If iCount > 1 then
			GetPrevInstance = True
		Else
			GetPrevInstance = False
		End if

	End Function

	' -----------------------------
	Private Function GetProcessID()

		Dim strComputer		: strComputer		= "."
		Dim oWMIService 	: Set oWMIService 	= GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
		Dim colItems 		: Set colItems 		= oWMIService.ExecQuery("Select * From Win32_Process")
		Dim sScriptCommandLine	: sScriptCommandLine	= wscript.ScriptFullName
		Dim sScriptHostPath	: sScriptHostPath	= WScript.FullName
		Dim sProcessID		: sProcessID		= 0

		For Each oItem in colItems

			If InStr(oItem.CommandLine, sScriptCommandLine) > 0 and InStr(oItem.CommandLine, sScriptHostPath) > 0 Then
				
				dProcessStartTime = oItem.CreationDate
				dProcessStartTime = DateSerial(Left(dProcessStartTime, 4), Mid(dProcessStartTime,  5, 2), Mid(dProcessStartTime,  7, 2) ) + TimeSerial(Mid(dProcessStartTime,  9, 2), Mid(dProcessStartTime, 11, 2), Mid(dProcessStartTime, 13, 2))
				dStartDateDiff = DateDiff("s", dStartTime, dProcessStartTime)

				Select Case dStartDateDiff
					Case 0, -1, 1
						sProcessID = oItem.ProcessID
				End Select

			End If

		Next

		Set oWMIService = Nothing
		Set colItems = Nothing

		If not sProcessID = 0 then
			GetProcessID = sProcessID
		Else
			GetProcessID = -1
		End if

	End Function


	Private Sub Class_Initialize()
		Set StartTime = Now
	End Sub
          
	Private Sub Class_Terminate()

	End Sub


End Class
