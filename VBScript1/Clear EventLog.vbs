' The Script clears the specified Eventlog
' Call the Script with the Eventlogname as Parameter. Like "Clear EventLog.vbs" Support

On Error Resume Next

Dim strArgs : strArgs = ""


Call GetArguments(ArgArray)


If IsArray(ArgArray) then

	For Each ArrayElement In ArgArray
		strArgs = strArgs & ArrayElement
	Next

	strArgs = Trim(strArgs)
	Call ClearEventLog(strArgs)

End if


' ----------------------------------
Private Function ClearEventLog(strEventLog)

On Error Resume Next

Dim strComputer : strComputer = "."
Dim oWMIService : Set oWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Dim colLogFiles : Set colLogFiles = oWMIService.ExecQuery("Select * from Win32_NTEventLogFile " & "Where LogFileName='Support'")

	For Each oLogfile in colLogFiles
		oLogFile.ClearEventLog()
	Next

	Set oWMIService = Nothing
	Set colLogFiles = Nothing

End Function


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
