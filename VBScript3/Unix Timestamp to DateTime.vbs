wscript.echo TimeStampToDateTime(342954300, false)
wscript.echo TimeStampToDateTime(342954300, true)

' -------------------------------
Private Function TimeStampToDateTime(iTimeStamp, bWithTimeZone)

	Select Case bWithTimeZone

		Case True
			TimeStampToDateTime = DateAdd("s", iTimeStamp + CurrentTimeZone(), CDate("01.01.1970"))
		Case False
			TimeStampToDateTime = DateAdd("s", iTimeStamp, CDate("01.01.1970"))

	End Select

End Function

' --------------------------------
Private Function CurrentTimeZone()

	Dim OS
	On Error Resume Next
	For Each OS in GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem")
      		CurrentTimeZone = OS.CurrentTimeZone * 60
 	Next

End Function

