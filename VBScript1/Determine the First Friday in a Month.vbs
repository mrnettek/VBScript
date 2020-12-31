dtmDate = #11/1/2006#

Do Until x = 1
    intDayOfWeek = Weekday(dtmDate)
    If intDayOfWeek = 6 Then
        Wscript.Echo "The first Friday of the month is " & dtmDate & "."
        Exit Do
    Else
        dtmDate = dtmDate + 1
    End If
Loop
  


