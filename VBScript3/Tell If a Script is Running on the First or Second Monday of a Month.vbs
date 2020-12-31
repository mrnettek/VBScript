Dim arrMondays(1)

dtmMonth = Month(Date)
dtmYear = Year(Date)
dtmDate = CDate(dtmMonth & "/1" & "/" & dtmYear)

Do Until i = 1
    intWeekDay = Weekday(dtmDate)
    If intWeekDay = 2 Then
        arrMondays(0) = dtmDate
        arrMondays(1) = dtmDate + 7
        Exit Do
    End If
    dtmDate = dtmDate + 1
Loop

For Each strMonday in arrMondays
    If Date = strMonday Then
        Wscript.Echo "Carry out the task."
    Else
        Wscript.Echo "Don't carry out the task."
    End If
Next
  


