dtmTargetDate = #12/19/2005#   

dtmDay = DatePart("d", dtmTargetDate)
dtmMonth = DatePart("m", dtmTargetDate)
dtmYear = DatePart("yyyy", dtmTargetDate)

dtmStartDate = dtmMonth & "/1/" & dtmYear
dtmStartDate = CDate(dtmStartDate)

intWeekday = Weekday(dtmStartDate)
intAddon = 8 - intWeekday

intWeek1 = intAddOn
intWeek2 = intWeek1 + 7
intWeek3 = intWeek2 + 7
intWeek4 = intWeek3 + 7
intWeek5 = intWeek4 + 7
intWeek6 = intWeek5 + 7

If dtmDay <= intWeek6 Then
    strWeek = "Week 6"
End If

If dtmDay <= intWeek5 Then
    strWeek = "Week 5"
End If

If dtmDay <= intWeek4 Then
    strWeek = "Week 4"
End If

If dtmDay <= intWeek3 Then
    strWeek = "Week 3"
End If

If dtmDay <= intWeek2 Then
    strWeek = "Week 2"
End If

If dtmDay <= intWeek1 Then
    strWeek = "Week 1"
End If

Wscript.Echo strWeek
  


