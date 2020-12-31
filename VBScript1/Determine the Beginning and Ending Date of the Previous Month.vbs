dtmTargetDate = Date
intDay = Day(dtmTargetDate)

dtmLastDay = dtmTargetDate - intDay
Wscript.Echo "Last day of previous month: " & dtmLastDay

dtmFirstDay = Month(dtmLastDay) & "/1/" & Year(dtmLastDay)
Wscript.Echo "First day of previous month: " & dtmFirstDay
  


