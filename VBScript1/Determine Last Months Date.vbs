intMonth = Month(Date)
intYear = Year(Date)

If intMonth = 1 Then
    intMonth = 12
    intYear = intYear - 1
Else
    intMonth = intMonth - 1
End If

If intMonth < 10 Then
    intMonth = "0" & intMonth
End If

intYear = Right(intYear, 2)

strName = intMonth & intYear

Wscript.Echo strName
  


