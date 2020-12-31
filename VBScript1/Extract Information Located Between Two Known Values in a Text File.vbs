strSearchString = "ABCDEFGHIJK, NDPSGW PORT=LPR HOSTNAME=R2333_HP_1100 ABCDEFGHIJK"

intStart = InStr(strSearchString, "HOSTNAME=")
intStart = intStart + 9

strText = Mid(strSearchString, intStart, 250)

For i = 1 to Len(strText)
    If Mid(strText, i, 1) = " " Then
        Exit For
    Else
        strData = strData & Mid(strText, i, 1)
    End If
Next

Wscript.Echo strData
  


