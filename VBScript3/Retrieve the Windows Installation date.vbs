'----- (c) Boris Toll 2007 -----


Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "." 

Set oReg	= GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
strKeyPath 	= "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName	= "InstallDate"


Call oReg.GetDWORDValue(HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)
MsgBox UnixTStampTo_Date(strValue)


'------------------------------------------ TimeStamp to Date
Function UnixTStampTo_Date(strValue)

    On Error Resume Next

    vStart 		= DateSerial(1970, 1, 1)
    UnixTStampTo_Date 	= DateAdd("s", strValue, vStart)
    Date_Array 		= Split(UnixTStampTo_Date, " ", -1, vbTextCompare)
    UnixTStampTo_Date 	= WeekdayName(Weekday(UnixTStampTo_Date, vbUseSystemDayOfWeek), False, vbUseSystemDayOfWeek) & ", " & Date_Array(0) & ", " & Date_Array(1)

End Function
