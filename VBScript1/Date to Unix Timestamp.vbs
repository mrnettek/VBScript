'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'
'/|									  |\\\\\\\\'
'//|									   |\\\\\\\'
'///|									    |\\\\\\'
'////|			Version 	1.0.0				     |\\\\\'
'/////|			Author:		Boris TOll 			      |\\\\'
'//////|		Last Update:	31.01.2008			       |\\\'
'///////|								        |\\'
'////////|									 |\'
'#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#*~#'



msgbox GetUnixTimeStamp("21.07.2007")
msgbox GetUnixTimeStamp("21.07.2007 00:00:01")


' ---------------------------------------- [Wandelt einen Date/Time-Wert in eine UnixTimeStamp um]
Private Function GetUnixTimeStamp(strDate)

Const cUnix 	= #1/1/1970#
Dim nTime	: nTime = DateAdd("n", 0, strDate)

	GetUnixTimeStamp = DateDiff("s", cUnix, nTime)

End Function
