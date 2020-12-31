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

Option Explicit

Dim AppExcel 	: 	Set AppExcel 	= CreateObject("Excel.Application")
Dim strURL 	: 	strURL 		= "http://www.google.com"
Dim strTarget 	: 	strTarget 	= "C:\test.log"
Dim iReturn

iReturn = AppExcel.ExecuteExcel4Macro("CALL(""urlmon"",""URLDownloadToFileA"",""JJCCJJ"",0,""" & strURL & """,""" & strTarget & """,0,0)")

If iReturn <> 0 then
	msgbox "ErrorNumber: " & iReturn
End if

Set AppExcel 	= Nothing
