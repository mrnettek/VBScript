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

Dim oIE			: Set oIE 		= CreateObject("InternetExplorer.Application")

Dim str_Test : str_Test = "my Text"

	oIE.Navigate("about:blank")
	oIE.document.parentwindow.clipboardData.SetData "text", str_Test
	oIE.Quit
	Msgbox "Text wurde in die Zwischenablage kopiert", 64
