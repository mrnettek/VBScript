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



Call GetArguments(ArgArray)


If IsArray(ArgArray) then

	For Each ArrayElement In ArgArray
		msgbox ArrayElement
	Next

End if


' ----------------------------------------
Private Function GetArguments(SourceArray)

Dim iCount : iCount = 0

	If wscript.arguments.count > 0 then

		ReDim ArgArray(wscript.arguments.count -1)

		For Each Argument in wscript.arguments

			ArgArray(iCount) = Argument
			iCount = iCount +1
		Next


	iCount = Null
	GetArguments = ArgArray
		

	End if

End Function
