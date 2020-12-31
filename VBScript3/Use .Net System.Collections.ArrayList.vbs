' .Net Required

Dim arrList	: Set arrList = CreateObject("System.Collections.ArrayList")

arrList.Add "Abc"
arrList.Add "cde"
arrList.Add "fgh"

arrList.Sort()
msgbox arrList.Count

arrList.RemoveAt(0)
msgbox arrList.Count

For Each strItem in arrList
	Wscript.Echo strItem
Next

'arrList.Sort()			-> Array sortieren
'arrList.Reverse()		-> Array Rückwärts neu anordnen
'arrList.Remove("CD")		-> entfernen eines Eintrags über desen Inhalt
'arrList.RemoveAt(0)		-> entfernen eines Eintrags über den Index
'arrList.Clear()		-> entfernt alle Einträge aus der Liste



