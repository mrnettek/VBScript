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
'arrList.Reverse()		-> Array R�ckw�rts neu anordnen
'arrList.Remove("CD")		-> entfernen eines Eintrags �ber desen Inhalt
'arrList.RemoveAt(0)		-> entfernen eines Eintrags �ber den Index
'arrList.Clear()		-> entfernt alle Eintr�ge aus der Liste



