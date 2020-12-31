Set objDictionary = CreateObject("Scripting.Dictionary")

arrItems = Array("a","b","b","c","c","c","d","e","e","e")

For Each strItem in arrItems
    If Not objDictionary.Exists(strItem) Then
        objDictionary.Add strItem, strItem   
    End If
Next

intItems = objDictionary.Count - 1

ReDim arrItems(intItems)

i = 0

For Each strKey in objDictionary.Keys
    arrItems(i) = strKey
    i = i + 1
Next

For Each strItem in arrItems
    Wscript.Echo strItem
Next
  


