strText = "aaa&bbb&ccc&aamp;ddd"

strText = Replace(strText, "&amp", "@@@@")

arrText = Split(strText, "&")

For i = 0 to Ubound(arrText) - 1
    arrText(i) = Replace(arrText(i), "@@@@", "&amp")
Next

For Each strItem in arrText
    Wscript.Echo strItem 
Next
  


