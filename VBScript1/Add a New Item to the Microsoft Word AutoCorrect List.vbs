Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set colEntries = objWord.AutoCorrect.Entries
colEntries.Add "Fabirkam", "Fabrikam"
  


