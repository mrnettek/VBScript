Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Open("C:\Scripts\Test.doc")

Set objNetwork = CreateObject("Wscript.Network")
strUser = objNetwork.UserName
strDomain = objNetwork.UserDomain
strComputer = objNetwork.ComputerName

Set objRange = objDoc.Bookmarks("UserBookmark").Range
objRange.Text = strUser
objDoc.Bookmarks.Add "UserBookmark",objRange

Set objRange = objDoc.Bookmarks("DomainBookmark").Range
objRange.Text = strDomain
objDoc.Bookmarks.Add "DomainBookmark",objRange

Set objRange = objDoc.Bookmarks("ComputerBookmark").Range
objRange.Text = strComputer
objDoc.Bookmarks.Add "ComputerBookmark",objRange
  


