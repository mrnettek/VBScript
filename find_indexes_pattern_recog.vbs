' MrNetTek
' eddiejackson.net/blog
' 10/14/2019
' free for public use 
' free to claim as your own

Option Explicit

Dim txtSource, thePattern, objMatchPattern, Match, Matches, count

txtSource = "The man in the moon is very far away"
thePattern = "\w+" ' you can change this part
count = 0

'Create the regular expression.
Set objMatchPattern = New RegExp
objMatchPattern.Pattern = thePattern
objMatchPattern.IgnoreCase = False
objMatchPattern.Global = True

'Perform the search.
Set Matches = objMatchPattern.Execute(txtSource)

'Iterate through the Matches collection.
For Each Match in Matches
   msgbox Match.FirstIndex & " " & Match.Value
   count = count + 1
Next

msgbox "total count: " & count


'clear session
set txtSource = Nothing
set thePattern = Nothing
set objMatchPattern = Nothing
set Match = Nothing
set Matches = Nothing

'The Index and Word Output:
' 0 The
' 4 man
' 8 in
' 11 the
' 15 moon
' 20 is
' 23 very
' 28 far
' 32 away
' total count: 9