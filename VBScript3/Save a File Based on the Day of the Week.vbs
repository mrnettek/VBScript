strWeekDay = Weekday(Date)
strDay = WeekdayName(strWeekDay)

strPath = "C:\Logs\" & strDay & "\Results.txt"

Set objFSo = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(strPath)

objFile.WriteLine "This is a test."
objFile.Close
  


