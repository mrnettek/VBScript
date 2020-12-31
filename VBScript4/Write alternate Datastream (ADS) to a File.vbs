Dim ofso	: Set ofso 		= Createobject("Scripting.FileSystemObject")

Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim oFile
Dim myADStream	: myADStream = ""


'---------- Create a new Textfile
Set oFile = ofso.CreateTextFile("testfile.log", True)
oFile.Close
Set oFile = Nothing

'---------- Write ADStream to the new Textfile
Set oFile = ofso.OpenTextFile("testfile.log:myADS", ForWriting, True)
oFile.Writeline "Testing Alternate Datastream"
oFile.Close
Set oFile = Nothing

'---------- Read ADStream from the Textfile
Set oFile = ofso.OpenTextFile("testfile.log:myADS", ForReading, True)
myADStream = oFile.Readall()
oFile.Close
Set oFile = Nothing


msgbox myADStream




