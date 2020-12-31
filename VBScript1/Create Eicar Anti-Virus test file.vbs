' Create -> THE ANTI-VIRUS OR ANTI-MALWARE TEST FILE
' before you use this Script read the TEST-FILE Information on http://www.eicar.org/anti_virus_test_file.htm
' http://www.eicar.org

' This sample script is not supported by Boris Toll, ScriptBox or Microsoft under any support program or service. The sample script is provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Boris Toll, ScriptBox, Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages. Microsoft Corporation in no way endorses or is affiliated with ScriptBox.


Dim ofso : Set ofso = Createobject("Scripting.FileSystemObject")
Const ForWriting = 2

Dim strEicarTest : strEicarTest = Chr(88) & Chr(53) & Chr(79) & Chr(33) & Chr(80) & Chr(37) & Chr(64) & Chr(65) & Chr(80) & Chr(91) & Chr(52) & Chr(92) & Chr(80) & Chr(90) & Chr(88) & Chr(53) & Chr(52) & Chr(40) & Chr(80) & Chr(94) & Chr(41) & Chr(55) & Chr(67) & Chr(67) & Chr(41) & Chr(55) & Chr(125) & Chr(36) & Chr(69) & Chr(73) & Chr(67) & Chr(65) & Chr(82) & Chr(45) & Chr(83) & Chr(84) & Chr(65) & Chr(78) & Chr(68) & Chr(65) & Chr(82) & Chr(68) & Chr(45) & Chr(65) & Chr(78) & Chr(84) & Chr(73) & Chr(86) & Chr(73) & Chr(82) & Chr(85) & Chr(83) & Chr(45) & Chr(84) & Chr(69) & Chr(83) & Chr(84) & Chr(45) & Chr(70) & Chr(73) & Chr(76) & Chr(69) & Chr(33) & Chr(36) & Chr(72) & Chr(43) & Chr(72) & Chr(42)


Dim oFile : Set oFile = ofso.OpenTextFile("TEST FILE www.eicar.org.txt",ForWriting,True)
oFile.WriteLine strEicarTest
oFile.Close

