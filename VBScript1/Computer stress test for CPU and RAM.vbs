'---------------------------
'Information (c) Boris Toll - 2008
'---------------------------
'Attention! Running this Script can produce unexpected errors on your PC.
'You run this Script at your own risk.

'The Author takes no liability for possible damages or errors by direct or indirect executing, modifying or distributing this Script.
'The Script uses a lot of physical memory and CPU capacity and it might be possible that you will have to hard reboot your PC to stop the Script.
'Normally the Script runs between one and five minutes and stops automatically.



On Error Resume Next

Dim ofso	: Set ofso 	= CreateObject("Scripting.FileSystemObject")
Dim oWshShell	: Set oWshShell = CreateObject("WScript.Shell")
Dim iRet

iRet = MsgBox("Attention! Running this Script can produce unexpected errors on your PC." &_
		VbCrLf & "You run this Script at your own risk." &_
		VbCrLf & VbCrLf & "The Author takes no liability for possible damages or errors by direct or indirect executing, modifying or distributing this Script." &_
		VbCrLf & "The Script uses a lot of physical memory and CPU capacity and it might be possible that you will have to hard reboot your PC to stop the Script." &_
		VbCrLf & "Normally the Script runs between one and five minutes and stops automatically.", 36, "Information")

Select Case iRet

	Case 6

	Case 7
		wscript.quit
	Case Else
		wscript.quit
End Select

If ofso.FileExists("exit.flg") then
	ofso.DeleteFile "exit.flg",True
End if


Dim oFile 	: Set oFile 	= ofso.OpenTextFile("proc.vbs",2,True)
oFile.WriteLine "On Error Resume Next"
oFile.WriteLine "Dim ofso : Set ofso = CreateObject(" & Chr(34) & "Scripting.FileSystemObject" & Chr(34) & ")"
oFile.WriteLine "wscript.sleep 15000"
oFile.WriteLine "Do"
oFile.WriteLine "If ofso.FileExists(" & Chr(34) & "exit.flg" & Chr(34) & ") then"
oFile.WriteLine "wscript.quit"
oFile.WriteLine "End if"
oFile.WriteLine "Loop"
oFile.Close

Set oFile 	= ofso.OpenTextFile("mem.vbs",2,True)
oFile.WriteLine "On Error Resume Next"
oFile.WriteLine "Dim ofso : Set ofso = CreateObject(" & Chr(34) & "Scripting.FileSystemObject" & Chr(34) & ")"
oFile.WriteLine "wscript.sleep 15000"
oFile.WriteLine "Do"
oFile.WriteLine "If ofso.FileExists(" & Chr(34) & "exit.flg" & Chr(34) & ") then"
oFile.WriteLine "wscript.quit"
oFile.WriteLine "End if"
oFile.WriteLine "testMem01 = Space(32000000)"
oFile.WriteLine "testMem02 = Space(32000000)"
oFile.WriteLine "testMem03 = Space(32000000)"
oFile.WriteLine "testMem04 = Space(32000000)"
oFile.WriteLine "testMem05 = Space(32000000)"
oFile.WriteLine "testMem06 = Space(32000000)"
oFile.WriteLine "testMem07 = Space(32000000)"
oFile.WriteLine "testMem08 = Space(32000000)"
oFile.WriteLine "testMem09 = Space(32000000)"
oFile.WriteLine "testMem10 = Space(32000000)"
oFile.WriteLine "testMem11 = Space(32000000)"
oFile.WriteLine "testMem12 = Space(32000000)"
oFile.WriteLine "testMem13 = Space(32000000)"
oFile.WriteLine "testMem14 = Space(32000000)"
oFile.WriteLine "testMem15 = Space(32000000)"
oFile.WriteLine "testMem16 = Space(32000000)"
oFile.WriteLine "testMem17 = Space(32000000)"
oFile.WriteLine "testMem18 = Space(32000000)"
oFile.WriteLine "testMem19 = Space(32000000)"
oFile.WriteLine "testMem20 = Space(32000000)"
oFile.WriteLine "testMem21 = Space(32000000)"
oFile.WriteLine "testMem22 = Space(32000000)"
oFile.WriteLine "testMem23 = Space(32000000)"
oFile.WriteLine "testMem24 = Space(32000000)"
oFile.WriteLine "testMem25 = Space(32000000)"
oFile.WriteLine "testMem26 = Space(32000000)"
oFile.WriteLine "testMem27 = Space(32000000)"
oFile.WriteLine "testMem28 = Space(32000000)"
oFile.WriteLine "testMem29 = Space(32000000)"
oFile.WriteLine "testMem30 = Space(32000000)"
oFile.WriteLine "testMem31 = Space(32000000)"
oFile.WriteLine "testMem32 = Space(32000000)"
oFile.WriteLine "testMem33 = Space(32000000)"
oFile.WriteLine "testMem34 = Space(32000000)"
oFile.WriteLine "testMem35 = Space(32000000)"
oFile.WriteLine "testMem36 = Space(32000000)"
oFile.WriteLine "testMem37 = Space(32000000)"
oFile.WriteLine "testMem38 = Space(32000000)"
oFile.WriteLine "testMem39 = Space(32000000)"
oFile.WriteLine "testMem40 = Space(32000000)"
oFile.WriteLine "testMem41 = Space(32000000)"
oFile.WriteLine "testMem42 = Space(32000000)"
oFile.WriteLine "testMem43 = Space(32000000)"
oFile.WriteLine "testMem44 = Space(32000000)"
oFile.WriteLine "testMem45 = Space(32000000)"
oFile.WriteLine "testMem46 = Space(32000000)"
oFile.WriteLine "testMem47 = Space(32000000)"
oFile.WriteLine "testMem48 = Space(32000000)"
oFile.WriteLine "testMem49 = Space(32000000)"
oFile.WriteLine "testMem50 = Space(32000000)"
oFile.WriteLine "Loop"
oFile.Close

For iCount = 0 to 50
	oWshShell.Run "mem.vbs",0,False
	oWshShell.Run "proc.vbs",0,False
Next

Set oWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcesses = oWMIService.ExecQuery("Select * from Win32_Process Where Name = 'wscript.exe'")
For Each objProcess in colProcesses
    objProcess.SetPriority(256)
Next
Set colProcesses = oWMIService.ExecQuery("Select * from Win32_Process Where Name = 'cscript.exe'")
For Each objProcess in colProcesses
    objProcess.SetPriority(256)
Next

wscript.sleep 50000
Set oFile 	= ofso.OpenTextFile("exit.flg",2,True)
oFile.Close

