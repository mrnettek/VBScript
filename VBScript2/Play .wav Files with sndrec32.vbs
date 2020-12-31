
Dim oShell : Set oShell = CreateObject("WScript.Shell")

sWavFile 	= "C:\WINDOWS\Media\Windows XP Startup.wav"
sCmd 		= "sndrec32 /play /close " & sWavFile

oShell.Run sCmd, 0, False

