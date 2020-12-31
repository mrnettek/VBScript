strSoundFile = "C:\Windows\Media\Notify.wav"
Set objShell = CreateObject("Wscript.Shell")
strCommand = "sndrec32 /play /close " & chr(34) & strSoundFile & chr(34)
objShell.Run strCommand, 0, True
  


