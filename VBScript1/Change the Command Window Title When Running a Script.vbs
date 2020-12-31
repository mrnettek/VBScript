Set objShell = CreateObject("Wscript.Shell")
objShell.Run _
    ("%comspec% /K title My Command Window |ping.exe 192.168.1.1"), _
        1, TRUE
  


