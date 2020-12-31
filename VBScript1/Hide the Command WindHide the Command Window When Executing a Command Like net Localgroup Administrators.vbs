Set objShell = CreateObject("Wscript.Shell")
strCommand = "net localgroup Administrators" 

Set objExec = objShell.Exec(strCommand) 
   
Do Until objExec.Status
    Wscript.Sleep 250
Loop 

Wscript.Echo objExec.StdOut.ReadAll()
  


