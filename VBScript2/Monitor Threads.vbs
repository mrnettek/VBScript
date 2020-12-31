' Description: Returns a list of threads and thread states for each process running on a computer.


Set objDictionary = CreateObject("Scripting.Dictionary")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")

For each objProcess in colProcesses 
    objDictionary.Add objProcess.ProcessID, objProcess.Name 
Next

Set colThreads = objWMIService.ExecQuery _
    ("Select * from Win32_Thread")
For each objThread in colThreads
    intProcessID = CInt(objThread.ProcessHandle)
    strProcessName = objDictionary.Item(intProcessID) 
    Wscript.Echo strProcessName & VbTab & objThread.ProcessHandle & _
        VbTab & objThread.Handle & VbTab & objThread.ThreadState 
Next

