' Description: Uses the Win32_PingStatus class to ping four computers and identify any computers that could not be reached over the network. This script was contributed by Maxim Stepin of Microsoft.


strMachines = "atl-dc-01;atl-win2k-01;atl-nt4-01;atl-dc-02"
aMachines = split(strMachines, ";")
 
For Each machine in aMachines
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        ExecQuery("select * from Win32_PingStatus where address = '"_
            & machine & "'")
    For Each objStatus in objPing
        If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then 
            WScript.Echo("Computer " & machine & " is not reachable") 
        End If
    Next
Next

