' Description: Lists accountant information for all the virtual machines on a computer.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
set colVMs = objVS.VirtualMachines

For Each objVM in colVMS
    Set colAccountants = objVM.Accountant
        Wscript.Echo "Virtual machine: " & objVM.Name
        Wscript.Echo "Allowable maximum system capacity: " & _
            colAccountants.AllowableMaximumSystemCapacity
        Wscript.Echo "Allowable reserved system capacity: " & _
            colAccountants.AllowableReservedSystemCapacity
        Wscript.Echo "CPU utilization: " & colAccountants.CPUUtilization

        i = 1
        Wscript.Echo "CPU utilization history:"
        For Each intCPUUtilization in colAccountants.CPUUtilizationHistory
            Wscript.Echo vbTab & i & " -- " & intCPUUtilization
            i = i + 1
        Next

        Wscript.Echo "Disk bytes read: " & colAccountants.DiskBytesRead
        Wscript.Echo "Disk bytes written: " & colAccountants.DiskBytesWritten
        Wscript.Echo "Host disk utilization: " & _
            colAccountants.HostDiskUtilization
        Wscript.Echo "Host memory utilization: " & _
            colAccountants.HostMemoryUtilization
        Wscript.Echo "Maximum system capacity: " & _
            colAccountants.MaximumSystemCapacity
        Wscript.Echo "Network bytes received: " & _
            colAccountants.NetworkBytesReceived
        Wscript.Echo "Network bytes sent: " & colAccountants.NetworkBytesSent
        Wscript.Echo "Relative weight: " & colAccountants.RelativeWeight
        Wscript.Echo "Reserved system capacity: " & _
            colAccountants.ReservedSystemCapacity
        Wscript.Echo "Uptime: " & colAccountants.Uptime
        Wscript.Echo
Next

