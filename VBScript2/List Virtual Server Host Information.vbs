' Description: Lists Virtual Server host computer  information.


On Error Resume Next

Set objVS = CreateObject("VirtualServer.Application")
Set objHost = objVS.HostInfo

Wscript.Echo "Logical processor count: " & objHost.LogicalProcessorCount
Wscript.Echo "Memory: " & objHost.Memory
Wscript.Echo "Memory available: " & objHost.MemoryAvail
Wscript.Echo "Memory available string: " & objHost.MemoryAvailString
Wscript.Echo "Memory total string: " & objHost.MemoryTotalString
Wscript.Echo "MMX: " & objHost.MMX
Wscript.Echo "Operating system: " & objHost.OperatingSystem
Wscript.Echo "OS major version: " & objHost.OSMajorVersion
Wscript.Echo "OS minor version: " & objHost.OSMinorVersion
Wscript.Echo "OS service pack string: " & objHost.OSServicePackString
Wscript.Echo "OS version string: " & objHost.OSVersionString
Wscript.Echo "Parallel port: " & objHost.ParallelPort
Wscript.Echo "Physical processor count: " & objHost.PhysicalProcessorCount
Wscript.Echo "Processor features string: " & objHost.ProcessorFeaturesString
Wscript.Echo "Processor manufacturer string: " & _
    objHost.ProcessorManufacturerString
Wscript.Echo "Processor speed: " & objHost.ProcessorSpeed
Wscript.Echo "Processor speed string: " & objHost.ProcessorSpeedString
Wscript.Echo "Processor version string: " & objHost.ProcessorVersionString
Wscript.Echo "SSE: " & objHost.SSE
Wscript.Echo "SSE2: " & objHost.SSE2
Wscript.Echo "3DNow!: " & objHost.ThreeDNow
Wscript.Echo "UTC time: " & objHost.UTCTime

