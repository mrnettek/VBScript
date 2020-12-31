On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_NETFramework_NETCLRLoading",,48)
For Each objItem in colItems
    Wscript.Echo "AssemblySearchLength: " & objItem.AssemblySearchLength
    Wscript.Echo "BytesinLoaderHeap: " & objItem.BytesinLoaderHeap
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Currentappdomains: " & objItem.Currentappdomains
    Wscript.Echo "CurrentAssemblies: " & objItem.CurrentAssemblies
    Wscript.Echo "CurrentClassesLoaded: " & objItem.CurrentClassesLoaded
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Frequency_Object: " & objItem.Frequency_Object
    Wscript.Echo "Frequency_PerfTime: " & objItem.Frequency_PerfTime
    Wscript.Echo "Frequency_Sys100NS: " & objItem.Frequency_Sys100NS
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PercentTimeLoading: " & objItem.PercentTimeLoading
    Wscript.Echo "Rateofappdomains: " & objItem.Rateofappdomains
    Wscript.Echo "Rateofappdomainsunloaded: " & objItem.Rateofappdomainsunloaded
    Wscript.Echo "RateofAssemblies: " & objItem.RateofAssemblies
    Wscript.Echo "RateofClassesLoaded: " & objItem.RateofClassesLoaded
    Wscript.Echo "RateofLoadFailures: " & objItem.RateofLoadFailures
    Wscript.Echo "Timestamp_Object: " & objItem.Timestamp_Object
    Wscript.Echo "Timestamp_PerfTime: " & objItem.Timestamp_PerfTime
    Wscript.Echo "Timestamp_Sys100NS: " & objItem.Timestamp_Sys100NS
    Wscript.Echo "TotalAppdomains: " & objItem.TotalAppdomains
    Wscript.Echo "Totalappdomainsunloaded: " & objItem.Totalappdomainsunloaded
    Wscript.Echo "TotalAssemblies: " & objItem.TotalAssemblies
    Wscript.Echo "TotalClassesLoaded: " & objItem.TotalClassesLoaded
    Wscript.Echo "TotalNumberofLoadFailures: " & objItem.TotalNumberofLoadFailures
Next

