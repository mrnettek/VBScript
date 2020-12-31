' Description: Returns information about compression scheme configuration on an IIS server.


On Error Resume Next
 
strComputer = "LocalHost"
Set colSchemes = GetObject _
    ("IIS://" & strComputer & "/W3SVC/Filters/Compression")
 
For Each objItem in colSchemes
    If objItem.Name <> "Parameters" Then
        wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Compression DLL: " & objItem.HcCompressionDll
        Wscript.Echo "Create Flags: " & objItem.HcCreateFlags
        Wscript.Echo "Do Dynamic Compression: " &  _
            objItem.HcDoDynamicCompression
        Wscript.Echo "Do Static Compression: " &  _
            objItem.HcDoStaticCompression
        Wscript.Echo "Do On-Demand Compression: " &  _
            objItem.HcDoOnDemandCompression
        Wscript.Echo "Dynamic Compression Level: " &  _
            objItem.HcDynamicCompressionLevel
        For Each strExtension in objItem.HcFileExtensions
            Wscript.Echo "File Extensions: " & strExtension
        Next
        Wscript.Echo "On-Demand Compression Level: " &  _
            objItem.HcOnDemandCompLevel
        Wscript.Echo "Priority: " & objItem.HcPriority
        For Each strExtension in objItem.HcScriptFileExtensions
            Wscript.Echo "Script File Extensions: " & strExtension
        Next
        Wscript.Echo
    End If
Next

