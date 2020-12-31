' Description: Displays IIS compression scheme setting information.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsCompressionSchemeSetting")
 
For Each objItem in colItems
    Wscript.Echo "Admin ACL Bin: " & objItem.AdminACLBin
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Hc Compression Dll: " & _
        objItem.HcCompressionDll
    Wscript.Echo "Hc Create Flags: " & objItem.HcCreateFlags
    Wscript.Echo "Hc Do Dynamic Compression: " & _
        objItem.HcDoDynamicCompression
    Wscript.Echo "Hc Do On-Demand Compression: " & _
        objItem.HcDoOnDemandCompression
    Wscript.Echo "Hc Do Static Compression: " & _
        objItem.HcDoStaticCompression
    Wscript.Echo "Hc Dynamic Compression Level: " & _
        objItem.HcDynamicCompressionLevel
    For Each strExtension in objItem.HcFileExtensions
        Wscript.Echo "Hc File Extensions: " & strExtension
    Next
    Wscript.Echo "Hc On-Demand Compression Level: " & _
        objItem.HcOnDemandCompLevel
    Wscript.Echo "Hc Priority: " & objItem.HcPriority
    For Each strExtension in objItem.HcScriptFileExtensions
        Wscript.Echo "Hc Script File Extensions: " & strExtension
    Next
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Setting ID: " & objItem.SettingID
Next

