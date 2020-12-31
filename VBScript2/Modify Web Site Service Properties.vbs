' Description: Demonstration script that modifies IIS compression schemes property values.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsCompressionSchemesSetting")

For Each objItem in colItems
    objItem.HcCompressionDirectory = "C:\Compressed_Files"
    objItem.HcDoDynamicCompression = True
    objItem.HcDoStaticCompression = True
    objItem.DoDiskSpaceLimiting = True
    objItem.HcMaxDiskSpaceUsage = 200000000
    objItem.Put_
Next

