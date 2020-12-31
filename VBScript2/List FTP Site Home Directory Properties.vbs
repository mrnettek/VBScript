' Description: Lists default home directory properties for FTP sites on an IIS server.


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from IIsFtpServiceSetting")

For Each objItem in colItems
    Wscript.Echo "MS-DOS Directory Output: " & _
        objItem.MSDOSDirOutput
Next

