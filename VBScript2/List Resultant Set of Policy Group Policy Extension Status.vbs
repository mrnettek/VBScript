' Description: Returns information about client-side processing of Group Policy extensions. To return information about Group Policy applied to the logged-on user, change the WMI class from root\rsop\computer to root\rsop\user.


On Error Resume Next
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
 
strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_ExtensionStatus")

For Each objItem in colItems  
    Wscript.Echo "Display name: " & objItem.DisplayName
    dtmConvertedDate.Value = objItem.BeginTime
    dtmTime = dtmConvertedDate.GetVarDate
    Wscript.Echo "Begin time: " & dtmTime
    dtmConvertedDate.Value = objItem.EndTime
    dtmTime = dtmConvertedDate.GetVarDate
    Wscript.Echo "End time: " & dtmTime
    Wscript.Echo "Error: " & objItem.Error
    Wscript.Echo "Extension GUID: " & objItem.ExtensionGuid
    Wscript.Echo "Logging Status: " & objItem.LoggingStatus
    Wscript.Echo
Next

