' Description: Returns information about the administrative template files (.ADM files) used in applying Group Policy to a computer. To return information about Group Policy applied to the logged-on user, change the WMI class from root\rsop\computer to root\rsop\user.


Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
 
strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery _
    ("Select * from RSOP_AdministrativeTemplateFile")

For Each objItem in colItems  
    Wscript.Echo "GPO ID: " & objItem.GPOID
    dtmConvertedDate.Value = objItem.LastWriteTime
    dtmCreationTime = dtmConvertedDate.GetVarDate
    Wscript.Echo "Last Write Time: " & dtmCreationTime 
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo
Next

