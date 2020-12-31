' Description: Retrieves and displays the textual representation of a WMI class definition in MOF (Managed Object Format) syntax.


strComputer = "."
strNameSpace = "root\cimv2"
strClass = "Win32_Service"
 
Const wbemFlagUseAmendedQualifiers = &h20000
 
Set objClass = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\" & strNameSpace)

Set objClass = objWMIService.Get(strClass, wbemFlagUseAmendedQualifiers)
strMOF = objClass.GetObjectText_
 
WScript.Echo strMOF

