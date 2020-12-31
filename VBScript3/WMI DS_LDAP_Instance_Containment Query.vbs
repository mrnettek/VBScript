On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\directory\LDAP")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM DS_LDAP_Instance_Containment", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "ChildInstance: " & objItem.ChildInstance
      WScript.Echo "ParentInstance: " & objItem.ParentInstance
      WScript.Echo
   Next
Next

