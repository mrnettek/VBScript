On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\ServiceModel")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Operation", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Action: " & objItem.Action
      WScript.Echo "AsyncPattern: " & objItem.AsyncPattern
      strBehaviors = Join(objItem.Behaviors, ",")
         WScript.Echo "Behaviors: " & strBehaviors
      WScript.Echo "IsCallback: " & objItem.IsCallback
      WScript.Echo "IsInitiating: " & objItem.IsInitiating
      WScript.Echo "IsOneWay: " & objItem.IsOneWay
      WScript.Echo "IsTerminating: " & objItem.IsTerminating
      WScript.Echo "MethodSignature: " & objItem.MethodSignature
      WScript.Echo "Name: " & objItem.Name
      strParameterTypes = Join(objItem.ParameterTypes, ",")
         WScript.Echo "ParameterTypes: " & strParameterTypes
      WScript.Echo "ReplyAction: " & objItem.ReplyAction
      WScript.Echo "ReturnType: " & objItem.ReturnType
      WScript.Echo
   Next
Next

