On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Msft_Providers", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "HostingGroup: " & objItem.HostingGroup
      WScript.Echo "HostingSpecification: " & objItem.HostingSpecification
      WScript.Echo "HostProcessIdentifier: " & objItem.HostProcessIdentifier
      WScript.Echo "Locale: " & objItem.Locale
      WScript.Echo "Namespace: " & objItem.Namespace
      WScript.Echo "provider: " & objItem.provider
      WScript.Echo "ProviderOperation_AccessCheck: " & objItem.ProviderOperation_AccessCheck
      WScript.Echo "ProviderOperation_CancelQuery: " & objItem.ProviderOperation_CancelQuery
      WScript.Echo "ProviderOperation_CreateClassEnumAsync: " & objItem.ProviderOperation_CreateClassEnumAsync
      WScript.Echo "ProviderOperation_CreateInstanceEnumAsync: " & objItem.ProviderOperation_CreateInstanceEnumAsync
      WScript.Echo "ProviderOperation_CreateRefreshableEnum: " & objItem.ProviderOperation_CreateRefreshableEnum
      WScript.Echo "ProviderOperation_CreateRefreshableObject: " & objItem.ProviderOperation_CreateRefreshableObject
      WScript.Echo "ProviderOperation_CreateRefresher: " & objItem.ProviderOperation_CreateRefresher
      WScript.Echo "ProviderOperation_DeleteClassAsync: " & objItem.ProviderOperation_DeleteClassAsync
      WScript.Echo "ProviderOperation_DeleteInstanceAsync: " & objItem.ProviderOperation_DeleteInstanceAsync
      WScript.Echo "ProviderOperation_ExecMethodAsync: " & objItem.ProviderOperation_ExecMethodAsync
      WScript.Echo "ProviderOperation_ExecQueryAsync: " & objItem.ProviderOperation_ExecQueryAsync
      WScript.Echo "ProviderOperation_FindConsumer: " & objItem.ProviderOperation_FindConsumer
      WScript.Echo "ProviderOperation_GetObjectAsync: " & objItem.ProviderOperation_GetObjectAsync
      WScript.Echo "ProviderOperation_GetObjects: " & objItem.ProviderOperation_GetObjects
      WScript.Echo "ProviderOperation_GetProperty: " & objItem.ProviderOperation_GetProperty
      WScript.Echo "ProviderOperation_NewQuery: " & objItem.ProviderOperation_NewQuery
      WScript.Echo "ProviderOperation_ProvideEvents: " & objItem.ProviderOperation_ProvideEvents
      WScript.Echo "ProviderOperation_PutClassAsync: " & objItem.ProviderOperation_PutClassAsync
      WScript.Echo "ProviderOperation_PutInstanceAsync: " & objItem.ProviderOperation_PutInstanceAsync
      WScript.Echo "ProviderOperation_PutProperty: " & objItem.ProviderOperation_PutProperty
      WScript.Echo "ProviderOperation_QueryInstances: " & objItem.ProviderOperation_QueryInstances
      WScript.Echo "ProviderOperation_SetRegistrationObject: " & objItem.ProviderOperation_SetRegistrationObject
      WScript.Echo "ProviderOperation_StopRefreshing: " & objItem.ProviderOperation_StopRefreshing
      WScript.Echo "ProviderOperation_ValidateSubscription: " & objItem.ProviderOperation_ValidateSubscription
      WScript.Echo "TransactionIdentifier: " & objItem.TransactionIdentifier
      WScript.Echo "User: " & objItem.User
      WScript.Echo
   Next
Next

