On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2\Applications\MicrosoftIE")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MicrosoftIE_Cache", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AvailableCacheSize: " & objItem.AvailableCacheSize
      WScript.Echo "AvailableDiskSpace: " & objItem.AvailableDiskSpace
      WScript.Echo "Caption: " & objItem.Caption
      WScript.Echo "Description: " & objItem.Description
      WScript.Echo "MaxCacheSize: " & objItem.MaxCacheSize
      WScript.Echo "PageRefreshType: " & objItem.PageRefreshType
      WScript.Echo "SettingID: " & objItem.SettingID
      WScript.Echo "TempInternetFilesFolder: " & objItem.TempInternetFilesFolder
      WScript.Echo "TotalDiskSpace: " & objItem.TotalDiskSpace
      WScript.Echo
   Next
Next

