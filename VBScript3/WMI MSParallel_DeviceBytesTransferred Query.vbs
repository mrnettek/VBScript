On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\WMI")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM MSParallel_DeviceBytesTransferred", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "Active: " & objItem.Active
      WScript.Echo "BoundedEcpReadCount: " & objItem.BoundedEcpReadCount
      WScript.Echo "BoundedEcpWriteCount: " & objItem.BoundedEcpWriteCount
      WScript.Echo "ByteReadCount: " & objItem.ByteReadCount
      WScript.Echo "ChannelNibbleReadCount: " & objItem.ChannelNibbleReadCount
      WScript.Echo "Flags1: " & objItem.Flags1
      WScript.Echo "Flags2: " & objItem.Flags2
      WScript.Echo "HwEcpReadCount: " & objItem.HwEcpReadCount
      WScript.Echo "HwEcpWriteCount: " & objItem.HwEcpWriteCount
      WScript.Echo "HwEppReadCount: " & objItem.HwEppReadCount
      WScript.Echo "HwEppWriteCount: " & objItem.HwEppWriteCount
      WScript.Echo "InstanceName: " & objItem.InstanceName
      WScript.Echo "NibbleReadCount: " & objItem.NibbleReadCount
      strspare = Join(objItem.spare, ",")
         WScript.Echo "spare: " & strspare
      WScript.Echo "SppWriteCount: " & objItem.SppWriteCount
      WScript.Echo "SwEcpReadCount: " & objItem.SwEcpReadCount
      WScript.Echo "SwEcpWriteCount: " & objItem.SwEcpWriteCount
      WScript.Echo "SwEppReadCount: " & objItem.SwEppReadCount
      WScript.Echo "SwEppWriteCount: " & objItem.SwEppWriteCount
      WScript.Echo
   Next
Next

