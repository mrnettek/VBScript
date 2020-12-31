Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."

Set objRegistry = GetObject _
    ("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "System\CurrentControlSet\Services\lanmanserver\parameters"
strValueName = "srvcomment"
strDescription = "Description changed programmatically"

objRegistry.SetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strDescription
  


