Const FILE_SHARE = 0
Const MAXIMUM_CONNECTIONS = 25

strComputer = "atl-ws-01"
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")

Set objNewShare = objWMIService.Get("Win32_Share")

errReturn = objNewShare.Create _
    ("C:\Public", "PublicShare", FILE_SHARE, _
        MAXIMUM_CONNECTIONS, "Public share for Fabrikam employees.")
  


