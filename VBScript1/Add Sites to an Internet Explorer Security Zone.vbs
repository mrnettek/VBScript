' Description: Adds the Web site Contoso.com to the Trusted sites zone and BenefitsWeb to the Local intranet zone on a computer running Internet Explorer Enhanced Security Configuration.


On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\ESCDomains\Contoso.com"
objReg.CreateKey HKEY_CURRENT_USER,strKeyPath
strValueName = "http"
dwValue = 2
objReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\ESCDomains\BenefitsWeb"
objReg.CreateKey HKEY_CURRENT_USER,strKeyPath
strValueName = "*"
dwValue = 1
objReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,dwValue

