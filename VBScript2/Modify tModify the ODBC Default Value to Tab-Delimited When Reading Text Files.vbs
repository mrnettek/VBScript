' Description: Modifies the registry so that ODBC assumes that text files -- unless otherwise indicated -- are saved in tab-delimited format.


Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Jet\4.0\Engines\Text"
strValueName = "Format"
strValue = "TabDelimited"
objReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

