Const HKEY_CURRENT_USER = &H80000001

strComputer = "."

Set objRegistry = GetObject _
    ("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

strKeyPath = "Software"
strValueName = "BinaryTest"
arrValues = Array(1,2,3,4,5,6,7,8,9,10)

errReturn = objRegistry.SetBinaryValue _
    (HKEY_CURRENT_USER, strKeyPath, strValueName, arrValues)
  


