strComputer = "."

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")

Set objDictionary = CreateObject("Scripting.Dictionary")

For Each objProcess in colProcessList
    objProcess.GetOwner strNameOfUser, strUserDomain
    strOwner = strUserDomain & "\" & strNameOfUser

    If Not objDictionary.Exists(strOwner) Then
        objDictionary.Add strOwner, strOwner
    End If
Next

For Each strKey in objDictionary.Keys
    Wscript.Echo strKey
Next
  


