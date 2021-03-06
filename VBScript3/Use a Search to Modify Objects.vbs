' Description: Demonstration script that returns a list of all the users in the Accounting Department, and then changes their department name to Finance.


On Error Resume Next

Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

objCommand.CommandText = _
    "SELECT ADsPath FROM 'LDAP://dc=fabrikam,dc=com' WHERE " _
        & "objectCategory='user' AND Department='Accounting'"  
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst

Do Until objRecordSet.EOF
    strPath = objRecordSet.Fields("AdsPath").Value
    Set objUser = GetObject(strPath)
    objUser.Department = "Finance"
    objUser.SetInfo
objRecordSet.MoveNext
Loop

