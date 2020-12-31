Const CreateLog = True

Set objAccess = CreateObject("Access.Application")

errReturn = objAccess.CompactRepair _
    ("c:\scripts\test.mdb", "c:\scripts\test2.mdb", CreateLog)

Wscript.Echo "Compact/repair succeeded: " & errReturn
  


