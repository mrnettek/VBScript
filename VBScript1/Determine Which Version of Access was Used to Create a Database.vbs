Set objAccess = CreateObject("Access.Application")
objAccess.OpenCurrentDatabase "C:\Scripts\Test.mdb"

intFormat = objAccess.CurrentProject.FileFormat

Select Case intFormat
    Case 2 Wscript.Echo "Microsoft Access 2" 
    Case 7 Wscript.Echo "Microsoft Access 95"
    Case 8 Wscript.Echo "Microsoft Access 97" 
    Case 9 Wscript.Echo "Microsoft Access 2000"
    Case 10 Wscript.Echo "Microsoft Access 2003"
End Select
  


