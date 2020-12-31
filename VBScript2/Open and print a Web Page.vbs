On Error Resume Next

Const OLECMDID_PRINT = 6
Const OLECMDEXECOPT_DONTPROMPTUSER = 2
Const PRINT_WAITFORCOMPLETION = 2

Dim oIExplorer : Set oIExplorer = CreateObject("InternetExplorer.Application")
oIExplorer.Navigate "http://www.scriptbox.at.tt/"
oIExplorer.Visible = 1

Do while oIExplorer.ReadyState <> 4
	wscript.sleep 1000
Loop

oIExplorer.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER


