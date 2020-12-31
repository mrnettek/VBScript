TnSc = "http://search.live.com/macros/scripter/searchscriptcenter/?FORM=MACWLG" 
MSDN = http://search.live.com/macros/scripter/searchmsdndocs/?FORM=OIJG
SC = http://www.microsoft.com/technet/scriptcenter/default.mspx
SR = http://www.microsoft.com/technet/scriptcenter/scripts/default.mspx?mfr=true
HSG = http://www.microsoft.com/technet/scriptcenter/resources/qanda/hsgarch.mspx
OS = http://www.microsoft.com/technet/scriptcenter/resources/officetips/archive.mspx

Set oIE = CreateObject("InternetExplorer.Application")
  oIE.Visible = True
  'open a new window
  oIE.Navigate2 Tnsc
  'open url In new tab
  oIE.Navigate2 MSDN, 2048
  oIE.Navigate2 SC, 2048
  oIE.Navigate2 SR, 2048
  oIE.Navigate2 HSG, 2048
  oIE.Navigate2 OS, 2048
Set oIE = Nothing

