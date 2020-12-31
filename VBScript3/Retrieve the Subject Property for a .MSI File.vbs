Set objInstaller = CreateObject("WindowsInstaller.Installer") 
Set objProduct = objInstaller.SummaryInformation("C:\Scripts\FP11.MSI")

Wscript.Echo "Subject: " & objProduct.Property(3)
