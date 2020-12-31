' Description: Displays service properties for Software Update Services.


Set objServiceManager = CreateObject("Microsoft.Update.ServiceManager")
Set colServices = objServiceManager.Services

For i = 0 to colServices.Count - 1
    Wscript.Echo "Name: " & colServices.Item(i).Name
    Wscript.Echo "Is managed: " & colServices.Item(i).IsManaged
    Wscript.Echo "Is registered with Automatic Updates: " & _
        colServices.Item(i).IsRegisteredWithAU
    Wscript.Echo "Issue date: " & colServices.Item(i).IssueDate
    Wscript.Echo "Offers Windows updates: " & _
        colServices.Item(i).OffersWindowsUpdates
    Wscript.Echo  "Redirection URL: " &colServices.Item(i).RedirectURL
    Wscript.Echo "Service ID: " & colServices.Item(i).ServiceID
    Wscript.Echo "UI Plugin Class ID: " & colServices.Item(i).UIPluginClsid
Next

