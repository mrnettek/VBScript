On Error Resume Next

Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strSitesContainer = "LDAP://cn=Sites," & strConfigurationNC
Set objSitesContainer = GetObject(strSitesContainer)
objSitesContainer.Filter = Array("site")
 
For Each objSite In objSitesContainer
    Wscript.Echo objSite.CN
    strSiteName = objSite.Name
    strServerPath = "LDAP://cn=Servers," & strSiteName & ",cn=Sites," & _
        strConfigurationNC
    Set colServers = GetObject(strServerPath)
 
    For Each objServer In colServers
        WScript.Echo vbTab & objServer.CN
    Next
    Wscript.Echo
Next
  


