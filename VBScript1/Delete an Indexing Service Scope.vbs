' Description: Removes the Indexing Service scope C:\Scripts from the Indexing Service catalog named Script Catalog.


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
Set objCatalog = objAdminIS.GetCatalogByName("Script Catalog")
objCatalog.RemoveScope("c:\scripts")

