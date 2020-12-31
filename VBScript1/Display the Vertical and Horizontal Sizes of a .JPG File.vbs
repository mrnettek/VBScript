Set objPlayer = CreateObject("WMPlayer.OCX" )
Set colMediaCollection = objPlayer.mediaCollection

Set objPhotos = colMediaCollection.getByAttribute("MediaType", "photo")

For i = 0 to objPhotos.Count - 1
    Set objPhoto = objPhotos.item(i)
    Wscript.Echo "Name: " & objPhoto.Name
    Wscript.Echo "Height: " & objPhoto.getItemInfo("WM/VideoHeight")
    Wscript.Echo "Width: " & objPhoto.getItemInfo("WM/VideoWidth")
Next
  


