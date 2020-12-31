Set objPlayer = CreateObject("WMPlayer.OCX" )

Set objMediaCollection = objPlayer.MediaCollection
Set colSongList = objMediaCollection.getByAttribute("MediaType", "audio")

For i = 0 to colSongList.Count - 1
    Set objSong = colSongList.Item(i)
    Wscript.Echo objSong.Name & " -- " & objSong.getItemInfo("WM/AlbumArtist")
Next
  


