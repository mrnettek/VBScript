Set objPlayer = CreateObject("WMPlayer.OCX" )

Set objMediaCollection = objPlayer.MediaCollection
Set objTempList = objMediaCollection.getByName("Scrambled Eggs")

Set objSong = objTempList.Item(0)
objSong.setItemInfo "Name", "Yesterday"
  


