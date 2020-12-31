' Description: Locates update ac94db3b-e1a8-4e92-9fd0-e86f355e6a44 and then displays the Software Update Service wizard.


Set objCollection = CreateObject("Microsoft.Update.UpdateColl")

Set objSearcher = CreateObject("Microsoft.Update.Searcher")
Set objResults = objSearcher.Search _
    ("UpdateID='ac94db3b-e1a8-4e92-9fd0-e86f355e6a44'")
Set colUpdates = objResults.Updates

objCollection.Add(colUpdates.Item(0))
Set objInstaller = CreateObject("Microsoft.Update.Installer")
objInstaller.Updates = objCollection
Set objResults = objInstaller.RunWizard

