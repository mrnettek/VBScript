' Description: Publishes a shared folder in Active Directory, assigning the folder a description and three keywords.


Set objComputer = GetObject _
    ("LDAP://OU=Finance, DC=fabrikam, DC=com")

Set objShare = objComputer.Create("volume", "CN=FinanceShare")

objShare.Put "uNCName", "\\atl-dc-02\FinanceShare"
objShare.Put "Description", "Public share for users in the Finance group."
objShare.Put "Keywords", Array("finance", "fiscal", "monetary") 
objShare.SetInfo

