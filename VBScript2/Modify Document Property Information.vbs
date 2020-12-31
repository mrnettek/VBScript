' Description: Modifies the Category property included in the summary information properties for a document named C:\Scripts\Test.doc.


Set objPropertyReader = CreateObject("DSOleFile.PropertyReader")
Set objDocument = objPropertyReader.GetDocumentProperties _
    ("C:\Scripts\Test.doc")

objDocument.Category = "Scripting Documents"

