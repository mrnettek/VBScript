On Error Resume Next

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20

arrComputers = Array(".")
For Each strComputer In arrComputers
   WScript.Echo
   WScript.Echo "=========================================="
   WScript.Echo "Computer: " & strComputer
   WScript.Echo "=========================================="

   Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\MSAPPS11")
   Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Word11Settings", "WQL", _
                                          wbemFlagReturnImmediately + wbemFlagForwardOnly)

   For Each objItem In colItems
      WScript.Echo "AllowDragAndDrop: " & objItem.AllowDragAndDrop
      WScript.Echo "AnimateScreenMovements: " & objItem.AnimateScreenMovements
      WScript.Echo "AutoHyphenation: " & objItem.AutoHyphenation
      WScript.Echo "BackgroundSave: " & objItem.BackgroundSave
      WScript.Echo "BlueScreen: " & objItem.BlueScreen
      WScript.Echo "ConfirmConversions: " & objItem.ConfirmConversions
      WScript.Echo "CreateBackup: " & objItem.CreateBackup
      WScript.Echo "DefaultFilePath: " & objItem.DefaultFilePath
      WScript.Echo "DefaultSaveFormat: " & objItem.DefaultSaveFormat
      WScript.Echo "DisplayAutoCompleteTips: " & objItem.DisplayAutoCompleteTips
      WScript.Echo "DisplayHorizontalScrollBar: " & objItem.DisplayHorizontalScrollBar
      WScript.Echo "DisplayRecentFiles: " & objItem.DisplayRecentFiles
      WScript.Echo "DisplayRulers: " & objItem.DisplayRulers
      WScript.Echo "DisplayScreenTips: " & objItem.DisplayScreenTips
      WScript.Echo "DisplayScrollBars: " & objItem.DisplayScrollBars
      WScript.Echo "DisplayStatusBar: " & objItem.DisplayStatusBar
      WScript.Echo "EnableSound: " & objItem.EnableSound
      WScript.Echo "FieldShading: " & objItem.FieldShading
      WScript.Echo "INSKeyForPaste: " & objItem.INSKeyForPaste
      WScript.Echo "MeasurementUnits: " & objItem.MeasurementUnits
      WScript.Echo "MinimumFontSize: " & objItem.MinimumFontSize
      WScript.Echo "Overtype: " & objItem.Overtype
      WScript.Echo "Pagination: " & objItem.Pagination
      WScript.Echo "PrintFormsData: " & objItem.PrintFormsData
      WScript.Echo "PrintPostScriptOverText: " & objItem.PrintPostScriptOverText
      WScript.Echo "PrintRevisions: " & objItem.PrintRevisions
      WScript.Echo "RecentFilesMaximum: " & objItem.RecentFilesMaximum
      WScript.Echo "ReplaceSelection: " & objItem.ReplaceSelection
      WScript.Echo "SaveFormsData: " & objItem.SaveFormsData
      WScript.Echo "SaveInterval: " & objItem.SaveInterval
      WScript.Echo "SaveNormalPrompt: " & objItem.SaveNormalPrompt
      WScript.Echo "SavePropertiesPrompt: " & objItem.SavePropertiesPrompt
      WScript.Echo "ShowAll: " & objItem.ShowAll
      WScript.Echo "ShowAnimation: " & objItem.ShowAnimation
      WScript.Echo "ShowBookmarks: " & objItem.ShowBookmarks
      WScript.Echo "ShowDrawings: " & objItem.ShowDrawings
      WScript.Echo "ShowFieldCodes: " & objItem.ShowFieldCodes
      WScript.Echo "ShowHiddenText: " & objItem.ShowHiddenText
      WScript.Echo "ShowHighlight: " & objItem.ShowHighlight
      WScript.Echo "ShowHyphens: " & objItem.ShowHyphens
      WScript.Echo "ShowMainTextLayer: " & objItem.ShowMainTextLayer
      WScript.Echo "ShowObjectAnchors: " & objItem.ShowObjectAnchors
      WScript.Echo "ShowParagraphs: " & objItem.ShowParagraphs
      WScript.Echo "ShowRevisions: " & objItem.ShowRevisions
      WScript.Echo "ShowSpaces: " & objItem.ShowSpaces
      WScript.Echo "ShowSummary: " & objItem.ShowSummary
      WScript.Echo "ShowTabs: " & objItem.ShowTabs
      WScript.Echo "ShowTextBoundaries: " & objItem.ShowTextBoundaries
      WScript.Echo "SmartCutPaste: " & objItem.SmartCutPaste
      WScript.Echo "TabIndentKey: " & objItem.TabIndentKey
      WScript.Echo "TableGridlines: " & objItem.TableGridlines
      WScript.Echo "TrackRevisions: " & objItem.TrackRevisions
      WScript.Echo
   Next
Next

