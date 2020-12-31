' Description: Deletes a metabase backup named ScriptedBackup.


Const BACKUP_VERSION = 0

strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set objComputer = _
    objWMIService.Get("IIsComputer.Name='LM'")

objComputer.DeleteBackup "ScriptedBackup", BACKUP_VERSION

