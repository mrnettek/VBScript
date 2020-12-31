' Description: Backs up the metabase on the local computer, using the password er$3qld9o.


Const MD_BACKUP_HIGHEST_VERSION = &HFFFFFFFE
Const MD_BACKUP_OVERWRITE = 1

strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:{authenticationLevel=pktPrivacy}\\" _
        & strComputer & "\root\microsoftiisv2")

Set objComputer = _
    objWMIService.Get("IIsComputer.Name='LM'")

objComputer.BackupWithPassword "ScriptedBackup", _
    MD_BACKUP_HIGHEST_VERSION, MD_BACKUP_OVERWRITE, _
        "er$3qld9o"

