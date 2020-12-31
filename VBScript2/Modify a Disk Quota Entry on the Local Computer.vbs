' Description: Modifies the disk quota warning threshold and disk space limit for a user named kenmyer. This script must be run on the local computer.


Set colDiskQuotas = CreateObject("Microsoft.DiskQuota.1")
colDiskQuotas.Initialize "C:\", True
set objUser = colDiskQuotas.FindUser("kenmyer")

objUser.QuotaThreshold = 90000000
objUser.QuotaLimit = 100000000

