
msgbox GetDriveSize("C:", 1, 2) & " Bytes"

msgbox GetDriveSize("C:", 2, 2) & " KB"

msgbox GetDriveSize("C:", 3, 2) & " MB"

msgbox GetDriveSize("C:", 4, 2) & " GB"



' --------------------------------------------
Function GetDriveSize(sDrive, iFormat, iRound)

Dim vDriveSize
Dim bObjectCreate : bObjectCreate = false

If not IsObject(ofso) then
	Dim ofso : Set ofso = CreateObject("Scripting.FileSystemObject")
	bObjectCreate = True	
End if

            If ofso.DriveExists(sDrive) then

                        Select Case iFormat

                                   Case 1  '---Return Bytes
                                               vDriveSize = Round(ofso.GetDrive(sDrive).TotalSize,iRound)

                                   Case 2 '---Return KB
                                               vDriveSize = Round((ofso.GetDrive(sDrive).TotalSize/1024),iRound)

                                   Case 3 '---Return MB
                                               vDriveSize = Round(((ofso.GetDrive(sDrive).TotalSize/1024)/1024),iRound)

                                   Case 4 '---Return GB
                                               vDriveSize = Round((((ofso.GetDrive(sDrive).TotalSize/1024)/1024)/1024),iRound)
                        End Select

            Else
                        If IsObject(ofso) and bObjectCreate then Set ofso = Nothing
                        GetDriveSize = -1
            End if

            If IsObject(ofso) and bObjectCreate then Set ofso = Nothing
            GetDriveSize = vDriveSize

End Function

