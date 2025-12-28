Module MDPartitionSeria
    Public Function ShowDriveInfo(ByVal drvpath)
        Dim fso, d, s, t
        fso = CreateObject("Scripting.FileSystemObject")
        d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(drvpath)))
        Select Case d.DriveType
            Case 0 : t = "Unknown"
            Case 1 : t = "Removable"
            Case 2 : t = "Fixed"
            Case 3 : t = "Network"
            Case 4 : t = "CD-ROM"
            Case 5 : t = "RAM Disk"
        End Select
        s = d.SerialNumber
        ShowDriveInfo = s
    End Function

    Public Function ShowDrivetype(ByVal drvpath)
        Dim fso, d, t
        fso = CreateObject("Scripting.FileSystemObject")
        d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(drvpath)))
        Select Case d.DriveType
            Case 0 : t = "Unknown"
            Case 1 : t = "Removable"
            Case 2 : t = "Fixed"
            Case 3 : t = "Network"
            Case 4 : t = "CD-ROM"
            Case 5 : t = "RAM Disk"
        End Select
        ShowDrivetype = 0
    End Function
End Module
