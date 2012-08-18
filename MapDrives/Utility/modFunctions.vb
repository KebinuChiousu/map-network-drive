Imports System.IO

Module modFunctions

    Public Function ListAllDrives() As String()
        'Use DriveInfo to gain access to the drives properties

        Dim allDrives() As DriveInfo = DriveInfo.GetDrives

        Dim volume As String = ""

        Dim idx As Integer

        Dim drv As String()

        ReDim drv(0 To UBound(allDrives))

        'loop through all the drives on the system
        For idx = 0 To UBound(allDrives)
            volume = ""
            On Error Resume Next
            volume = allDrives(idx).VolumeLabel
            drv(idx) = allDrives(idx).Name + " - " + volume
            On Error GoTo 0
        Next

        Return drv

    End Function

End Module
