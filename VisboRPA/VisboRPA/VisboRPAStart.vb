Imports ProjectBoardBasic
Imports WebServerAcc
Imports ProjectBoardDefinitions
Imports ProjectboardReports
Imports DBAccLayer
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel

Public Class VisboRPAStart
    Private Sub btn_start_Click(sender As Object, e As EventArgs) Handles btn_start.Click

        'this is the path we want to monitor
        watchFolder.Path = My.Computer.FileSystem.CombinePath(rpaPath, "RPA")

        'Add a list of Filter we want to specify
        'make sure you use OR for each Filter as we need to
        'all of those 


        'Set this property to true to start watching
        watchFolder.EnableRaisingEvents = True


    End Sub

    Private Sub btn_stop_Click(sender As Object, e As EventArgs) Handles btn_stop.Click

        'Set this property to true to start watching
        watchFolder.EnableRaisingEvents = False


        ' now store User Login Data
        My.Settings.userNamePWD = awinSettings.userNamePWD

        ' speichern 
        My.Settings.Save()


    End Sub

    Private Sub watchFolder_Changed(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Changed

        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Changed")
        ''code here for newly changed file Or directory

        logfileNamePath = createLogfileName(rpaFolder)

        Call logger(ptErrLevel.logInfo, "WatchFolder_changed", "File '" & e.FullPath & "' was changed at: " & Date.Now().ToLongDateString)

        Dim fullFileName As String = e.FullPath
        Dim myName As String = ""
        Dim rpaCategory As New PTRpa
        Dim result As Boolean = False

        'FileExtension ansehen
        Dim fileExt As String = My.Computer.FileSystem.GetFileInfo(fullFileName).Extension
        Select Case fileExt
            Case ".xlsx"
                myName = My.Computer.FileSystem.GetName(fullFileName)

                ' Bestimme den Import-Typ der zu importierenden Daten
                rpaCategory = bestimmeRPACategory(fullFileName)

                If rpaCategory = PTRpa.visboUnknown Then
                    ' move file to unknown Folder ... 
                    Dim newDestination As String = My.Computer.FileSystem.CombinePath(unknownFolder, myName)
                    My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                    Call logger(ptErrLevel.logInfo, "unknown file / category: ", myName)
                Else
                    result = importOneProject(fullFileName, rpaCategory, Date.Now())
                    If result Then
                        Call logger(ptErrLevel.logInfo, "WatchFolder_changed", "File '" & e.FullPath & "' was imported successfully at: " & Date.Now().ToLongDateString)
                    End If
                End If
            Case ".mpp"

                myName = My.Computer.FileSystem.GetName(fullFileName)

                ' Import Typ ist Microsoft Project File
                rpaCategory = PTRpa.visboMPP

                ' Import wird durchgeführt
                result = importOneProject(fullFileName, rpaCategory, Date.Now())
                If result Then
                    Call logger(ptErrLevel.logInfo, "WatchFolder_changed", "File '" & e.FullPath & "' was imported successfully at: " & Date.Now().ToLongDateString)
                End If
            Case Else

        End Select
    End Sub

    Private Sub watchFolder_Created(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Created
        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Created")
        ''code here for newly changed file Or directory

        logfileNamePath = createLogfileName(rpaFolder)

        Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & e.FullPath & "' was created at: " & Date.Now().ToLongDateString)

        Dim fullFileName As String = e.FullPath
        Dim myName As String = ""
        Dim rpaCategory As New PTRpa
        Dim result As Boolean = False

        ' Completion-File delivered?
        completedOK = LCase(fullFileName).Contains(LCase("Timesheet_completed"))
        If completedOK Then
            'Einlesen der TimeSheets - Telair
            ' nachsehen ob collect vollständig
            myName = My.Computer.FileSystem.GetName(fullFileName)
            result = processVisboActualData2(myName, myActivePortfolio, collectFolder, Date.Now())

        End If


        If My.Computer.FileSystem.FileExists(fullFileName) And Not fullFileName.Contains("~$") Then
            'FileExtension ansehen
            Dim fileExt As String = My.Computer.FileSystem.GetFileInfo(fullFileName).Extension
            Select Case fileExt
                Case ".xlsx"

                    myName = My.Computer.FileSystem.GetName(fullFileName)

                    ' Bestimme den Import-Typ der zu importierenden Daten
                    rpaCategory = bestimmeRPACategory(fullFileName)

                    If rpaCategory = PTRpa.visboUnknown Then
                        ' move file to unknown Folder ... 
                        Dim newDestination As String = My.Computer.FileSystem.CombinePath(unknownFolder, myName)
                        My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                        Call logger(ptErrLevel.logInfo, "unknown file / category: ", myName)
                    Else
                        result = importOneProject(fullFileName, rpaCategory, Date.Now())
                        If result Then
                            Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & e.FullPath & "' was imported successfully at: " & Date.Now().ToLongDateString)
                        End If
                    End If
                Case ".mpp"

                    myName = My.Computer.FileSystem.GetName(fullFileName)

                    ' Import Typ ist Microsoft Project File
                    rpaCategory = PTRpa.visboMPP

                    ' Import wird durchgeführt
                    result = importOneProject(fullFileName, rpaCategory, Date.Now())
                    If result Then
                        Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & e.FullPath & "' was imported successfully at: " & Date.Now().ToLongDateString)
                    End If
                Case Else

            End Select
        Else
            Dim a As String = ""
        End If


    End Sub

    Private Sub watchFolder_Deleted(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Deleted

        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Deleted")
        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "File '" & e.FullPath & "' was deleted at: " & Date.Now().ToLongDateString)
    End Sub

    Private Sub watchFolder_Renamed(sender As Object, e As IO.RenamedEventArgs) Handles watchFolder.Renamed
        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Renamed")
        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "File '" & e.FullPath & "' was renamed at: " & Date.Now().ToLongDateString)
    End Sub

    Private Sub VisboRPAStart_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'this is the path we want to monitor
        watchFolder.Path = My.Computer.FileSystem.CombinePath(rpaPath, "RPA")

        'Add a list of Filter we want to specify
        'make sure you use OR for each Filter as we need to
        'all of those 


        'Set this property to true to start watching
        watchFolder.EnableRaisingEvents = True
    End Sub

    Private Sub VisboRPAStart_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        ' now store User Login Data
        My.Settings.userNamePWD = awinSettings.userNamePWD

        ' speichern 
        My.Settings.Save()

    End Sub
End Class