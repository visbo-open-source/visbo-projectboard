Imports ProjectBoardBasic
Imports WebServerAcc
Imports ProjectBoardDefinitions
Imports ProjectboardReports
Imports DBAccLayer
Imports Microsoft.Office.Interop.Excel



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

    End Sub

    Private Sub watchFolder_Changed(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Changed
        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Changed")
    End Sub

    Private Sub watchFolder_Created(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Created
        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Created")
    End Sub

    Private Sub watchFolder_Deleted(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Deleted
        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Deleted")
    End Sub

    Private Sub watchFolder_Renamed(sender As Object, e As IO.RenamedEventArgs) Handles watchFolder.Renamed
        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Renamed")
    End Sub

    Private Sub FileSystemWatcher1_Changed(sender As Object, e As IO.FileSystemEventArgs)

    End Sub
End Class