Public Class Form1
    Private Sub btn_start_Click(sender As Object, e As EventArgs) Handles btn_start.Click
        watchFolder = New System.IO.FileSystemWatcher()


        'this is the path we want to monitor
        watchFolder.Path = My.Computer.FileSystem.CombinePath(rpaPath, "RPA")

        'Add a list of Filter we want to specify
        'make sure you use OR for each Filter as we need to
        'all of those 

        watchFolder.NotifyFilter = IO.NotifyFilters.DirectoryName
        watchFolder.NotifyFilter = watchFolder.NotifyFilter Or
                            IO.NotifyFilters.FileName
        watchFolder.NotifyFilter = watchFolder.NotifyFilter Or
                            IO.NotifyFilters.Attributes

        ' add the handler to each event
        AddHandler watchFolder.Changed, AddressOf logchange
        AddHandler watchFolder.Created, AddressOf logchange
        AddHandler watchFolder.Deleted, AddressOf logchange

        ' add the rename handler as the signature is different
        AddHandler watchFolder.Renamed, AddressOf logrename

        'Set this property to true to start watching
        watchFolder.EnableRaisingEvents = True


    End Sub

    Private Sub btn_stop_Click(sender As Object, e As EventArgs) Handles btn_stop.Click

        'Set this property to true to start watching
        watchFolder.EnableRaisingEvents = False

    End Sub
End Class