Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Call MsgBox("hello world")
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Call MsgBox("byebye world")
    End Sub

End Class
