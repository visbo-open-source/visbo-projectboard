Public Class ucProperties

    Private Sub ucProperties_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ' label resize
        eleName.MaximumSize = New Drawing.Size(Me.Width - eleName.Margin.Left - eleName.Margin.Right - eleName.Location.X, eleName.MaximumSize.Height)

    End Sub

    Private Sub ucProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Call MsgBox("load")
    End Sub

    Private Sub ucProperties_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        'Call MsgBox("visibleChanged")
    End Sub

End Class
