Public Class ucProperties

    Private Sub ucProperties_Leave(sender As Object, e As EventArgs) Handles Me.Leave

    End Sub


    Private Sub ucProperties_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ' label resize
        eleName.MaximumSize = New Drawing.Size(Me.Width - eleName.Margin.Left - eleName.Margin.Right - eleName.Location.X, eleName.MaximumSize.Height)

    End Sub

    Private Sub ucProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If englishLanguage Then
            With Me
                .labelAmpel.Text = "Traffic Light:"
                .labelDate.Text = "Date:"
                .labelDeliver.Text = "Deliverables:"
                .labelRespons.Text = "Responsible:"
            End With
        Else
            With Me
                .labelAmpel.Text = "Ampel:"
                .labelDate.Text = "Datum:"
                .labelDeliver.Text = "Leistungsumfänge:"
                .labelRespons.Text = "Verantwortlich:"
            End With
        End If
    End Sub

    Private Sub eleAmpelText_MouseDoubleClick(sender As Object, e As Windows.Forms.MouseEventArgs) Handles eleAmpelText.MouseDoubleClick

    End Sub


    Private Sub eleAmpelText_TextChanged(sender As Object, e As EventArgs) Handles eleAmpelText.TextChanged

    End Sub
End Class
