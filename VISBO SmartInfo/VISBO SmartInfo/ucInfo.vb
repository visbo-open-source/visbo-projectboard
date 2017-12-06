Public Class ucInfo
    Private Sub ucInfo_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        eleAmpelText.MaximumSize = New Drawing.Size(TableLayoutPanel1.Width - eleAmpelText.Margin.Left - eleAmpelText.Margin.Right, 0)
        Console.WriteLine(eleAmpelText.MaximumSize.ToString())
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles eleDeliverables.TextChanged
        ' TODO: adjust size of textbox and show or hide scrollbar
    End Sub

End Class
