Public Class ucInfo
    Private Sub ucInfo_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ' label resize
        eleAmpelText.MaximumSize = New Drawing.Size(TableLayoutPanel1.Width - eleAmpelText.Margin.Left - eleAmpelText.Margin.Right, 0)

        ' textbox resize
        eleDeliverables.Height = eleDeliverables.CreateGraphics().MeasureString(eleDeliverables.Text, eleDeliverables.Font, eleDeliverables.Width).Height
    End Sub


    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles eleDeliverables.TextChanged
        ' TODO: adjust size of textbox and show or hide scrollbar

        ' textbox resize
        ' eleDeliverables.Height = eleDeliverables.CreateGraphics().MeasureString(eleDeliverables.Text, eleDeliverables.Font, eleDeliverables.Width).Height

    End Sub
End Class
