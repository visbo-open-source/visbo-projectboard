Public Class frmSettings

    
    Private Sub schriftSize_TextChanged(sender As Object, e As EventArgs) Handles schriftSize.TextChanged

        Try
            schriftGroesse = CDbl(schriftSize.Text)
        Catch ex As Exception
            schriftSize.Text = schriftGroesse.ToString
        End Try
    End Sub

    Private Sub abstandseinheit_SelectedIndexChanged(sender As Object, e As EventArgs) Handles abstandseinheit.SelectedIndexChanged

        If abstandseinheit.Text = "Tagen" Then
            absEinheit = pptAbsUnit.tage
        ElseIf abstandseinheit.Text = "Wochen" Then
            absEinheit = pptAbsUnit.wochen
        Else
            absEinheit = pptAbsUnit.monate
        End If

    End Sub

    Private Sub showInfoBC_CheckedChanged(sender As Object, e As EventArgs) Handles showInfoBC.CheckedChanged

        showBreadCrumbField = showInfoBC.Checked

    End Sub

    Private Sub extendedSearch_CheckedChanged(sender As Object, e As EventArgs) Handles extendedSearch.CheckedChanged
        extSearch = extendedSearch.Checked
    End Sub

    Private Sub protectShapes_CheckedChanged(sender As Object, e As EventArgs) Handles protectShapes.CheckedChanged

        Dim sichtbar As Boolean = Not protectShapes.Checked

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            If tmpShape.Tags.Count > 0 Then
                If isRelevantShape(tmpShape) Then
                    ' Sichtbarkeit setzen ....
                    tmpShape.Visible = sichtbar
                End If
            End If
        Next

    End Sub
End Class