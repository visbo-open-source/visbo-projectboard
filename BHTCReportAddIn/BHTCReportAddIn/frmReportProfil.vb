Public Class frmReportProfil

    Private Sub RepProfilListbox_load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer
       
        For i = 0 To 30

            Try

                RepProfilListbox.Items.Add("aaa" & CStr(i))

            Catch ex As Exception

            End Try

        Next i
       
    End Sub
    Private Sub RepProfilListbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RepProfilListbox.SelectedIndexChanged
      
    End Sub

    Private Sub vonDate_ValueChanged(sender As Object, e As EventArgs) Handles vonDate.ValueChanged

    End Sub

    Private Sub bisDate_ValueChanged(sender As Object, e As EventArgs) Handles bisDate.ValueChanged

    End Sub

    Private Sub ReportErstellen_Click(sender As Object, e As EventArgs) Handles ReportErstellen.Click

    End Sub

    Private Sub changeProfil_Click(sender As Object, e As EventArgs) Handles changeProfil.Click

    End Sub
End Class