Public Class frmLoadConstellation

    Private formerselect As String
    Private Sub frmLoadConstellation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            ListBox1.Items.Add(kvp.Key)

        Next
        formerselect = ""

    End Sub

    Private Sub Abbrechen_Click(sender As Object, e As EventArgs) Handles Abbrechen.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        If ListBox1.SelectedItems.Count >= 1 Then
            DialogResult = System.Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            Call MsgBox("bitte einen Eintrag selektieren")
        End If

    End Sub

    Private Sub addToSession_CheckedChanged(sender As Object, e As EventArgs) Handles addToSession.CheckedChanged


    End Sub
End Class