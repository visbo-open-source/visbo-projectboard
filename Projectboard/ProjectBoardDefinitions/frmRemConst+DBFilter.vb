Public Class frmRemoveConstellation
    Private formerselect As String
    Public frmOption As String
    Private Sub frmRemoveConstellation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ListBox1.Items.Clear()

        If frmOption = "DBFilter" Then

            Me.Text = "DB-Filter löschen"
            For Each kvp As KeyValuePair(Of String, clsFilter) In filterDefinitions.filterListe

                ListBox1.Items.Add(kvp.Key)

            Next
            formerselect = ""
        End If

        If frmOption = "ProjConstellation" Then
            Me.Text = "Portfolio löschen"

            For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

                ListBox1.Items.Add(kvp.Key)

            Next
            formerselect = ""
        End If


    End Sub

    Private Sub Abbrechen_Click(sender As Object, e As EventArgs) Handles Abbrechen.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        If ListBox1.Text <> "" Then
            If ListBox1.Text = formerselect Then
                Call MsgBox("ist bereits geladen ...")
                DialogResult = System.Windows.Forms.DialogResult.Cancel
                MyBase.Close()
            Else
                DialogResult = System.Windows.Forms.DialogResult.OK
                MyBase.Close()

            End If

        Else
            Call MsgBox("bitte einen Eintrag selektieren")
            DialogResult = System.Windows.Forms.DialogResult.Retry
        End If


    End Sub

End Class