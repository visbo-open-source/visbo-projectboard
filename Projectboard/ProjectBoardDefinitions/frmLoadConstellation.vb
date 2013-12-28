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

        If ListBox1.Text <> "" Then
            If ListBox1.Text = formerselect Then
                Call MsgBox("ist bereits geladen ...")
                DialogResult = System.Windows.Forms.DialogResult.Cancel
                MyBase.Close()
            Else
                DialogResult = System.Windows.Forms.DialogResult.OK
                MyBase.Close()
                'formerselect = ListBox1.Text
                'Call awinLoadConstellation(ListBox1.Text)

                'appInstance.ScreenUpdating = False
                'Call diagramsVisible(False)
                'Call awinClearPlanTafel()
                'Call awinZeichnePlanTafel()
                'Call awinNeuZeichnenDiagramme(2)
                'Call diagramsVisible(True)
                ''Call awinScrollintoView()
                'appInstance.ScreenUpdating = True

                'Call MsgBox(formerselect & " wurde geladen ...")
            End If

        Else
            Call MsgBox("bitte einen Eintrag selektieren")
        End If

        'DialogResult = System.Windows.Forms.DialogResult.OK


    End Sub
End Class