Public Class frmCompareConstellation

    Private formerIndex As Integer

    Private Sub frmCompareConstellation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            CheckedListBox1.Items.Add(kvp.Key)

        Next

        formerIndex = -1

    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Dim value As Integer = formerIndex

        If formerIndex <> CheckedListBox1.SelectedIndex Then
            ' eine anderer Eintrag wurde selektiert 
            If formerIndex > 0 Then
                CheckedListBox1.SetItemChecked(formerIndex, False)
            End If
            formerIndex = CheckedListBox1.SelectedIndex
        Else
            ' ein De-Select hat stattgefunden 
            formerIndex = -1
        End If

        If formerIndex <> -1 Then
            Call MsgBox(CheckedListBox1.Text & " wird geladen ...")
        Else
            Call MsgBox("der alte Zustand wird wieder hergestellt ...")
        End If


    End Sub

   
End Class