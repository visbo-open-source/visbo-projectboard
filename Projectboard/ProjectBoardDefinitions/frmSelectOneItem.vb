Public Class frmSelectOneItem

    Public itemsCollection As New Collection
    Private Sub frmSelectOneItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call languageSettings()

        For Each itemName As String In itemsCollection
            itemList.Items.Add(itemName)
        Next

        ' Den ersten eintrag by default selektieren 
        itemList.SelectedIndex = 0

    End Sub

    Private Sub languageSettings()
        If awinSettings.englishLanguage Then
            Me.Text = "choose Visbo Center"
        Else
            Me.Text = "Visbo Center wählen"
        End If
    End Sub

    Private Sub itemList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles itemList.SelectedIndexChanged

    End Sub
End Class