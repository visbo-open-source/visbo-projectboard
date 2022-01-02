Imports System.Windows.Forms

Public Class frmSelectOneItem

    Public itemsCollection As New List(Of String)
    Private Sub frmSelectOneItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)

        Call languageSettings()

        For Each itemName As String In itemsCollection
            itemList.Items.Add(itemName)
        Next

        ' Den ersten eintrag by default selektieren 
        If itemList.Items.Count > 0 Then
            itemList.SelectedIndex = 0
        End If

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

    Private Sub frmSelectOneItem_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try
    End Sub
End Class