Imports System.Windows.Forms

Public Class frmStoreConstellation



    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim cName As String

        If IsNothing(ComboBox1.SelectedItem) Then
            cName = ComboBox1.Text
        Else
            cName = ComboBox1.SelectedItem.ToString
        End If

        'Call awinStoreConstellation(cName)
        'Call MsgBox(" jetzt wird die Konstellation " & cName & " geschrieben ...")

        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()
    End Sub

    Private Sub frmStoreConstellation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)

        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste
            If kvp.Key <> "Start" Then
                ComboBox1.Items.Add(kvp.Key)
            End If
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        'Call MsgBox("Text changed")
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub frmStoreConstellation_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try
    End Sub
End Class