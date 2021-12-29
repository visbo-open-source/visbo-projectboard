Imports System.Windows.Forms

Public Class frmconfirmDeletePrj

    Public botschaft As String
    Private Sub OK_Button_Click(sender As System.Object, e As System.EventArgs) Handles OK_Button.Click

        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()

    End Sub

    Private Sub Cancel_Button_Click(sender As System.Object, e As System.EventArgs) Handles Cancel_Button.Click

        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub


    Private Sub frmconfirmDeletePrj_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Call getFrmPosition(PTfrm.other, Top, Left)
        Label1.Text = botschaft

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub frmconfirmDeletePrj_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try
    End Sub
End Class