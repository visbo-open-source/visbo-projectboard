Imports System.Windows.Forms

Public Class frmconfirmOptimierung

    Private Sub ButtonJA_Click(sender As System.Object, e As System.EventArgs) Handles ButtonJA.Click

        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()

    End Sub

    Private Sub ButtonNEIN_Click(sender As System.Object, e As System.EventArgs) Handles ButtonNEIN.Click

        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    Private Sub frmconfirmOptimierung_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call getFrmPosition(PTfrm.other, Top, Left)
    End Sub

    Private Sub frmconfirmOptimierung_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try
    End Sub
End Class