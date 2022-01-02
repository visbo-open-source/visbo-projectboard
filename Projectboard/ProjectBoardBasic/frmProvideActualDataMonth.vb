Imports System.Windows.Forms
Imports ProjectBoardDefinitions
Public Class frmProvideActualDataMonth


    Private Sub frmProvideActualDataMonth_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call getFrmPosition(PTfrm.other, Top, Left)

        Dim tmpMonth As Integer = Date.Now.Month - 1

        If tmpMonth < 1 Then
            tmpMonth = 12
        End If

        valueMonth.Text = tmpMonth.ToString("#0")


    End Sub

    Private Sub cancelBtn_Click(sender As Object, e As EventArgs) Handles cancelBtn.Click

    End Sub

    Private Sub frmProvideActualDataMonth_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Try
            frmCoord(PTfrm.other, PTpinfo.top) = Me.Top
            frmCoord(PTfrm.other, PTpinfo.left) = Me.Left
        Catch ex As Exception

        End Try
    End Sub
End Class