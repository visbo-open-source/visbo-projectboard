Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports ClassLibrary1
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Public Class frmMEhryRoleCost


    Private Sub frmMEhryRoleCost_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If frmCoord(PTfrm.rolecostME, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.rolecostME, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

    End Sub

    Private Sub frmMEhryRoleCost_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        frmCoord(PTfrm.rolecostME, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.rolecostME, PTpinfo.left) = Me.Left
    End Sub
End Class