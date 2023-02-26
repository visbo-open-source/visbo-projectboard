Imports System.ComponentModel

Public Class frmAddOrDeleteALine

    Public position As Excel.Range
    Private frmtop As Integer
    Private frmleft As Integer
    Public addLine As Boolean
    Public deleteLine As Boolean
    Private Sub frmAddOrDeleteALine_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'aktuell = (Selection.ColumnWidth + 0.71) / 5.1425

        'aktuell = Selection.RowHeight / 29.5
        'Selection.RowHeight = hoehe * 29.5
        'Position für Formular bestimmen
        Dim cw As Double = position.ColumnWidth
        Dim startPos As FormStartPosition = MyBase.StartPosition

        'If frmtop = 0 And frmleft = 0 Then
        '    frmtop = Me.Top - 100
        '    frmleft = Me.Left - 500
        'End If

        Me.Top = 45
        Me.Left = Me.Left - 400


    End Sub

    Private Sub AddALine_Click(sender As Object, e As EventArgs) Handles AddALine.Click
        addLine = True
        Me.Close()
    End Sub

    Private Sub DeleteALine_Click(sender As Object, e As EventArgs) Handles DeleteALine.Click
        deleteLine = True
        Me.Close()
    End Sub

    Private Sub frmAddOrDeleteALine_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        frmtop = Me.Top
        frmleft = Me.Left
    End Sub


End Class