Imports System.ComponentModel

Public Class frmAddOrDeleteALine

    Public position As Excel.Range
    Public addLine As Boolean
    Public deleteLine As Boolean
    Private Sub frmAddOrDeleteALine_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'aktuell = (Selection.ColumnWidth + 0.71) / 5.1425

        'aktuell = Selection.RowHeight / 29.5
        'Selection.RowHeight = hoehe * 29.5
        'Position für Formular bestimmen
        Dim cw As Double = position.ColumnWidth

        Me.Top = position.Top + 200
        Me.Left = position.Left + position.Width + cw + 20


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
        Dim top As Double = Me.Top
        Dim left As Double = Me.Left

    End Sub


End Class