Imports System.ComponentModel

Public Class frmAddOrDeleteALine

    Public position As Excel.Range
    Private frmtop As Integer
    Private frmleft As Integer
    Public addLine As Boolean
    Public deleteLine As Boolean
    Public isRoleSkill As Boolean
    Public isCost As Boolean
    Public isEmpty As Boolean
    Public enableDeleteLine As Boolean

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        addLine = False
        deleteLine = False
        isRoleSkill = False
        isCost = False
        isEmpty = False
        enableDeleteLine = False

    End Sub

    Private Sub frmAddOrDeleteALine_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Top = 45
        Me.Left = Me.Left - 400


        Try
            Me.AddALine.Text = "Add empty row"
            If isRoleSkill Then
                Me.DeleteALine.Text = "Delete resource"
            ElseIf isCost Then
                Me.DeleteALine.Text = "Delete cost"
            End If

            If isEmpty Then
                Me.DeleteALine.Text = "Delete row"
            End If

            Me.DeleteALine.Enabled = enableDeleteLine

        Catch ex As Exception

        End Try


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