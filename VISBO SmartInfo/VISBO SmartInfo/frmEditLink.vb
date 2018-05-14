Imports ProjectBoardDefinitions
Public Class frmEditLink

    Friend titleExtension As String

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        titleExtension = ""
    End Sub

    Private Sub frmEditLink_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = Me.Text & titleExtension
    End Sub

    Private Sub okBtn_Click(sender As Object, e As EventArgs) Handles okBtn.Click
        If isValidURL(linkValue.Text) Then
            Me.Close()
        Else
            Call MsgBox("invalid Url: " & linkValue.Text)
            linkValue.Text = ""
        End If
    End Sub

    Private Sub clearBtn_Click(sender As Object, e As EventArgs) Handles clearBtn.Click
        linkValue.Text = ""
        Me.Close()
    End Sub
End Class