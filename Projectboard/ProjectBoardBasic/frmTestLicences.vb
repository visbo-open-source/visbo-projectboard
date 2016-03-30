Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Public Class frmTestLicences

    Private VisboLic As New clsLicences
    Private clientLic As New clsLicences

    Private Sub frmTestLicences_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            Me.statusLabel.Visible = False

            clientLic = XMLImportLicences(licFileName)

            Dim i As Integer
            ' vorhandene Komponenten laden
            For i = 0 To LizenzKomponenten.Length - 1

                ListKomponenten.Items.Add(LizenzKomponenten(i))

            Next i

        Catch ex As Exception
            Me.statusLabel.Text = ex.Message
            Me.statusLabel.Visible = False
        End Try
    End Sub
    Private Sub UserName_TextChanged(sender As Object, e As EventArgs) Handles UserName.TextChanged
        Me.statusLabel.Visible = False
    End Sub

    Private Sub ListKomponenten_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListKomponenten.SelectedIndexChanged
        Me.statusLabel.Visible = False
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
       
        Dim isvalid As Boolean = False
        ''Dim Lic As New clsLicences

        ''Lic = XMLImportLicences(licFileName)

        Dim user As String = UserName.Text
        Dim komponente As String = ListKomponenten.SelectedItem
        Dim testerg As Boolean = False

        isvalid = clientLic.validLicence(user, komponente)

        If isvalid Then

            Me.statusLabel.Text = "Lizenz für User: " & UserName.Text & " und Komponente " & komponente & " ist gültig"
            Me.statusLabel.Visible = True
        Else
            Me.statusLabel.Text = "Fehler:  Lizenz für User: " & UserName.Text & " und Komponente: '" & komponente & "' ist ungültig"
            Me.statusLabel.Visible = True
        End If
    End Sub

End Class