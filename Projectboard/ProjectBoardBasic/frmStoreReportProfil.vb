Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports System.Windows.Forms

Public Class frmStoreReportProfil

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim cName As String

        If IsNothing(NameReportProfil.SelectedItem) Then
            cName = NameReportProfil.Text
        Else
            cName = NameReportProfil.SelectedItem.ToString
        End If
        ' Beschreibung für Report Profil speichern
        currentReportProfil.description = profilDescription.Text

        If cName <> "" Then

            currentReportProfil.name = cName
            Call XMLExportReportProfil(currentReportProfil)

            DialogResult = System.Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            Call MsgBox("Bitte geben Sie einen Namen für das ReportProfil ein! ")
        End If

    End Sub

    Private Sub AbbruchButton_Click(sender As Object, e As EventArgs) Handles AbbruchButton.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    Private Sub NameReportProfil_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NameReportProfil.SelectedIndexChanged
        Dim hilfsReportProfil As New clsReportAll

        hilfsReportProfil = XMLImportReportProfil(NameReportProfil.SelectedItem.ToString)
        If Not IsNothing(hilfsReportProfil) Then
            profilDescription.Text = hilfsReportProfil.description
        Else
            profilDescription.Text = ""
        End If

    End Sub

    Private Sub frmStoreReportProfil_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dirName As String
        Dim dateiName As String
        Dim profilName As String = ""

        dirName = awinPath & ReportProfileOrdner


        If My.Computer.FileSystem.DirectoryExists(dirName) Then


            Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirName)
            If listOfFiles.Count > 0 Then
                For k As Integer = 1 To listOfFiles.Count

                    dateiName = listOfFiles.Item(k - 1)
                    If dateiName.Contains(".xml") Then

                        Try

                            Dim hstr() As String
                            hstr = Split(dateiName, ".xml", 2)
                            Dim hhstr() As String
                            hhstr = Split(hstr(0), "\")
                            profilName = hhstr(hhstr.Length - 1)
                            NameReportProfil.Items.Add(profilName)

                        Catch ex As Exception

                        End Try

                    End If
                Next k
            End If

        End If
        profilDescription.Text = ""
    End Sub

    Private Sub NameReportProfil_KeyDown(sender As Object, e As KeyEventArgs) Handles NameReportProfil.KeyDown
        
        profilDescription.Text = ""
    End Sub
End Class