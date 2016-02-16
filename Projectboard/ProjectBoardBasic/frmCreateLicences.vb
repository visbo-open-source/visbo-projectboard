Imports ProjectBoardDefinitions
Imports ProjectBoardBasic

Public Class frmCreateLicences

    Private VisboLic As New clsLicences
    Private clientLic As New clsLicences

    Private Sub ListKomponenten_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListKomponenten.SelectedIndexChanged

    End Sub

    Private Sub untilDate_ValueChanged(sender As Object, e As EventArgs) Handles untilDate.ValueChanged

    End Sub


    Private Sub frmCreateLicences_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim i As Integer
        ' vorhandene Komponenten laden
        For i = 0 To LizenzKomponenten.Length - 1

            ListKomponenten.Items.Add(LizenzKomponenten(i))

        Next i


        untilDate.Value = Date.Now

    End Sub

    Private Sub UserName_TextChanged(sender As Object, e As EventArgs) Handles UserName.TextChanged

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        ' Lizenzen in XML-Dateien speichern
        Call XMLExportLicences(VisboLic, requirementsOrdner & "visboLicfile.xml")

        Call XMLExportLicences(clientLic, licFileName)
    End Sub

    Private Sub AddLicences_Click(sender As Object, e As EventArgs) Handles AddLicences.Click

        Dim i As Integer
        Dim k As Integer

        Dim benutzer As String = UserName.Text
        If benutzer = "" Then
            Call MsgBox("Username muss angegeben werden!")
        Else

            Dim komponenten(ListKomponenten.SelectedItems.Count - 1) As String

            For i = 0 To ListKomponenten.SelectedItems.Count - 1
                komponenten(i) = ListKomponenten.SelectedItems(i)
            Next
            If ListKomponenten.SelectedItems.Count < 1 Then
                Call MsgBox("Bitte wählen Sie die Softwarekomponenten aus!")
            Else

                Dim endDate As Date = untilDate.Value
                If DateDiff(DateInterval.Day, Date.Now, endDate) < 0 Then
                    Call MsgBox("Gültigkeitsdatum muss nach dem heutigen Datum liegen!")
                Else

                    ' Licensen erzeugen und in die Liste aufnehmen
                    For k = 0 To komponenten.Length - 1

                        ' Lizenzkey berechnen
                        Dim licString As String = VisboLic.berechneKey(endDate, benutzer, komponenten(k))

                        ' VsisboListe mit Angabe von username, komponente, endDate
                        Dim visbokey As String = benutzer & "-" & komponenten(k) & "-" & endDate.ToString
                        If VisboLic.Liste.ContainsKey(visbokey) Then
                            Dim ok As Boolean = VisboLic.Liste.Remove(visbokey)
                        End If
                        VisboLic.Liste.Add(visbokey, licString)

                        ' Liste von Lizenzen für den Kunden 
                        If clientLic.Liste.ContainsKey(licString) Then
                            Dim ok As Boolean = clientLic.Liste.Remove(licString)
                        End If
                        clientLic.Liste.Add(licString, licString)

                    Next k               'nächste Komponente

                End If

            End If

        End If
    End Sub
End Class