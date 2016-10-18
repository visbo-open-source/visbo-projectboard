Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmCreateLicences

    Private VisboLic As New clsLicences
    Private clientLic As New clsLicences
    Dim vorhClientlic As New clsLicences
    Dim vorhVisbolic As New clsLicences

    Private Sub ListKomponenten_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListKomponenten.SelectedIndexChanged

    End Sub

    Private Sub untilDate_ValueChanged(sender As Object, e As EventArgs) Handles untilDate.ValueChanged

    End Sub
 

    Private Sub frmCreateLicences_FormClosing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.FormClosing

        If Not (clientLic.Liste.Count = 0 And VisboLic.Liste.Count = 0) Then
            If Not MsgBox("Wollen Sie das Formular wirklich schließen ohne das Lizenzfile zu schreiben?", _
                     vbYesNo, "Formular schließen") = vbYes Then

                e.cancel = True
                ' '' '' Lizenzen in XML-Dateien speichern
                '' ''Call XMLExportLicences(VisboLic, requirementsOrdner & "visboLicfile.xml")

                '' ''Call XMLExportLicences(clientLic, licFileName)

                '' ''Call MsgBox("Lizenzen wurden in '" & licFileName & "'" & vbLf & "und '" & requirementsOrdner & "visboLicfile.xml' geschrieben")
                '' ''clientLic.clear()
                '' ''VisboLic.clear()

                '' ''UserName.Visible = True
                '' ''UserName.Enabled = True
                '' ''LabelUser.Visible = True
            Else
                ' Formular wird geschlossen, ohne das Lizenzfile wegzuschreiben
            End If

        End If

    End Sub

    Private Sub frmCreateLicences_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim i As Integer
        ' vorhandene Komponenten laden
        For i = 0 To LizenzKomponenten.Length - 1

            ListKomponenten.Items.Add(LizenzKomponenten(i))

        Next i

        untilDate.Value = Date.Now

        UserName.Text = myWindowsName

        ' einlesen der vorhandenen Lizenzen
        Try

            vorhClientlic = XMLImportLicences(licFileName)
            vorhVisbolic = XMLImportLicences(requirementsOrdner & "visboLicfile.xml")

        Catch ex As Exception

        End Try
    End Sub

    Private Sub UserName_TextChanged(sender As Object, e As EventArgs) Handles UserName.TextChanged

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        ' Lizenzen in XML-Dateien speichern
        Call XMLExportLicences(VisboLic, requirementsOrdner & "visboLicfile.xml")

        Call XMLExportLicences(clientLic, licFileName)

        Call MsgBox("Lizenzen wurden in '" & licFileName & "'" & vbLf & "und '" & requirementsOrdner & "visboLicfile.xml' geschrieben")
        clientLic.clear()
        VisboLic.clear()

        UserName.Visible = True
        UserName.Enabled = True
        LabelUser.Visible = True
    End Sub

    Private Sub AddLicences_Click(sender As Object, e As EventArgs) Handles AddLicences.Click
        Try


            Dim lastrow As Integer
            Dim lastcolumn As Integer
            Dim rowOffset As Integer = 1
            Dim columnOffset As Integer = 1
            Dim UserRange As Excel.Range
            Dim UserDomain As String = ""
            Dim endDate As Date
            Dim UserListFile As Microsoft.Office.Interop.Excel.Workbook
            Dim wsUserList As Microsoft.Office.Interop.Excel.Worksheet


            Dim i As Integer
            Dim k As Integer
            Dim angabenOK As Boolean = True

            Dim komponenten(ListKomponenten.SelectedItems.Count - 1) As String
            For i = 0 To ListKomponenten.SelectedItems.Count - 1
                komponenten(i) = ListKomponenten.SelectedItems(i)
            Next
            If ListKomponenten.SelectedItems.Count < 1 Then
                angabenOK = False
                Call MsgBox("Bitte wählen Sie die Softwarekomponenten aus!")

            Else
                endDate = untilDate.Value
                If DateDiff(DateInterval.Day, Date.Now, endDate) < 0 Then
                    angabenOK = False
                    Call MsgBox("Gültigkeitsdatum muss nach dem heutigen Datum liegen!")
                End If
            End If

            If angabenOK Then

                If Not UserName.Enabled Then
                    ' Lizenzen werden über die Datei UserList.xlsx erzeugt

                    Try
                        If My.Computer.FileSystem.FileExists(FileNameUserList.Text) Then
                            ' Lizenzen werden über die Datei UserList.xlsx geprüft
                            UserListFile = appInstance.Workbooks.Open(FileNameUserList.Text)

                        Else
                            Dim message As String = "Datei " & FileNameUserList.Text & " existiert nicht auf diesem Computer"
                            Call MsgBox(message)
                            Throw New ArgumentException(message)
                        End If


                    Catch ex As Exception
                        Throw New ArgumentException(ex.Message & "Fehler beim Öffnen der Eingabedatei:" & FileNameUserList.Text)

                    End Try

                    wsUserList = CType(UserListFile.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)

                    With wsUserList

                        lastrow = CInt(CType(.Cells(2000, columnOffset), Excel.Range).End(Excel.XlDirection.xlUp).Row)
                        lastcolumn = CInt(CType(.Cells(rowOffset, 2000), Excel.Range).End(Excel.XlDirection.xlToLeft).Column)

                        UserRange = .Range(.Cells(rowOffset, columnOffset + 2), .Cells(lastrow, columnOffset + 2))
                        For Each zelle In UserRange

                            UserDomain = CType(CType(zelle, Excel.Range).Value, String).Trim

                            ' Licensen erzeugen und in die Liste aufnehmen
                            For k = 0 To komponenten.Length - 1

                                ' Lizenzkey berechnen
                                Dim licString As String = VisboLic.berechneKey(endDate, UserDomain, komponenten(k))

                                ' VsisboListe mit Angabe von username, komponente, endDate
                                Dim visbokey As String = UserDomain & "-" & komponenten(k) & "-" & endDate.ToString
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

                        Next

                    End With
                    Call MsgBox(" Die Lizenzen der Datei '" & FileNameUserList.Text & "' wurden hinzugefügt! ")

                    ' Eingabefile wieder schließen
                    UserListFile.Close(SaveChanges:=False)
                Else

                    If UserName.Text.Length < 1 Then
                        Call MsgBox("Bitte geben Sie den User (mit Domain) an, für den die Lizenz ausgestellt werden soll !")
                    Else

                        UserDomain = UserName.Text

                        ' Licensen erzeugen und in die Liste aufnehmen
                        For k = 0 To komponenten.Length - 1

                            ' Lizenzkey berechnen
                            Dim licString As String = VisboLic.berechneKey(endDate, UserDomain, komponenten(k))

                            ' VsisboListe mit Angabe von username, komponente, endDate
                            Dim visbokey As String = UserDomain & "-" & komponenten(k) & "-" & endDate.ToString
                            If VisboLic.Liste.ContainsKey(visbokey) Then
                                Dim ok As Boolean = VisboLic.Liste.Remove(visbokey)
                            End If
                            VisboLic.Liste.Add(visbokey, licString)

                            ' Liste von Lizenzen für den Kunden 
                            If clientLic.Liste.ContainsKey(licString) Then
                                Dim ok As Boolean = clientLic.Liste.Remove(licString)
                            End If
                            clientLic.Liste.Add(licString, licString)

                            Call MsgBox("License: " & licString & " wurde hinzugefügt!")

                        Next k               'nächste Komponente

                    End If

                End If     ' ende von If UserName.enabled

            End If   ' Ende if angabenOK



            FileNameUserList.Text = ""
            UserName.Visible = True
            UserName.Enabled = True
            LabelUser.Visible = True
        Catch ex As Exception
            Call MsgBox("Fehler beim Hinzufügen der Lizenzen " & vbLf & ex.Message)
        End Try
    End Sub

    Private Sub FileNameUserList_TextChanged(sender As Object, e As EventArgs) Handles FileNameUserList.TextChanged
        Try
            If My.Computer.FileSystem.FileExists(FileNameUserList.Text) Then
                UserName.Visible = False
                UserName.Enabled = False
                LabelUser.Visible = False
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles SaveButton.Click
        Try

            ' neue VISBOLicensen werden um alte erweitert
            For Each kvp As KeyValuePair(Of String, String) In vorhVisbolic.Liste

                If VisboLic.Liste.ContainsKey(kvp.Key) Then
                    Dim ok As Boolean = VisboLic.Liste.Remove(kvp.Key)
                End If
                VisboLic.Liste.Add(kvp.Key, kvp.Value)
            Next

            ' neue ClientLicensen werden um alte erweitert
            For Each kvp As KeyValuePair(Of String, String) In vorhClientlic.Liste

                If clientLic.Liste.ContainsKey(kvp.Key) Then
                    Dim ok As Boolean = clientLic.Liste.Remove(kvp.Key)
                End If
                clientLic.Liste.Add(kvp.Key, kvp.Value)
            Next

            ' Lizenzen in XML-Dateien speichern
            Call XMLExportLicences(VisboLic, requirementsOrdner & "visboLicfile.xml")

            Call XMLExportLicences(clientLic, licFileName)

            Call MsgBox("Lizenzen wurden in '" & licFileName & "'" & vbLf & "und '" & requirementsOrdner & "visboLicfile.xml' geschrieben")
            clientLic.clear()
            VisboLic.clear()

            UserName.Visible = True
            UserName.Enabled = True
            LabelUser.Visible = True

        Catch ex As Exception
            Call MsgBox("Fehler beim Erzeugen der Lizenzfiles " & vbLf & ex.Message)
        End Try
    End Sub
End Class