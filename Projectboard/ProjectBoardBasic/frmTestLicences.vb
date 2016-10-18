Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmTestLicences

    Private VisboNoLic As New clsLicences
    Private clientLic As New clsLicences

    Private Sub frmTestLicences_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        ' löschen der LizenzFehlerliste
        VisboNoLic.clear()
        ' Löschen der Liste mit den Lizenzen in licFileName
        clientLic.clear()
    End Sub



    Private Sub frmTestLicences_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            Me.statusLabel.Visible = False

            clientLic = XMLImportLicences(licFileName)

            Dim i As Integer
            ' vorhandene Komponenten laden
            For i = 0 To LizenzKomponenten.Length - 1

                ListKomponenten.Items.Add(LizenzKomponenten(i))

            Next i

            UserName.Text = myWindowsName

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
        Try

            Dim isvalid As Boolean = False

            Dim lastrow As Integer
            Dim lastcolumn As Integer
            Dim rowOffset As Integer = 1
            Dim columnOffset As Integer = 1
            Dim UserRange As Excel.Range
            Dim UserDomain As String = ""
            Dim UserListFile As Microsoft.Office.Interop.Excel.Workbook
            Dim wsUserList As Microsoft.Office.Interop.Excel.Worksheet
            Dim zelle As Object

            Dim i As Integer
            Dim k As Integer
            Dim angabenOK As Boolean = True

            VisboNoLic.clear()

            Dim komponenten(ListKomponenten.SelectedItems.Count - 1) As String
            For i = 0 To ListKomponenten.SelectedItems.Count - 1
                komponenten(i) = ListKomponenten.SelectedItems(i)
            Next
            If ListKomponenten.SelectedItems.Count < 1 Then

                angabenOK = False
                Call MsgBox("Bitte wählen Sie die Softwarekomponenten aus!")

            End If

            If angabenOK Then

                If Not UserName.Enabled Then
                    Try
                        If My.Computer.FileSystem.FileExists(FileNameUserList.Text) Then
                            Me.statusLabel.Visible = False
                            ' Lizenzen werden über die Datei UserList.xlsx geprüft
                            UserListFile = appInstance.Workbooks.Open(FileNameUserList.Text)

                        Else
                            Me.statusLabel.Text = "Datei " & FileNameUserList.Text & " existiert nicht auf diesem Computer"
                            Me.statusLabel.Visible = True
                            Throw New ArgumentException(Me.statusLabel.Text)
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

                                isvalid = clientLic.validLicence(UserDomain, komponenten(k))

                                If Not isvalid Then

                                    ' VsisboListe mit Angabe von username, komponente, endDate
                                    Dim visbokey As String = UserDomain & "-" & komponenten(k) & "-" & Date.Now.ToString
                                    If VisboNoLic.Liste.ContainsKey(visbokey) Then
                                        Dim ok As Boolean = VisboNoLic.Liste.Remove(visbokey)
                                    End If
                                    VisboNoLic.Liste.Add(visbokey, ": Lizenz nicht vorhanden")

                                End If
                            Next k               'nächste Komponente

                        Next

                    End With
                    If Not VisboNoLic.Liste.Count = 0 Then
                        ' Schreiben der nicht vorhandenen Lizenzen
                        Call XMLExportLicences(VisboNoLic, "LizenzFehler.xml")
                        Me.statusLabel.Text = "Im File " & awinPath & "LizenzFehler.xml sind die fehlerhaften Lizenzen enthalten"
                        Me.statusLabel.Visible = True
                    Else
                        Me.statusLabel.Text = "Alle Lizenzen für die User der Eingabedatei: " & FileNameUserList.Text & " sind korrekt in '" & licFileName & "' enthalten!"
                        Me.statusLabel.Visible = True
                    End If

                    ' Eingabefile wieder schließen
                    UserListFile.Close(SaveChanges:=False)


                Else   ' einzelne Lizenz wird geprüft
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
                End If

            End If

            FileNameUserList.Text = ""
            UserName.Visible = True
            UserName.Enabled = True
            LabelUser.Visible = True

        Catch ex As Exception

        End Try

    End Sub

    Private Sub FileNameUserList_TextChanged(sender As Object, e As EventArgs) Handles FileNameUserList.TextChanged
        Try
            If My.Computer.FileSystem.FileExists(FileNameUserList.Text) Then
                UserName.Text = ""
                UserName.Visible = False
                UserName.Enabled = False
                LabelUser.Visible = False
            Else
                UserName.Visible = True
                UserName.Enabled = True
                LabelUser.Visible = True
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class