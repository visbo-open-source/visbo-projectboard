
Imports ProjectBoardDefinitions
Imports xlNS = Microsoft.Office.Interop.Excel



Public Class frmLoadCockpit

    Private xlsCockpits As xlNS.Workbook = Nothing

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim cName As String
        If ListBox1.Text <> "" Then
            If IsNothing(ListBox1.SelectedItem) Then
                cName = ListBox1.Text
            Else
                cName = ListBox1.SelectedItem.ToString
            End If
            DialogResult = System.Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            Call MsgBox("bitte einen Eintrag selektieren")
        End If

    End Sub

    Private Sub frmStoreCockpit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer
        Dim fileName As String

        Dim wsSheet As xlNS.Worksheet = Nothing

        fileName = awinPath & cockpitsFile

        If My.Computer.FileSystem.FileExists(fileName) Then

            Try

                xlsCockpits = appInstance.Workbooks.Open(fileName)

            Catch ex As Exception

                i = 1
                While i <= appInstance.Workbooks.Count
                    If appInstance.Workbooks(i).Name = fileName Then
                        xlsCockpits = appInstance.Workbooks(i)
                    Else
                        i = i + 1
                    End If
                End While

                'If Not fileIsOpen Then
                '    logMessage = "Öffnen von " & fileName & " fehlgeschlagen" & vbLf & _
                '                                "falls die Datei bereits geöffnet ist: Schließen Sie sie bitte"

                '    Throw New ArgumentException(logMessage)
                'End If

            End Try
        Else
            Throw New ArgumentException("Die Datei " & fileName & " existiert nicht.")
        End If

        ' alle vorhandenen Cockpits (=Tabellenblätter) zur Auswahl anzeigen

        i = 1
        While i <= xlsCockpits.Worksheets.Count
            wsSheet = xlsCockpits.Worksheets.Item(i)
            ListBox1.Items.Add(wsSheet.Name)
            i = i + 1
        End While

        xlsCockpits.Close(SaveChanges:=False)

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As EventArgs) Handles ListBox1.TextChanged
        'Call MsgBox("Text changed")
    End Sub

End Class