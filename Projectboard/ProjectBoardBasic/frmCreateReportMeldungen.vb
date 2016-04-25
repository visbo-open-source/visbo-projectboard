Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmCreateReportMeldungen


    Private reportMessages_de As New clsReportMessages
    Private reportMessages_en As New clsReportMessages


    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click


        Dim lastrow As Integer = 1
        Dim lastcolumn As Integer = 1
        Dim rowOffset As Integer = 1
        Dim columnOffset As Integer = 1
        Dim msgDE As String = ""
        Dim msgEN As String = ""
        Dim msgRange As Excel.Range

        Dim wsReportMessages As Microsoft.Office.Interop.Excel.Worksheet
        Dim ReportMessagesFile As Microsoft.Office.Interop.Excel.Workbook
        Try
            If My.Computer.FileSystem.FileExists(FileReportMessages.Text) Then
                ' Lizenzen werden über die Datei UserList.xlsx geprüft
                ReportMessagesFile = appInstance.Workbooks.Open(FileReportMessages.Text)

            Else
                Dim message As String = "Datei " & FileReportMessages.Text & " existiert nicht auf diesem Computer"
                Call MsgBox(message)
                Throw New ArgumentException(message)
            End If


        Catch ex As Exception
            Throw New ArgumentException(ex.Message & "Fehler beim Öffnen der Eingabedatei:" & FileReportMessages.Text)

        End Try

        wsReportMessages = CType(ReportMessagesFile.Worksheets(1), Global.Microsoft.Office.Interop.Excel.Worksheet)

        With wsReportMessages

            lastrow = CInt(CType(.Cells(2000, columnOffset), Excel.Range).End(Excel.XlDirection.xlUp).Row)
            lastcolumn = CInt(CType(.Cells(rowOffset, 2000), Excel.Range).End(Excel.XlDirection.xlToLeft).Column)

            msgRange = .Range(.Cells(rowOffset, columnOffset), .Cells(lastrow, lastcolumn))

            For zeile = rowOffset To lastrow

                If lastcolumn > 0 Then

                    msgDE = CType(.Cells(zeile, columnOffset + PTSprache.deutsch).Value, String).Trim

                End If

                If lastcolumn > 1 Then

                    msgEN = CType(.Cells(zeile, columnOffset + PTSprache.englisch).Value, String).Trim

                End If


                reportMessages_de.Liste.Add(zeile, msgDE)
                reportMessages_en.Liste.Add(zeile, msgEN)

            Next zeile

        End With
        Call MsgBox(" Die Message der Datei '" & FileReportMessages.Text & "' wurden eingelesen! ")

        ' Eingabefile wieder schließen
        ReportMessagesFile.Close(SaveChanges:=False)

        Call XMLExportReportMsg(reportMessages_de, repMsgFileName, ReportLang(PTSprache.deutsch))
        Call XMLExportReportMsg(reportMessages_en, repMsgFileName, ReportLang(PTSprache.englisch))

    End Sub

    Private Sub FileReportMessages_TextChanged(sender As Object, e As EventArgs) Handles FileReportMessages.TextChanged

    End Sub

    Private Sub frmCreateReportMeldungen_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            FileReportMessages.Text = "C:\Users\tom\Documents\Visual Studio 2013\Projects\ProjectBoard\Projectboard\ClassLibrary1\My Project\ReportTexte.xlsx"
        Catch ex As Exception

        End Try
    End Sub
End Class