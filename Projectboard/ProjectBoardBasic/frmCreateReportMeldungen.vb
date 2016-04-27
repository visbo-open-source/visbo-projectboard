Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Excel = Microsoft.Office.Interop.Excel
Public Class frmCreateReportMeldungen


    Private report_DE_messages As New clsReportMessages
    Private report_en_messages As New clsReportMessages
    Private report_fr_messages As New clsReportMessages
    Private report_es_messages As New clsReportMessages


    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click


        Dim lastrow As Integer = 1
        Dim lastcolumn As Integer = 1
        Dim rowOffset As Integer = 1
        Dim columnOffset As Integer = 1
        Dim msgDE As String = ""
        Dim msgEN As String = ""
        Dim msgFR As String = ""
        Dim msgES As String = ""
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

                    'msgDE = CType(.Cells(zeile, columnOffset + PTSprache.deutsch).Value, String).Trim
                    msgDE = CType(.Cells(zeile, columnOffset + PTSprache.deutsch).Value, String)
                    report_DE_messages.Liste.Add(zeile, msgDE)

                End If

                If lastcolumn > 1 Then

                    'msgEN = CType(.Cells(zeile, columnOffset + PTSprache.englisch).Value, String).Trim
                    msgEN = CType(.Cells(zeile, columnOffset + PTSprache.englisch).Value, String)
                    report_en_messages.Liste.Add(zeile, msgEN)
                End If
                If lastcolumn > 2 Then

                    'msgFR = CType(.Cells(zeile, columnOffset + PTSprache.fanzösisch).Value, String).Trim
                    msgFR = CType(.Cells(zeile, columnOffset + PTSprache.französisch).Value, String)
                    report_fr_messages.Liste.Add(zeile, msgFR)
                End If

                If lastcolumn > 3 Then

                    'msgES = CType(.Cells(zeile, columnOffset + PTSprache.spanisch).Value, String).Trim
                    msgES = CType(.Cells(zeile, columnOffset + PTSprache.spanisch).Value, String)
                    report_es_messages.Liste.Add(zeile, msgES)
                End If
            Next zeile

        End With
        _
        _
        ' Eingabefile wieder schließen
        ReportMessagesFile.Close(SaveChanges:=False)

        Call XMLExportReportMsg(report_DE_messages, repMsgFileName, ReportLang(PTSprache.deutsch).Name)
        Call XMLExportReportMsg(report_en_messages, repMsgFileName, ReportLang(PTSprache.englisch).Name)
        Call XMLExportReportMsg(report_fr_messages, repMsgFileName, ReportLang(PTSprache.französisch).Name)
        Call XMLExportReportMsg(report_es_messages, repMsgFileName, ReportLang(PTSprache.spanisch).Name)

        Call MsgBox(" Die Message der Datei '" & FileReportMessages.Text & "' wurden eingelesen " & vbLf & _
                    " und übersetzt!")

        Me.Close()

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