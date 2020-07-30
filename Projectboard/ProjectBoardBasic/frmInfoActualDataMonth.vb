Imports ProjectBoardDefinitions
Public Class frmInfoActualDataMonth
    Public Sub MonatJahr_ValueChanged(sender As Object, e As EventArgs) Handles MonatJahr.ValueChanged

    End Sub

    Private Sub okBtn_Click(sender As Object, e As EventArgs) Handles okBtn.Click

    End Sub

    Private Sub cancelBtn_Click(sender As Object, e As EventArgs) Handles cancelBtn.Click

    End Sub

    Private Sub frmInfoActualDataMonth_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' zugelassenen Min und Max-Werte für das Datum setzen 
        MonatJahr.MinDate = StartofCalendar
        MonatJahr.MaxDate = Date.Now

        ' Default setzen 
        ' Vorbesetzung des Datums für Istdaten ist aktuelles Datum
        MonatJahr.Value = Date.Now

        ' jetzt die Referenz-Portfolio Dropbox mit Namen besetzen 
        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
            Dim Err As New clsErrorCodeMsg
            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, Err)

            For Each kvp As KeyValuePair(Of String, String) In dbPortfolioNames
                listOfPortfolioNames.Items.Add(kvp.Key)
            Next

        End If

        Call languageSettings()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            Label1.Text = "Actual data including last month to"
            lbl_refPortfolioName.Text = "Reference-Portfolio"
            okBtn.Text = "Import Data"
        End If


    End Sub
End Class