Public Class frmLoadConstellation

    Private formerselect As String
    Public retrieveFromDB As Boolean
    Public earliestDate As Date
    Public constellationsToShow As clsConstellations
    Private Sub frmLoadConstellation_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        Call languageSettings()
        For Each kvp As KeyValuePair(Of String, clsConstellation) In constellationsToShow.Liste

            ListBox1.Items.Add(kvp.Key)

        Next
        formerselect = ""

        If Not retrieveFromDB Then
            requiredDate.Visible = False
            lblStandvom.Visible = False
        Else

            'Try

            '    dropBoxTimeStamps.Items.Clear()

            '    For k As Integer = 1 To listOfTimeStamps.Count
            '        Dim tmpDate As Date = CDate(listOfTimeStamps.Item(k))
            '        dropBoxTimeStamps.Items.Add(tmpDate)
            '    Next

            'Catch ex As Exception

            'End Try

            ' jetzt ist dropBoxTimeStamps.selecteditem = Nothing ..
        End If

    End Sub

    Private Sub languageSettings()

        If awinSettings.englishLanguage Then
            lblStandvom.Text = "Version"
            addToSession.Text = "add to session"
            OKButton.Text = "OK"
            Abbrechen.Text = "Cancel"
            loadAsSummary.Text = "load summary project"
        End If

    End Sub

    Private Sub Abbrechen_Click(sender As Object, e As EventArgs) Handles Abbrechen.Click
        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        If ListBox1.SelectedItems.Count >= 1 Then
            DialogResult = System.Windows.Forms.DialogResult.OK
            MyBase.Close()
        Else
            Call MsgBox("bitte einen Eintrag selektieren")
        End If

    End Sub

    Private Sub addToSession_CheckedChanged(sender As Object, e As EventArgs) Handles addToSession.CheckedChanged


    End Sub

    Public Sub New()

        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        retrieveFromDB = False
        constellationsToShow = New clsConstellations

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.

    End Sub

    Private Sub dropBoxTimeStamps_SelectedIndexChanged(sender As Object, e As EventArgs)

        ' den Fokus von diesem Element wegnehmen 
        ListBox1.Focus()
        Try
            ListBox1.SelectedItems.Clear()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub requiredDate_ValueChanged(sender As Object, e As EventArgs) Handles requiredDate.ValueChanged

        If Not IsNothing(requiredDate) Then

            If requiredDate.Value >= earliestDate Then
                requiredDate.Value = requiredDate.Value.Date.AddHours(23).AddMinutes(59)
            Else
                Call MsgBox("es gibt vor dem " & earliestDate.ToShortDateString & " keine Projekte in der Datenbank ")
                requiredDate.Value = Date.Now.Date.AddHours(23).AddMinutes(59)
            End If

        Else
            ' nichts tun ...
        End If

    End Sub
End Class