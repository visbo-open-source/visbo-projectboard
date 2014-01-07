Imports System.Windows.Forms

Public Class frmProjektEingabe1

    'Private dateIsStart As Boolean = False
    Private vorlagenDauer As Integer
    Public calcProjektStart As Date

    Private Sub frmProjektEingabe1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        frmCoord(PTfrm.eingabeProj, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.eingabeProj, PTpinfo.left) = Me.Left

    End Sub



    Private Sub frmProjektEingabe1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim randomValue As Double
        
        With Me


            With vorlagenDropbox
                For Each kvp As KeyValuePair(Of String, clsProjektvorlage) In Projektvorlagen.Liste
                    If kvp.Key <> "Projekt ohne Vorlage" Then
                        .Items.Add(kvp.Key)
                    End If
                Next kvp
            End With

            If awinSettings.lastProjektTyp <> "" Then
                '
                ' zuletzt gewählten Typ anzeigen
                '
                vorlagenDropbox.Text = awinSettings.lastProjektTyp
            Else
                '
                ' Voreinstellungg auf Projekt-Typ 1
                '
                vorlagenDropbox.Text = vorlagenDropbox.Items(1)
                awinSettings.lastProjektTyp = vorlagenDropbox.Items(1)
            End If


            ' Jetzt den Wert für den Erlös bestimmen 

            Dim hvalue As Integer
            Try
                hvalue = CType(System.Math.Round(Projektvorlagen.getProject(vorlagenDropbox.Text).getGesamtKostenBedarf.Sum / 10, _
                                                                     mode:=MidpointRounding.ToEven) * 10, Integer)
            Catch ex As Exception

            End Try

            Me.Erloes.Text = hvalue.ToString("N0")


            If dateIsStart.Checked Then
                .kennzeichnungDate.Text = "Start"
                .DateTimeProject.Value = Date.Now.AddMonths(1)

            Else
                .kennzeichnungDate.Text = "Ende"
                .DateTimeProject.Value = Date.Now.AddDays(vorlagenDauer).AddMonths(1)

            End If

            Try
                vorlagenDauer = Projektvorlagen.getProject(vorlagenDropbox.SelectedIndex).dauerInDays
            Catch ex As Exception
                vorlagenDauer = Projektvorlagen.getProject(1).dauerInDays
            End Try


            .Top = frmCoord(PTfrm.eingabeProj, PTpinfo.top)
            .Left = frmCoord(PTfrm.eingabeProj, PTpinfo.left)

            '.selectedMonth.Value = DateDiff(DateInterval.Month, StartofCalendar, Date.Now) + 2

            randomValue = appInstance.WorksheetFunction.RandBetween(1, 100) / 10
            .risiko.Text = randomValue.ToString("0.0")
            randomValue = appInstance.WorksheetFunction.RandBetween(1, 100) / 10
            .sFit.Text = randomValue.ToString("0.0")

            .volume.Text = "150"

            '.calcMonth.Text = Date.Now.AddMonths(1).ToString("MMM yy")



        End With
    End Sub

    Private Sub selectedMonth_ValueChanged(sender As Object, e As EventArgs) Handles selectedMonth.ValueChanged

        calcMonth.Text = StartofCalendar.AddMonths(CType(selectedMonth.Value, Integer) - 1).ToString("MMM yy")

    End Sub



    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        With projectName
            If Len(.Text) < 1 Then
                'MsgBox("Name muss mindestens 1 Zeichen lang sein")
                .Text = ""
                .Focus()
                Exit Sub
            ElseIf IsNumeric(.Text) Then
                MsgBox("Zahlen sind nicht zugelassen")
                .Text = ""
                .Focus()
                Exit Sub
            ElseIf inProjektliste(.Text) Then
                MsgBox("Name bereits vorhanden !")
                .Text = ""
                .Focus()
                Exit Sub
            End If
        End With

        If dateIsStart.Checked Then
            calcProjektStart = DateTimeProject.Value
        Else
            'vorlagendauer As Integer = Projektvorlagen.getProject(vorlagenDropbox.SelectedIndex).dauerInDays
            calcProjektStart = DateTimeProject.Value.AddDays(-1 * vorlagenDauer)
        End If

        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()

    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    Private Sub Erloes_LostFocus(sender As Object, e As EventArgs) Handles Erloes.LostFocus

        With Me.Erloes
            If Not IsNumeric(.Text) Then
                MsgBox("bitte eine Zahl eingeben ")
                .Text = ""
                .Focus()
            ElseIf CType(.Text, Double) < 0 Then
                Call MsgBox(" der Erlös muss eine positive Dezimal-Zahl sein")
                .Text = ""
                .Focus()
            End If
        End With

    End Sub


    Private Sub sFit_LostFocus(sender As Object, e As EventArgs) Handles sFit.LostFocus

        With Me.sFit

            If Not IsNumeric(.Text) Then
                MsgBox("bitte eine Zahl zwischen 0. und 10 eingeben ")
                .Text = ""
                .Focus()
            ElseIf CType(.Text, Double) < 0 Or CType(.Text, Double) > 10 Then
                Call MsgBox(" der strategische Fit muss eine positive Dezimal-Zahl zwischen 0. und 10 sein")
                .Text = ""
                .Focus()
            Else
                Dim hfit As Double
                hfit = CType(.Text, Double)
                .Text = hfit.ToString("0.0")
            End If

        End With

    End Sub


    Private Sub risiko_LostFocus(sender As Object, e As EventArgs) Handles risiko.LostFocus

        With Me.risiko

            If Not IsNumeric(.Text) Then
                MsgBox("bitte eine Zahl zwischen 0. und 10 eingeben ")
                .Text = ""
                .Focus()
            ElseIf CType(.Text, Double) < 0 Or CType(.Text, Double) > 10 Then
                Call MsgBox(" das Risiko muss eine positive Dezimal-Zahl zwischen 0. und 10 sein")
                .Text = ""
                .Focus()
            Else
                Dim hrisk As Double
                hrisk = CType(.Text, Double)
                .Text = hrisk.ToString("0.0")
            End If

        End With

    End Sub

    Private Sub vorlagenDropbox_LostFocus(sender As Object, e As EventArgs) Handles vorlagenDropbox.LostFocus

        If Projektvorlagen.Liste.ContainsKey(vorlagenDropbox.Text) Then
            awinSettings.lastProjektTyp = vorlagenDropbox.Text

            Dim hvalue As Integer
            Try
                hvalue = CType(System.Math.Round(Projektvorlagen.getProject(vorlagenDropbox.Text).getGesamtKostenBedarf.Sum / 10, _
                                                     mode:=MidpointRounding.ToEven) * 10, Integer)

            Catch ex As Exception

            End Try

            Me.Erloes.Text = hvalue.ToString("N0")
        Else
            Call MsgBox("Vorlage " & vorlagenDropbox.Text & " nicht vorhanden!")

            With vorlagenDropbox
                .Text = awinSettings.lastProjektTyp
                .Focus()
            End With

        End If

    End Sub


    Private Sub vorlagenDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles vorlagenDropbox.SelectedIndexChanged


        Dim oldVorlagenDauer As Integer = vorlagenDauer
        Dim diff As Integer

        Try
            vorlagenDauer = Projektvorlagen.getProject(vorlagenDropbox.SelectedIndex).dauerInDays
            diff = vorlagenDauer - oldVorlagenDauer

            If Not dateIsStart.Checked Then
                DateTimeProject.Value = DateTimeProject.Value.AddDays(diff)
            End If

        Catch ex As Exception
            Call MsgBox("Vorlagen Dauer konnte nicht bestimmt werden ...")
        End Try


    End Sub



    Private Sub DateTimeProject_ValueChanged(sender As Object, e As EventArgs) Handles DateTimeProject.ValueChanged



        If dateIsStart.Checked Then
            If DateDiff(DateInterval.Month, StartofCalendar, DateTimeProject.Value) < 0 Then
                Call MsgBox("Start-Datum kann nicht vor dem Start des Projekt-Tafel Kalenders liegen ...")
                DateTimeProject.Value = Date.Now.AddMonths(1)
            End If
        Else
            If DateDiff(DateInterval.Month, StartofCalendar.AddDays(vorlagenDauer), DateTimeProject.Value) < 0 Then
                Call MsgBox("Start-Datum kann nicht vor dem Start des Projekt-Tafel Kalenders liegen ...")
                DateTimeProject.Value = Date.Now.AddDays(vorlagenDauer).AddMonths(1)
            End If
        End If


    End Sub

    Private Sub dateIsStart_CheckedChanged(sender As Object, e As EventArgs) Handles dateIsStart.CheckedChanged

        If dateIsStart.Checked Then
            ' es war vorher auf Datum = End-Datum
            kennzeichnungDate.Text = "Start"
            DateTimeProject.Value = DateTimeProject.Value.AddDays(-1 * vorlagenDauer)
        Else
            ' es war vorher auf Datum = Start-Datum
            kennzeichnungDate.Text = "Ende"
            DateTimeProject.Value = DateTimeProject.Value.AddDays(vorlagenDauer)

        End If
    End Sub
End Class