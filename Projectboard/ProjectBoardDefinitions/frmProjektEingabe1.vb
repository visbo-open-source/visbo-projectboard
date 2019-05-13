Imports System.Windows.Forms

Public Class frmProjektEingabe1


    ' notwendig, weil sonst eine Fehlermeldung kommt bezgl ValueChanged und zugelassenen Werten 
    Private dauerVorlage As Integer = 365
    Private listOFMilestones As New SortedList(Of Date, String)
    Private startMsOffset As Integer = 0
    Private endMsOffset As Integer = 0
    Private vproj As clsProjektvorlage

    Private dontFire As Boolean = False

    Public calcProjektStart As Date = Date.Now
    Public calcProjektEnde As Date = Date.Now.AddMonths(6)
    Public newProjektDauer As Integer = 0

    Private Sub frmProjektEingabe1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        frmCoord(PTfrm.eingabeProj, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.eingabeProj, PTpinfo.left) = Me.Left


    End Sub

    Private Sub defineButtonVisibility()
        With Me
            ' Sprach-Einstellungen ...
            If awinSettings.englishLanguage Then
                .Text = "create a new project"
                .lbl_pName.Text = "Project-Name"
                .lblVorlage.Text = "Template"
                .lbl_Number.Text = "Number"
                .lbl_Description.Text = "Goals"
                .lblProfitField.Text = "Margin(%)"
                .dauerUnverändert.Text = "duration like template"
                .lbl_Laufzeit.Text = "Duration: "
                .lbl_Referenz1.Text = "Milestone 1"
                .lbl_Referenz2.Text = "Milestone 2"
                .AbbrButton.Text = "Cancel"
            Else
                ' Texte sind bereits deutsch im Formular hinterlegt ... 
            End If

            ' Sichtbarkeit und Voreinstellungen 
            .lbl_Referenz2.Visible = False
            .endMilestoneDropbox.Visible = False
            .txtbx_pNr.Text = ""

        End With
    End Sub

    Private Sub frmProjektEingabe1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call defineButtonVisibility()



        With Me

            dontFire = True

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
                vorlagenDropbox.Text = CStr(vorlagenDropbox.Items(1))
                awinSettings.lastProjektTyp = CStr(vorlagenDropbox.Items(1))
            End If


            Try
                Call setParametersOfVorlage()
            Catch ex As Exception
                Call MsgBox(ex.Message)
                Exit Sub
            End Try

            ' Projekt-Dauer setzen 
            newProjektDauer = dauerVorlage



            ' Jetzt den Wert für den Erlös bestimmen 

            Dim hvalue As Integer = 0
            Try
                hvalue = CType(System.Math.Round(vproj.getGesamtKostenBedarf.Sum / 10,
                                                                     mode:=MidpointRounding.ToEven) * 10, Integer)
            Catch ex As Exception

            End Try

            .Erloes.Text = hvalue.ToString("N0")


            .txtbx_description.Text = ""
            .dauerUnverändert.Checked = True
            .calcProjektStart = Date.Now.AddMonths(1)
            .calcProjektEnde = .calcProjektStart.AddDays(dauerVorlage - 1)

            .DateTimeStart.Value = .calcProjektStart
            .DateTimeEnde.Value = .calcProjektEnde


            .lbl_Laufzeit.Text = "Laufzeit von " & calcProjektStart.ToShortDateString & " - " &
                                    calcProjektEnde.ToShortDateString

            .lbl_Referenz1.Text = "Referenz"
            .lbl_Referenz2.Text = "Referenz 2"
            .lbl_Referenz2.Visible = False
            .endMilestoneDropbox.Visible = False
            .DateTimeEnde.Visible = False

            ' das Formular an die letzte / Default-Position setzen 
            .Top = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.top))
            .Left = CInt(frmCoord(PTfrm.eingabeProj, PTpinfo.left))

            dontFire = False

        End With
    End Sub



    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim msgtxt As String = ""

        With projectName


            ' Änderung tk 1.7.14: andernfalls kann ein Blank am Ende angehängt sein - dann kommt es im Nachgang zu einem Fehler 
            Try
                .Text = .Text.Trim
            Catch ex As Exception
                .Text = ""
            End Try


            If Len(.Text) < 2 Then

                If awinSettings.englishLanguage Then
                    msgtxt = "Projektname has to be at least 2 characters!"
                Else
                    msgtxt = "Projektname muss mindestens zwei Zeichen haben!"
                End If
                Call MsgBox(msgtxt)
                .Text = ""
                .Undo()
                DialogResult = System.Windows.Forms.DialogResult.None

            Else
                If IsNumeric(.Text) Then
                    If awinSettings.englishLanguage Then
                        msgtxt = "numbers are not permitted as projectnames"
                    Else
                        msgtxt = "Zahlen sind nicht zugelassen"
                    End If
                    Call MsgBox(msgtxt)

                    .Text = ""
                    .Undo()
                    DialogResult = System.Windows.Forms.DialogResult.None

                ElseIf inProjektliste(.Text) Then
                    If awinSettings.englishLanguage Then
                        msgtxt = "projectname does already exist"
                    Else
                        msgtxt = "Projekt-Name bereits vorhanden !"
                    End If
                    Call MsgBox(msgtxt)

                    .Text = ""
                    .Undo()
                    DialogResult = System.Windows.Forms.DialogResult.None
                ElseIf Not isValidProjectName(.Text) Then
                    If awinSettings.englishLanguage Then
                        msgtxt = "projectname must not contain any special characters"
                    Else
                        msgtxt = "Der Projekt-Name darf keine Sonderzeichen Zeichen enthalten"
                    End If
                    Call MsgBox(msgtxt)

                    .Text = ""
                    .Undo()
                    DialogResult = System.Windows.Forms.DialogResult.None
                Else


                    If Not AlleProjekte.containsPNr(txtbx_pNr.Text) Then
                        DialogResult = System.Windows.Forms.DialogResult.OK
                        MyBase.Close()
                    Else
                        If awinSettings.englishLanguage Then
                            msgtxt = "Project Nr already exists ... "
                        Else
                            msgtxt = "Die Projekt-Nummer existiert bereits ..."
                        End If

                        Call MsgBox(msgtxt)

                        txtbx_pNr.Text = ""
                        txtbx_pNr.Undo()
                        DialogResult = System.Windows.Forms.DialogResult.None

                    End If


                End If
            End If
        End With

    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

        'DialogResult = System.Windows.Forms.DialogResult.Cancel
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


    Private Sub vorlagenDropbox_LostFocus(sender As Object, e As EventArgs) Handles vorlagenDropbox.LostFocus

        If Projektvorlagen.Liste.ContainsKey(vorlagenDropbox.Text) Then
            awinSettings.lastProjektTyp = vorlagenDropbox.Text

            Dim hvalue As Integer
            Try
                hvalue = CType(System.Math.Round(Projektvorlagen.getProject(vorlagenDropbox.Text).getGesamtKostenBedarf.Sum / 10,
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


        Dim oldVorlagenDauer As Integer = dauerVorlage
        Dim diff As Integer

        Try
            vproj = Projektvorlagen.getProject(vorlagenDropbox.SelectedIndex)
            dauerVorlage = vproj.dauerInDays
            diff = dauerVorlage - oldVorlagenDauer


        Catch ex As Exception
            Call MsgBox("Vorlagen Dauer konnte nicht bestimmt werden ...")
        End Try

        Call setParametersOfVorlage()


        If dauerUnverändert.Checked Then
            'StartDatum muss gemäß Vorlagendauer errechnet werden

            calcProjektStart = DateTimeStart.Value.AddDays(-1 * startMsOffset)
            calcProjektEnde = calcProjektStart.AddDays(dauerVorlage - 1)

            DateTimeEnde.Value = calcProjektStart.AddDays(endMsOffset)


        Else

            calcProjektStart = DateTimeStart.Value.AddDays(-1 * startMsOffset * faktorfuerDauer)
            calcProjektEnde = calcProjektStart.AddDays((dauerVorlage - 1) * faktorfuerDauer)

            DateTimeEnde.Value = calcProjektStart.AddDays(endMsOffset)

        End If

        lbl_Laufzeit.Text = "Laufzeit von " & calcProjektStart.ToShortDateString & " - " &
                                    calcProjektEnde.ToShortDateString

        Dim hvalue As Integer = 0
        Try
            hvalue = CType(System.Math.Round(vproj.getGesamtKostenBedarf.Sum / 10,
                                                                 mode:=MidpointRounding.ToEven) * 10, Integer)
        Catch ex As Exception

        End Try

        Erloes.Text = hvalue.ToString("N0")

    End Sub




    Private Sub DateTimeEnde_ValueChanged(sender As Object, e As EventArgs) Handles DateTimeEnde.ValueChanged

        If dontFire Then
            ' nichts tun 
        Else
            If dauerUnverändert.Checked Then
                'StartDatum muss gemäß Vorlagendauer errechnet werden
                If DateDiff(DateInterval.Month, StartofCalendar, DateTimeEnde.Value) < 0 Or DateDiff(DateInterval.Month, DateTimeStart.Value, DateTimeEnde.Value) < 0 Then
                    Call MsgBox("Ende-Datum kann nicht vor dem Start des Projekt-Tafel Kalenders" & vbLf & "und nicht vor dem Start des Projektes liegen ...")
                    DateTimeEnde.Value = DateTimeStart.Value.AddDays(dauerVorlage - 1)
                Else
                    calcProjektEnde = DateTimeEnde.Value.AddDays(dauerVorlage - 1 - endMsOffset)
                    calcProjektStart = calcProjektEnde.AddDays(-1 * (dauerVorlage - 1))

                    DateTimeStart.Value = calcProjektStart.AddDays(startMsOffset)

                End If
            Else
                If DateDiff(DateInterval.Month, StartofCalendar, DateTimeEnde.Value) < 0 Or DateDiff(DateInterval.Month, DateTimeStart.Value, DateTimeEnde.Value) < 0 Then

                    Call MsgBox("Ende-Datum kann nicht vor dem Start des Projekt-Tafel Kalenders" & vbLf & "und nicht vor dem Start des Projektes liegen ...")
                    DateTimeEnde.Value = DateTimeStart.Value.AddMonths(6)

                Else
                    calcProjektEnde = DateTimeEnde.Value.AddDays((dauerVorlage - 1 - endMsOffset) * faktorfuerDauer)
                    calcProjektStart = calcProjektEnde.AddDays(-1 * (dauerVorlage - 1) * faktorfuerDauer)

                End If
            End If

            lbl_Laufzeit.Text = "Laufzeit von " & calcProjektStart.ToShortDateString & " - " &
                                        calcProjektEnde.ToShortDateString
        End If


    End Sub

    Private Sub DateTimeStart_ValueChanged(sender As Object, e As EventArgs) Handles DateTimeStart.ValueChanged

        If dontFire Then
            ' nichts tun
        Else
            If dauerUnverändert.Checked Then
                'StartDatum muss gemäß Vorlagendauer errechnet werden
                If DateDiff(DateInterval.Month, StartofCalendar, DateTimeStart.Value) < 0 Then
                    Call MsgBox("Start-Datum kann nicht vor dem Start des Projekt-Tafel Kalenders liegen ...")
                    DateTimeStart.Value = Date.Now.AddMonths(1)
                Else
                    calcProjektStart = DateTimeStart.Value.AddDays(-1 * startMsOffset)
                    calcProjektEnde = calcProjektStart.AddDays(dauerVorlage - 1)

                    DateTimeEnde.Value = calcProjektStart.AddDays(endMsOffset)


                End If
            Else
                If DateDiff(DateInterval.Month, StartofCalendar, DateTimeStart.Value) < 0 Then
                    Call MsgBox("Start-Datum kann nicht vor dem Start des Projekt-Tafel Kalenders liegen ...")
                    DateTimeStart.Value = Date.Now.AddMonths(1)
                    'DateTimeProject.Value = Date.Now.AddDays(vorlagenDauer - 1).AddMonths(1)
                Else
                    calcProjektStart = DateTimeStart.Value.AddDays(-1 * startMsOffset * faktorfuerDauer)
                    calcProjektEnde = calcProjektStart.AddDays((dauerVorlage - 1) * faktorfuerDauer)
                End If
            End If

            lbl_Laufzeit.Text = "Laufzeit von " & calcProjektStart.ToShortDateString & " - " &
                                        calcProjektEnde.ToShortDateString
        End If


    End Sub

    Private Sub dauerUnverändert_CheckedChanged(sender As Object, e As EventArgs) Handles dauerUnverändert.CheckedChanged

        If dontFire Then
            ' nichts tun
        Else
            If dauerUnverändert.Checked Then

                If awinSettings.englishLanguage Then
                    lbl_Referenz1.Text = "Reference"
                Else
                    lbl_Referenz1.Text = "Referenz"
                End If

                lbl_Referenz2.Visible = False
                endMilestoneDropbox.Visible = False
                DateTimeEnde.Visible = False

                calcProjektStart = DateTimeStart.Value.AddDays(-1 * startMsOffset)
                calcProjektEnde = calcProjektStart.AddDays(dauerVorlage - 1)

                DateTimeEnde.Value = calcProjektStart.AddDays(endMsOffset)
                calcProjektEnde = calcProjektStart.AddDays(dauerVorlage - 1)
            Else
                If awinSettings.englishLanguage Then
                    lbl_Referenz1.Text = "Reference 1"
                    lbl_Referenz2.Text = "Reference 2"
                Else
                    lbl_Referenz1.Text = "Referenz 1"
                    lbl_Referenz2.Text = "Referenz 2"
                End If

                lbl_Referenz2.Visible = True
                endMilestoneDropbox.Visible = True
                DateTimeEnde.Visible = True

            End If

            If awinSettings.englishLanguage Then
                lbl_Laufzeit.Text = "Duration: " & calcProjektStart.ToShortDateString & " - " &
                                        calcProjektEnde.ToShortDateString
            Else
                lbl_Laufzeit.Text = "Laufzeit von " & calcProjektStart.ToShortDateString & " - " &
                                        calcProjektEnde.ToShortDateString
            End If

        End If


    End Sub

    Public Sub New()

        ' This call is required by the designer.
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        InitializeComponent()

        appInstance.EnableEvents = formerEE

        ' Add any initialization after the InitializeComponent() call.

    End Sub



    Private Sub startMilestoneDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles startMilestoneDropbox.SelectedIndexChanged


        If dontFire Then
            ' nichts tun
        Else
            If startMilestoneDropbox.Text = "Projektstart" Then
                startMsOffset = 0
            Else
                startMsOffset = CInt(vproj.getMilestoneOffsetToProjectStart(startMilestoneDropbox.Text))
            End If

            If dauerUnverändert.Checked Then
                calcProjektStart = DateTimeStart.Value.AddDays(-1 * startMsOffset)
                calcProjektEnde = calcProjektStart.AddDays(dauerVorlage - 1)
                DateTimeEnde.Value = calcProjektStart.AddDays(endMsOffset)
            Else
                calcProjektStart = DateTimeStart.Value.AddDays(-1 * startMsOffset * faktorfuerDauer)
                calcProjektEnde = calcProjektStart.AddDays((dauerVorlage - 1) * faktorfuerDauer)
                DateTimeEnde.Value = calcProjektStart.AddDays(endMsOffset * faktorfuerDauer)
            End If

            lbl_Laufzeit.Text = "Laufzeit von " & calcProjektStart.ToShortDateString & " - " &
                                        calcProjektEnde.ToShortDateString

        End If

    End Sub

    Private Sub endMilestoneDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles endMilestoneDropbox.SelectedIndexChanged

        If dontFire Then
            ' nichts tun
        Else
            If endMilestoneDropbox.Text = "Projektende" Then
                endMsOffset = dauerVorlage - 1
            Else
                endMsOffset = CInt(vproj.getMilestoneOffsetToProjectStart(endMilestoneDropbox.Text))
            End If

            If dauerUnverändert.Checked Then
                calcProjektEnde = DateTimeEnde.Value.AddDays(dauerVorlage - 1 - endMsOffset)
                calcProjektStart = calcProjektEnde.AddDays(-1 * (dauerVorlage - 1))

                DateTimeStart.Value = calcProjektStart.AddDays(startMsOffset)
            Else
                calcProjektEnde = DateTimeEnde.Value.AddDays((dauerVorlage - 1 - endMsOffset) * faktorfuerDauer)
                calcProjektStart = calcProjektEnde.AddDays(-1 * (dauerVorlage - 1) * faktorfuerDauer)

            End If

            lbl_Laufzeit.Text = "Laufzeit von " & calcProjektStart.ToShortDateString & " - " &
                                        calcProjektEnde.ToShortDateString

        End If

    End Sub

    ''' <summary>
    ''' bestimmt den Abstand in Tagen zwischen Start-Meilenstein und Ende-Meilenstein in der Vorlage
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property abschnittsDauerVorlage As Integer
        Get
            abschnittsDauerVorlage = endMsOffset - startMsOffset + 1
        End Get
    End Property

    ''' <summary>
    ''' betimmt den Abstand in Tagen zwischen Start- und Ende-Datum
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property abschnittsDauerNeu As Integer
        Get
            abschnittsDauerNeu = CInt(DateDiff(DateInterval.Day, CDate(DateTimeStart.Text), CDate(DateTimeEnde.Text)))
        End Get
    End Property

    Private ReadOnly Property faktorfuerDauer As Double
        Get
            faktorfuerDauer = abschnittsDauerNeu / abschnittsDauerVorlage
        End Get
    End Property

    Private Sub businessUnitDropBox_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Erloes_TextChanged(sender As Object, e As EventArgs) Handles Erloes.TextChanged
        If IsNumeric(Erloes.Text) Then
            If CDbl(Erloes.Text) < 0.0 Then
                Call MsgBox("Budget kann nicht negativ sein")
                Erloes.Text = "0"
            End If
        Else
            Call MsgBox("Budget muss eine positive Zahl sein ")
            Erloes.Text = "0"
        End If
    End Sub



    ''' <summary>
    ''' wird aufgerufen, wenn die Vorlage wechselt: dann muss die Meilenstein Liste neu aufgebaut werden  
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub setParametersOfVorlage()

        ' jetzt das Vorlagen Projekt bestimmen 
        vproj = Projektvorlagen.getProject(vorlagenDropbox.Text)


        If vproj.getSummeKosten > 0 Then
            Me.lblProfitField.Visible = True
            Me.profitAskedFor.Visible = True
            Me.profitAskedFor.Text = "0.0"
        Else
            Me.lblProfitField.Visible = False
            Me.profitAskedFor.Visible = False
            Me.profitAskedFor.Text = "0.0"
        End If

        If IsNothing(vproj) Then
            Throw New ArgumentException("Vorlage" & vorlagenDropbox.Text & " existiert nicht ...")
        End If

        ' jetzt die Dauer der Vorlage bestimmen 
        dauerVorlage = vproj.dauerInDays


        ' jetzt die listOfMilestones bestimmen
        Try
            listOFMilestones = Projektvorlagen.getProject(vorlagenDropbox.Text).getMilestones
        Catch ex As Exception

        End Try

        ' jetzt die Start- und End-Milestone Dropboxen aufbauen 
        startMilestoneDropbox.Items.Clear()
        endMilestoneDropbox.Items.Clear()

        startMilestoneDropbox.Items.Add("Projektstart")
        For Each kvp As KeyValuePair(Of Date, String) In listOFMilestones
            Dim msName As String = elemNameOfElemID(kvp.Value)
            startMilestoneDropbox.Items.Add(msName)
            endMilestoneDropbox.Items.Add(msName)
        Next kvp
        endMilestoneDropbox.Items.Add("Projektende")

        startMilestoneDropbox.Text = "Projektstart"
        endMilestoneDropbox.Text = "Projektende"

        ' die Offsets bestimmen 
        startMsOffset = 0
        endMsOffset = dauerVorlage - 1


    End Sub

    Private Sub txtbx_pNr_TextChanged(sender As Object, e As EventArgs) Handles txtbx_pNr.TextChanged

    End Sub
End Class