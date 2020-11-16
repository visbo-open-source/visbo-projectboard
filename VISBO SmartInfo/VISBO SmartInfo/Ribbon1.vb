Imports System.Drawing
Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Core
Imports PPTNS = Microsoft.Office.Interop.PowerPoint
Imports DBAccLayer
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic

Public Class Ribbon1


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Try

            If englishLanguage Then
                With Me
                    .Group2.Label = "Update"
                    .Group3.Label = "Time Machine"
                    .Group4.Label = "Actions"
                    .btnUpdate.Label = "Update"
                    .btnStart.Label = "First  "
                    .btnFastBack.Label = "Backward"
                    .btnDate.Label = "Date"
                    .btnShowChanges.Label = "Difference"
                    .btnFastForward.Label = "Forward"
                    .btnEnd2.Label = "Last"
                    .btnToggle.Label = "Toggle"
                    .activateInfo.Label = "Properties"
                    .activateSearch.Label = "Search"
                    .activateTab.Label = "Annotate"
                    .btnFreeze.Label = "Freeze/Defreeze"
                    .settingsTab.Label = "Settings"
                End With
            Else
                With Me
                    .Group2.Label = "Aktualisieren"
                    .Group3.Label = "Time Machine"
                    .Group4.Label = "Aktionen"
                    .btnUpdate.Label = "Aktuell"
                    .btnStart.Label = "Erste Version"
                    .btnFastBack.Label = "Vorgänger Version"
                    .btnDate.Label = "Datum"
                    .btnShowChanges.Label = "Veränderung"
                    .btnFastForward.Label = "Nächste Version"
                    .btnEnd2.Label = "Neueste Version"
                    .btnToggle.Label = "hin- und herschalten"
                    .activateInfo.Label = "Eigenschaften"
                    .activateSearch.Label = "Suche"
                    .activateTab.Label = "Beschriften"
                    .btnFreeze.Label = "Konservieren/Freigeben"
                    .settingsTab.Label = "Einstellungen"
                End With
            End If

            ' password by default merken ...
            awinSettings.rememberUserPwd = True

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try
    End Sub




    Private Sub settingsTab_Click(sender As Object, e As RibbonControlEventArgs) Handles settingsTab.Click

        Dim msg As String = ""
        Try

            ' tk 11.1217 nur aktiv machen, wenn man Slides zur Weitergabe komplett strippen möchte ... um zu verhindern, dass die Re-Engineering machen ...
            'Call stripOffAllSmartInfo()

            Dim settingsfrm As New frmSettingsNew
            With settingsfrm
                Dim res As System.Windows.Forms.DialogResult = .ShowDialog()
            End With

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

    End Sub



    Private Sub activateTab_Click(sender As Object, e As RibbonControlEventArgs) Handles activateTab.Click

        Dim msg As String = ""
        Try

            'If userIsEntitled(msg) Then

            ' wird das Formular aktuell angezeigt ? 
            If IsNothing(infoFrm) And Not formIsShown Then
                infoFrm = New frmInfo
                formIsShown = True
                infoFrm.Show()
            End If

            'Else
            '    Call MsgBox(msg)
            'End If

        Catch ex As Exception
            'Call MsgBox(ex.StackTrace)
        End Try
    End Sub

    ''' <summary>
    ''' hier wird der Zustand, ob eine Slide frozen ist oder nicht gesteuert
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnFreeze_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFreeze.Click
        Try


            Dim freeze As Boolean = True


            With currentSlide

                ' Slide - Markierung frozen wieder entfernen, auch das Wasserzeichen-Shape
                If .Tags.Item("FROZEN").Length > 0 Then

                    .Tags.Delete("FROZEN")
                    currentSlide.Shapes("FreezeShape").Delete()

                Else

                    ' Slide als frozen markieren, d.h. beim Update aller Slides einer Präsi wird dieses Slide
                    ' nicht mit auf den neusten Stand gebracht
                    .Tags.Add("FROZEN", freeze.ToString)

                    Dim csWidth As Single = currentSlide.CustomLayout.Width
                    Dim csHeigth As Single = currentSlide.CustomLayout.Height
                    Dim freezeShape As PowerPoint.Shape

                    'Symbol - snowflake aus Resources holen, in File auf Temp-Dir schreiben und von dort ins Shape holen
                    Dim snowflake As Image = My.Resources.snowflake
                    Dim fileSnowflake As String = Path.Combine(Path.GetTempPath(), "snowflake.png")
                    snowflake.Save(fileSnowflake)

                    freezeShape = currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
                                                          Left:=CSng(csWidth * 0.75),
                                                          Top:=8,
                                                          Width:=32,
                                                          Height:=32)

                    With freezeShape
                        .LockAspectRatio = MsoTriState.msoTrue
                        .Name = "FreezeShape"
                        .Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                        .Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        .Fill.UserPicture(fileSnowflake)
                        .Fill.TextureTile = MsoTriState.msoFalse
                        .Fill.RotateWithObject = MsoTriState.msoTrue
                    End With

                    ' File mit dem Symbol - snowflake wieder löschen
                    File.Delete(fileSnowflake)
                End If
            End With


        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

    End Sub


    Private Sub activateSearch_Click(sender As Object, e As RibbonControlEventArgs) Handles activateSearch.Click

        Try
            If Not IsNothing(searchPane) Then
                If searchPane.Visible Then
                    searchPane.Visible = False
                Else
                    searchPane.Visible = True
                End If
            End If

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try



    End Sub

    Private Sub activateInfo_Click(sender As Object, e As RibbonControlEventArgs) Handles activateInfo.Click
        Try

            If propertiesPane.Visible Then
                propertiesPane.Visible = False
            Else
                propertiesPane.Visible = True
            End If


        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

    End Sub




    ''' <summary>
    ''' zeitgt die Veränderungen zweier Versionen an
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnShowChanges_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShowChanges.Click

        Try
            Dim key As String = CType(currentSlide.Parent, PowerPoint.Presentation).Name
            ' das Formular aufschalten 
            If IsNothing(changeFrm) Then
                changeFrm = New frmChanges
                changeFrm.changeliste = Nothing

                If chgeLstListe.ContainsKey(key) Then
                    If chgeLstListe.Item(key).ContainsKey(currentSlide.SlideID) Then
                        changeFrm.changeliste = chgeLstListe.Item(key).Item(currentSlide.SlideID)
                    End If
                End If

                'changeFrm.changeliste = chgeLstListe(currentSlide.SlideID)
                changeFrm.Show()
            Else

                changeFrm.changeliste.clearChangeList()

                If chgeLstListe.ContainsKey(key) Then
                    If chgeLstListe.Item(key).ContainsKey(currentSlide.SlideID) Then
                        changeFrm.changeliste = chgeLstListe.Item(key).Item(currentSlide.SlideID)
                    End If
                End If

                changeFrm.neuAufbau()
            End If

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try
    End Sub


    ''' <summary>
    ''' zeigt die letzte Version an
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnEnd2_Click(sender As Object, e As RibbonControlEventArgs) Handles btnEnd2.Click

        Dim presName As String = ""
        Try
            presName = pptAPP.ActivePresentation.Name

            Dim tmpDate As Date = Date.MinValue
            Call updateSelectedSlide(ptNavigationButtons.update, tmpDate)



        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnEnd2")
        End If

    End Sub


    ''' <summary>
    ''' geht einen Schritt in die Zukunft 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFastForward_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFastForward.Click

        Dim presName As String = ""
        Try
            presName = pptAPP.ActivePresentation.Name
            Dim tmpDate As Date = Date.MinValue
            Call updateSelectedSlide(ptNavigationButtons.nachher, tmpDate)

            'Dim msg As String = ""

            '' Prüfen, ob Login noch passt ...
            'If userIsEntitled(msg) Then
            '    Call btnUpdateAction(ptNavigationButtons.nachher, tmpDate)
            'Else
            '    Call MsgBox(msg)
            'End If



            ' tk 18.10.18 durch obigen Aufruf ersetzt 
            'Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            'Dim formerSlide As PowerPoint.Slide = currentSlide

            'For i As Integer = 1 To pres.Slides.Count
            '    Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
            '    If Not IsNothing(sld) Then
            '        If Not (sld.Tags.Item("FROZEN").Length > 0) _
            '            And (sld.Tags.Item("SMART") = "visbo") Then
            '            Call pptAPP_UpdateOneSlide(sld)
            '            Call visboUpdate(ptNavigationButtons.nachher, , False)
            '        End If
            '    End If
            'Next

            'currentSlide = formerSlide
            '' smartSlideLists für die aktuelle currentslide wieder aufbauen
            '' tk 22.8.18
            'Call pptAPP_UpdateOneSlide(currentSlide)
            ''Call buildSmartSlideLists()

            '' das Formular ggf, also wenn aktiv,  updaten 
            'If Not IsNothing(changeFrm) Then
            '    changeFrm.neuAufbau()
            'End If



        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnFastForward")
        End If

    End Sub

    ''' <summary>
    ''' zeigt die vorige Version an
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFastBack_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFastBack.Click

        Dim presName As String = ""
        Try
            presName = pptAPP.ActivePresentation.Name
            Dim tmpDate As Date = Date.MinValue
            Call updateSelectedSlide(ptNavigationButtons.vorher, tmpDate)

            'Dim msg As String = ""

            '' Prüfen, ob Login noch passt ...
            'If userIsEntitled(msg) Then
            '    Call btnUpdateAction(ptNavigationButtons.vorher, tmpDate)
            'Else
            '    Call MsgBox(msg)
            'End If



            ' tk 18.10.18 durch obigen Aufruf ersetzt 
            'Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            'Dim formerSlide As PowerPoint.Slide = currentSlide

            'For i As Integer = 1 To pres.Slides.Count
            '    Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
            '    If Not IsNothing(sld) Then
            '        If Not (sld.Tags.Item("FROZEN").Length > 0) _
            '            And (sld.Tags.Item("SMART") = "visbo") Then
            '            Call pptAPP_UpdateOneSlide(sld)
            '            Call visboUpdate(ptNavigationButtons.vorher, , False)
            '        End If
            '    End If
            'Next

            'currentSlide = formerSlide
            '' smartSlideLists für die aktuelle currentslide wieder aufbauen
            '' tk 22.8.18
            'Call pptAPP_UpdateOneSlide(currentSlide)
            ''Call buildSmartSlideLists()

            '' das Formular ggf, also wenn aktiv,  updaten 
            'If Not IsNothing(changeFrm) Then
            '    changeFrm.neuAufbau()
            'End If

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnFastBack")
        End If


    End Sub
    ''' <summary>
    ''' positioniert alle Slides auf den ersten Timestamp 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnStart_Click(sender As Object, e As RibbonControlEventArgs) Handles btnStart.Click

        Dim presName As String = ""
        Try
            presName = pptAPP.ActivePresentation.Name
            Dim tmpDate As Date = Date.MinValue

            Call updateSelectedSlide(ptNavigationButtons.erster, tmpDate)

            'Dim msg As String = ""

            '' Prüfen, ob Login noch passt ...
            'If userIsEntitled(msg) Then
            '    Call btnUpdateAction(ptNavigationButtons.erster, tmpDate)
            'Else
            '    Call MsgBox(msg)
            'End If


            ' tk 18.10.18 durch obigen Aufruf ersetzt 
            'Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            'Dim formerSlide As PowerPoint.Slide = currentSlide

            'For i As Integer = 1 To pres.Slides.Count
            '    Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
            '    If Not IsNothing(sld) Then
            '        If Not (sld.Tags.Item("FROZEN").Length > 0) _
            '            And (sld.Tags.Item("SMART") = "visbo") Then
            '            Call pptAPP_UpdateOneSlide(sld)
            '            Call visboUpdate(ptNavigationButtons.erster, , False)
            '        End If

            '    End If
            'Next
            'currentSlide = formerSlide
            '' smartSlideLists für die aktuelle currentslide wieder aufbauen
            '' tk 22.8.18
            'Call pptAPP_UpdateOneSlide(currentSlide)
            ''Call buildSmartSlideLists()

            '' das Formular ggf, also wenn aktiv,  updaten 
            'If Not IsNothing(changeFrm) Then
            '    changeFrm.neuAufbau()
            'End If

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnStart")
        End If

    End Sub
    Private Sub btnUpdate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdate.Click

        Dim presName As String = ""
        Try
            presName = pptAPP.ActivePresentation.Name
            'ur: 2019-06-04
            Dim control As IRibbonControl = e.Control

            Dim tmpDate As Date = Date.MinValue
            Call updateSelectedSlide(ptNavigationButtons.update, tmpDate)

            'Dim msg As String = ""

            '' Prüfen, ob Login noch passt ...
            'If userIsEntitled(msg) Then
            '    Call btnUpdateAction(ptNavigationButtons.update, tmpDate)
            'Else
            '    Call MsgBox(msg)
            'End If


            ' durch obigen Aufruf ersetzt ... 
            'Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            'Dim formerSlide As PowerPoint.Slide = currentSlide
            'Dim newestVersion As Boolean = False
            'Dim newdate As Date
            'Dim formerCurrentTimestamp As Date

            'For i As Integer = 1 To pres.Slides.Count
            '    Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
            '    newdate = Nothing
            '    If Not IsNothing(sld) Then
            '        If Not (sld.Tags.Item("FROZEN").Length > 0) _
            '            And (sld.Tags.Item("SMART") = "visbo") Then
            '            Call pptAPP_UpdateOneSlide(sld)
            '            formerCurrentTimestamp = currentTimestamp
            '            'Call visboUpdate(ptNavigationButtons.letzter, newdate, False)
            '            Call visboUpdate(ptNavigationButtons.update, newdate, False)
            '            If formerCurrentTimestamp = newdate Then
            '                newestVersion = True
            '            End If
            '        End If
            '    End If
            'Next
            'If newestVersion Then
            '    If englishLanguage Then
            '        Call MsgBox("Report is already up-to-date: (" & newdate.ToLongDateString & " " & newdate.TimeOfDay.ToString & ") ")
            '    Else
            '        Call MsgBox("Report hat den aktuellen Stand: (" & newdate.ToLongDateString & " " & newdate.TimeOfDay.ToString & ")")
            '    End If
            'End If
            'currentSlide = formerSlide
            '' smartSlideLists für die aktuelle currentslide wieder aufbauen
            '' tk 22.8.18
            'Call pptAPP_UpdateOneSlide(currentSlide)
            ''Call buildSmartSlideLists()

            '' das Formular ggf, also wenn aktiv,  updaten 
            'If Not IsNothing(changeFrm) Then
            '    changeFrm.neuAufbau()
            'End If


        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnUpdate")
        End If


    End Sub

    Private Sub varianten_Tab_Click(sender As Object, e As RibbonControlEventArgs) Handles varianten_Tab.Click
        Dim msg As String = ""

        Try

            If userIsEntitled(msg, currentSlide) Then
                Dim anzahlProjekte As Integer = smartSlideLists.countProjects
                ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
                If anzahlProjekte > 0 Then

                    ' muss noch eingeloggt werden ? 
                    ' wird inzwischen in isUserIsEntitled gemacht ... 
                    'If noDBAccessInPPT Then

                    '    noDBAccessInPPT = Not logInToMongoDB(True)

                    '    If noDBAccessInPPT Then
                    '        If englishLanguage Then
                    '            msg = "no database access ... "
                    '        Else
                    '            msg = "kein Datenbank Zugriff ... "
                    '        End If
                    '        Call MsgBox(msg)
                    '    Else

                    '        ' hier müssen jetzt die Role- & Cost-Definitions gelesen werden 
                    '        RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(Date.Now)
                    '        CostDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCostsFromDB(Date.Now)

                    '    End If

                    'End If

                    If Not noDBAccessInPPT Then

                        ' die MArker, falls welche sichtbar sind , wegmachen ... 
                        Call deleteMarkerShapes()

                        ' aktuell nur für ein Projekt implementiert 
                        If anzahlProjekte = 1 Then
                            Dim tmpName As String = smartSlideLists.getPVName(1)

                            ' jetzt wird das Formular Varianten  aufgerufen ...
                            Dim variantFormular As New frmSelectVariant
                            With variantFormular
                                .pName = getPnameFromKey(tmpName)
                                .vName = getVariantnameFromKey(tmpName)
                            End With

                            Dim dgRes As Windows.Forms.DialogResult = variantFormular.ShowDialog

                        Else
                            Call MsgBox("method not yet implemented ...")

                        End If


                    End If

                Else
                    Call MsgBox("es gibt auf dieser Seite keine Datenbank-relevanten Informationen ...")
                End If
            Else
                Call MsgBox(msg)
            End If

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try
    End Sub

    Private Sub btnDate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDate.Click

        Dim userResult As Windows.Forms.DialogResult

        Dim presName As String = ""
        Try
            presName = pptAPP.ActivePresentation.Name
            Try
                ' das Formular für Kalender aufschalten 
                calendarFrm = New frmCalendar
                userResult = calendarFrm.ShowDialog()

            Catch ex As Exception
                Throw New ArgumentException("Fehler bei der Datumseingabe: " & ex.Message)
            End Try

            If userResult = Windows.Forms.DialogResult.OK Then
                Dim specDate As Date = calendarFrm.DateTimePicker1.Value
                Call updateSelectedSlide(ptNavigationButtons.individual, specDate)

                'Dim msg As String = ""

                '' Prüfen, ob Login noch passt ...
                'If userIsEntitled(msg) Then
                '    Call btnUpdateAction(ptNavigationButtons.individual, specDate)
                'Else
                '    Call MsgBox(msg)
                'End If


                ' tk 18.10.18 ersetzt durch obigen Aufruf ... 
                'Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
                'Dim formerSlide As PowerPoint.Slide = currentSlide

                'For i As Integer = 1 To pres.Slides.Count
                '    Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
                '    If Not IsNothing(sld) Then
                '        If Not (sld.Tags.Item("FROZEN").Length > 0) _
                '            And (sld.Tags.Item("SMART") = "visbo") Then
                '            Call pptAPP_UpdateOneSlide(sld)
                '            Call visboUpdate(ptNavigationButtons.individual, specDate, False)
                '        End If
                '    End If
                'Next
                'If specDate > Date.Now Then
                '    If englishLanguage Then
                '        Call MsgBox("Last Version in Database: (" & varPPTTM.timeStamps.Last.Key.ToLongDateString & " " & varPPTTM.timeStamps.Last.Key.TimeOfDay.ToString & ")")
                '    Else
                '        Call MsgBox("aktuellster Stand in der Datenbank:  (" & varPPTTM.timeStamps.Last.Key.ToLongDateString & " " & varPPTTM.timeStamps.Last.Key.TimeOfDay.ToString & ")")
                '    End If
                'End If
                'If specDate < varPPTTM.timeStamps.First.Key Then
                '    If englishLanguage Then
                '        Call MsgBox("First Version in Database: (" & varPPTTM.timeStamps.First.Key.ToLongDateString & " " & varPPTTM.timeStamps.First.Key.TimeOfDay.ToString & ")")
                '    Else
                '        Call MsgBox("erster Stand in der Datenbank:  (" & varPPTTM.timeStamps.First.Key.ToLongDateString & " " & varPPTTM.timeStamps.First.Key.TimeOfDay.ToString & ")")
                '    End If
                'End If


                'currentSlide = formerSlide
                '' smartSlideLists für die aktuelle currentslide wieder aufbauen
                '' tk 22.8.18
                'Call pptAPP_UpdateOneSlide(currentSlide)
                ''Call buildSmartSlideLists()

                '' das Formular ggf, also wenn aktiv,  updaten 
                'If Not IsNothing(changeFrm) Then
                '    changeFrm.neuAufbau()
                'End If

            End If

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnDate")
        End If
    End Sub


    Private Sub btnToggle_Click(sender As Object, e As RibbonControlEventArgs) Handles btnToggle.Click

        Dim presName As String = ""
        Try
            presName = pptAPP.ActivePresentation.Name

            Dim tmpDate As Date = Date.MinValue
            Call updateSelectedSlide(ptNavigationButtons.previous, tmpDate)

            'Dim msg As String = ""

            '' Prüfen, ob Login noch passt ...
            'If userIsEntitled(msg) Then
            '    Call btnUpdateAction(ptNavigationButtons.previous, tmpDate)
            'Else
            '    Call MsgBox(msg)
            'End If


            ' tk , jetzt durch obigen Aufruf ersetzt 
            'Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            'Dim formerSlide As PowerPoint.Slide = currentSlide
            'For i As Integer = 1 To pres.Slides.Count
            '    Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
            '    If Not IsNothing(sld) Then
            '        If Not (sld.Tags.Item("FROZEN").Length > 0) _
            '            And (sld.Tags.Item("SMART") = "visbo") Then
            '            Call pptAPP_UpdateOneSlide(sld)
            '            Call visboUpdate(ptNavigationButtons.previous, tmpDate, False)
            '        End If
            '    End If
            'Next


            'currentSlide = formerSlide
            '' smartSlideLists für die aktuelle currentslide wieder aufbauen
            '' tk 22.8.18
            'Call pptAPP_UpdateOneSlide(currentSlide)
            ''Call buildSmartSlideLists()

            '' das Formular ggf, also wenn aktiv,  updaten 
            'If Not IsNothing(changeFrm) Then
            '    changeFrm.neuAufbau()
            'End If


        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnToggle")
        End If
    End Sub

    'Private Sub Create_Button_Click(sender As Object, e As RibbonControlEventArgs) Handles Create_Button.Click
    '    Dim deletedProj As Integer = 0
    '    Dim returnValue As Windows.Forms.DialogResult

    '    'Dim deleteProjects As New frmDeleteProjects
    '    Dim loadProjectsForm As New frmProjPortfolioAdmin

    '    Try

    '        With loadProjectsForm

    '            .aKtionskennung = PTTvActions.loadPVInPPT

    '            '' '' ''.portfolioName.Visible = False
    '            '' '' ''.Label1.Visible = False
    '        End With

    '        returnValue = loadProjectsForm.ShowDialog

    '        If returnValue = Windows.Forms.DialogResult.OK Then
    '            'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

    '            ' tk 7.10.19 jetzt werden die Platzhalter umgewandelt ...
    '            Dim hproj As clsProjekt = Nothing
    '            If selectedProjekte.Count = 1 Then
    '                hproj = selectedProjekte.getProject(1)

    '                Dim tmpCollection As New Collection
    '                Call fillReportingComponentWithinPPT(hproj, tmpCollection, tmpCollection, tmpCollection, tmpCollection, tmpCollection, tmpCollection, 0.0, 12.0)
    '                ' tk 7.10 selectedProjekte wieder zurücksetzen ..
    '                selectedProjekte.Clear(False)
    '            Else
    '                Dim msgtxt As String = "kein Projekt ausgewählt ... Abbruch"
    '                If awinSettings.englishLanguage Then
    '                    msgtxt = "no project selected ... Exit"
    '                End If
    '                Call MsgBox(msgtxt)
    '            End If



    '        Else
    '            ' returnValue = DialogResult.Cancel

    '        End If

    '    Catch ex As Exception

    '        Call MsgBox(ex.Message)
    '    End Try

    '    ' hier wird ja nix geladen, deshalb soll das nicht gemacht werden .. 
    '    'If currentConstellationName <> calcLastSessionScenarioName() Then
    '    '    currentConstellationName = calcLastSessionScenarioName()
    '    'End If

    'End Sub



    Private Function loginAndReadApearances(ByRef errMsg As String) As Boolean
        Dim wasSuccessful As Boolean = False
        Dim err As New clsErrorCodeMsg
        Dim VCId As String = ""

        ' tk wenn die jetzt noch nicht gesetzt sind , dann müssen die jetzt gesetzt werden 
        If awinSettings.databaseURL = "" Then
            awinSettings.databaseURL = "https://my.visbo.net/api"
            awinSettings.visboServer = True
            awinSettings.proxyURL = ""
            awinSettings.DBWithSSL = True
            awinSettings.databaseName = "MS Project"
        End If

        ' tk das muss beim Login gemacht werden 
        'awinSettings.databaseURL = My.Settings.dbURL
        'awinSettings.databaseName = My.Settings.dbName
        'awinSettings.visboServer = True
        'awinSettings.proxyURL = My.Settings.proxyURL
        'awinSettings.DBWithSSL = My.Settings.mongoDBSSL
        awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        awinSettings.userNamePWD = My.Settings.userNamePWD

        If logInToMongoDB(True) Then
            ' weitermachen ...

            Try
                ' die dem User zugeodneten Visbo Center lesen ...
                ' jetzt muss geprüft werden, ob es mehr als ein zugelassenes VISBO Center gibt , ist dann der Fall wenn es ein # im awinsettings.databaseNAme gibt 
                Dim listOfVCs As List(Of String) = CType(databaseAcc, DBAccLayer.Request).retrieveVCsForUser(err)

                If listOfVCs.Count > 1 Then
                    Dim chooseVC As New frmSelectOneItem
                    chooseVC.itemsCollection = listOfVCs
                    If chooseVC.ShowDialog = System.Windows.Forms.DialogResult.OK Then
                        ' alles ok 
                        awinSettings.databaseName = chooseVC.itemList.SelectedItem.ToString
                        Dim changeOK As Boolean = CType(databaseAcc, DBAccLayer.Request).updateActualVC(awinSettings.databaseName, VCId, err)
                        awinSettings.VCid = VCId

                        If Not changeOK Then
                            Throw New ArgumentException("bad Selection of VISBO project Center ... program ends  ...")
                        End If
                    Else
                        Throw New ArgumentException("no Selection of VISBO project Center ... program ends  ...")
                    End If

                End If

                ' lesen der Customization und Appearance Classes; hier wird der SOC , der StartOfCalendar gesetzt ...  

                appearanceDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveAppearancesFromDB("", Date.Now, False, err)
                If IsNothing(appearanceDefinitions) Then
                    Throw New ArgumentException("Appearance classes do not exist")
                End If

                Dim customizations As clsCustomization = CType(databaseAcc, DBAccLayer.Request).retrieveCustomizationFromDB("", Date.Now, False, err)
                If IsNothing(customizations) Then
                    Throw New ArgumentException("Customization does not exist")
                Else
                    ' alle awinSettings... mit den customizations... besetzen
                    'For Each kvp As KeyValuePair(Of Integer, clsBusinessUnit) In businessUnitDefinitions
                    '    customizations.businessUnitDefinitions.Add(kvp.Key, kvp.Value)
                    'Next
                    businessUnitDefinitions = customizations.businessUnitDefinitions

                    'For Each kvp As KeyValuePair(Of String, clsPhasenDefinition) In PhaseDefinitions.liste
                    '    customizations.phaseDefinitions.Add(kvp.Value)
                    'Next
                    PhaseDefinitions = customizations.phaseDefinitions

                    'For Each kvp As KeyValuePair(Of String, clsMeilensteinDefinition) In MilestoneDefinitions.liste
                    '    customizations.milestoneDefinitions.Add(kvp.Value)
                    'Next
                    MilestoneDefinitions = customizations.milestoneDefinitions
                    ' die Struktur clsCustomization besetzen und in die DB dieses VCs eintragen

                    showtimezone_color = customizations.showtimezone_color
                    noshowtimezone_color = customizations.noshowtimezone_color
                    calendarFontColor = customizations.calendarFontColor
                    nrOfDaysMonth = customizations.nrOfDaysMonth
                    farbeInternOP = customizations.farbeInternOP
                    farbeExterne = customizations.farbeExterne
                    iProjektFarbe = customizations.iProjektFarbe
                    iWertFarbe = customizations.iWertFarbe
                    vergleichsfarbe0 = customizations.vergleichsfarbe0
                    vergleichsfarbe1 = customizations.vergleichsfarbe1
                    'customizations.vergleichsfarbe2 = vergleichsfarbe2

                    awinSettings.SollIstFarbeB = customizations.SollIstFarbeB
                    awinSettings.SollIstFarbeL = customizations.SollIstFarbeL
                    awinSettings.SollIstFarbeC = customizations.SollIstFarbeC
                    awinSettings.AmpelGruen = customizations.AmpelGruen
                    'tmpcolor = CType(.Range("AmpelGruen").Interior.Color, Microsoft.Office.Interop.Excel.ColorFormat)
                    awinSettings.AmpelGelb = customizations.AmpelGelb
                    awinSettings.AmpelRot = customizations.AmpelRot
                    awinSettings.AmpelNichtBewertet = customizations.AmpelNichtBewertet
                    awinSettings.glowColor = customizations.glowColor

                    awinSettings.timeSpanColor = customizations.timeSpanColor
                    'awinSettings.showTimeSpanInPT = customizations.showTimeSpanInPT
                    awinSettings.showTimeSpanInPT = False

                    awinSettings.gridLineColor = customizations.gridLineColor

                    awinSettings.missingDefinitionColor = customizations.missingDefinitionColor

                    awinSettings.ActualdataOrgaUnits = customizations.allianzIstDatenReferate

                    awinSettings.autoSetActualDataDate = customizations.autoSetActualDataDate

                    awinSettings.actualDataMonth = customizations.actualDataMonth
                    ergebnisfarbe1 = customizations.ergebnisfarbe1
                    ergebnisfarbe2 = customizations.ergebnisfarbe2
                    weightStrategicFit = customizations.weightStrategicFit
                    awinSettings.kalenderStart = customizations.kalenderStart
                    awinSettings.zeitEinheit = customizations.zeitEinheit
                    awinSettings.kapaEinheit = customizations.kapaEinheit
                    awinSettings.offsetEinheit = customizations.offsetEinheit
                    awinSettings.EinzelRessExport = customizations.EinzelRessExport
                    awinSettings.zeilenhoehe1 = customizations.zeilenhoehe1
                    awinSettings.zeilenhoehe2 = customizations.zeilenhoehe2
                    awinSettings.spaltenbreite = customizations.spaltenbreite
                    awinSettings.autoCorrectBedarfe = customizations.autoCorrectBedarfe
                    awinSettings.propAnpassRess = customizations.propAnpassRess
                    awinSettings.showValuesOfSelected = customizations.showValuesOfSelected

                    awinSettings.mppProjectsWithNoMPmayPass = customizations.mppProjectsWithNoMPmayPass
                    awinSettings.fullProtocol = customizations.fullProtocol
                    awinSettings.addMissingPhaseMilestoneDef = customizations.addMissingPhaseMilestoneDef
                    awinSettings.alwaysAcceptTemplateNames = customizations.alwaysAcceptTemplateNames
                    awinSettings.eliminateDuplicates = customizations.eliminateDuplicates
                    awinSettings.importUnknownNames = customizations.importUnknownNames
                    awinSettings.createUniqueSiblingNames = customizations.createUniqueSiblingNames

                    awinSettings.readWriteMissingDefinitions = customizations.readWriteMissingDefinitions
                    awinSettings.meExtendedColumnsView = customizations.meExtendedColumnsView
                    awinSettings.meDontAskWhenAutoReduce = customizations.meDontAskWhenAutoReduce
                    awinSettings.readCostRolesFromDB = customizations.readCostRolesFromDB

                    awinSettings.importTyp = customizations.importTyp

                    awinSettings.meAuslastungIsInclExt = customizations.meAuslastungIsInclExt

                    awinSettings.englishLanguage = customizations.englishLanguage

                    awinSettings.showPlaceholderAndAssigned = customizations.showPlaceholderAndAssigned
                    awinSettings.considerRiskFee = customizations.considerRiskFee

                    ' noch zu tun, sonst in readOtherdefinitions
                    StartofCalendar = awinSettings.kalenderStart
                    'StartofCalendar = StartofCalendar.ToLocalTime()

                    historicDate = StartofCalendar
                    Try
                        If awinSettings.englishLanguage Then
                            menuCult = ReportLang(PTSprache.englisch)
                            repCult = menuCult
                            awinSettings.kapaEinheit = "PD"
                        Else
                            awinSettings.kapaEinheit = "PT"
                            menuCult = ReportLang(PTSprache.deutsch)
                            repCult = menuCult
                        End If
                    Catch ex As Exception
                        awinSettings.englishLanguage = False
                        awinSettings.kapaEinheit = "PT"
                        menuCult = ReportLang(PTSprache.deutsch)
                        repCult = menuCult
                    End Try
                End If

                ' Lesen der CustomField-Definitions
                ' Auslesen der Custom Field Definitions aus den VCSettings über ReST-Server

                customFieldDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCustomFieldsFromDB(err)

                If IsNothing(customFieldDefinitions) Then
                    customFieldDefinitions = New clsCustomFieldDefinitions
                    'Call MsgBox("no Custom-Field-Definitions in database")
                End If


                ' lesen der Organisation und Kapazitäten
                Dim currentOrga As clsOrganisation = CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", Date.Now, False, err)
                If IsNothing(currentOrga) Then

                ElseIf currentOrga.count > 0 Then
                    validOrganisations.addOrga(currentOrga)
                    CostDefinitions = currentOrga.allCosts
                    RoleDefinitions = currentOrga.allRoles
                Else
                    RoleDefinitions = New clsRollen
                    CostDefinitions = New clsKostenarten
                End If

                ' lesen der Custom User Roles 
                Dim meldungen As New Collection
                Try

                    Call setUserRoles(meldungen)
                Catch ex As Exception
                    ' hier bekommt der Nutzer die Rolle Projektleiter 
                    myCustomUserRole = New clsCustomUserRole

                    With myCustomUserRole
                        .customUserRole = ptCustomUserRoles.ProjektLeitung
                        .specifics = ""
                        .userName = dbUsername
                    End With
                    ' jetzt gibt es eine currentUserRole: myCustomUserRole - die gelten aktuell nur für Excel Projectboard, haben aber keine auswirkungen auf PPT Report Creation Addin
                    Call myCustomUserRole.setNonAllowances()
                End Try


                wasSuccessful = True
                appearancesWereRead = True

            Catch ex As Exception
                wasSuccessful = False
                errMsg = ex.Message
            End Try

        Else
            wasSuccessful = False
        End If
        ' tk 13.11.20 dem Programm klar machen, dass die Appearances gelesen wurden ...


        loginAndReadApearances = wasSuccessful
    End Function



    Private Sub btn_CreateReport_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_CreateReport.Click

        Dim returnValue As Windows.Forms.DialogResult
        Dim errMsg As String = ""

        Dim singleProjectSelect As Boolean = True

        ' check whether or not there are any reporting Components on current page. 
        ' IF Not , do nothing 

        If Not slideHasReportComponents(currentSlide) Then
            Call MsgBox("no reporting components found on current slide! -> Exit")
            Exit Sub
        End If

        ' check on valid combinations 
        If currentSldHasProjectTemplates And Not (currentSldHasMultiProjectTemplates Or currentSldHasPortfolioTemplates) Then
            singleProjectSelect = True
        ElseIf Not currentSldHasProjectTemplates And (currentSldHasMultiProjectTemplates Or currentSldHasPortfolioTemplates) Then
            singleProjectSelect = False
        Else
            Call MsgBox("no combination of project and multiproject/Portfolio components allowed! -> Exit")
            Exit Sub
        End If

        Dim loadProjectsForm As New frmProjPortfolioAdmin
        Dim weitermachen As Boolean = True
        If Not appearancesWereRead Then
            ' einloggen, dann Visbo Center wählen, dann Orga einlesen, dann user roles, dann customization und appearance classes ... 
            weitermachen = loginAndReadApearances(errMsg)
        End If

        If weitermachen Then
            ' jetzt hat ja alles geklappt: login, Settings lesen, ... 
            appearancesWereRead = True
            noDBAccessInPPT = False

            ' tk 13.11 jetzt bestimmen ob die Slide PRoject / Multiproject / Portfolio Reporting Komponenten hat
            ' Constante in projectboardDefinitions mit den Schlüsselwörtern für Projekte, Multiprojekte, Portfolios ...
            ' tk noch nicht bestimmt ... 

            Try
                If Not currentSldHasPortfolioTemplates Then

                    With loadProjectsForm
                        ' if it is a project reporting template such as Swimlanes, then only allow selection of one item 
                        If currentSldHasProjectTemplates Then
                            .aKtionskennung = PTTvActions.loadPVInPPT
                        ElseIf currentSldHasMultiProjectTemplates Then
                            ' if it is a multiproject reporting template such as Multiprojektsicht, then allow multi-selection of several items 
                            .aKtionskennung = PTTvActions.loadMultiPVInPPT
                        End If

                    End With

                    returnValue = loadProjectsForm.ShowDialog

                    If returnValue = Windows.Forms.DialogResult.OK Then

                        ' tk 7.10.19 jetzt werden die Platzhalter umgewandelt ...
                        Dim hproj As clsProjekt = Nothing
                        Dim anzP As Integer = ShowProjekte.Count
                        If selectedProjekte.Count >= 1 Then
                            hproj = selectedProjekte.getProject(1)

                            Dim tmpCollection As New Collection

                            ' hier müssen jetzt die Module alle zu smartInfo transferiert werden ... 
                            Call fillReportingComponentWithinPPT(hproj, tmpCollection, tmpCollection, tmpCollection, tmpCollection, tmpCollection, tmpCollection, 0.0, 12.0)
                            ' tk 7.10 selectedProjekte wieder zurücksetzen ..
                            ShowProjekte.Clear(False)
                            selectedProjekte.Clear(False)
                            showRangeLeft = 0
                            showRangeRight = 0


                            Try
                                ' jetzt den Namen auf das Projekt setzen, wenn er nicht schon vorher gesetzt wurde .. 

                                Dim savePath As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                                Dim fullFileName As String = My.Computer.FileSystem.CombinePath(savePath, hproj.name)
                                If anzP > 1 Then
                                    fullFileName = My.Computer.FileSystem.CombinePath(savePath, "Multiprojekt-Report")
                                End If

                                pptAPP.ActivePresentation.SaveAs(fullFileName)

                            Catch ex As Exception

                            End Try


                        Else
                            Dim msgtxt As String = "kein Projekt ausgewählt ... Abbruch"
                            If awinSettings.englishLanguage Then
                                msgtxt = "no project selected ... Exit"
                            End If
                            Call MsgBox(msgtxt)
                        End If



                    Else
                        ' returnValue = DialogResult.Cancel

                    End If


                Else
                    Call MsgBox("not yet implemented ... -> Exit")
                End If

            Catch ex As Exception

                Call MsgBox(ex.Message)
            End Try
        Else
            Call MsgBox("Login Cancelled ... - no further action")
        End If

    End Sub
End Class

