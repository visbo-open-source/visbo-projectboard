Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Core
Imports PPTNS = Microsoft.Office.Interop.PowerPoint
Imports DBAccLayer
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic

Public Class Ribbon1


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
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
                .btnPrevious.Label = "Previous"
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
                .btnPrevious.Label = "zuletzt gezeigte Version"
                .activateInfo.Label = "Eigenschaften"
                .activateSearch.Label = "Suche"
                .activateTab.Label = "Beschriften"
                .btnFreeze.Label = "Konservieren/Freigeben"
                .settingsTab.Label = "Einstellungen"
            End With
        End If

        ' password by default merken ...
        awinSettings.rememberUserPwd = True

    End Sub




    Private Sub settingsTab_Click(sender As Object, e As RibbonControlEventArgs) Handles settingsTab.Click

        Dim msg As String = ""

        ' tk 11.1217 nur aktiv machen, wenn man Slides zur Weitergabe komplett strippen möchte ... um zu verhindern, dass die Re-Engineering machen ...
        'Call stripOffAllSmartInfo()

        If userIsEntitled(msg) Then
            Dim settingsfrm As New frmSettings
            With settingsfrm
                Dim res As System.Windows.Forms.DialogResult = .ShowDialog()
            End With
        Else
            Call MsgBox(msg)
        End If

    End Sub



    Private Sub activateTab_Click(sender As Object, e As RibbonControlEventArgs) Handles activateTab.Click

        Dim msg As String = ""
        If userIsEntitled(msg) Then

            ' wird das Formular aktuell angezeigt ? 
            If IsNothing(infoFrm) And Not formIsShown Then
                infoFrm = New frmInfo
                formIsShown = True
                infoFrm.Show()
            End If

        Else
            Call MsgBox(msg)
        End If

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
                freezeShape = currentSlide.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle,
                                                          Left:=csWidth * 0.75,
                                                          Top:=8,
                                                          Width:=32,
                                                          Height:=32)
                    With freezeShape
                        .LockAspectRatio = MsoTriState.msoTrue
                        .Name = "FreezeShape"
                        .Line.Visible = False
                        .Fill.Visible = True
                        .Fill.UserPicture(waterSign)
                        .Fill.TextureTile = MsoTriState.msoFalse
                        .Fill.RotateWithObject = MsoTriState.msoTrue
                    End With
                End If
            End With

        Catch ex As Exception

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
            ' das Formular aufschalten 
            If IsNothing(changeFrm) Then
                changeFrm = New frmChanges
                changeFrm.changeliste = chgeLstListe(currentSlide.SlideID)
                changeFrm.Show()
            Else
                changeFrm.changeliste = chgeLstListe(currentSlide.SlideID)
                changeFrm.neuAufbau()
            End If
        Catch ex As Exception

        End Try

    End Sub


    ''' <summary>
    ''' zeigt die letzte Version an
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnEnd2_Click(sender As Object, e As RibbonControlEventArgs) Handles btnEnd2.Click


        Try
            Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            Dim formerSlide As PowerPoint.Slide = currentSlide
            Dim newestVersion As Boolean = False
            Dim newdate As Date
            Dim formerCurrentTimestamp As Date

            For i As Integer = 1 To pres.Slides.Count
                Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
                newdate = Nothing
                If Not IsNothing(sld) Then
                    If Not (sld.Tags.Item("FROZEN").Length > 0) Then
                        Call pptAPP_UpdateOneSlide(sld)
                        formerCurrentTimestamp = currentTimestamp
                        Call visboUpdate(ptNavigationButtons.letzter, newdate, False)
                        If formerCurrentTimestamp = newdate Then
                            newestVersion = True
                        End If
                    End If
                End If
            Next

            If newestVersion Then
                If englishLanguage Then
                    Call MsgBox("newest TimeStamp: " & newdate.ToLongDateString & " " & newdate.TimeOfDay.ToString & " is already shown!")
                Else
                    Call MsgBox("neuester TimeStamp: " & newdate.ToLongDateString & " " & newdate.TimeOfDay.ToString & " wird bereits angezeigt")
                End If
            End If

            currentSlide = formerSlide
            ' smartSlideLists für die aktuelle currentslide wieder aufbauen
            ' tk 22.8.18
            Call pptAPP_UpdateOneSlide(currentSlide)
            'Call buildSmartSlideLists()

            ' das Formular ggf, also wenn aktiv,  updaten 
            If Not IsNothing(changeFrm) Then
                changeFrm.neuAufbau()
            End If

        Catch ex As Exception

        End Try

    End Sub


    ''' <summary>
    ''' geht einen Schritt in die Zukunft 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFastForward_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFastForward.Click

        Try

            Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            Dim formerSlide As PowerPoint.Slide = currentSlide

            For i As Integer = 1 To pres.Slides.Count
                Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
                If Not IsNothing(sld) Then
                    If Not (sld.Tags.Item("FROZEN").Length > 0) Then
                        Call pptAPP_UpdateOneSlide(sld)
                        Call visboUpdate(ptNavigationButtons.nachher, , False)
                    End If
                End If
            Next

            currentSlide = formerSlide
            ' smartSlideLists für die aktuelle currentslide wieder aufbauen
            ' tk 22.8.18
            Call pptAPP_UpdateOneSlide(currentSlide)
            'Call buildSmartSlideLists()

            ' das Formular ggf, also wenn aktiv,  updaten 
            If Not IsNothing(changeFrm) Then
                changeFrm.neuAufbau()
            End If


        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' zeigt die vorige Version an
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFastBack_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFastBack.Click
        Try

            Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            Dim formerSlide As PowerPoint.Slide = currentSlide

            For i As Integer = 1 To pres.Slides.Count
                Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
                If Not IsNothing(sld) Then
                    If Not (sld.Tags.Item("FROZEN").Length > 0) Then
                        Call pptAPP_UpdateOneSlide(sld)
                        Call visboUpdate(ptNavigationButtons.vorher, , False)
                    End If
                End If
            Next

            currentSlide = formerSlide
            ' smartSlideLists für die aktuelle currentslide wieder aufbauen
            ' tk 22.8.18
            Call pptAPP_UpdateOneSlide(currentSlide)
            'Call buildSmartSlideLists()

            ' das Formular ggf, also wenn aktiv,  updaten 
            If Not IsNothing(changeFrm) Then
                changeFrm.neuAufbau()
            End If
        Catch ex As Exception

        End Try

    End Sub
    ''' <summary>
    ''' positioniert alle Slides auf den ersten Timestamp 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnStart_Click(sender As Object, e As RibbonControlEventArgs) Handles btnStart.Click
        Try
            Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            Dim formerSlide As PowerPoint.Slide = currentSlide

            For i As Integer = 1 To pres.Slides.Count
                Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
                If Not IsNothing(sld) Then
                    If Not (sld.Tags.Item("FROZEN").Length > 0) Then
                        Call pptAPP_UpdateOneSlide(sld)
                        Call visboUpdate(ptNavigationButtons.erster, , False)
                    End If

                End If
            Next
            currentSlide = formerSlide
            ' smartSlideLists für die aktuelle currentslide wieder aufbauen
            ' tk 22.8.18
            Call pptAPP_UpdateOneSlide(currentSlide)
            'Call buildSmartSlideLists()

            ' das Formular ggf, also wenn aktiv,  updaten 
            If Not IsNothing(changeFrm) Then
                changeFrm.neuAufbau()
            End If
        Catch ex As Exception

        End Try

    End Sub
    Private Sub btnUpdate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdate.Click
        Try

            Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            Dim formerSlide As PowerPoint.Slide = currentSlide
            Dim newestVersion As Boolean = False
            Dim newdate As Date
            Dim formerCurrentTimestamp As Date

            For i As Integer = 1 To pres.Slides.Count
                Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
                newdate = Nothing
                If Not IsNothing(sld) Then
                    If Not (sld.Tags.Item("FROZEN").Length > 0) Then
                        Call pptAPP_UpdateOneSlide(sld)
                        formerCurrentTimestamp = currentTimestamp
                        Call visboUpdate(ptNavigationButtons.letzter, newdate, False)
                        If formerCurrentTimestamp = newdate Then
                            newestVersion = True
                        End If
                    End If
                End If
            Next
            If newestVersion Then
                If englishLanguage Then
                    Call MsgBox("newest TimeStamp: " & newdate.ToLongDateString & " " & newdate.TimeOfDay.ToString & " is already shown!")
                Else
                    Call MsgBox("neuester TimeStamp: " & newdate.ToLongDateString & " " & newdate.TimeOfDay.ToString & " wird bereits angezeigt")
                End If
            End If
            currentSlide = formerSlide
            ' smartSlideLists für die aktuelle currentslide wieder aufbauen
            ' tk 22.8.18
            Call pptAPP_UpdateOneSlide(currentSlide)
            'Call buildSmartSlideLists()

            ' das Formular ggf, also wenn aktiv,  updaten 
            If Not IsNothing(changeFrm) Then
                changeFrm.neuAufbau()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub varianten_Tab_Click(sender As Object, e As RibbonControlEventArgs) Handles varianten_Tab.Click
        Dim msg As String = ""

        If userIsEntitled(msg) Then
            Dim anzahlProjekte As Integer = smartSlideLists.countProjects
            ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
            If anzahlProjekte > 0 Then

                ' muss noch eingeloggt werden ? 
                If noDBAccessInPPT Then

                    noDBAccessInPPT = Not logInToMongoDB(True)

                    If noDBAccessInPPT Then
                        If englishLanguage Then
                            msg = "no database access ... "
                        Else
                            msg = "kein Datenbank Zugriff ... "
                        End If
                        Call MsgBox(msg)
                    Else

                        ' hier müssen jetzt die Role- & Cost-Definitions gelesen werden 
                        RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(Date.Now)
                        CostDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCostsFromDB(Date.Now)

                    End If

                End If

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
    End Sub

    Private Sub btnDate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDate.Click

        Dim userResult As Windows.Forms.DialogResult
        Try
            Try
                ' das Formular für Kalender aufschalten 
                calendarFrm = New frmCalendar
                userResult = calendarFrm.ShowDialog()

            Catch ex As Exception
                Throw New ArgumentException("Fehler bei der Datumseingabe: " & ex.Message)
            End Try

            If userResult = Windows.Forms.DialogResult.OK Then
                Dim specDate As Date = calendarFrm.DateTimePicker1.Value

                Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
                Dim formerSlide As PowerPoint.Slide = currentSlide

                For i As Integer = 1 To pres.Slides.Count
                    Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
                    If Not IsNothing(sld) Then
                        If Not (sld.Tags.Item("FROZEN").Length > 0) Then
                            Call pptAPP_UpdateOneSlide(sld)
                            Call visboUpdate(ptNavigationButtons.individual, specDate, False)
                        End If
                    End If
                Next
                If specDate > varPPTTM.timeStamps.Last.Key Then
                    If englishLanguage Then
                        Call MsgBox("TimeStamp: " & specDate.ToLongDateString & " " & specDate.TimeOfDay.ToString & " does not exist: Now the newest is shown")
                    Else
                        Call MsgBox("TimeStamp: " & specDate.ToLongDateString & " " & specDate.TimeOfDay.ToString & " existiert nicht: Es wird der neueste angezeigt")
                    End If
                End If
                If specDate < varPPTTM.timeStamps.First.Key Then
                    If englishLanguage Then
                        Call MsgBox("TimeStamp: " & specDate.ToLongDateString & " " & specDate.TimeOfDay.ToString & " does not exist")
                    Else
                        Call MsgBox("TimeStamp: " & specDate.ToLongDateString & " " & specDate.TimeOfDay.ToString & " existiert nicht")
                    End If
                End If


                currentSlide = formerSlide
                ' smartSlideLists für die aktuelle currentslide wieder aufbauen
                ' tk 22.8.18
                Call pptAPP_UpdateOneSlide(currentSlide)
                'Call buildSmartSlideLists()

                ' das Formular ggf, also wenn aktiv,  updaten 
                If Not IsNothing(changeFrm) Then
                    changeFrm.neuAufbau()
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub


    Private Sub btnPrevious_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPrevious.Click
        Try

            Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation
            Dim formerSlide As PowerPoint.Slide = currentSlide


            For i As Integer = 1 To pres.Slides.Count
                Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
                If Not IsNothing(sld) Then
                    If Not (sld.Tags.Item("FROZEN").Length > 0) Then
                        Call pptAPP_UpdateOneSlide(sld)
                        Call visboUpdate(ptNavigationButtons.previous, Nothing, False)
                    End If
                End If
            Next


            currentSlide = formerSlide
            ' smartSlideLists für die aktuelle currentslide wieder aufbauen
            ' tk 22.8.18
            Call pptAPP_UpdateOneSlide(currentSlide)
            'Call buildSmartSlideLists()

            ' das Formular ggf, also wenn aktiv,  updaten 
            If Not IsNothing(changeFrm) Then
                changeFrm.neuAufbau()
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class

