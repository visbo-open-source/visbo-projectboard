Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Core
Imports PPTNS = Microsoft.Office.Interop.PowerPoint
Imports MongoDbAccess
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
                .btnFastBack.Label = "Prev."
                .btnDate.Label = "Date"
                .btnShowChanges.Label = "Difference"
                .btnFastForward.Label = "Next"
                .btnEnd2.Label = "Last"
                .activateInfo.Label = "Properties"
                .activateSearch.Label = "Search"
                .activateTab.Label = "Annotate"
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
                .activateInfo.Label = "Eigenschaften"
                .activateSearch.Label = "Suche"
                .activateTab.Label = "Beschriften"
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

    ''Private Sub timeMachineTab_Click(sender As Object, e As RibbonControlEventArgs)
    ''    Dim msg As String = ""

    ''    If userIsEntitled(msg) Then
    ''        ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
    ''        If smartSlideLists.countProjects > 0 Then

    ''            ' muss noch eingeloggt werden ? 
    ''            If noDBAccessInPPT Then

    ''                noDBAccessInPPT = Not logInToMongoDB(True)

    ''                If noDBAccessInPPT Then
    ''                    If englishLanguage Then
    ''                        msg = "no database access ... "
    ''                    Else
    ''                        msg = "kein Datenbank Zugriff ... "
    ''                    End If
    ''                    Call MsgBox(msg)
    ''                Else
    ''                    ' hier müssen jetzt die Role- & Cost-Definitions gelesen werden 
    ''                    Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
    ''                    'RoleDefinitions = request.retrieveRolesFromDB(currentTimestamp)
    ''                    'CostDefinitions = request.retrieveCostsFromDB(currentTimestamp)
    ''                    RoleDefinitions = request.retrieveRolesFromDB(Date.Now)
    ''                    CostDefinitions = request.retrieveCostsFromDB(Date.Now)
    ''                End If

    ''            End If

    ''            If Not noDBAccessInPPT Then

    ''                If Not smartSlideLists.historiesExist Then


    ''                    Dim anzahlProjekte As Integer = smartSlideLists.countProjects
    ''                    ' größter kleinster Wert 
    ''                    Dim gkw As Date = Date.MinValue

    ''                    For i As Integer = 1 To anzahlProjekte
    ''                        Dim tmpName As String = smartSlideLists.getPVName(i)
    ''                        Dim pName As String = getPnameFromKey(tmpName)
    ''                        Dim vName As String = getVariantnameFromKey(tmpName)
    ''                        Dim pvName As String = calcProjektKeyDB(pName, vName)
    ''                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
    ''                        Dim tsCollection As Collection = request.retrieveZeitstempelFromDB(pvName)
    ''                        ' ermitteln des größten kleinstern Wertes ...
    ''                        ' stellt sicher, dass , wenn mehrere Projekte dargesteltl sind, nur TimeStamps abgerufen werden, die jedes Projekt hat ... 

    ''                        Dim kleinsterWert As Date = Date.Now
    ''                        If Not IsNothing(tsCollection) Then
    ''                            If tsCollection.Count > 0 Then
    ''                                ' tsCollection ist absteigend sortiert ... 
    ''                                kleinsterWert = tsCollection.Item(tsCollection.Count)
    ''                            End If
    ''                        End If
    ''                        If kleinsterWert > gkw Then
    ''                            gkw = kleinsterWert
    ''                        End If

    ''                        smartSlideLists.addToListOfTS(tsCollection)
    ''                    Next

    ''                    If anzahlProjekte > 1 Then
    ''                        ' jetzt werden aus der TimeStampListe alle TimeStamps rausgeworfen, die kleiner als der gkw sind ... 
    ''                        smartSlideLists.adjustListOfTS(gkw)
    ''                    End If

    ''                End If

    ''                ' jetzt wird das Formular TimeStamps aufgerufen ...
    ''                Dim tmFormular As New frmPPTTimeMachine
    ''                Dim dgRes As Windows.Forms.DialogResult = tmFormular.ShowDialog
    ''                'tmFormular.Show()
    ''            End If

    ''        Else
    ''            Call MsgBox("es gibt auf dieser Seite keine Datenbank-relevanten Informationen ...")
    ''        End If
    ''    Else
    ''        Call MsgBox(msg)
    ''    End If

    ''End Sub


    ''Private Sub variantTab_Click_Click(sender As Object, e As RibbonControlEventArgs)
    ''    Dim msg As String = ""

    ''    If userIsEntitled(msg) Then
    ''        Dim anzahlProjekte As Integer = smartSlideLists.countProjects
    ''        ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
    ''        If anzahlProjekte > 0 Then

    ''            ' muss noch eingeloggt werden ? 
    ''            If noDBAccessInPPT Then

    ''                noDBAccessInPPT = Not logInToMongoDB(True)

    ''                If noDBAccessInPPT Then
    ''                    If englishLanguage Then
    ''                        msg = "no database access ... "
    ''                    Else
    ''                        msg = "kein Datenbank Zugriff ... "
    ''                    End If
    ''                    Call MsgBox(msg)
    ''                Else
    ''                    ' hier müssen jetzt die Role- & Cost-Definitions gelesen werden 
    ''                    Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
    ''                    'RoleDefinitions = request.retrieveRolesFromDB(currentTimestamp)
    ''                    'CostDefinitions = request.retrieveCostsFromDB(currentTimestamp)
    ''                    RoleDefinitions = request.retrieveRolesFromDB(Date.Now)
    ''                    CostDefinitions = request.retrieveCostsFromDB(Date.Now)
    ''                End If

    ''                Dim tmpKey As String = ""

    ''            End If

    ''            If Not noDBAccessInPPT Then

    ''                ' die MArker, falls welche sichtbar sind , wegmachen ... 
    ''                Call deleteMarkerShapes()

    ''                ' aktuell nur für ein Projekt implementiert 
    ''                If anzahlProjekte = 1 Then
    ''                    Dim tmpName As String = smartSlideLists.getPVName(1)

    ''                    ' jetzt wird das Formular Varianten  aufgerufen ...
    ''                    Dim variantFormular As New frmSelectVariant
    ''                    With variantFormular
    ''                        .pName = getPnameFromKey(tmpName)
    ''                        .vName = getVariantnameFromKey(tmpName)
    ''                    End With

    ''                    Dim dgRes As Windows.Forms.DialogResult = variantFormular.ShowDialog

    ''                Else
    ''                    Call MsgBox("method not yet implemented ...")

    ''                End If


    ''            End If

    ''        Else
    ''            Call MsgBox("es gibt auf dieser Seite keine Datenbank-relevanten Informationen ...")
    ''        End If
    ''    Else
    ''        Call MsgBox(msg)
    ''    End If
    ''End Sub

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

        If propertiesPane.Visible Then
            propertiesPane.Visible = False
        Else
            propertiesPane.Visible = True
        End If

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
                changeFrm.Show()
            Else
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

        Call visboUpdate()

    End Sub

    '''' <summary>
    '''' führt den Code gehe-zum-letzten bzw Visbo-Update aus 
    '''' </summary>
    '''' <remarks></remarks>
    'Private Sub visboUpdate()

    '    If IsNothing(varPPTTM) Then
    '        Call initPPTTimeMachine(varPPTTM)
    '    End If

    '    If Not IsNothing(varPPTTM.timeStamps) Then

    '        If varPPTTM.timeStamps.Count > 0 Then

    '            Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.letzter)

    '            If newDate <> currentTimestamp Then

    '                Call performBtnAction(newDate)

    '            End If
    '        End If
    '    End If

    'End Sub
    ''' <summary>
    ''' geht einen Schritt in die Zukunft 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFastForward_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFastForward.Click

        Dim newDate As Date
        Dim found As Boolean = False
        Dim weitermachen As Boolean = False


        If IsNothing(varPPTTM) Then
            Call initPPTTimeMachine(varPPTTM)
        End If

        If Not IsNothing(varPPTTM.timeStamps) Then
            If varPPTTM.timeStamps.Count > 0 Then

                newDate = getNextNavigationDate(ptNavigationButtons.nachher)

                If newDate <> currentTimestamp Then

                    Call performBtnAction(newDate)

                End If

            End If
        End If

    End Sub

    ''' <summary>
    ''' zeigt die vorige Version an
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnFastBack_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFastBack.Click

        If IsNothing(varPPTTM) Then
            Call initPPTTimeMachine(varPPTTM)
        End If

        If Not IsNothing(varPPTTM.timeStamps) Then

            If varPPTTM.timeStamps.Count > 0 Then

                Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.vorher)

                If newDate <> currentTimestamp Then

                    Call performBtnAction(newDate)


                End If


            End If
        End If

    End Sub
    ''' <summary>
    ''' positioniert auf den ersten Timestamp 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnStart_Click(sender As Object, e As RibbonControlEventArgs) Handles btnStart.Click

        If IsNothing(varPPTTM) Then
            Call initPPTTimeMachine(varPPTTM)
        End If
        If Not IsNothing(varPPTTM.timeStamps) Then
            If varPPTTM.timeStamps.Count > 0 Then

                Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.erster)

                If newDate <> currentTimestamp Then

                    Call performBtnAction(newDate)

                End If

            End If
        End If

    End Sub
    Private Sub btnUpdate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdate.Click
        'Call visboUpdate()

        Dim pres As PowerPoint.Presentation = pptAPP.ActivePresentation

        For i As Integer = 1 To pres.Slides.Count
            Dim sld As PowerPoint.Slide = pres.Slides.Item(i)
            If Not IsNothing(sld) Then
                If Not (sld.Tags.Item("FROZEN").Length > 0) Then
                    Call pptAPP_UpdateSpecSlide(sld)
                    Call visboUpdate(False)
                End If
            End If


        Next
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
                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                        'RoleDefinitions = request.retrieveRolesFromDB(currentTimestamp)
                        'CostDefinitions = request.retrieveCostsFromDB(currentTimestamp)
                        RoleDefinitions = request.retrieveRolesFromDB(Date.Now)
                        CostDefinitions = request.retrieveCostsFromDB(Date.Now)
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
        Try

            If IsNothing(varPPTTM) Then
                Call initPPTTimeMachine(varPPTTM)
            End If
            If Not IsNothing(varPPTTM.timeStamps) Then
                If varPPTTM.timeStamps.Count > 0 Then
                    Try
                        ' das Formular für Kalender aufschalten 
                        If IsNothing(calendarFrm) Then
                            calendarFrm = New frmCalendar
                            calendarFrm.ShowDialog()
                        Else
                            calendarFrm = New frmCalendar
                            calendarFrm.ShowDialog()
                        End If
                    Catch ex As Exception

                    End Try

                    Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.individual, calendarFrm.DateTimePicker1.Value)

                    If newDate <> currentTimestamp Then

                        Call performBtnAction(newDate)

                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub


    Private Sub btnPrevious_Click(sender As Object, e As RibbonControlEventArgs) Handles btnPrevious.Click

        If IsNothing(varPPTTM) Then
            Call initPPTTimeMachine(varPPTTM)
        End If

        If Not IsNothing(varPPTTM) Then

            If Not IsNothing(varPPTTM.timeStamps) Then
                If varPPTTM.timeStamps.Count > 0 Then

                    Dim newDate As Date = getNextNavigationDate(ptNavigationButtons.previous)

                    If newDate <> currentTimestamp Then

                        Call performBtnAction(newDate)

                    End If

                End If

            End If

        End If
    End Sub
End Class

