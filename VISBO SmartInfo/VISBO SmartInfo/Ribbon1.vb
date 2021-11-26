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
            With Me
                .Settings.ShowLabel = False
                .settingsTab.ShowImage = False
                .btn_ImportAppCust.ShowImage = False
            End With

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
                    .btn_ImportAppCust.Label = "Import customizable settings"

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
                    .btn_ImportAppCust.Label = "spezifische Einstellungen importieren"
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
            If Not IsNothing(propertiesPane) Then
                If propertiesPane.Visible Then
                    propertiesPane.Visible = False
                Else
                    propertiesPane.Visible = True
                End If
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


        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnUpdate")
        End If


    End Sub

    'Private Sub varianten_Tab_Click(sender As Object, e As RibbonControlEventArgs) Handles varianten_Tab.Click
    '    Dim msg As String = ""

    '    Try

    '        If userIsEntitled(msg, currentSlide) Then
    '            Dim anzahlProjekte As Integer = smartSlideLists.countProjects
    '            ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
    '            If anzahlProjekte > 0 Then

    '                ' muss noch eingeloggt werden ? 
    '                ' wird inzwischen in isUserIsEntitled gemacht ... 
    '                'If noDBAccessInPPT Then

    '                '    noDBAccessInPPT = Not logInToMongoDB(True)

    '                '    If noDBAccessInPPT Then
    '                '        If englishLanguage Then
    '                '            msg = "no database access ... "
    '                '        Else
    '                '            msg = "kein Datenbank Zugriff ... "
    '                '        End If
    '                '        Call MsgBox(msg)
    '                '    Else

    '                '        ' hier müssen jetzt die Role- & Cost-Definitions gelesen werden 
    '                '        RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(Date.Now)
    '                '        CostDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCostsFromDB(Date.Now)

    '                '    End If

    '                'End If

    '                If Not noDBAccessInPPT Then

    '                    ' die MArker, falls welche sichtbar sind , wegmachen ... 
    '                    Call deleteMarkerShapes()

    '                    ' aktuell nur für ein Projekt implementiert 
    '                    If anzahlProjekte = 1 Then
    '                        Dim tmpName As String = smartSlideLists.getPVName(1)

    '                        ' jetzt wird das Formular Varianten  aufgerufen ...
    '                        Dim variantFormular As New frmSelectVariant
    '                        With variantFormular
    '                            .pName = getPnameFromKey(tmpName)
    '                            .vName = getVariantnameFromKey(tmpName)
    '                        End With

    '                        Dim dgRes As Windows.Forms.DialogResult = variantFormular.ShowDialog

    '                    Else
    '                        Call MsgBox("method not yet implemented ...")

    '                    End If


    '                End If

    '            Else
    '                Call MsgBox("es gibt auf dieser Seite keine Datenbank-relevanten Informationen ...")
    '            End If
    '        Else
    '            Call MsgBox(msg)
    '        End If

    '    Catch ex As Exception
    '        Call MsgBox(ex.StackTrace)
    '    End Try
    'End Sub

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

        Catch ex As Exception
            Call MsgBox(ex.StackTrace)
        End Try

        pptAPP.Presentations(presName).Windows(1).Activate()

        'ur:2019-06-04
        If awinSettings.visboDebug Then
            Call MsgBox("ende btnToggle")
        End If
    End Sub


    Private Function loginAndReadApearances(ByVal dbNameIsKnown As Boolean, ByRef errMsg As String) As Boolean
        Dim wasSuccessful As Boolean = False
        Dim err As New clsErrorCodeMsg
        Dim VCId As String = ""

        ' tk wenn die jetzt noch nicht gesetzt sind , dann müssen die jetzt gesetzt werden 
        'If awinSettings.databaseURL = "" Then
        '    awinSettings.databaseURL = "https://my.visbo.net/api"
        '    awinSettings.visboServer = True
        '    awinSettings.proxyURL = ""
        '    awinSettings.DBWithSSL = True
        '    awinSettings.databaseName = "MS Project"
        'End If

        ' ur:2020.12.1: Einstellungen für direkt MongoDB oder ReST-Server Zugriff
        ' ur: 2020.12.04: werden nun in readSettings gelesen

        'awinSettings.databaseURL = My.Settings.mongoDBURL
        'awinSettings.visboServer = My.Settings.VISBOServer
        'awinSettings.proxyURL = My.Settings.proxyServerURL
        'awinSettings.DBWithSSL = My.Settings.mongoDBWithSSL
        'awinSettings.databaseName = My.Settings.mongoDBname
        'awinSettings.awinPath = My.Settings.awinPath

        ' Lesen aller userSettings
        Call readSettings(dbNameIsKnown)

        ' tk das muss beim Login gemacht werden 
        awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        awinSettings.userNamePWD = My.Settings.userNamePWD


        logfileNamePath = createLogfileName()


        If awinSettings.visboServer Then

            If logInToMongoDB(True) Then
                ' weitermachen ...

                Try

                    ' die dem User zugeodneten Visbo Center lesen ...
                    ' jetzt muss geprüft werden, ob es mehr als ein zugelassenes VISBO Center gibt , ist dann der Fall wenn es ein # im awinsettings.databaseNAme gibt 

                    If Not dbNameIsKnown Then
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
                        ElseIf listOfVCs.Count = 1 Then
                            awinSettings.databaseName = listOfVCs.First
                            Dim changeOK As Boolean = CType(databaseAcc, DBAccLayer.Request).updateActualVC(awinSettings.databaseName, VCId, err)
                            awinSettings.VCid = VCId

                            ' keine VC-Abfrage, da User nur für ein VC Zugriff hat
                        ElseIf awinSettings.visboServer Then
                            Throw New ArgumentException("no access to any VISBO project Center ... program ends  ...")
                        Else
                            ' hier direkter MongoDB-Zugriff - alles ok

                        End If
                    End If

                    ' lesen der Customization und Appearance Classes; hier wird der SOC , der StartOfCalendar gesetzt ...  
                    appearanceDefinitions.liste = CType(databaseAcc, DBAccLayer.Request).retrieveAppearancesFromDB("", Date.Now, False, err)
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

        Else ' direkter MongoDB-Zugriff und lesen der appearances und customizationSettings from File

            Try
                Dim customizations As New clsCustomization

                If Not awinsetTypen_Performed Then

                    dbUsername = ""
                    dbPasswort = ""

                    If logInToMongoDB(True) Then
                        ' weitermachen ...

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
                        ElseIf listOfVCs.Count = 1 Then
                            ' keine VC-Abfrage, da User nur für ein VC Zugriff hat
                        ElseIf awinSettings.visboServer Then
                            Throw New ArgumentException("no access to any VISBO project Center ... program ends  ...")
                        Else
                            ' hier direkter MongoDB-Zugriff - alles ok

                        End If


                        ' lesen der Customization und Appearance Classes; hier wird der SOC , der StartOfCalendar gesetzt ...  


                        appearanceDefinitions.liste = CType(databaseAcc, DBAccLayer.Request).retrieveAppearancesFromDB("", Date.Now, False, err)
                        If IsNothing(appearanceDefinitions) Then

                            If awinSettings.englishLanguage Then
                                Call MsgBox("There are no appearances defined!" & vbCrLf & "Please ask your administrator")
                            Else
                                Call MsgBox("Es sind keine Darstellungsklassen definiert!" & vbCrLf & "Bitte kontaktieren Sie Ihren Administrator")

                            End If
                        Else

                        End If

                        ' tk 14.1.2020
                        ' jetzt muss gleich die Customization ausgelesen werden und der StartOfCalendar gesetzt werden 

                        customizations = CType(databaseAcc, DBAccLayer.Request).retrieveCustomizationFromDB("", Date.Now, False, err)
                        If IsNothing(customizations) Then
                            If awinSettings.englishLanguage Then
                                Call MsgBox("There are no customizations defined!" & vbCrLf & "Please ask your administrator")
                            Else
                                Call MsgBox("Es sind keine benutzerspezifischen Einstellungen definiert!" & vbCrLf & "Bitte kontaktieren Sie Ihren Administrator")

                            End If
                        Else

                            StartofCalendar = customizations.kalenderStart

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

                    End If
                End If

                ' UserName - Password merken
                If awinSettings.rememberUserPwd Then
                    My.Settings.userNamePWD = awinSettings.userNamePWD
                End If


                If Not IsNothing(appearanceDefinitions) And Not IsNothing(customizations) Then
                    ' tk 13.11.20 dem Programm klar machen, dass die Appearances gelesen wurden ...
                    wasSuccessful = True
                    awinsetTypen_Performed = True
                    appearancesWereRead = True
                Else
                    wasSuccessful = False
                    awinsetTypen_Performed = False
                    appearancesWereRead = False
                    If awinSettings.englishLanguage Then
                        Call MsgBox("There are no customizations defined!" & vbCrLf & "Please ask your administrator")
                    Else
                        Call MsgBox("Es sind keine benutzerspezifischen Einstellungen definiert!" & vbCrLf & "Bitte kontaktieren Sie Ihren Administrator")
                    End If
                End If

            Catch ex As Exception
                Call MsgBox("Fehler beim lesen der Appearances and customizations from MongoDB")
            End Try

        End If      ' visboServer = true/false

        loginAndReadApearances = wasSuccessful
    End Function



    Private Sub btn_CreateReport_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_CreateReport.Click

        Dim returnValue As Windows.Forms.DialogResult
        Dim errMsg As String = ""

        Dim singleProjectSelect As Boolean = True
        Dim portfolioSelect As Boolean = False

        ' jetzt die ShowProjekte und soweiter löschen 
        Call emptyAllVISBOStructures(calledFromPPT:=True)

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
            If currentSldHasPortfolioTemplates Then
                portfolioSelect = True
            End If
        Else
            Call MsgBox("no combination of project and multiproject/Portfolio components allowed! -> Exit")
            Exit Sub
        End If


        Dim weitermachen As Boolean = True
        ' check out whether project portfolio selection was successful, then do action 
        Dim continueOnComponents As Boolean = False

        If Not appearancesWereRead Then
            ' einloggen, dann Visbo Center wählen, dann Orga einlesen, dann user roles, dann customization und appearance classes ... 
            weitermachen = loginAndReadApearances(False, errMsg)
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

                    Dim loadProjectsForm As New frmProjPortfolioAdmin
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
                        continueOnComponents = True
                    Else
                        ' returnValue = DialogResult.Cancel
                    End If

                Else

                    ' loads all
                    continueOnComponents = loadPortfolioProjectsInPPT()

                End If

                ' now do the action 

                If continueOnComponents Then

                    ' tk 7.10.19 jetzt werden die Platzhalter umgewandelt ...
                    Dim hproj As clsProjekt = Nothing

                    Dim projectType As ptPRPFType = ptPRPFType.project

                    If currentSldHasPortfolioTemplates Then
                        projectType = ptPRPFType.portfolio
                    End If


                    Dim anzP As Integer = ShowProjekte.Count
                    If ShowProjekte.Count >= 1 Then

                        If projectType = ptPRPFType.portfolio Then
                            ' calculate summary Project 
                            hproj = calcUnionProject(currentSessionConstellation, False, Date.Now)
                        Else
                            hproj = ShowProjekte.getProject(1)
                        End If


                        Dim tmpCollection As New Collection
                        Dim outputCollection As New Collection


                        ' hier müssen jetzt die Module alle zu smartInfo transferiert werden ... 
                        Call fillReportingComponentWithinPPT(hproj, projectType, tmpCollection, tmpCollection, tmpCollection, tmpCollection, tmpCollection, tmpCollection, 0.0, 12.0, outputCollection)

                        ' now show messgaes if there are any ... 
                        If outputCollection.Count > 0 Then
                            Call showOutPut(outputCollection, "Create Report", "Errors & Warnings")
                        End If

                        ' smartSlideLists und slidcoordInfo aufbauen
                        Call pptAPP_AufbauSmartSlideLists(currentSlide)

                        ' tk 7.10 selectedProjekte wieder zurücksetzen ..
                        ShowProjekte.Clear(False)
                        selectedProjekte.Clear(False)

                        showRangeLeft = 0
                        showRangeRight = 0

                        Dim savePath As String
                        Dim fullFileName As String
                        Try
                            savePath = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                        Catch ex As Exception
                            savePath = My.Computer.FileSystem.SpecialDirectories.Temp
                        End Try

                        Try
                            If anzP > 1 Then
                                fullFileName = My.Computer.FileSystem.CombinePath(savePath, "Multiproject-Report")
                            Else
                                fullFileName = My.Computer.FileSystem.CombinePath(savePath, hproj.name)
                            End If

                        Catch ex As Exception
                            fullFileName = My.Computer.FileSystem.CombinePath(savePath, "Single Project Report")
                        End Try

                        Try
                            pptAPP.ActivePresentation.SaveAs(fullFileName)
                        Catch ex As Exception
                            Call MsgBox("Could not save powerpoint to " & fullFileName)
                        End Try

                    Else
                        Dim msgtxt As String = "kein Projekt ausgewählt ... Abbruch"
                            If awinSettings.englishLanguage Then
                                msgtxt = "no project selected ... Exit"
                            End If
                            Call MsgBox(msgtxt)
                        End If
                    End If


            Catch ex As Exception

                Call MsgBox(ex.Message)
            End Try
        Else
            Call MsgBox("Login Cancelled ... - no further action" & vbCrLf & errMsg)
        End If

    End Sub

    Private Sub setButtonVisibility(ByVal modus As Integer)
        ' to set Visbility resp enablement of Menu-buttons, 
        ' dependent on whether it is a smartVisbo Slide , a smart template or none of the above
    End Sub

    Private Sub addElement_Click(sender As Object, e As RibbonControlEventArgs) Handles addElement.Click

        If isVisboSlide(currentSlide) Then

            Dim errmsg As String = ""
            Dim weiterMachen As Boolean = True

            ' true, if at least one milestone or phase has been added 
            Dim atleastOneAddedElement As Boolean = False

            If Not appearancesWereRead Then
                ' einloggen, dann Visbo Center wählen, dann Orga einlesen, dann user roles, dann customization und appearance classes ... 
                ' tk 5.12.20 an dieser Stelle sidn die awinsetitngs.dbname bereits gesetzt
                weiterMachen = loginAndReadApearances(True, errmsg)
            End If


            If weiterMachen Then    ' User ist bereits eingeloggt 

                Dim outPutCollection As New Collection

                ' jetzt die ShowProjekte und soweiter löschen 
                Call emptyAllVISBOStructures(calledFromPPT:=True)

                Dim pvNames As Collection = smartSlideLists.getPVNames
                If pvNames.Count > 0 Then
                    ' jetzt werden diese Projekte in AlleProjekte geladen ... 
                    ' einfach deswegen, weill evtl ja mehrere Varianten ein und desselben Projektes darunter sind 
                    For Each pvName As String In pvNames
                        Dim pName As String = getPnameFromKey(pvName)
                        Dim vName As String = getVariantnameFromKey(pvName)

                        Call loadProjectfromDB(outPutCollection, pName, vName, False, Date.Now, True)

                    Next
                End If

                If outPutCollection.Count > 0 Then
                    Dim header As String = ""
                    If englishLanguage Then
                        header = "Error Loading Project/s! "
                    Else
                        header = "Fehler beim Laden der/des Projekte/s !"
                    End If
                    Call showOutPut(outPutCollection, header, "")

                Else   ' keine Fehler beim Laden des Projekts

                    ' jetzt wird showrangeLeft und showrangeRight bestimmt 
                    Try
                        showRangeLeft = getColumnOfDate(slideCoordInfo.PPTStartOFCalendar)
                        showRangeRight = getColumnOfDate(slideCoordInfo.PPTEndOFCalendar)
                    Catch ex As Exception

                    End Try


                    ' jetzt werden die Meilensteine / Phasen ausgewählt 

                    Dim selectedPhases As New Collection
                    Dim selectedMilestones As New Collection

                    ' jetzt die selectedProjekte auf ein Projekt setzen, das wird nämlich dann verwendet , um im TreeView bei 
                    ' die Struktur Auswahl zu machen 
                    selectedProjekte.Clear(False)
                    If ShowProjekte.Count > 0 Then
                        selectedProjekte.Add(ShowProjekte.getProject(1), False)
                    End If

                    ' set the datepicker boxes in the form to invisible
                    ' because timeframe is defined by report which is currently shown
                    Dim frmSelectionPhMs As New frmSelectPhasesMilestones
                    frmSelectionPhMs.addElementMode = True


                    If frmSelectionPhMs.ShowDialog = Windows.Forms.DialogResult.OK Then

                        If Not IsNothing(frmSelectionPhMs.selectedPhases) Then
                            selectedPhases = frmSelectionPhMs.selectedPhases
                        Else
                            selectedPhases = New Collection
                        End If

                        If Not IsNothing(frmSelectionPhMs.selectedMilestones) Then
                            selectedMilestones = frmSelectionPhMs.selectedMilestones
                        Else
                            selectedMilestones = New Collection
                        End If
                    Else
                        Exit Sub
                    End If


                    For i As Integer = 1 To 2
                        Dim nameCollection As Collection

                        Dim isMilestones As Boolean = False
                        If i = 1 Then
                            nameCollection = selectedPhases
                        Else
                            nameCollection = selectedMilestones
                            isMilestones = True
                        End If

                        Dim hproj As clsProjekt = Nothing


                        ' jetzt die Phasen bzw Meilensteine zeichnen
                        ' change 
                        For Each PhaseMilestoneName As String In nameCollection

                            Dim pvName As String = ""
                            Dim pName As String = ""
                            Dim vName As String = ""
                            Dim breadcrumb As String = ""
                            Dim type As Integer = -1
                            Dim elemName As String = ""

                            Call splitHryFullnameTo2(PhaseMilestoneName, elemName, breadcrumb, type, pvName)
                            pName = getPnameFromKey(pvName)
                            vName = getVariantnameFromKey(pvName)

                            Dim msgText As String = ""
                            Dim header As String = ""

                            If type = PTItemType.projekt Or type = -1 Then

                                hproj = ShowProjekte.getProject(pName)

                                If Not IsNothing(hproj) Then
                                    Dim searchString As String = smartSlideLists.bestimmeFullBreadcrumb(pvName, breadcrumb, elemName)

                                    If Not smartSlideLists.containsFullBreadCrumb(searchString) Then
                                        Call drawElemOfProject(hproj, pvName, breadcrumb, elemName, isMilestones, atleastOneAddedElement)
                                    Else
                                        msgText = pvName & " : " & breadcrumb & "-" & elemName
                                        outPutCollection.Add(msgText)
                                    End If

                                End If

                            ElseIf type = PTItemType.vorlage Then

                                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                                    hproj = kvp.Value
                                    pvName = calcProjektKey(hproj)

                                    If Not IsNothing(hproj) Then
                                        Dim searchString As String = smartSlideLists.bestimmeFullBreadcrumb(pvName, breadcrumb, elemName)

                                        If Not smartSlideLists.containsFullBreadCrumb(searchString) Then
                                            Call drawElemOfProject(hproj, pvName, breadcrumb, elemName, isMilestones, atleastOneAddedElement)
                                        Else
                                            msgText = pvName & " : " & breadcrumb & "-" & elemName
                                            outPutCollection.Add(msgText)
                                        End If
                                    End If

                                Next


                            End If

                        Next
                    Next

                    If outPutCollection.Count > 0 Then
                        Dim header As String = ""
                        If englishLanguage Then
                            header = "Element already there - no drawing occurred!"
                        Else
                            header = "Element bereits vorhanden - nicht gezeichnet! "
                        End If
                        Call showOutPut(outPutCollection, header, "")
                    End If

                End If

            Else
                ' hier ggf auf invisible setzen, wenn erforderlich 
                If englishLanguage Then
                    Call MsgBox("sorry, you are not entitled ... ")
                Else
                    Call MsgBox("Tut uns leid, aber Sie sind nicht berechtigt ... ")
                End If

                Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoFalse)
            End If

            If atleastOneAddedElement Then
                Call pptAPP_AufbauSmartSlideLists(currentSlide)
            End If

        Else
            If englishLanguage Then
                Call MsgBox("no Smart VISBO elements found - so nothing to add ...")
            Else
                Call MsgBox("keine Smart-Phasen oder Meilensteine gefunden - Abbruch ...")
            End If
        End If

    End Sub

    Private Sub btn_ImportAppCust_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_ImportAppCust.Click

        Dim err As New clsErrorCodeMsg
        Dim VCId As String = ""
        Dim wasSuccessful As Boolean = False

        Try
            Call readSettings(False)

            pseudoappInstance = New Microsoft.Office.Interop.Excel.Application

            dbUsername = ""
            dbPasswort = ""

            If logInToMongoDB(True) Then
                ' weitermachen ...

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
                ElseIf listOfVCs.Count = 1 Then
                    ' keine VC-Abfrage, da User nur für ein VC Zugriff hat
                ElseIf awinSettings.visboServer Then
                    Throw New ArgumentException("no access to any VISBO project Center ... program ends  ...")
                Else
                    ' hier direkter MongoDB-Zugriff - alles ok

                End If


                ' lesen der Customization und Appearance Classes; hier wird der SOC , der StartOfCalendar gesetzt ...  


                Dim xlsCustomization As Excel.Workbook = Nothing


                Dim customFile As String = My.Computer.FileSystem.CombinePath(awinSettings.awinPath, customizationFile)

                If Not My.Computer.FileSystem.FileExists(customFile) Then
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Error: Couldn't find this file: '" & customFile & "'")
                    Else
                        Call MsgBox("Fehler: Folgende Datei konnte nicht gefunden werden '" & customFile & "'")
                    End If
                    Exit Sub
                End If

                'appearanceDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveAppearancesFromDB("", Date.Now, False, Err)
                appearanceDefinitions = Nothing
                If IsNothing(appearanceDefinitions) Then


                    Dim wsName7810 As Excel.Worksheet = Nothing

                    appearanceDefinitions = New clsAppearances
                    'appearanceDefinitions = New SortedList(Of String, clsAppearance)
                    ' hier muss jetzt das Customization File aufgemacht werden ...
                    Try
                        xlsCustomization = pseudoappInstance.Workbooks.Open(Filename:=customFile, [ReadOnly]:=True, Editable:=False)
                        myCustomizationFile = pseudoappInstance.ActiveWorkbook.Name

                        If Not IsNothing(xlsCustomization) Then
                            Try
                                wsName7810 = CType(xlsCustomization.Worksheets("Darstellungsklassen"),
                                                                  Global.Microsoft.Office.Interop.Excel.Worksheet)
                                If awinSettings.visboDebug Then
                                    Call MsgBox("wsName7810 angesprochen")
                                End If
                            Catch ex As Exception
                                wsName7810 = Nothing
                            End Try
                        End If
                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            Call MsgBox("Error: Couldn't find this file: '" & customFile & "'")
                        Else
                            Call MsgBox("Fehler: Folgende Datei konnte nicht gefunden werden '" & customFile & "'")
                        End If
                        Exit Sub
                    End Try

                    If Not IsNothing(wsName7810) Then   ' es existiert das Customization-File auf Platte

                        ' Aufbauen der Darstellungsklassen  aus Customizationfile
                        Call aufbauenAppearanceDefinitions(wsName7810)

                        If Not IsNothing(appearanceDefinitions) And appearanceDefinitions.liste.Count > 0 Then
                            ' jetzt wird die Appearances als Setting weggespeichert ... 
                            ' alles ok 

                            Dim result As Boolean = False
                            result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(appearanceDefinitions.liste,
                                                                                                CStr(settingTypes(ptSettingTypes.appearance)),
                                                                                                CStr(settingTypes(ptSettingTypes.appearance)),
                                                                                                Date.Now,
                                                                                                err)

                            If result = True Then
                                Call MsgBox("ok, appearances stored ...")
                                Call logger(ptErrLevel.logInfo, "appearances stored ...", "loginAndReadApearances", -1)
                            Else
                                Call MsgBox("Error when writing appearances")
                                Call logger(ptErrLevel.logError, "Error when writing appearances ...", "loginAndReadApearances", -1)
                            End If
                        Else
                            If awinSettings.englishLanguage Then
                                Call MsgBox("There are no appearances defined!" & vbCrLf & "Please ask your administrator")
                            Else
                                Call MsgBox("Es sind keine Darstellungsklassen definiert!" & vbCrLf & "Bitte kontaktieren Sie Ihren Administrator")

                            End If
                        End If
                    End If


                Else

                End If

                ' für den Fall, dass aus dem File gelesen werden muss
                Dim wsName4 As Excel.Worksheet = Nothing

                Dim customizations As clsCustomization = Nothing

                Try
                    'xlsCustomization = pseudoappInstance.Workbooks.Open(Filename:=customFile, [ReadOnly]:=True, Editable:=False)
                    'myCustomizationFile = pseudoappInstance.ActiveWorkbook.Name

                    If Not IsNothing(xlsCustomization) Then
                        wsName4 = CType(xlsCustomization.Worksheets("Einstellungen"),
                                                  Global.Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                    If awinSettings.visboDebug Then
                        Call MsgBox("wsName4 angesprochen")
                    End If
                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Error: Couldn't find this file: '" & customFile & "'")
                    Else
                        Call MsgBox("Fehler: Folgende Datei konnte nicht gefunden werden '" & customFile & "'")
                    End If
                    Exit Sub
                End Try


                Try
                    ' ur:2019-07-18: hier werden nun die Customizations-Einstellungen aus der DB gelesen, wenn allerdings nicht vorhanden, 
                    ' so aus dem Customization-File lesen, wenn auch kein Customization-File vorhanden, dann Abbruch

                    Dim noCustomizationFound As Boolean = False   ' zeigt an, dass keine Einstellungen, entweder in DB oder auf Platte, gefunden wurden

                    If IsNothing(customizations) And Not IsNothing(wsName4) Then

                        ' nur wenn der User Orga-Admin ist, kann das Customization-File gelesen werden
                        If (myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin) Then

                            ' Auslesen der BusinessUnit Definitionen
                            Call readBusinessUnitDefinitions(wsName4)

                            ' Auslesen der Phasen Definitionen 
                            Call readPhaseDefinitions(wsName4)

                            ' Auslesen der Meilenstein Definitionen 
                            Call readMilestoneDefinitions(wsName4)

                            If awinSettings.visboDebug Then
                                Call MsgBox("readMilestoneDefinitions")
                            End If

                            ' auslesen der anderen Informationen 
                            Call readOtherDefinitions(wsName4)

                            customizations = New clsCustomization

                            ' Einstellungen aus CustomizationFile und awinSettings übernehmen in customizations
                            customizations = get_customSettings()

                            If Not IsNothing(customizations) Then
                                ' jetzt werden die benutzerspez. Einstellungen als Setting weggespeichert ... 
                                ' alles ok 

                                Dim result As Boolean = False
                                result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(customizations,
                                                                                                    CStr(settingTypes(ptSettingTypes.customization)),
                                                                                                    CStr(settingTypes(ptSettingTypes.customization)),
                                                                                                    Date.Now,
                                                                                                    err)

                                If result = True Then
                                    Call MsgBox("ok, customizations stored ...")
                                    Call logger(ptErrLevel.logInfo, "customizations stored ...", "loginAndReadApearances", -1)
                                Else
                                    Call MsgBox("Error when writing customizations")
                                    Call logger(ptErrLevel.logError, "Error when writing customizations ...", "loginAndReadApearances", -1)
                                End If
                            Else
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("There are no customizations defined!" & vbCrLf & "Please ask your administrator")
                                Else
                                    Call MsgBox("Es sind keine benutzerspezifischen Einstellungen definiert!" & vbCrLf & "Bitte kontaktieren Sie Ihren Administrator")

                                End If
                            End If
                        Else
                            If awinSettings.englishLanguage Then
                                Call MsgBox("You do not have the rights setting up a new Visbo Center")

                            Else
                                Call MsgBox("Nur der OrgaAdmin kann ein VC initialisieren")
                            End If
                        End If

                    Else
                        If Not IsNothing(customizations) Then
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
                        Else
                            noCustomizationFound = True
                        End If


                    End If

                    If awinSettings.visboDebug Then
                        Call MsgBox("readOtherDefinitions")
                    End If

                    If noCustomizationFound Then
                        Throw New ArgumentException("Aktuell sind keine Einstellungen vorhanden." & vbCrLf &
                                                                "Bitte kontaktieren Sie ihren Administator!")
                    End If

                Catch ex As Exception
                    Call MsgBox("Fehler beim lesen der Appearances and customizations from MongoDB")
                End Try

                ' UserName - Password merken
                If awinSettings.rememberUserPwd Then
                    My.Settings.userNamePWD = awinSettings.userNamePWD
                End If

                If Not IsNothing(appearanceDefinitions) And Not IsNothing(customizations) Then
                    ' tk 13.11.20 dem Programm klar machen, dass die Appearances gelesen wurden ...
                    wasSuccessful = True
                    awinsetTypen_Performed = True
                    appearancesWereRead = True
                End If

            End If

        Catch ex As Exception
            Call MsgBox("Fehler beim lesen der Appearances and customizations from MongoDB")
        End Try


    End Sub

    Private Sub ImportAppCust_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub
End Class

