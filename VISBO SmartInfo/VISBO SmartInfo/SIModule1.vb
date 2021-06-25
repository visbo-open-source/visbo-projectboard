Imports ProjectBoardDefinitions
'Imports ProjectboardReports
'Imports DBAccLayer
Imports ProjectBoardBasic
Imports xlNS = Microsoft.Office.Interop.Excel
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
'Imports Microsoft.Office.Core.MsoThemeColorIndex

Module SIModule1
    ' benötigt zum Lesen customization
    Public pseudoappInstance As Microsoft.Office.Interop.Excel.Application = Nothing

    'Friend Const hiddenExcelSheetName As String = "visboupdate"
    Friend WithEvents pptAPP As PowerPoint.Application

    Friend ucPropertiesView As ucProperties
    Friend ucSearchView As ucSearch
    Friend WithEvents propertiesPane As Microsoft.Office.Tools.CustomTaskPane
    Friend WithEvents searchPane As Microsoft.Office.Tools.CustomTaskPane

    'Friend visboInfoActivated As Boolean = False
    Friend formIsShown As Boolean = False
    Friend Const markerName As String = "VisboMarker"
    Friend Const shadowName As String = "VisboShadow"
    Friend Const protectionTag As String = "VisboProtection"
    Friend Const protectionValue As String = "VisboValue"
    Friend Const noVariantName As String = "-9999999"


    Friend myPPTWindow As PowerPoint.DocumentWindow = Nothing

    Friend Const changeColor As Integer = Excel.XlRgbColor.rgbSteelBlue
    Friend currentSlide As PowerPoint.Slide

    Friend VisboProtected As Boolean = False
    Friend protectionSolved As Boolean = False

    ' bestimmt, ob in englisch oder auf deutsch ..
    Friend englishLanguage As Boolean = True

    ' wird in Activate_Window gesetzt bzw. in After_presentation
    Friend currentPresHasVISBOElements As Boolean = False


    ' die VISBO TimeMachine nimmt alle PRojekte und Timestamps auf 
    Friend timeMachine As New clsPPTTimeMachine

    ' alle meine customUserRoles 
    Dim allMyCustomUserRoles As Collection = Nothing

    ' in der Liste werden für jede Präsentation die beiden Timestamps Previous und current gemerkt 
    ' 0 = previous, 1 = current
    Friend rememberListOfCPTimeStamps As SortedList(Of String, Date()) = Nothing

    ' was ist der aktuelle Timestamp der Slide 
    Friend currentTimestamp As Date = Date.MinValue
    Friend previousTimeStamp As Date = Date.MinValue



    ' in der Liste werden für jede Präsentation die beiden Varianten-Namen Previous und current gemerkt 
    ' 0 = previous, 1 = current
    Friend rememberListOfCPVariantNames As SortedList(Of String, String()) = Nothing

    Friend currentVariantname As String = ""
    Friend previousVariantName As String = noVariantName

    ' der Key ist der Name des Referenz-Shapes, zu dem der Marker gezeichnet wird , der Value ist der Name des Marker-Shapes 
    Friend markerShpNames As New SortedList(Of String, String)

    ' wird gesetzt in Einstellungen 
    ' steuert, ob extended seach gemacht werden kann; wirkt auf Suchfeld (NAme, Original Name, Abkürzung, ..)  
    Friend extSearch As Boolean = False
    ' wird gesetzt in Einstellungen 
    ' gibt an, mit welcher Schriftgroesse der Text geschrieben wird 
    Friend schriftGroesse As Double = 10.0

    ' gibt an, ob die Charts editable sein sollen 
    Friend smartChartsAreEditable As Boolean = False

    ' wird gesetzt in Einstellungen 
    ' gibt an, ob das Breadcrumb Feld gezeigt werden soll 
    Friend showBreadCrumbField As Boolean = False
    ' gibt die MArker-Höhe und Breite an 
    Friend markerHeight As Double = 19
    Friend markerWidth As Double = 13



    ' globale Variable, die angibt, ob ShortName gezeichnet werden soll 
    Friend showShortName As Boolean = False
    ' globale Variable, die anzeigt, ob Orginal Name gezeigt werden soll 
    Friend showOrigName As Boolean = False
    ' globale Varianle, die angibt, ob der Best-Name, also der eindeutige Name gezeigt werden soll 
    Friend showBestName As Boolean = True
    ' globale Variable, die angibt, ob für Meilenstein und/oder Phase mit KW beschriftet wird
    Friend showWeek As Boolean = False

    Friend protectType As Integer
    Friend protectFeld1 As String = ""
    Friend protectFeld2 As String = ""

    Friend noDBAccessInPPT As Boolean = True
    Friend myUserName As String = ""
    Friend myUserPWD As String = ""

    Friend defaultSprache As String = "Original"
    Friend selectedLanguage As String = defaultSprache

    Friend absEinheit As Integer = 0

    Friend selectedPlanShapes As PowerPoint.ShapeRange = Nothing

    ' hier werden PPTCalendar, linker Rand etc gehalten
    ' mit dieser Klasse können auch die Berechnungen Koord->Datum und umgekehrt durchgeführt werden 
    Friend slideCoordInfo As clsPPTShapes = Nothing

    Friend infoFrm As frmInfo = Nothing
    ' wird automatisch gesetzt, wenn in einer Slide Smart-Infos sind ... 
    Friend slideHasSmartElements As Boolean = False


    ' diese Listen enthalten die Infos welche Shapes Ampel grün, gelb etc haben
    ' welche welchen Namen tragen, ...
    Friend smartSlideLists As New clsSmartSlideListen

    ' diese Liste enthält die Veränderungen nach einem TimeStamp oder Varianten Wechsel 
    'Friend changeListe As New clsChangeListe

    ' diese Liste enthält für jede Slide der Presentation die changeListe, sortiert nach WindowID und dann nach SlideNr.
    Friend chgeLstListe As New SortedList(Of String, SortedList(Of Integer, clsChangeListe))

    ' dieses Formular gibt die Changes, die sich bei den Elementen ergeben haben 
    Friend changeFrm As frmChanges = Nothing

    ' dieses Formular gibt soll die Eingabe im Kalender für TimeMachine Konkretes Datum ermöglichen
    Friend calendarFrm As frmCalendar = Nothing

    ' see msdn: https://social.msdn.microsoft.com/Forums/sqlserver/en-US/b1c610bf-82ab-4d9e-b425-de21b45ea3fb/same-taskpane-in-multiple-powerpoint-windows?forum=vsto 
    Friend listOfWindows As New List(Of Integer)
    Friend listOfucProperties As New SortedList(Of Integer, Microsoft.Office.Tools.CustomTaskPane)
    Friend listOfucSearch As New SortedList(Of Integer, Microsoft.Office.Tools.CustomTaskPane)
    Friend listOfucPropView As New SortedList(Of Integer, ucProperties)
    Friend listOfucSearchView As New SortedList(Of Integer, ucSearch)

    ' dieser array nimmt die Koordinaten der Formulare auf 
    ' die Koordinaten werden in der Reihenfolge gespeichert: top, left, width, height 
    Public frmCoord(2, 3) As Integer

    ' Enumeration Formulare - muss in Korrelation sein mit frmCoord: Dim von frmCoord muss der Anzahl Elemente entsprechen
    Public Enum PTfrm
        changes = 0
        calendar = 1
    End Enum

    Public Enum PTpinfo
        top = 0
        left = 1
        width = 2
        height = 3
    End Enum
    ' wird verwendet um in SlideSelectionChange wieder zu bestimmen, ob es eine TodayLine und ein Version-Field gibt
    Public Enum ptImportantShapes
        todayline = 0
        version = 1
    End Enum

    Public importantShapes() As PowerPoint.Shape

    Friend languages As New clsLanguages

    ' diese Variablen geben an, ob es irgendwo Shapes gibt, die verschoben wurden 
    ' bzw. Shapes, die zwar am Home sind, aber einen Changed Wert haben ... 
    Friend homeButtonRelevance As Boolean = False
    Friend changedButtonRelevance As Boolean = False

    Friend initialHomeButtonRelevance As Boolean = False
    Friend initialChangedButtonRelevance As Boolean = False

    Friend bekannteIDs As SortedList(Of Integer, String)

    Friend trafficLightColors(4) As Long
    Friend showTrafficLights(4) As Boolean

    ' wird verwendet, um zu jeder Presentation die eindeutige ID und damit die zugehörigen currentTimestamps, variantNames, varPPTTM, chgelst 's zu finden   
    Friend listOfPresentations As New SortedList(Of String, Integer)

    ' nicht mehr notwendig ... 
    '' muss bei jedem SlideSelection Change auf Nothing gesetzt werden ...
    'Friend varPPTTM As New SortedList(Of String, clsPPTTimeMachine)

    Friend Enum ptNavigationButtons
        letzter = 0
        erster = 1
        nachher = 2
        vorher = 3
        individual = 4
        previous = 5
        update = 6
    End Enum


    Friend Enum pptAbsUnit
        tage = 0
        wochen = 1
        monate = 2
    End Enum

    Friend Enum pptAnnotationType
        text = 0
        datum = 1
        ampelText = 2
        lieferumfang = 3
        movedExplanation = 4
        resourceCost = 5
        responsible = 6
    End Enum

    Friend Enum pptInfoType
        cName = 0
        oName = 1
        sName = 2
        bCrumb = 3
        aColor = 4
        aExpl = 5
        appClass = 6
        lUmfang = 7
        mvElement = 8
        resources = 9
        costs = 10
        responsible = 11
        overDue = 12
        noProgress = 13
    End Enum

    Friend Enum pptPositionType
        center = 0
        aboveCenter = 1
        aboveRight = 2
        centerRight = 3
        belowRight = 4
        belowCenter = 5
        belowLeft = 6
        centerLeft = 7
        aboveLeft = 8
        asis = 9
    End Enum


    ''' <summary>
    ''' berechnet eine Integer Zahl, die Auskunft gibt, wie die vier TrafficLights gesetzt sind 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function calcColorCode(ByVal showTrafficLights() As Boolean) As Integer

        Dim tmpNumber As Integer = 0

        If showTrafficLights(0) Then
            tmpNumber = tmpNumber + 1 ' 2 hoch 0 
        End If

        If showTrafficLights(1) Then
            tmpNumber = tmpNumber + 2 ' 2 hoch 1 
        End If

        If showTrafficLights(2) Then
            tmpNumber = tmpNumber + 4 ' 2 hoch 2 
        End If

        If showTrafficLights(3) Then
            tmpNumber = tmpNumber + 8 ' 2 hoch 3 
        End If

        calcColorCode = tmpNumber

    End Function


    ''' <summary>
    ''' hier wird unterscheiden zwischen Meilensteinen und Phasen: der Schatten wird für die beiden Typen anders dargestellt
    ''' in der alten Verison (siehe oben) wurde immer 160% drauf gepackt , das hat bei Phasen extrem komisch ausgesehen 
    ''' </summary>
    ''' <param name="ampelColor"></param>
    ''' <param name="show"></param>
    ''' <remarks></remarks>
    Friend Sub faerbeShapes(ByVal ampelColor As Integer, ByVal show As Boolean)

        Dim tmpCollection As Collection = smartSlideLists.getShapeNamesWithColor(ampelColor)
        Dim anzSelected As Integer = tmpCollection.Count
        Dim nameArray() As String

        If ampelColor >= 0 And ampelColor <= 3 Then
            'alles ok 
        Else
            ' sicherstellen, es kommt zu keinem Absturz .... 
            ampelColor = 0
        End If


        Dim shapesToBeColored As PowerPoint.ShapeRange


        If anzSelected >= 1 Then
            ReDim nameArray(anzSelected - 1)

            For i As Integer = 0 To anzSelected - 1
                nameArray(i) = CStr(tmpCollection.Item(i + 1))
            Next

            Try
                shapesToBeColored = currentSlide.Shapes.Range(nameArray)

                If Not show Then
                    ' Schatten wieder wegnehmen 
                    With shapesToBeColored.Shadow
                        .Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    End With


                Else
                    For i As Integer = 0 To anzSelected - 1
                        shapesToBeColored = currentSlide.Shapes.Range(nameArray(i))
                        If pptShapeIsMilestone(shapesToBeColored(1)) Then
                            With shapesToBeColored.Shadow
                                .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                                .Type = Microsoft.Office.Core.MsoShadowType.msoShadow25
                                .Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow
                                .Blur = 0
                                '.Size = 160
                                .Size = 140
                                .OffsetX = 0
                                .OffsetY = 0
                                .Transparency = 0
                                .ForeColor.RGB = CInt(trafficLightColors(ampelColor))
                            End With
                        Else
                            With shapesToBeColored.Shadow
                                .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                                .Type = Microsoft.Office.Core.MsoShadowType.msoShadow25
                                .Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow
                                .Blur = 0
                                '.Size = 160
                                .Size = 100
                                .OffsetX = 3
                                .OffsetY = -3
                                .Transparency = 0
                                .ForeColor.RGB = CInt(trafficLightColors(ampelColor))
                            End With
                        End If
                    Next

                    ' mit Schatten einfärben 


                End If


            Catch ex As Exception

            End Try

        Else
            ' nichts tun ...

        End If


    End Sub

    ''' <summary>
    ''' färbt das übergebene Shape mit der AmpelFarbe bzw. löscht die angezeigte AmpelFarbe
    ''' </summary>
    ''' <param name="ampelColor"></param>
    ''' <param name="show"></param>
    ''' <remarks></remarks>
    Friend Sub faerbeShape(ByRef tmpShape As PowerPoint.Shape,
                           ByVal ampelColor As Integer, ByVal show As Boolean)


        If ampelColor >= 0 And ampelColor <= 3 Then
            'alles ok 
            Try
                If show Then
                    ' mit Schatten einfärben 
                    With tmpShape.Shadow
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        .Type = Microsoft.Office.Core.MsoShadowType.msoShadow25
                        .Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow
                        .Blur = 0
                        .Size = 160
                        .OffsetX = 0
                        .OffsetY = 0
                        .Transparency = 0
                        .ForeColor.RGB = CInt(trafficLightColors(ampelColor))
                    End With
                Else
                    ' Schatten wieder wegnehmen 
                    With tmpShape.Shadow
                        .ForeColor.RGB = CInt(trafficLightColors(ampelColor))
                        .Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    End With
                End If
            Catch ex As Exception

            End Try
        Else
            ' andernfalls nichts machen .... 
        End If



    End Sub

    ''' <summary>
    ''' prüft, ob es sich um eine geschützte Präsentation handelt
    ''' kann über pwd, Computer, oder valid Login geschützt werden ... 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function userIsEntitled(ByRef msg As String, ByVal sld As PowerPoint.Slide) As Boolean

        Dim err As New clsErrorCodeMsg

        Dim tmpResult As Boolean = False
        Dim meldungen As New Collection

        ' sind die Zugangsdaten mit den aktuellen identisch ?
        ' wenn nein, dann wird noDBAccessInPPT auf false gesetzt 
        Call getDBsettings(sld)

        ' hier muss demnächst die Abfrage rein, ob der anwender auch die entsprechende customUserRole hat, um die Slide zu aktualisieren / zu sehen 
        ' ... Code ...
        ' Ende CustomUser Role Behandlung

        If noDBAccessInPPT Then

            noDBAccessInPPT = Not logInToMongoDB(True)

            If noDBAccessInPPT Then

                tmpResult = False
                If englishLanguage Then
                    msg = "no database access : " & awinSettings.databaseName
                Else
                    msg = "kein Datenbank Zugriff : " & awinSettings.databaseName
                End If
                Call MsgBox(msg)
            Else

                ' Username/Pwd in den Settings merken, falls Remember Me gecheckt
                Try
                    My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
                    If My.Settings.rememberUserPWD Then
                        My.Settings.userNamePWD = awinSettings.userNamePWD
                    End If
                    My.Settings.Save()
                Catch ex As Exception
                    Call MsgBox(ex.StackTrace)
                End Try

                ' CustomUserRoles holen 
                ' aber nur wenn es sich um eine Visbo-Server Version handelt ... 
                If awinSettings.visboServer = True Then
                    Dim allCustomUserRoles As clsCustomUserRoles = CType(databaseAcc, DBAccLayer.Request).retrieveCustomUserRoles(err)
                    allMyCustomUserRoles = allCustomUserRoles.getCustomUserRoles(dbUsername)


                Else
                    With myCustomUserRole
                        .userName = dbUsername
                        .customUserRole = ptCustomUserRoles.Alles
                        .specifics = ""
                    End With
                End If




                ' jetzt werden der Proxy Wert eingetragen, der beim letzten Mal funktioniert hat 
                If sld.Tags.Item("PRXYL").Length > 0 Then
                    sld.Tags.Delete("PRXYL")
                End If

                If awinSettings.proxyURL.Length > 0 Then
                    sld.Tags.Add("PRXYL", awinSettings.proxyURL)
                End If

                'Start of Calendar auslesen, damit Orga richtig interpretiert wird
                If sld.Tags.Item("SOC").Length > 0 Then
                    StartofCalendar = CDate(sld.Tags.Item("SOC"))
                End If

                tmpResult = True

                ' Lesen der Organisation aus der Datenbank direkt oder auch von DB
                Dim currentOrga As New clsOrganisation
                currentOrga = CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", Date.Now, False, err)

                If Not IsNothing(currentOrga) Then


                    If currentOrga.count > 0 Then
                        validOrganisations.addOrga(currentOrga)
                    End If

                End If


                If Not IsNothing(currentOrga) Then
                    ' hier müssen jetzt die Role- & Cost-Definitions gelesen werden 
                    RoleDefinitions = currentOrga.allRoles
                    CostDefinitions = currentOrga.allCosts
                End If

                ' Behandeln der myUserRole 
                ' jetzt wird die für die Slide passende Rolle gesucht 
                If awinSettings.visboServer = True Then
                    myCustomUserRole = getAppropriateUserRole(dbUsername, sld.Tags.Item("CURS"), meldungen)
                    If IsNothing(myCustomUserRole) Then

                        msg = "Error: Keine Berechtigung"
                        tmpResult = False
                    Else
                        tmpResult = True
                    End If


                    If meldungen.Count > 0 Then
                        Call MsgBox(meldungen.Item(1))
                        tmpResult = False
                    End If

                    If tmpResult Then
                        '' tk 5.2.20 evtl jetzt noch machen:  mit dieser USerRole nochmal die Top Nodes bauen 
                        'Call RoleDefinitions.buildTopNodes()
                        'Call RoleDefinitions.buildOrgaTeamChilds()

                        ' Auslesen der Custom Field Definitions aus den VCSettings über ReST-Server
                        Try
                            customFieldDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCustomFieldsFromDB(err)

                            If IsNothing(customFieldDefinitions) Then
                                customFieldDefinitions = New clsCustomFieldDefinitions
                                'Call MsgBox(err.errorMsg)
                            End If
                        Catch ex As Exception

                        End Try
                    End If
                End If

                ' in allen Slides den Sicht Schutz aufheben 
                protectionSolved = True
                Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoTrue)

            End If

        Else
            ' jetzt wird die für die Slide passende Rolle gesucht 
            If awinSettings.visboServer = True Then
                myCustomUserRole = getAppropriateUserRole(dbUsername, sld.Tags.Item("CURS"), meldungen)
                If IsNothing(myCustomUserRole) Then

                    msg = "Error: Keine Berechtigung"
                    tmpResult = False
                Else
                    tmpResult = True
                End If
            Else
                tmpResult = True
            End If

        End If


        userIsEntitled = tmpResult

    End Function


    ''' <summary>
    ''' gibt aus der Liste von customUserRoles, die dem Nutzer zugewiesen sind, die zurück, die zur Slide passt , was also die Slide erfordert ...
    ''' </summary>
    ''' <param name="myUserName"></param>
    ''' <param name="encryptedUserRoleString"></param>
    ''' <returns></returns>
    Public Function getAppropriateUserRole(ByVal myUserName As String, ByVal encryptedUserRoleString As String,
                                           ByRef meldungen As Collection) As clsCustomUserRole

        ' tk 25.1 kein Abprüfen der customUserRoles mehr ...

        Dim result As New clsCustomUserRole
        With result
            .userName = dbUsername
            .customUserRole = ptCustomUserRoles.OrgaAdmin
        End With

        'Dim result As clsCustomUserRole = Nothing

        'Dim err As New clsErrorCodeMsg

        'If IsNothing(allMyCustomUserRoles) Then
        '    Dim allCustomUserRoles As clsCustomUserRoles = CType(databaseAcc, DBAccLayer.Request).retrieveCustomUserRoles(err)
        '    If Not IsNothing(allCustomUserRoles) Then
        '        allMyCustomUserRoles = allCustomUserRoles.getCustomUserRoles(dbUsername)
        '    Else
        '        allMyCustomUserRoles = Nothing
        '    End If
        'End If

        'If encryptedUserRoleString = "" Then
        '    result = New clsCustomUserRole
        '    With result
        '        .userName = dbUsername
        '        .customUserRole = ptCustomUserRoles.Alles
        '    End With
        'Else
        '    If Not IsNothing(allMyCustomUserRoles) Then
        '        Dim pptUserRole As New clsCustomUserRole
        '        Call pptUserRole.decrypt(encryptedUserRoleString)

        '        Dim found As Boolean = False
        '        Dim ix As Integer = 0

        '        Do While ix <= allMyCustomUserRoles.Count - 1 And Not found
        '            If pptUserRole.customUserRole = ptCustomUserRoles.RessourceManager Then
        '                found = (CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = pptUserRole.customUserRole And
        '                    CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).specifics = pptUserRole.specifics) Or
        '                    CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = ptCustomUserRoles.PortfolioManager Or
        '                        CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = ptCustomUserRoles.ProjektLeitung Or
        '                        CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = ptCustomUserRoles.Alles Or
        '                        CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = ptCustomUserRoles.InternalViewer

        '            ElseIf CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = pptUserRole.customUserRole Then
        '                found = True

        '            ElseIf pptUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Or pptUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Then
        '                found = CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = ptCustomUserRoles.PortfolioManager Or
        '                        CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = ptCustomUserRoles.ProjektLeitung Or
        '                        CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = ptCustomUserRoles.Alles Or
        '                        CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole).customUserRole = ptCustomUserRoles.InternalViewer

        '            End If

        '            If found Then
        '                result = CType(allMyCustomUserRoles.Item(ix + 1), clsCustomUserRole)
        '                If result.customUserRole = ptCustomUserRoles.PortfolioManager Then
        '                    Dim IdArray() As Integer = result.getAggregationRoleIDs
        '                End If
        '            Else
        '                ix = ix + 1
        '            End If
        '        Loop

        '    Else
        '        result = Nothing
        '    End If
        'End If

        'If IsNothing(result) Then
        '    Dim msg As String = ""
        '    If awinSettings.englishLanguage Then
        '        msg = "Sorry, you don't have the required user role"
        '    Else
        '        msg = "Leider haben Sie nicht die geforderte Berechtigung"
        '    End If
        '    meldungen.Add(msg)
        'End If

        getAppropriateUserRole = result

    End Function



    Private Function userHasValidLicence() As Boolean
        userHasValidLicence = True
    End Function

    ''' <summary>
    ''' hier wird bestimmt, ob es sich um eine VisboProtected Präsentation handelt 
    ''' </summary>
    ''' <param name="Pres"></param>
    ''' <remarks></remarks>
    Private Sub pptAPP_AfterPresentationOpen(Pres As PowerPoint.Presentation) Handles pptAPP.AfterPresentationOpen


        ' hier muss nur weitergemacht werden, wenn es sich überhaupt um eine Presentation mit Smart Elements handelt 
        ' hier wird davon ausgegange, dass Setzen von CurrentPreHAsVisboElements zuverlässig in Window_Activate erfolgt 

        If presentationHasAnySmartSlides(Pres) Then


            ' ein ggf. vorhandener Schutz  muss wieder aktiviert werden ... 
            protectionSolved = False


            ' gibt es eine Sprachen-Tabelle ? 
            Dim langGUID As String = Pres.Tags.Item("langGUID")
            If langGUID.Length > 0 Then

                Dim langXMLpart As Office.CustomXMLPart = Pres.CustomXMLParts.SelectByID(langGUID)

                Dim langXMLstring = langXMLpart.XML
                languages = xml_deserialize(langXMLstring)

            End If

            Dim key As String = Pres.Name
            Dim hwinid As Integer = pptAPP.ActiveWindow.HWND

            ' Erweitern der listOFPresentation
            If Not listOfPresentations.ContainsKey(key) Then
                listOfPresentations.Add(key, hwinid)
            Else
                listOfPresentations.Item(key) = hwinid
            End If

            ' jetzt muss die Time Machine für diese Presentation mit leer angelegt werden 

            'Dim tmpTM As clsPPTTimeMachine = Nothing
            'If Not varPPTTM.ContainsKey(key) Then
            '    varPPTTM.Add(key, tmpTM)
            'Else
            '    varPPTTM.Item(key) = tmpTM
            'End If

            ' Anlegen einer leeren changeliste für jede Slide in der activePresentation
            ' key ist die SlideID in der besagten Presentation, clsChangeListe die Liste der Veränderungen zu dieser Seite 
            Dim slideChgListe As New SortedList(Of Integer, clsChangeListe)

            For Each slide As PowerPoint.Slide In Pres.Slides
                Dim chgelst As New clsChangeListe
                slideChgListe.Add(slide.SlideID, chgelst)
            Next

            ' in chgeLstListe sind für jede Presentation die slideChgListen 
            ' jetzt muss die chgListe ergänzt werden 
            If Not chgeLstListe.ContainsKey(key) Then
                chgeLstListe.Add(key, slideChgListe)
            Else
                chgeLstListe.Item(key) = slideChgListe
            End If


            ' tk 17.10.18 jetzt muss geprüft werden, ob eine der Slides smart-Infos enthält, wenigstens eine Slide nicht frozen ist und das aktuelle Datum der Slide vor dem heutigen Tag liegt 
            Dim atleastOne As Boolean = False
            Dim anzSlides As Integer = Pres.Slides.Count
            Dim ix As Integer = 1

            Do While ix <= anzSlides And Not atleastOne

                atleastOne = isSlideWithNeedToBeUpdated(Pres.Slides.Item(ix))
                ix = ix + 1

            Loop

            ' den Timestamp auslesen , ggf wird der ja nachher wieder beim Update umgesetzt 
            currentTimestamp = getCurrentTimestampFromPresentation(Pres)


            ' wenn evtl wenigstens eine Slide ge-updated werden muss
            ' die muss genau dann aktualisiert werden, wenn sie smart-Elements enthält, nicht bereits heute aktualisiert wurde und nicht frozen ist
            If atleastOne Then

                currentSlide = Pres.Slides.Item(ix - 1)
                Try
                    currentSlide.Select()
                Catch ex As Exception

                End Try

                Dim msgtxt As String = "there might be a newer Version" & vbLf & "than " & currentTimestamp.ToShortDateString & "." & vbLf & vbLf & "Do you want to update?"


                'Dim updateFrm As New frmUpdateInfo
                'With updateFrm
                '    .updateMsg.Text = msgtxt
                '    Dim diagResult As Windows.Forms.DialogResult = updateFrm.ShowDialog

                '    If diagResult = Windows.Forms.DialogResult.OK Then
                '        Dim tmpDate As Date = Date.MinValue
                '        Call updateSelectedSlide(ptNavigationButtons.update, tmpDate)

                '    End If
                'End With

            Else
                ' es wird jetzt gleich aus der Presi ausgelesen 
                ' trotzdem muss eine Slide zur CurrentSlide gemacht werden
                If anzSlides > 0 Then
                    currentSlide = Pres.Slides.Item(1)
                    currentSlide.Select()
                End If


            End If
        End If




    End Sub


    Private Sub pptAPP_PresentationBeforeClose(Pres As PowerPoint.Presentation, ByRef Cancel As Boolean) Handles pptAPP.PresentationBeforeClose

        If presentationHasAnySmartSlides(Pres) = True Then

            ' Id des aktiven Windows
            Dim key As String = Pres.Name
            Dim hwinid As Integer = pptAPP.ActiveWindow.HWND

            '' die Time Machine Settings löschen 
            'If varPPTTM.ContainsKey(key) Then
            '    varPPTTM.Remove(key)
            'End If

            ' die chgeliste aktualisieren , das heisst die 
            If chgeLstListe.ContainsKey(key) Then
                chgeLstListe.Remove(key)
            End If

            ' globale Variablen für Eigenschaften Pane und das Pane selbst löschen
            If listOfucProperties.ContainsKey(hwinid) Then
                listOfucProperties.Remove(hwinid)
            End If
            If listOfucPropView.ContainsKey(hwinid) Then
                listOfucPropView.Remove(hwinid)
            End If

            If Not IsNothing(currentSlide) Then

                ' changeliste der vorigen Slide (hier noch currentslide) in die chgeLstListe einfügen
                'If chgeLstListe.ContainsKey(currentSlide.SlideID) Then
                '    chgeLstListe.Remove(currentSlide.SlideID)
                '    chgeLstListe.Add(currentSlide.SlideID, changeListe)
                'Else
                '    chgeLstListe.Add(currentSlide.SlideID, changeListe)
                'End If

            End If

            If VisboProtected Then
                Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoFalse)
            End If

        End If

        ' tk 5.12.20 damit beim Schliessen nicht die alten Listen erhalten bleiben 
        smartSlideLists = New clsSmartSlideListen

        My.Settings.Save()


    End Sub

    Private Sub pptAPP_PresentationBeforeSave(Pres As PowerPoint.Presentation, ByRef Cancel As Boolean) Handles pptAPP.PresentationBeforeSave

        Dim a As Integer = 0
        Dim beforeSlideTS As Date = Date.MinValue
        Dim activeSlideTS As Date = Date.MinValue
        Dim sld As PowerPoint.Slide
        Dim canBeSaved As Boolean = True

        If Not IsNothing(Pres) And Pres.Slides.Count > 1 Then
            ' Vorbesetzung
            sld = Pres.Slides.Item(1)
            If isVisboSlide(sld) Then
                activeSlideTS = getCurrentTimeStampFromSlide(sld)
            End If

            For i = 2 To Pres.Slides.Count

                beforeSlideTS = activeSlideTS
                sld = Pres.Slides.Item(i)

                If isVisboSlide(sld) Then

                    activeSlideTS = getCurrentTimeStampFromSlide(sld)

                    If activeSlideTS <> beforeSlideTS Then
                        canBeSaved = False
                        Exit For
                    End If

                End If

            Next

            If beforeSlideTS > Date.MinValue And Not canBeSaved Then
                Cancel = True
                If awinSettings.englishLanguage Then
                    Call MsgBox("Not saved!" & vbCrLf &
                                "This presentation contains slides with different timestamps!")
                Else
                    Call MsgBox("Speichern nicht sinnvoll!" & vbCrLf &
                                "Die Präsentation enthält Seiten mit unterschiedlichem Timestamp")
                End If
            ElseIf beforeSlideTS = Date.MinValue Then
                If awinSettings.visboDebug Then
                    Call MsgBox("An Error has occurred in PresentationBeforeSave")
                End If
            ElseIf canBeSaved Then
                ' nothing to do
                ' presentation will be saved
            End If
        End If

        ' wenn VisboProtected, dann müssen jetzt alle relevanten Shapes auf invisible gesetzt werden ...
        If VisboProtected Then
            Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoFalse)
        End If

    End Sub


    ''' <summary>
    ''' ein VISBO Protected File kann nur als pptx gespeichert werden ...
    ''' </summary>
    ''' <param name="Pres"></param>
    ''' <remarks></remarks>
    Private Sub pptAPP_PresentationSave(Pres As PowerPoint.Presentation) Handles pptAPP.PresentationSave
        If VisboProtected And Not Pres.Name.EndsWith(".pptx") Then
            If englishLanguage Then
                Call MsgBox("Save only possible with file extension .pptx !")
            Else
                Call MsgBox("Speichern nur als .pptx möglich!")
            End If

            Dim vollerName As String = Pres.FullName
            Dim correctName As String = Pres.Name & ".pptx"

            Pres.SaveAs(correctName)
            My.Computer.FileSystem.DeleteFile(vollerName)
        End If
    End Sub

    Private Sub pptAPP_PresentationPrint(Pres As PowerPoint.Presentation) Handles pptAPP.PresentationPrint

        Dim beforeSlideTS As Date = Date.MinValue
        Dim activeSlideTS As Date = Date.MinValue
        Dim sld As PowerPoint.Slide
        Dim canBeSaved As Boolean = True

        If Not IsNothing(Pres) And Pres.Slides.Count > 1 Then
            ' Vorbesetzung
            sld = Pres.Slides.Item(1)
            activeSlideTS = getCurrentTimeStampFromSlide(sld)

            For i = 2 To Pres.Slides.Count

                beforeSlideTS = activeSlideTS
                sld = Pres.Slides.Item(i)
                activeSlideTS = getCurrentTimeStampFromSlide(sld)

                If activeSlideTS <> beforeSlideTS Then
                    canBeSaved = False
                    Exit For
                End If
            Next

            If beforeSlideTS > Date.MinValue And Not canBeSaved Then

                If awinSettings.englishLanguage Then
                    Call MsgBox("Warning!" & vbCrLf &
                                "This presentation contains slides with different timestamps!")
                Else
                    Call MsgBox("Achtung!" & vbCrLf &
                                "Die Präsentation enthält Seiten mit unterschiedlichem Timestamp")
                End If
            ElseIf beforeSlideTS = Date.MinValue Then
                If awinSettings.visboDebug Then
                    Call MsgBox("An Error has occurred in PresentationBeforeSave")
                End If
            ElseIf canBeSaved Then
                ' nothing to do
                ' presentation will be saved
            End If
        End If
    End Sub


    ''' <summary>
    ''' wird aufgerufen, sobald der User eine andere Slide wählt
    ''' wenn er mehrere selektiert, wird nichts gemacht ... 
    ''' </summary>
    ''' <param name="SldRange"></param>
    ''' <remarks></remarks>
    Private Sub pptAPP_SlideSelectionChanged(SldRange As PowerPoint.SlideRange) Handles pptAPP.SlideSelectionChanged

        ' die aktuelle Slide setzen 
        If SldRange.Count = 1 Then
            Call own_SlideSelectionChanged(SldRange.Item(1))
        End If


        'If SldRange.Count = 1 Then

        '    If currentPresHasVISBOElements Then
        '        ' nur dann muss irgendwas weitergemacht werden ..

        '        Dim afterSlideID As Integer = SldRange.Item(1).SlideID ' aktuell selektierte SlideID

        '        ' hier muss nur weitergemacht werden, wenn es sich um eine VISBO slide handelt 
        '        If isVisboSlide(SldRange.Item(1)) Then

        '            Dim afterSlideKennung As String = CType(SldRange.Item(1).Parent, PowerPoint.Presentation).Name & afterSlideID.ToString
        '            Dim beforeSlideKennung As String = ""

        '            Dim key As String = CType(SldRange.Item(1).Parent, PowerPoint.Presentation).Name

        '            Dim beforeSlideID As Integer = 0               ' zuvor selektierte SlideID

        '            If Not IsNothing(currentSlide) Then
        '                Try
        '                    beforeSlideID = currentSlide.SlideID
        '                    beforeSlideKennung = CType(currentSlide.Parent, PowerPoint.Presentation).Name & beforeSlideID.ToString
        '                Catch ex As Exception

        '                End Try

        '            End If

        '            '' jetzt die CurrentSlide setzen , denn evtl kommt man ja gar nicht in pptAPP_UpdateOneSlide
        '            currentSlide = SldRange.Item(1)

        '            If beforeSlideKennung <> afterSlideKennung Or smartSlideLists.countProjects = 0 Then
        '                Call pptAPP_AufbauSmartSlideLists(SldRange.Item(1))

        '                'If varPPTTM.ContainsKey(key) Then
        '                '    ' fertig ... 
        '                'Else
        '                '    Dim tmpTM As clsPPTTimeMachine = Nothing
        '                '    varPPTTM.Add(key, tmpTM)
        '                'End If

        '            End If

        '            ' jetzt die currentTimeStamp setzen 
        '            With currentSlide
        '                If .Tags.Item("CRD").Length > 0 Then
        '                    currentTimestamp = getCurrentTimeStampFromSlide(currentSlide)
        '                End If
        '            End With


        '            ' nur wenn die SlideID gewechselt hat, muss agiert werden
        '            ' dabei auch berücksichtigen, ob sich Presentation geändert hat 
        '            If beforeSlideKennung <> afterSlideKennung Then
        '                Try
        '                    ' das Change-Formular aktualisieren, wenn es gezeigt wird  
        '                    Dim hwind As Integer = pptAPP.ActiveWindow.HWND
        '                    If Not IsNothing(changeFrm) Then

        '                        changeFrm.changeliste.clearChangeList()

        '                        If chgeLstListe.ContainsKey(key) Then
        '                            If chgeLstListe.Item(key).ContainsKey(currentSlide.SlideID) Then
        '                                changeFrm.changeliste = chgeLstListe.Item(key).Item(currentSlide.SlideID)
        '                            Else
        '                                ' eine Liste für die neue SlideID einfügen ..
        '                            End If
        '                        End If

        '                        changeFrm.neuAufbau()
        '                    End If
        '                Catch ex As Exception

        '                End Try

        '            End If       'Ende ob SlideIDs ungleich sind
        '        Else
        '            'ur: ???
        '            'currentSlide = Nothing
        '        End If
        '    Else
        '        'ur: ???
        '        'currentSlide = Nothing
        '    End If ' if currentPresHasVisboElements

        'Else
        '    ' nichts tun, das heisst auch nichts verändern ...
        'End If

    End Sub

    ''' <summary>
    ''' bestimmt die Settings der Datenbank, sofern welche da sind 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub getDBsettings(ByVal sld As PowerPoint.Slide)

        With sld

            If .Tags.Item("DBURL").Length > 0 And
                .Tags.Item("DBNAME").Length > 0 Then



                If ((.Tags.Item("DBURL") = awinSettings.databaseURL And
                        .Tags.Item("DBNAME") = awinSettings.databaseName) Or
                        (.Tags.Item("DBURL") = awinSettings.databaseURL And
                        .Tags.Item("VCid") = awinSettings.VCid)) And
                        Not noDBAccessInPPT Then
                    ' nichts machen, user ist schon berechtigt ...
                Else
                    noDBAccessInPPT = True
                    awinSettings.proxyURL = .Tags.Item("PRXYL")
                    awinSettings.databaseURL = .Tags.Item("DBURL")
                    awinSettings.databaseName = .Tags.Item("DBNAME")
                    awinSettings.VCid = .Tags.Item("VCid")
                    awinSettings.DBWithSSL = (.Tags.Item("DBSSL") = "True")
                    awinSettings.visboServer = (.Tags.Item("REST") = "True")
                End If


            End If
        End With
    End Sub

    ''' <summary>
    ''' liefert den current Timestamp einer Präsentation zurück 
    ''' dabei wird der Timestamp der ersten Folie zurück geleifertm die Smart Elements enthält und nicht frozen ist 
    ''' </summary>
    ''' <param name="pres"></param>
    ''' <returns></returns>
    Friend Function getCurrentTimestampFromPresentation(ByVal pres As PowerPoint.Presentation) As Date

        Dim tmpresult As Date = currentTimestamp

        For Each sld As PowerPoint.Slide In pres.Slides

            With sld
                If .Tags.Item("SMART") = "visbo" Then
                    If .Tags.Item("FROZEN").Length = 0 Then
                        If .Tags.Item("CRD").Length > 0 Then
                            tmpresult = CDate(.Tags.Item("CRD"))
                            Exit For
                        End If
                    End If

                End If
            End With

        Next

        getCurrentTimestampFromPresentation = tmpresult

    End Function

    ''' <summary>
    ''' liefetr den currentTimestamp der Seite zurück, wenn er existiert
    ''' wenn die Seite keinen enthält bleibt der Wert unveräändert auf dem bisherigen currentTimestamp
    ''' </summary>
    ''' <param name="sld"></param>
    ''' <returns></returns>
    Friend Function getCurrentTimeStampFromSlide(ByVal sld As PowerPoint.Slide) As Date

        Dim tmpresult As Date = currentTimestamp

        With sld
            If .Tags.Item("SMART") = "visbo" Then
                If .Tags.Item("FROZEN").Length = 0 Then
                    If .Tags.Item("CRD").Length > 0 Then
                        tmpresult = CDate(.Tags.Item("CRD"))
                    End If
                End If

            End If
        End With

        getCurrentTimeStampFromSlide = tmpresult
    End Function

    ''' <summary>
    ''' setzt in der aktuellen Slide den Timestamp 
    ''' </summary>
    ''' <param name="ts"></param>
    ''' <remarks></remarks>
    Friend Sub setCurrentTimestampInSlide(ByVal ts As Date)
        ' jetzt in der currentSlide den CRD setzen ..
        With currentSlide
            ' currentTimeStamp setzen 
            If .Tags.Item("CRD").Length > 0 Then
                .Tags.Delete("CRD")
            End If
            .Tags.Add("CRD", ts.ToString)
        End With
    End Sub
    ''' <summary>
    ''' setzt in der aktuellen Slide den previous Timestamp 
    ''' </summary>
    ''' <param name="ts"></param>
    ''' <remarks></remarks>
    Friend Sub setPreviousTimestampInSlide(ByVal ts As Date)
        ' jetzt in der currentSlide den CRD setzen ..
        With currentSlide

            ' PreviousTimeStamp setzen 
            If .Tags.Item("PREV").Length > 0 Then
                .Tags.Delete("PREV")
            End If
            .Tags.Add("PREV", ts.ToString)

        End With
    End Sub
    ''' <summary>
    ''' erstellt die SmartSlideListen neu ... 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub buildSmartSlideLists()

        Dim err As New clsErrorCodeMsg
        Dim vpid As String = ""

        '' vorherige smartSlideLists zwischenspeichern
        'Dim former_smartSlideLists As clsSmartSlideListen = smartSlideLists

        ' zurücksetzen der SmartSlideLists
        smartSlideLists = New clsSmartSlideListen

        ' das ist jetzt nicht mehr notwendig - die Projekte und Timestamps werden in der visboTimeMachine Variablen gehalten 
        '' wenn bereits die tsCollection existiert, müssen ListOfTS und ListOfProjektHistorien gesichert werden
        'If tsCollExists Then
        '    smartSlideLists.ListOfProjektHistorien = former_smartSlideLists.ListOfProjektHistorien
        '    smartSlideLists.ListOfTS = former_smartSlideLists.ListOfTS
        'End If

        bekannteIDs = New SortedList(Of Integer, String)

        Dim aktSlideId As Integer = currentSlide.SlideID

        '
        ' Definition der importantShapes und der entsprechenden Enumertaion in Module 1 
        ReDim importantShapes(1)
        '
        ' zurücksetzen der importantShapes 

        importantShapes(ptImportantShapes.todayline) = Nothing
        importantShapes(ptImportantShapes.version) = Nothing

        ' jetzt todayLine suchen ...
        Try
            importantShapes(ptImportantShapes.todayline) = currentSlide.Shapes.Item("todayLine")
        Catch ex As Exception
            importantShapes(ptImportantShapes.todayline) = Nothing
        End Try
        'ur: 2019-05-29: TryCatch vermeiden
        'For i = 1 To currentSlide.Shapes.Count
        '    If currentSlide.Shapes.Item(i).Name = "todayLine" Then

        '        importantShapes(ptImportantShapes.todayline) = currentSlide.Shapes.Item("todayLine")
        '        Exit For
        '    Else
        '        importantShapes(ptImportantShapes.todayline) = Nothing
        '    End If
        'Next


        With currentSlide
            If .Tags.Item("CRD").Length > 0 Then
                smartSlideLists.creationDate = CDate(.Tags.Item("CRD"))
            End If
            If .Tags.Item("PREV").Length > 0 Then
                smartSlideLists.prevDate = CDate(.Tags.Item("PREV"))
            End If
            If .Tags.Item("VCid").Length > 0 Then
                smartSlideLists.slideVCid = .Tags.Item("VCid")
            End If

            If .Tags.Item("DBURL").Length > 0 And
                .Tags.Item("DBNAME").Length > 0 Then

                smartSlideLists.slideDBName = .Tags.Item("DBNAME")
                smartSlideLists.slideDBUrl = .Tags.Item("DBURL")

                If awinSettings.databaseURL <> smartSlideLists.slideDBUrl Or
                    (awinSettings.databaseName <> smartSlideLists.slideDBName And
                    awinSettings.VCid <> smartSlideLists.slideVCid) Then

                    noDBAccessInPPT = True
                    awinSettings.databaseURL = smartSlideLists.slideDBUrl
                    awinSettings.databaseName = smartSlideLists.slideDBName
                    awinSettings.VCid = smartSlideLists.slideVCid

                End If
            End If


            If .Tags.Item("REST").Length > 0 Then
                awinSettings.visboServer = .Tags.Item("REST") = "True"
            End If



        End With

        Dim anzShapes As Integer = currentSlide.Shapes.Count
        ' jetzt werden die ganzen Listen aufgebaut 

        Dim bigToDoList As New Collection
        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            bigToDoList.Add(tmpShape.Name)
        Next

        For Each tmpShpName As String In bigToDoList
            Try
                Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                If Not IsNothing(tmpShape) Then

                    If tmpShape.Tags.Item("BID").Length > 0 And tmpShape.Tags.Item("DID").Length > 0 Then

                        ' handelt es sich um das Version Field Shape? 
                        If tmpShape.Tags.Item("BID") = CStr(CInt(ptReportBigTypes.components)) _
                            And (tmpShape.Tags.Item("DID") = CStr(CInt(ptReportComponents.prStand)) Or tmpShape.Tags.Item("DID") = CStr(CInt(ptReportComponents.pfStand))) Then
                            importantShapes(ptImportantShapes.version) = tmpShape
                        End If

                        Dim projType As ptPRPFType = ptPRPFType.project

                        If tmpShape.Tags.Item("PRPF").Length > 0 Then
                            projType = CType(tmpShape.Tags.Item("PRPF"), ptPRPFType)
                        Else

                        End If

                        Dim pvName As String = ""

                        If projType = ptPRPFType.project Then
                            If tmpShape.Tags.Item("PNM").Length > 0 Then
                                Dim pName As String = tmpShape.Tags.Item("PNM")
                                Dim vName As String = tmpShape.Tags.Item("VNM")
                                pvName = calcProjektKey(pName, vName)
                            End If
                            If tmpShape.Tags.Item("VPID").Length > 0 Then
                                vpid = tmpShape.Tags.Item("VPID")
                            End If
                        ElseIf projType = ptPRPFType.portfolio Then
                            If tmpShape.Tags.Item("PNM").Length > 0 Then
                                Dim pName As String = tmpShape.Tags.Item("PNM")
                                Dim vName As String = tmpShape.Tags.Item("VNM")
                                pvName = calcPortfolioKey(pName, vName)
                            End If
                            If tmpShape.Tags.Item("VPID").Length > 0 Then
                                vpid = tmpShape.Tags.Item("VPID")
                            End If
                        End If

                        If tmpShape.Tags.Item("BID") = CStr(CInt(ptReportBigTypes.components)) _
                                And (tmpShape.Tags.Item("DID") = CStr(CInt(ptReportComponents.pfName))) Then

                            'tmpShape ist ein Componente , wenn Portfoliochart, dann muss verwendetes Porfolio (TAG: PNM und/oder VPID) in _portfoliolist aufgenommen werden
                            ' das ist in einem Tag im tmpshape enthalten

                            If pvName <> "" Then
                                If tmpShape.Tags.Item("VPID").Length > 0 Then
                                    vpid = tmpShape.Tags.Item("VPID")
                                End If
                                If smartSlideLists.containsPortfolio(pvName, vpid) Then
                                    ' nichts tun, ist schon drin ..
                                Else
                                    smartSlideLists.addPortfolio(pvName, vpid)
                                End If
                            End If


                        ElseIf tmpShape.Tags.Item("BID") = CStr(CInt(ptReportBigTypes.components)) _
                                And (tmpShape.Tags.Item("DID") = CStr(CInt(ptReportComponents.prName))) Then


                            ' um zu berücksichtigen, dass auch Slides ohne Meilensteine / Phasen als Smart-Slides aufgefasst werden ...

                            If pvName <> "" Then
                                If smartSlideLists.containsProject(pvName, vpid) Then
                                    ' nichts tun, ist schon drin ..
                                Else
                                    smartSlideLists.addProject(pvName, vpid)
                                End If
                            End If

                        End If
                    End If


                    If tmpShape.Tags.Count > 0 Then

                        If isRelevantMSPHShape(tmpShape) Or isProjectCard(tmpShape) Then


                            Dim isPcardInvisible As Boolean = isProjectCardInvisible(tmpShape)
                            If isPcardInvisible Then
                                Dim a As Integer = 10
                            End If

                            bekannteIDs.Add(tmpShape.Id, tmpShape.Name)

                            Call aktualisiereSortedLists(tmpShape)

                            ' tk 27.3.20 darf nicht wieder auf visible gesetzt werden ... 
                            ' da sonst das unsichtbar machen von Phasen / Meilensteinen konterkariert wird  
                            'If Not isPcardInvisible Then
                            '    If protectionSolved And tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse Then
                            '        tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                            '    End If
                            'End If


                        ElseIf isVISBOChartElement(tmpShape) Then
                            If protectionSolved And tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse Then
                                tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                            End If

                            'tmpShape ist ein Chart , wenn Portfoliochart, dann muss verwendetes Porfolio (TAG: PNM und/oder VPID) in _portfoliolist aufgenommen werden
                            ' das ist in einem Tag im tmpshape enthalten
                            If tmpShape.Tags.Item("PRPF").Length > 0 Then
                                If CType(tmpShape.Tags.Item("PRPF"), ptPRPFType) = ptPRPFType.portfolio Then
                                    Dim pfName As String = ""
                                    If tmpShape.Tags.Item("PNM").Length > 0 Then
                                        Dim pName As String = tmpShape.Tags.Item("PNM")
                                        Dim vName As String = tmpShape.Tags.Item("VNM")
                                        'pvName = calcProjektKey(pName, vName)
                                        pfName = pName
                                    End If
                                    If tmpShape.Tags.Item("VPID").Length > 0 Then
                                        vpid = tmpShape.Tags.Item("VPID")
                                    End If
                                    If pfName <> "" Then
                                        If smartSlideLists.containsPortfolio(pfName, vpid) Then
                                            ' nichts tun, ist schon drin ..
                                        Else
                                            smartSlideLists.addPortfolio(pfName, vpid)
                                        End If
                                    End If
                                Else
                                End If
                            End If

                        End If
                    End If
                End If

            Catch ex As Exception

            End Try
        Next



        If Not noDBAccessInPPT Then

            ' hier müssen jetzt die Timestamps noch aufgebaut werden 
            For i As Integer = 1 To smartSlideLists.countProjects

                Dim pvName As String = smartSlideLists.getPVName(i)
                vpid = smartSlideLists.getvpID(pvName)

                If Not timeMachine.containsProject(pvName, vpid) Then
                    timeMachine.addProject(pvName, vpid)
                End If
                'Dim tsCollection As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveZeitstempelFirstLastFromDB(pvName, err)
                'smartSlideLists.addToListOfTS(tsCollection)
            Next


            For Each tmpShpName As String In bigToDoList
                Try
                    Dim pvname As String = getPVnameFromShpName(tmpShpName)
                    If pvname <> "" Then
                        Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                        If Not IsNothing(tmpShape) Then
                            If tmpShape.Tags.Count > 0 Then
                                If isRelevantMSPHShape(tmpShape) Then

                                    Call aktualisiereRoleCostLists(tmpShape)

                                End If
                            End If
                        End If
                    End If

                Catch ex As Exception

                End Try

            Next

        End If

    End Sub

    Private Sub pptAPP_WindowActivate(Pres As Microsoft.Office.Interop.PowerPoint.Presentation, Wn As PowerPoint.DocumentWindow) Handles pptAPP.WindowActivate
        ' Id des aktiven DocumentWindow

        Dim key As String = Pres.Name
        currentPresHasVISBOElements = presentationHasAnySmartSlides(Pres)

        ' die müssen zurück gesetzt werden , weil neue PResentation 
        ' müssen auch erst gesetzt werden, wenn neue Slide
        currentSldHasProjectTemplates = False
        currentSldHasPortfolioTemplates = False


        If currentPresHasVISBOElements Then

            currentTimestamp = getCurrentTimestampFromPresentation(Pres)

            Try
                '
                ' setzen der current und previous timestamps 
                If IsNothing(rememberListOfCPTimeStamps) Then
                    ' ... sind die curent und previous Timestamps ja initial gesetzt ...
                    rememberListOfCPTimeStamps = New SortedList(Of String, Date())
                    Dim tmpDates(1) As Date
                    tmpDates(0) = previousTimeStamp
                    tmpDates(1) = currentTimestamp
                    rememberListOfCPTimeStamps.Add(key, tmpDates)
                Else
                    If rememberListOfCPTimeStamps.ContainsKey(key) Then
                        previousTimeStamp = rememberListOfCPTimeStamps.Item(key)(0)
                        currentTimestamp = rememberListOfCPTimeStamps.Item(key)(1)
                    Else
                        ' das setzen, was initial gesetzt wird ... 
                        currentTimestamp = Date.MinValue
                        previousTimeStamp = Date.MinValue
                    End If
                End If

                '
                ' setzen der current und previous VariantNames  
                If IsNothing(rememberListOfCPVariantNames) Then
                    ' ... sind die curent und previous Timestamps ja initial gesetzt ...
                    rememberListOfCPVariantNames = New SortedList(Of String, String())
                    Dim tmpVnames(1) As String
                    tmpVnames(0) = previousVariantName
                    tmpVnames(1) = currentVariantname
                    rememberListOfCPVariantNames.Add(key, tmpVnames)
                Else
                    If rememberListOfCPVariantNames.ContainsKey(key) Then
                        previousVariantName = rememberListOfCPVariantNames.Item(key)(0)
                        currentVariantname = rememberListOfCPVariantNames.Item(key)(1)
                    Else
                        ' das setzen, was initial gesetzt wird ...
                        currentVariantname = ""
                        previousVariantName = noVariantName
                    End If
                End If

            Catch ex As Exception

            End Try


            ' globale Variablen für Eigenschaften Pane umsetzen
            If listOfucProperties.ContainsKey(Wn.HWND) Then
                propertiesPane = listOfucProperties.Item(Wn.HWND)
            End If
            If listOfucPropView.ContainsKey(Wn.HWND) Then
                ucPropertiesView = listOfucPropView.Item(Wn.HWND)
            End If

            ' globale Variable für search pane umsetzen
            If listOfucSearch.ContainsKey(Wn.HWND) Then
                searchPane = listOfucSearch.Item(Wn.HWND)
            End If
            If listOfucSearchView.ContainsKey(Wn.HWND) Then
                ucSearchView = listOfucSearchView.Item(Wn.HWND)
            End If


        End If


    End Sub

    Private Sub pptAPP_WindowDeactivate(Pres As PowerPoint.Presentation, Wn As PowerPoint.DocumentWindow) Handles pptAPP.WindowDeactivate

        Dim key As String = Pres.Name

        If currentPresHasVISBOElements Then

            Try
                ' setzen der current und previous timestamps 
                If Not IsNothing(rememberListOfCPTimeStamps) Then
                    ' ... sind die curent und previous Timestamps ja initial gesetzt ...
                    If rememberListOfCPTimeStamps.ContainsKey(key) Then
                        rememberListOfCPTimeStamps.Item(key)(0) = previousTimeStamp
                        rememberListOfCPTimeStamps.Item(key)(1) = currentTimestamp
                    Else
                        ' einfügen 
                        Dim tmpDates(1) As Date
                        tmpDates(0) = previousTimeStamp
                        tmpDates(1) = currentTimestamp
                        rememberListOfCPTimeStamps.Add(key, tmpDates)
                    End If

                Else
                    ' ... sind die curent und previous Timestamps ja initial gesetzt ...
                    rememberListOfCPTimeStamps = New SortedList(Of String, Date())
                    Dim tmpDates(1) As Date
                    tmpDates(0) = previousTimeStamp
                    tmpDates(1) = currentTimestamp
                    rememberListOfCPTimeStamps.Add(key, tmpDates)
                End If

                '
                ' setzen der current und previous VariantNames  
                If Not IsNothing(rememberListOfCPVariantNames) Then
                    ' ... sind die curent und previous Timestamps ja initial gesetzt ...
                    If rememberListOfCPVariantNames.ContainsKey(key) Then
                        rememberListOfCPVariantNames.Item(key)(0) = previousVariantName
                        rememberListOfCPVariantNames.Item(key)(1) = currentVariantname
                    Else
                        ' einfügen 
                        Dim tmpVnames(1) As String
                        tmpVnames(0) = previousVariantName
                        tmpVnames(1) = currentVariantname
                        rememberListOfCPVariantNames.Add(key, tmpVnames)
                    End If

                Else
                    ' ... sind die curent und previous Timestamps ja initial gesetzt ...
                    rememberListOfCPVariantNames = New SortedList(Of String, String())
                    Dim tmpVnames(1) As String
                    tmpVnames(0) = previousVariantName
                    tmpVnames(1) = currentVariantname
                    rememberListOfCPVariantNames.Add(key, tmpVnames)
                End If

            Catch ex As Exception

            End Try

            ' wenn geschützt, dann unsichtbar machen der relecanten Shapes 
            If VisboProtected Then
                Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoFalse)
            End If
        Else
            ' auf false setzen, weil das in der nächsten Activate Routine bestimmt wird ... 
            currentPresHasVISBOElements = False
        End If

    End Sub

    Private Sub pptAPP_WindowSelectionChange(Sel As PowerPoint.Selection) Handles pptAPP.WindowSelectionChange

        'Dim relevantShape As PowerPoint.Shape
        Dim arrayOfNames() As String
        Dim relevantShapeNames As New Collection


        selectedPlanShapes = Nothing

        ' alles weitere nur machen, wenn überhaupt Smart-Element enthalten sind 
        If currentPresHasVISBOElements Then

            Try
                Select Case Sel.Type
                    Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes

                        Dim shpRange As PowerPoint.ShapeRange = Sel.ShapeRange

                        If Not IsNothing(shpRange) And slideHasSmartElements Then

                            ' jetzt muss hier die Behandlung für Office 2010 rein 
                            Dim correctErrorShape1 As PowerPoint.Shape = Nothing
                            Dim correctErrorShape2 As PowerPoint.Shape = Nothing

                            ' nur was machen, wenn es sich um Office 2010 handelt ... 
                            ' werden temporäre Shapes erzeugt und selektiert, die wiederum einen SelectionChange erzeugen
                            ' dabei wird das ursprünglich selektierte Shape gemerkt udn am Schluss, wenn das Property Window angezeigt ist, 
                            ' wieder selektiert .. das alles muss aber nur im Fall Version = 14.0 gemacht werden 
                            If pptAPP.Version = "14.0" Then
                                Try
                                    correctErrorShape1 = currentSlide.Shapes("visboCorrectError1")
                                Catch ex As Exception

                                End Try

                                Try
                                    correctErrorShape2 = currentSlide.Shapes("visboCorrectError2")
                                Catch ex As Exception

                                End Try
                            End If


                            If ((pptAPP.Version = "14.0") And
                                (((Not propertiesPane.Visible) Or
                                (propertiesPane.Visible And Not IsNothing(correctErrorShape1)) Or
                                (propertiesPane.Visible And Not IsNothing(correctErrorShape2))))) Then
                                ' Erzeugen eines Hilfs-Elements

                                ' Ist es 
                                If IsNothing(correctErrorShape1) And IsNothing(correctErrorShape2) And Not isRelevantMSPHShape(shpRange(1)) Then
                                    ' nichts machen 
                                Else
                                    If IsNothing(correctErrorShape1) Then
                                        ' erzeugen und selektieren der beiden Shapes  
                                        Dim oldShape As PowerPoint.Shape = shpRange(1)

                                        Dim helpShape1 As PowerPoint.Shape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                                                                   0, 0, 50, 50)



                                        If Not IsNothing(helpShape1) Then
                                            helpShape1.Name = "visboCorrectError1"
                                            helpShape1.Tags.Add("formerSN", oldShape.Name)
                                            helpShape1.Select()
                                        End If



                                    ElseIf IsNothing(correctErrorShape2) Then

                                        ' jetzt die zweite Welle 
                                        propertiesPane.Visible = True

                                        Dim helpShape2 As PowerPoint.Shape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                                                                   0, 0, 50, 50)
                                        If Not IsNothing(helpShape2) Then
                                            helpShape2.Name = "visboCorrectError2"
                                            helpShape2.Select()
                                        End If
                                    Else


                                        ' Selektieren des vorher geklickten shapes 
                                        Dim formerShapeName As String = correctErrorShape1.Tags.Item("formerSN")
                                        Dim formerSelShape As PowerPoint.Shape = Nothing

                                        If formerShapeName.Length > 0 Then
                                            Try

                                                formerSelShape = currentSlide.Shapes(formerShapeName)

                                                ' Löschen der Hilfs-Shapes 
                                                correctErrorShape1.Delete()
                                                correctErrorShape2.Delete()

                                                ' Selektieren des formerShapes
                                                formerSelShape.Select()

                                            Catch ex As Exception

                                            End Try

                                        End If

                                    End If
                                End If

                            Else
                                ' es sind ein oder mehrere Shapes selektiert worden 
                                Dim i As Integer = 0
                                If shpRange.Count = 1 Then

                                    ' prüfen, ob inzwischen was selektiert wurde, was nicht zu der Selektion in der 
                                    ' Listbox passt 

                                    '' '' prüfen, ob das Info Fenster offen ist und der Search bereich sichtbar - 
                                    '' '' dann muss der Klarheit wegen die Listbox neu aufgebaut werden 
                                    ' ''If Not IsNothing(infoFrm) And formIsShown Then
                                    ' ''    If infoFrm.rdbName.Visible Then
                                    ' ''        If infoFrm.listboxNames.SelectedItems.Count > 0 Then
                                    ' ''            'Call infoFrm.listboxNames.SelectedItems.Clear()
                                    ' ''        End If
                                    ' ''    End If
                                    ' ''End If


                                    ' jetzt ggf die angezeigten Marker löschen 
                                    If Not markerShpNames.ContainsKey(shpRange(1).Name) Then
                                        Call deleteMarkerShapes()
                                    ElseIf markerShpNames.Count > 1 Then
                                        Call deleteMarkerShapes(shpRange(1).Name)
                                    End If

                                    ' prüfen, ob es ein Kommentar ist 
                                    Dim tmpShape As PowerPoint.Shape = shpRange(1)
                                    If isCommentShape(tmpShape) Then
                                        Call markReferenceShape(tmpShape.Name)
                                    End If
                                ElseIf shpRange.Count > 1 Then
                                    ' für jedes Shape prüfen, ob es ein Comment Shape ist .. 
                                    For Each tmpShape As PowerPoint.Shape In shpRange
                                        If isCommentShape(tmpShape) Then
                                            Call markReferenceShape(tmpShape.Name)
                                        End If
                                    Next
                                ElseIf shpRange.Count = 0 Then

                                    Call deleteMarkerShapes()

                                End If


                                For Each tmpShape As PowerPoint.Shape In shpRange

                                    If tmpShape.Tags.Count > 0 Then

                                        'If tmpShape.AlternativeText <> "" And tmpShape.Title <> "" Then

                                        If isRelevantShape(tmpShape) Then
                                            If bekannteIDs.ContainsKey(tmpShape.Id) Or
                                        tmpShape.Name.EndsWith(shadowName) Then

                                                If Not relevantShapeNames.Contains(tmpShape.Name) Then
                                                    relevantShapeNames.Add(tmpShape.Name, tmpShape.Name)
                                                End If

                                            Else
                                                ' die vorhandenen Tags löschen ... und den Namen ändern 
                                                Call deleteShpTags(tmpShape)
                                            End If

                                        End If

                                    End If


                                Next

                                '' Anfang ... das war vorher innerhalb der next Schleife .. 
                                ' jetzt muss geprüft werden, ob relevantShapeNames mindestens ein Element enthält ..
                                If relevantShapeNames.Count >= 1 Then

                                    ReDim arrayOfNames(relevantShapeNames.Count - 1)

                                    For ix As Integer = 1 To relevantShapeNames.Count
                                        arrayOfNames(ix - 1) = CStr(relevantShapeNames(ix))
                                    Next

                                    selectedPlanShapes = currentSlide.Shapes.Range(arrayOfNames)

                                ElseIf isSymbolShape(shpRange(1)) Then

                                    selectedPlanShapes = shpRange
                                    Call aktualisiereInfoPane(shpRange(1))

                                Else
                                    ' in diesem Fall wurden nur nicht-relevante Shapes selektiert 
                                    Call checkHomeChangeBtnEnablement()
                                    Try
                                        If propertiesPane.Visible Then
                                            Call aktualisiereInfoPane(Nothing)
                                        End If
                                    Catch ex As Exception

                                    End Try

                                    ' ur: wegen Pane
                                    ' ''If formIsShown Then
                                    ' ''    Call aktualisiereInfoFrm(Nothing)
                                    ' ''End If
                                End If
                                '' Ende ...

                                If Not isSymbolShape(shpRange(1)) Then
                                    If Not IsNothing(selectedPlanShapes) Then

                                        Dim tmpShape As PowerPoint.Shape = Nothing
                                        Dim elemWasMoved As Boolean = False

                                        Dim isPCard As Boolean = isProjectCard(shpRange(1))

                                        If Not isPCard Then
                                            For Each tmpShape In selectedPlanShapes
                                                ' hier sind nur noch richtige Shapes  

                                                ' sollen Home- bzw. Change-Button angezeigt werden ? 
                                                elemWasMoved = isMovedElement(tmpShape) Or elemWasMoved
                                                If elemWasMoved Then
                                                    homeButtonRelevance = True
                                                Else
                                                    If tmpShape.Tags.Item("MVD").Length > 0 Then
                                                        changedButtonRelevance = True
                                                    End If
                                                End If

                                            Next
                                        Else
                                            tmpShape = selectedPlanShapes(1)
                                        End If


                                        ' hier wird die Information zu dem selektierten Shape angezeigt 
                                        If Not IsNothing(propertiesPane) Then
                                            Call aktualisiereInfoPane(tmpShape, elemWasMoved)
                                        End If
                                        ' ur: wegen Pane
                                        If formIsShown Then
                                            If isPCard Then
                                                Call aktualisiereInfoFrm(Nothing)
                                            Else
                                                Call aktualisiereInfoFrm(tmpShape, elemWasMoved)
                                            End If

                                        End If


                                        ' jetzt den Window Ausschnitt kontrollieren: ist das oder die selectedPlanShapes überhaupt sichtbar ? 
                                        ' wenn nein, dann sicherstellen, dass sie sichtbar werden 
                                        Call ensureVisibilityOfSelection(selectedPlanShapes)

                                        ' kann jetzt wieder aktiviert werden ...
                                        If Not IsNothing(propertiesPane) Then
                                            propertiesPane.Visible = True
                                        End If
                                    Else

                                        Call checkHomeChangeBtnEnablement()
                                        If propertiesPane.Visible Then
                                            Call aktualisiereInfoPane(Nothing)
                                        End If


                                    End If

                                End If

                            End If

                        End If
                    Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionSlides
                        For i = 0 To 1000
                            i = i + 1
                        Next

                    Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText
                        Dim shpRange As PowerPoint.ShapeRange = Sel.ShapeRange

                        If Not IsNothing(shpRange) And slideHasSmartElements Then

                            ' jetzt muss hier die Behandlung für Office 2010 rein 
                            Dim correctErrorShape1 As PowerPoint.Shape = Nothing
                            Dim correctErrorShape2 As PowerPoint.Shape = Nothing

                            ' nur was machen, wenn es sich um Office 2010 handelt ... 
                            ' werden temporäre Shapes erzeugt und selektiert, die wiederum einen SelectionChange erzeugen
                            ' dabei wird das ursprünglich selektierte Shape gemerkt udn am Schluss, wenn das Property Window angezeigt ist, 
                            ' wieder selektiert .. das alles muss aber nur im Fall Version = 14.0 gemacht werden 
                            If pptAPP.Version = "14.0" Then
                                Try
                                    correctErrorShape1 = currentSlide.Shapes("visboCorrectError1")
                                Catch ex As Exception

                                End Try

                                Try
                                    correctErrorShape2 = currentSlide.Shapes("visboCorrectError2")
                                Catch ex As Exception

                                End Try
                            End If


                            If ((pptAPP.Version = "14.0") And
                                (((Not propertiesPane.Visible) Or
                                (propertiesPane.Visible And Not IsNothing(correctErrorShape1)) Or
                                (propertiesPane.Visible And Not IsNothing(correctErrorShape2))))) Then
                                ' Erzeugen eines Hilfs-Elements

                                ' Ist es 
                                If IsNothing(correctErrorShape1) And IsNothing(correctErrorShape2) And Not isRelevantMSPHShape(shpRange(1)) Then
                                    ' nichts machen 
                                Else
                                    If IsNothing(correctErrorShape1) Then
                                        ' erzeugen und selektieren der beiden Shapes  
                                        Dim oldShape As PowerPoint.Shape = shpRange(1)

                                        Dim helpShape1 As PowerPoint.Shape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                                                                   0, 0, 50, 50)



                                        If Not IsNothing(helpShape1) Then
                                            helpShape1.Name = "visboCorrectError1"
                                            helpShape1.Tags.Add("formerSN", oldShape.Name)
                                            helpShape1.Select()
                                        End If



                                    ElseIf IsNothing(correctErrorShape2) Then

                                        ' jetzt die zweite Welle 
                                        propertiesPane.Visible = True

                                        Dim helpShape2 As PowerPoint.Shape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                                                                   0, 0, 50, 50)
                                        If Not IsNothing(helpShape2) Then
                                            helpShape2.Name = "visboCorrectError2"
                                            helpShape2.Select()
                                        End If
                                    Else


                                        ' Selektieren des vorher geklickten shapes 
                                        Dim formerShapeName As String = correctErrorShape1.Tags.Item("formerSN")
                                        Dim formerSelShape As PowerPoint.Shape = Nothing

                                        If formerShapeName.Length > 0 Then
                                            Try

                                                formerSelShape = currentSlide.Shapes(formerShapeName)

                                                ' Löschen der Hilfs-Shapes 
                                                correctErrorShape1.Delete()
                                                correctErrorShape2.Delete()

                                                ' Selektieren des formerShapes
                                                formerSelShape.Select()

                                            Catch ex As Exception

                                            End Try

                                        End If

                                    End If
                                End If

                            Else
                                ' es sind ein oder mehrere Shapes selektiert worden 
                                Dim i As Integer = 0
                                If shpRange.Count = 1 Then

                                    ' prüfen, ob inzwischen was selektiert wurde, was nicht zu der Selektion in der 
                                    ' Listbox passt 

                                    '' '' prüfen, ob das Info Fenster offen ist und der Search bereich sichtbar - 
                                    '' '' dann muss der Klarheit wegen die Listbox neu aufgebaut werden 
                                    ' ''If Not IsNothing(infoFrm) And formIsShown Then
                                    ' ''    If infoFrm.rdbName.Visible Then
                                    ' ''        If infoFrm.listboxNames.SelectedItems.Count > 0 Then
                                    ' ''            'Call infoFrm.listboxNames.SelectedItems.Clear()
                                    ' ''        End If
                                    ' ''    End If
                                    ' ''End If


                                    ' jetzt ggf die angezeigten Marker löschen 
                                    If Not markerShpNames.ContainsKey(shpRange(1).Name) Then
                                        Call deleteMarkerShapes()
                                    ElseIf markerShpNames.Count > 1 Then
                                        Call deleteMarkerShapes(shpRange(1).Name)
                                    End If

                                    ' prüfen, ob es ein Kommentar ist 
                                    Dim tmpShape As PowerPoint.Shape = shpRange(1)
                                    If isCommentShape(tmpShape) Then
                                        Call markReferenceShape(tmpShape.Name)
                                    End If
                                ElseIf shpRange.Count > 1 Then
                                    ' für jedes Shape prüfen, ob es ein Comment Shape ist .. 
                                    For Each tmpShape As PowerPoint.Shape In shpRange
                                        If isCommentShape(tmpShape) Then
                                            Call markReferenceShape(tmpShape.Name)
                                        End If
                                    Next
                                ElseIf shpRange.Count = 0 Then

                                    Call deleteMarkerShapes()

                                End If


                                For Each tmpShape As PowerPoint.Shape In shpRange

                                    If tmpShape.Tags.Count > 0 Then

                                        'If tmpShape.AlternativeText <> "" And tmpShape.Title <> "" Then

                                        If isRelevantShape(tmpShape) Then
                                            If bekannteIDs.ContainsKey(tmpShape.Id) Or
                                        tmpShape.Name.EndsWith(shadowName) Then

                                                If Not relevantShapeNames.Contains(tmpShape.Name) Then
                                                    relevantShapeNames.Add(tmpShape.Name, tmpShape.Name)
                                                End If

                                            Else
                                                ' die vorhandenen Tags löschen ... und den Namen ändern 
                                                Call deleteShpTags(tmpShape)
                                            End If

                                        End If

                                    End If


                                Next

                                '' Anfang ... das war vorher innerhalb der next Schleife .. 
                                ' jetzt muss geprüft werden, ob relevantShapeNames mindestens ein Element enthält ..
                                If relevantShapeNames.Count >= 1 Then

                                    ReDim arrayOfNames(relevantShapeNames.Count - 1)

                                    For ix As Integer = 1 To relevantShapeNames.Count
                                        arrayOfNames(ix - 1) = CStr(relevantShapeNames(ix))
                                    Next

                                    selectedPlanShapes = currentSlide.Shapes.Range(arrayOfNames)

                                ElseIf isSymbolShape(shpRange(1)) Then

                                    selectedPlanShapes = shpRange
                                    Call aktualisiereInfoPane(shpRange(1))

                                Else
                                    ' in diesem Fall wurden nur nicht-relevante Shapes selektiert 
                                    Call checkHomeChangeBtnEnablement()
                                    Try
                                        If propertiesPane.Visible Then
                                            Call aktualisiereInfoPane(Nothing)
                                        End If
                                    Catch ex As Exception

                                    End Try

                                    ' ur: wegen Pane
                                    ' ''If formIsShown Then
                                    ' ''    Call aktualisiereInfoFrm(Nothing)
                                    ' ''End If
                                End If
                                '' Ende ...

                                If Not isSymbolShape(shpRange(1)) Then
                                    If Not IsNothing(selectedPlanShapes) Then

                                        Dim tmpShape As PowerPoint.Shape = Nothing
                                        Dim elemWasMoved As Boolean = False

                                        Dim isPCard As Boolean = isProjectCard(shpRange(1))

                                        If Not isPCard Then
                                            For Each tmpShape In selectedPlanShapes
                                                ' hier sind nur noch richtige Shapes  

                                                ' sollen Home- bzw. Change-Button angezeigt werden ? 
                                                elemWasMoved = isMovedElement(tmpShape) Or elemWasMoved
                                                If elemWasMoved Then
                                                    homeButtonRelevance = True
                                                Else
                                                    If tmpShape.Tags.Item("MVD").Length > 0 Then
                                                        changedButtonRelevance = True
                                                    End If
                                                End If

                                            Next
                                        Else
                                            tmpShape = selectedPlanShapes(1)
                                        End If


                                        ' hier wird die Information zu dem selektierten Shape angezeigt 
                                        If Not IsNothing(propertiesPane) Then
                                            Call aktualisiereInfoPane(tmpShape, elemWasMoved)
                                        End If
                                        ' ur: wegen Pane
                                        If formIsShown Then
                                            If isPCard Then
                                                Call aktualisiereInfoFrm(Nothing)
                                            Else
                                                Call aktualisiereInfoFrm(tmpShape, elemWasMoved)
                                            End If

                                        End If


                                        ' jetzt den Window Ausschnitt kontrollieren: ist das oder die selectedPlanShapes überhaupt sichtbar ? 
                                        ' wenn nein, dann sicherstellen, dass sie sichtbar werden 
                                        Call ensureVisibilityOfSelection(selectedPlanShapes)

                                        ' kann jetzt wieder aktiviert werden ...
                                        If Not IsNothing(propertiesPane) Then
                                            propertiesPane.Visible = True
                                        End If
                                    Else

                                        Call checkHomeChangeBtnEnablement()
                                        If propertiesPane.Visible Then
                                            Call aktualisiereInfoPane(Nothing)
                                        End If


                                    End If

                                End If

                            End If

                        End If
                        'For i = 0 To 1000
                        '    i = i + 1
                        'Next


                    Case Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionNone
                        If Not IsNothing(propertiesPane) Then
                            If propertiesPane.Visible Then
                                Call aktualisiereInfoPane(Nothing)
                            End If
                        End If
                End Select


            Catch ex As Exception

                'Call MsgBox("in windowSelectionChange: sel.type = " & Sel.Type.ToString)
                'If Not IsNothing(propertiesPane) Then
                '    If propertiesPane.Visible Then
                '        Call aktualisiereInfoPane(Nothing)
                '    End If
                'End If

            End Try

        End If

    End Sub

    ''' <summary>
    ''' stellt sicher, dass der ausschnitt im dargestellten View sichtbar ist, 
    ''' wenn nicht wird dahin gescrollt ... 
    ''' </summary>
    ''' <param name="selectedShapes"></param>
    ''' <remarks></remarks>
    Private Sub ensureVisibilityOfSelection(ByVal selectedShapes As PowerPoint.ShapeRange)

        If IsNothing(selectedShapes) Then
            ' nichts tun 
        Else
            Dim selectionLeft As Single = CSng(slideCoordInfo.drawingAreaRight + 1000)
            Dim selectionTop As Single = CSng(slideCoordInfo.drawingAreaBottom + 1000)
            Dim selectionBottom As Single = 0.0
            Dim selectionRight As Single = 0.0
            Dim markerTol As Double = markerHeight + 5

            Dim selectionWidth As Single = 0.0
            Dim selectionHeight As Single = 0.0

            For Each tmpShape As PowerPoint.Shape In selectedShapes
                With tmpShape
                    selectionLeft = System.Math.Min(selectionLeft, .Left)
                    selectionTop = System.Math.Min(selectionTop, .Top)
                    selectionBottom = System.Math.Max(selectionBottom, .Top + .Height)
                    selectionRight = System.Math.Max(selectionRight, .Left + .Width)
                End With
            Next

            ' jetzt sicherstellen, dass der Marker auch immer zu sehen  ist ... 
            selectionTop = CSng(selectionTop - markerTol)
            selectionWidth = selectionRight - selectionLeft
            selectionHeight = selectionBottom - selectionTop

            With slideCoordInfo
                If selectionLeft >= .drawingAreaLeft And
                    selectionTop >= .drawingAreaTop - markerTol And
                    selectionWidth <= .drawingAreaWidth And
                    selectionHeight <= .drawingAreaBottom - .drawingAreaTop Then
                    ' zulässig ... 

                    pptAPP.ActiveWindow.ScrollIntoView(selectionLeft, selectionTop,
                                               selectionWidth, selectionHeight)

                End If
            End With





        End If
    End Sub


    ''' <summary>
    ''' gibt aus der Enum pptinfoType den entsprechenden Wert zurück, je nachdem welcher Radiobutton gesetzt ist 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function calcRDB() As Integer
        Dim tmpResult As Integer = pptInfoType.cName
        'If formIsShown And Not IsNothing(infoFrm) Then
        '    With infoFrm
        '        If .rdbName.Checked Then
        '            tmpResult = pptInfoType.cName
        '        ElseIf .rdbOriginalName.Checked Then
        '            tmpResult = pptInfoType.oName
        '        ElseIf .rdbAbbrev.Checked Then
        '            tmpResult = pptInfoType.sName
        '        ElseIf .rdbBreadcrumb.Checked Then
        '            tmpResult = pptInfoType.bCrumb
        '        ElseIf .rdbLU.Checked Then
        '            tmpResult = pptInfoType.lUmfang
        '        ElseIf .rdbMV.Checked Then
        '            tmpResult = pptInfoType.mvElement
        '        ElseIf .rdbResources.Checked Then
        '            tmpResult = pptInfoType.resources
        '        ElseIf .rdbCosts.Checked Then
        '            tmpResult = pptInfoType.costs
        '        ElseIf .rdbVerantwortlichkeiten.Checked Then
        '            tmpResult = pptInfoType.responsible
        '        Else
        '            tmpResult = pptInfoType.cName
        '        End If
        '    End With
        'Else
        '    tmpResult = pptInfoType.cName
        'End If

        calcRDB = tmpResult

    End Function



    ''' <summary>
    ''' wird nur für relevante Shapes aufgerufen
    ''' baut die intelligenten Listen für das Slide auf 
    ''' wenn das Shape keine Abkürzung hat, so wird eine aus der laufenden Nummer erzeugt ...
    ''' 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <remarks></remarks>
    Private Sub aktualisiereSortedLists(ByVal tmpShape As PowerPoint.Shape)

        Dim shapeName As String = tmpShape.Name
        Dim checkIT As Boolean = False
        Dim isMilestone As Boolean
        Dim pvName As String = ""
        ' neu ergänzt wegen dem Element projectCard
        Dim isPCard As Boolean = False

        ' neu ergänzt - prüft ob es sich um eine PRojekt- btw Phasen Linie handelt 
        ' neu ergänzt wegen zu verwendender vpid

        ' following checks whether the line is a project resp Phase Line
        Dim isProjectOrSwimlaneLine As Boolean = (tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoLine) And
            isRelevantMSPHShape(tmpShape)

        Dim vpid As String = getVPIDFromTags(tmpShape)

        Dim shapeHeight As Single = 0.0

        If isProjectCard(tmpShape) Then
            isPCard = True
            pvName = getPVnameFromTags(tmpShape)
            vpid = getVPIDFromTags(tmpShape)
        Else
            pvName = getPVnameFromShpName(tmpShape.Name)
        End If

        'ur: 9.8.2019
        'If pvName <> "" Or vpid <> "" Then
        If pvName <> "" Then
            If smartSlideLists.containsProject(pvName) Then
                ' nichts tun, ist schon drin ..
            Else
                smartSlideLists.addProject(pvName, vpid)
            End If
        End If

        'If (tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Or
        'tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoLine) And Not isPCard Then
        '

        If (tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Or
            ((tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoLine) And Not isProjectOrSwimlaneLine)) And Not isPCard Then
            ' nichts tun 
        Else
            ' es werden nur die aufgebaut, die Meilensteine oder Phasen sind ...  
            If isPCard Then
                checkIT = True
                isMilestone = False
                shapeHeight = 0.0
            ElseIf pptShapeIsMilestone(tmpShape) Then
                checkIT = True
                isMilestone = True
                shapeHeight = tmpShape.Height
                ' nichts tun 
            ElseIf pptShapeIsPhase(tmpShape) Then
                checkIT = True
                isMilestone = False
                Try
                    If tmpShape.TextFrame2.TextRange.Text <> "" Then
                        ' dann handelt es sich nicht um einen echtes Phasen Shape, sondern um die Swimlnae Beschriftung
                        shapeHeight = 0.0
                    Else
                        shapeHeight = tmpShape.Height
                    End If
                Catch ex As Exception

                End Try

            Else
                ' nichts tun 
                checkIT = False
            End If
        End If



        If checkIT Then

            ' den classified Name behandeln ...
            Dim tmpName As String = tmpShape.Tags.Item("CN")
            If tmpName.Trim.Length = 0 Then
                Exit Sub
            End If

            Call smartSlideLists.addCN(tmpName, shapeName, isMilestone, shapeHeight)


            If Not isPCard Then
                ' den original Name behandeln ...
                tmpName = tmpShape.Tags.Item("ON")
                If tmpName.Trim.Length > 0 Then
                    Call smartSlideLists.addON(tmpName, shapeName, isMilestone)
                End If


                ' den Short Name behandeln ...
                tmpName = tmpShape.Tags.Item("SN")
                If tmpName.Trim.Length = 0 Then
                    ' es gibt keinen Short-Name, also soll einer aufgrund der laufenden Nummer erzeugt werden ...
                    tmpName = smartSlideLists.getUID(shapeName).ToString
                End If
                Call smartSlideLists.addSN(tmpName, shapeName, isMilestone)

                ' den BreadCrumb behandeln 
                tmpName = tmpShape.Tags.Item("BC")
                If tmpName.Trim.Length > 0 Then
                    Call smartSlideLists.addBC(tmpName, shapeName, isMilestone)
                End If
            End If


            ' AmpelColor behandeln
            Dim ampelColor As Integer = 0
            tmpName = tmpShape.Tags.Item("AC")
            'If tmpName.Trim.Length > 0 And pptShapeIsMilestone(tmpShape) Then
            ' tk, ab 17.11 werden jetzt auch die Phasen mit Ampel dargestellt ... 
            If tmpName.Trim.Length > 0 Then
                Try
                    If IsNumeric(tmpName) Then
                        ampelColor = CInt(tmpName)
                        Call smartSlideLists.addAC(ampelColor, shapeName, isMilestone)
                    End If

                Catch ex As Exception

                End Try

            End If

            ' Lieferumfänge behandeln
            tmpName = tmpShape.Tags.Item("LU")
            If tmpName.Trim.Length > 0 Then
                Try
                    Call smartSlideLists.addLU(tmpName, shapeName, isMilestone)
                Catch ex As Exception

                End Try
            End If

            ' Verantwortlichkeiten behandeln ...
            tmpName = tmpShape.Tags.Item("VE")
            If tmpName.Trim.Length > 0 Then
                Try
                    Call smartSlideLists.addVE(tmpName, shapeName, isMilestone)
                Catch ex As Exception

                End Try
            End If

            ' Overdue behandeln 
            tmpName = tmpShape.Tags.Item("ED")
            If tmpName.Trim.Length > 0 Then
                Try
                    Dim tmpName2 As String = tmpShape.Tags.Item("PD")
                    Dim finishDate As Date = CDate(tmpName)
                    Dim anzTageOVD As Integer = CInt(DateDiff(DateInterval.Day, finishDate, currentTimestamp))
                    If anzTageOVD > 0 Then
                        ' ist abgeschlossen, sollte also auf 100% sein

                        Dim percentDone As Double = 0.0
                        If tmpName2.Trim.Length > 0 Then
                            percentDone = CDbl(tmpName2)
                        End If
                        If percentDone > 1 Then
                            percentDone = percentDone / 100
                        End If
                        ' jetzt wird es als überfällig eingestuft, da das Finish Datum in der Vergangenheit liegt und ausserdem PercentDone nicht gleich 1 ist
                        If percentDone < 1.0 Then
                            Call smartSlideLists.addOvd(anzTageOVD, shapeName, isMilestone)
                        End If
                    End If
                Catch ex As Exception

                End Try

            End If

            ' Document Links behandeln ...
            tmpName = tmpShape.Tags.Item("DUC")


            ' wurde das Element verschoben ? 
            ' SmartslideLists werden auch gleich mit aktualisiert ... 
            If Not isPCard Then
                Call checkShpOnManualMovement(tmpShape.Name)
            End If

            ' wenn Datenbank Zugang vorliegt und es sich um eine Phase handelt, 
            ' denn nur die können Resourcen und Kostenbedarfe haben 
            ' das wird jetzt in der Routine aktualisiereRoleCostLists 
            ''If Not noDBAccessInPPT And pptShapeIsPhase(tmpShape) Then

            ''    Dim hproj As clsProjekt = smartSlideLists.getTSProject(pvName, currentTimestamp)
            ''    Dim phNameID As String = getElemIDFromShpName(tmpShape.Name)
            ''    Dim cPhase As clsPhase = hproj.getPhaseByID(phNameID)
            ''    Dim roleInformations As SortedList(Of String, Double) = cPhase.getRoleNamesAndValues
            ''    Dim costInformations As SortedList(Of String, Double) = cPhase.getCostNamesAndValues

            ''    Try
            ''        Call smartSlideLists.addRoleAndCostInfos(roleInformations, _
            ''                                                 costInformations, _
            ''                                                 shapeName)
            ''    Catch ex As Exception

            ''    End Try

            ''End If

            ' jetzt wird noch die Liste der Projekt-Varianten aufgebaut 

        End If


    End Sub

    ''' <summary>
    ''' wird nur für relevante Shapes aufgerufen
    ''' baut die intelligenten Listen für das Slide auf 
    ''' wenn das Shape keine Abkürzung hat, so wird eine aus der laufenden Nummer erzeugt ...
    ''' 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <remarks></remarks>
    Private Sub aktualisiereRoleCostLists(ByVal tmpShape As PowerPoint.Shape)
        Dim shapeName As String = tmpShape.Name
        Dim checkIT As Boolean = False
        Dim isMilestone As Boolean


        If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Or
            tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoLine Then
            ' nichts tun 
        Else
            ' es werden nur die aufgebaut, die Meilensteine oder Phasen sind ...  
            If pptShapeIsMilestone(tmpShape) Then
                checkIT = True
                isMilestone = True
                ' nichts tun 
            ElseIf pptShapeIsPhase(tmpShape) Then
                checkIT = True
                isMilestone = False
            Else
                ' nichts tun 
                checkIT = False
            End If
        End If


        If checkIT Then

            ' wenn Datenbank Zugang vorliegt und es sich um eine Phase handelt, 
            ' denn nur die können Resourcen und Kostenbedarfe haben 
            If Not noDBAccessInPPT And Not isMilestone Then
                Dim pvName As String = getPVnameFromShpName(tmpShape.Name)
                Dim vpid As String = smartSlideLists.getvpID(pvName)

                If pvName <> "" Then
                    'Dim hproj As clsProjekt = smartSlideLists.getTSProject(pvName, currentTimestamp)
                    Dim hproj As clsProjekt = timeMachine.getProjectVersion(pvName, currentTimestamp, vpid)

                    If Not IsNothing(hproj) Then
                        Dim phNameID As String = getElemIDFromShpName(tmpShape.Name)
                        Dim cPhase As clsPhase = hproj.getPhaseByID(phNameID)
                        Dim roleInformations As SortedList(Of String, Double) = cPhase.getRoleNamesAndValues
                        Dim costInformations As SortedList(Of String, Double) = cPhase.getCostNamesAndValues

                        Try
                            Call smartSlideLists.addRoleAndCostInfos(roleInformations,
                                                                     costInformations,
                                                                     shapeName,
                                                                     isMilestone)
                        Catch ex As Exception

                        End Try
                    End If

                End If


            End If


        End If


    End Sub

    ''' <summary>
    ''' prüft, ob ein Shape manuell verschoben wurde; 
    ''' wenn ja, wird dem Shape die Movement Info gleich in Tags mitgegeben und die SmartSlideLists werden aktualisiert  
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Private Sub checkShpOnManualMovement(ByVal shapeName As String)

        Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes(shapeName)
        Dim defaultExplanation As String = "manuell verschoben durch " & My.Computer.Name
        Dim isMilestone As Boolean

        If englishLanguage Then
            defaultExplanation = "moved manually by " & My.Computer.Name
        End If

        If IsNothing(tmpShape) Then
            Exit Sub
        Else

            If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                ' die Swimlane Texte sollen nicht berücksichtigt werden ...
            Else
                If pptShapeIsMilestone(tmpShape) Then

                    isMilestone = True
                    If isMovedElement(tmpShape) Then

                        homeButtonRelevance = True

                        Dim pptDate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left + 0.5 * tmpShape.Width)

                        With tmpShape
                            If .Tags.Item("MVD").Length > 0 Then
                                ' nichts tun, wenn sich das Element auf der bereits dokumentierten Veränderungs-Position befindet ... 
                                If Not isMovedElement(tmpShape, True) Then
                                    ' do nothing
                                Else
                                    ' Tags entsprechend ändern, wenn sich das Element nicht mehr auf der dokumentierten Position befindet 
                                    .Tags.Delete("MVD")
                                    .Tags.Add("MVD", pptDate.ToString)

                                    If .Tags.Item("MVE").Length > 0 Then
                                        .Tags.Delete("MVE")
                                    End If
                                    .Tags.Add("MVE", defaultExplanation)
                                End If

                            Else
                                .Tags.Add("MVD", pptDate.ToString)
                                If .Tags.Item("MVE").Length > 0 Then
                                    .Tags.Delete("MVE")
                                End If
                                .Tags.Add("MVE", defaultExplanation)
                            End If

                        End With

                        Call smartSlideLists.addMV(tmpShape.Name, isMilestone)
                    Else
                        ' das Shape wurde nicht verschoben, aber hat es einen MVD Teil ? 
                        ' dann muss der ChangedButton gezeigt werden 
                        If tmpShape.Tags.Item("MVD").Length > 0 Then
                            changedButtonRelevance = True
                        End If
                    End If


                Else
                    isMilestone = False
                    If isMovedElement(tmpShape) Then

                        homeButtonRelevance = True

                        Dim pptSDate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left)
                        Dim pptEDate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left + tmpShape.Width)

                        With tmpShape
                            If .Tags.Item("MVD").Length > 0 Then
                                ' nichts tun, wenn sich das Element auf der bereits dokumentierten Veränderungs-Position befindet ... 
                                If Not isMovedElement(tmpShape, True) Then
                                    ' do nothing
                                Else
                                    ' Tags entsprechend ändern, wenn sich das Element nicht mehr auf der dokumentierten Position befindet 
                                    .Tags.Delete("MVD")
                                    .Tags.Add("MVD", pptSDate.ToString & "#" & pptEDate.ToString)

                                    If .Tags.Item("MVE").Length > 0 Then
                                        .Tags.Delete("MVE")
                                    End If
                                    .Tags.Add("MVE", defaultExplanation)
                                End If

                            Else
                                .Tags.Add("MVD", pptSDate.ToString & "#" & pptEDate.ToString)
                                ' wenn bereits eine Explanation existiert, soll die erhalten bleiben 
                                If .Tags.Item("MVE").Length > 0 Then
                                    .Tags.Delete("MVE")
                                Else
                                    .Tags.Add("MVE", defaultExplanation)
                                End If

                            End If


                        End With

                        Call smartSlideLists.addMV(tmpShape.Name, isMilestone)
                    Else
                        ' das Shape wurde nicht verschoben, aber hat es einen MVD Teil ? 
                        ' dann muss der ChangedButton gezeigt werden 
                        If tmpShape.Tags.Item("MVD").Length > 0 Then
                            changedButtonRelevance = True
                        End If
                    End If

                End If
            End If

        End If




    End Sub

    ''' <summary>
    ''' wird nur aufgerufen für relevant Shapes
    ''' positioniert ein Shape auf seine "Home"-Position, wenn es nicht ohnehin schon dort ist ... 
    ''' </summary>
    ''' <param name="tmpShapeName"></param>
    ''' <remarks></remarks>
    Friend Sub sentToHomePosition(ByVal tmpShapeName As String)

        Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes(tmpShapeName)
        If Not IsNothing(tmpShape) Then

            Dim homeSDate As Date
            Dim homeEDate As Date
            Dim x1Pos As Double, x2Pos As Double

            ' Prüfen, ob Text Box , wenn ja, gleich Exit 
            If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                ' nichts tun 
            Else
                If pptShapeIsMilestone(tmpShape) Then

                    With tmpShape
                        If .Tags.Item("MVD").Length > 0 Then
                            ' nur dann muss was nach Hause geschickt werden 
                            Try
                                ' ED existiert - das wird in pptShapeisMilestone geprüft 
                                homeEDate = CDate(.Tags.Item("ED"))
                                Call slideCoordInfo.calculatePPTx1x2(homeEDate, homeEDate, x1Pos, x2Pos)

                                ' Positionieren auf Home Position und aktualisieren des Info-Formulars..
                                If .Left <> CSng(x1Pos) - .Width / 2 Then

                                    .Left = CSng(x1Pos) - .Width / 2
                                    changedButtonRelevance = True

                                    If formIsShown Then
                                        Call aktualisiereInfoFrm(tmpShape, True)
                                    End If



                                End If

                            Catch ex As Exception

                            End Try
                        End If
                    End With

                ElseIf pptShapeIsPhase(tmpShape) Then

                    With tmpShape
                        If .Tags.Item("MVD").Length > 0 Then
                            ' nur dann muss was nach Hause geschickt werden 
                            Try
                                ' SD, ED existieren - das wird in pptShapeisPhase geprüft 
                                homeSDate = CDate(.Tags.Item("SD"))
                                homeEDate = CDate(.Tags.Item("ED"))
                                Call slideCoordInfo.calculatePPTx1x2(homeSDate, homeEDate, x1Pos, x2Pos)

                                ' Positionieren auf Home Position und aktualisieren des Info-Formulars..
                                If ((.Left <> CSng(x1Pos)) Or (.Width <> CSng(x2Pos - x1Pos))) Then

                                    changedButtonRelevance = True

                                    .Left = CSng(x1Pos)
                                    .Width = CSng(x2Pos - x1Pos)

                                    If formIsShown Then
                                        Call aktualisiereInfoFrm(tmpShape, True)
                                    End If
                                End If

                            Catch ex As Exception

                            End Try

                        End If
                    End With
                End If

            End If

        End If



    End Sub

    ''' <summary>
    ''' wird nur aufgerufen für relevant Shapes
    ''' positioniert ein Shape auf seine "Changed"-Position, wenn es denn eine gibt  ... 
    ''' aktualisiert das info-Fenster, wenn nur ein Shape selektiert ist 
    ''' verschiebt evtl vorhandene Text und Datums-Beschriftungen mit 
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Friend Sub sentToChangedPosition(ByVal shapeName As String)

        Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes(shapeName)
        If Not IsNothing(tmpShape) Then

            Dim homeSDate As Date
            Dim homeEDate As Date
            Dim x1Pos As Double, x2Pos As Double
            Dim tmpstr() As String
            'Dim diff As Double

            ' Prüfen, ob Text Box , wenn ja, gleich Exit 
            If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                ' nichts tun 
            Else
                If pptShapeIsMilestone(tmpShape) Then

                    With tmpShape
                        If .Tags.Item("MVD").Length > 0 Then
                            ' nur dann kann was zur Changed Position geschickt werden 
                            Try
                                ' ED existiert - das wird in pptShapeisMilestone geprüft 
                                tmpstr = .Tags.Item("MVD").Split(New Char() {CType("#", Char)})
                                homeEDate = CDate(tmpstr(0))
                                Call slideCoordInfo.calculatePPTx1x2(homeEDate, homeEDate, x1Pos, x2Pos)

                                ' Positionieren auf Changed Position und aktualisieren des Info-Formulars..
                                If .Left <> CSng(x1Pos) - .Width / 2 Then

                                    homeButtonRelevance = True

                                    .Left = CSng(x1Pos) - .Width / 2
                                    If formIsShown Then
                                        Call aktualisiereInfoFrm(tmpShape, True)
                                    End If

                                End If

                            Catch ex As Exception

                            End Try
                        End If
                    End With

                ElseIf pptShapeIsPhase(tmpShape) Then

                    With tmpShape
                        If .Tags.Item("MVD").Length > 0 Then
                            ' nur dann kann was zur Changed Position geschickt werden 
                            Try
                                ' SD, ED existieren - das wird in pptShapeisPhase geprüft 
                                tmpstr = .Tags.Item("MVD").Split(New Char() {CType("#", Char)})

                                If tmpstr.Length = 2 Then
                                    homeSDate = CDate(tmpstr(0))
                                    homeEDate = CDate(tmpstr(1))
                                    Call slideCoordInfo.calculatePPTx1x2(homeSDate, homeEDate, x1Pos, x2Pos)

                                    ' Positionieren auf Changed Position und aktualisieren des Info-Formulars..
                                    If ((.Left <> CSng(x1Pos)) Or (.Width <> CSng(x2Pos - x1Pos))) Then

                                        homeButtonRelevance = True

                                        .Left = CSng(x1Pos)
                                        .Width = CSng(x2Pos - x1Pos)

                                        If formIsShown Then
                                            Call aktualisiereInfoFrm(tmpShape, True)
                                        End If
                                    End If

                                End If

                            Catch ex As Exception

                            End Try

                        End If
                    End With
                End If

            End If

        End If


    End Sub

    ''' <summary>
    ''' aktualisiert alle VISBO Charts, VISBO Platzhalter und VISBO Tabellen ...
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="curTimeStamp">der aktuelle TimeStamp</param>
    ''' <param name="prevTimeStamp">der vorherig gültige TimeStamp</param>
    ''' <remarks></remarks>
    Friend Sub updateVisboComponent(ByRef pptShape As PowerPoint.Shape, ByVal curTimeStamp As Date, ByVal prevTimeStamp As Date,
                                    Optional ByVal showOtherVariant As Boolean = False)
        Dim chtObjName As String = ""
        Dim bigType As Integer = -1
        Dim detailID As Integer = -1
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim bProj As clsProjekt = Nothing ' nimmt das erste beauftragte Projekt auf ..
        Dim lProj As clsProjekt = Nothing ' nimmt das zuletzt beauftragte Projekt auf 
        Dim bProj1 As clsProjekt = Nothing
        Dim lProj1 As clsProjekt = Nothing
        Try

            If Not IsNothing(pptShape) Then

                If pptShape.Tags("BID").Length > 0 And pptShape.Tags("DID").Length > 0 Then
                    If IsNumeric(pptShape.Tags("BID")) And IsNumeric(pptShape.Tags("DID")) Then
                        bigType = CInt(pptShape.Tags("BID"))
                        detailID = CInt(pptShape.Tags("DID"))
                    End If
                End If

                If bigType = ptReportBigTypes.charts Then

                    If pptShape.Tags.Item("CHON").Length > 0 Then
                        ' es handelt sich um ein Projekt- oder Portfolio Chart 


                        If pptShape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            Dim pptChart As PowerPoint.Chart = pptShape.Chart

                            Call updateVisboChart(pptShape, curTimeStamp, prevTimeStamp, showOtherVariant)

                        End If

                    End If


                ElseIf bigType = ptReportBigTypes.components Then

                    Call updatePPTComponent(pptShape, detailID, curTimeStamp)

                ElseIf bigType = ptReportBigTypes.tables Then

                    Dim pName As String = pptShape.Tags.Item("PNM")
                    Dim vName As String = pptShape.Tags.Item("VNM")
                    Dim vpid As String = pptShape.Tags.Item("VPID")

                    If showOtherVariant Then
                        vName = currentVariantname
                        If pptShape.Tags.Item("VNM").Length > 0 Then
                            pptShape.Tags.Delete("VNM")
                        End If
                        pptShape.Tags.Add("VNM", vName)
                        Dim chck As String = pptShape.Tags.Item("VNM")
                    End If


                    If vpid = "" Then

                        If pName <> "" Then
                            Dim pvName As String = calcProjektKey(pName, vName)

                            ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                            Dim tsProj As clsProjekt = timeMachine.getProjectVersion(pvName, curTimeStamp)

                            If Not IsNothing(tsProj) Then

                                ' bei normalen Projekten wird immer mit der Basis-Variante verglichen, bei Portfolio Projekten mit dem Portfolio Name

                                Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString


                                'ElseIf bigType = ptReportBigTypes.tables Then

                                If detailID = PTpptTableTypes.prZiele Then
                                    Call updatePPTProjektTabelleZiele(pptShape, tsProj)

                                ElseIf detailID = PTpptTableTypes.prBudgetCostAPVCV Then

                                    Dim continueOperation As Boolean = False
                                    Try
                                        'bProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(tsProj.name, vorgabeVariantName)
                                        'bProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).beauftragung

                                        'lProj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(tsProj.name, vorgabeVariantName, curTimeStamp.AddMinutes(-1))
                                        'lProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).lastBeauftragung(curTimeStamp.AddMinutes(-1))
                                        bProj = timeMachine.getFirstContractedVersion(pvName)
                                        lProj = timeMachine.getLastContractedVersion(pvName, curTimeStamp)

                                        ' hier unterscheiden, ob Summary-Projekt oder normales
                                        If tsProj.projectType = ptPRPFType.portfolio Then

                                            ' lade das Portfolio 
                                            Dim err As New clsErrorCodeMsg
                                            Dim realTimestamp As Date

                                            Dim aktConst As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(pName, vpid, realTimestamp, err, vName, curTimeStamp)

                                            Dim hproj As clsProjekt = calcUnionProject(aktConst, False, curTimeStamp)

                                            ' das neu berechnete SummaryProjekt hat den curTimeStamp
                                            hproj.timeStamp = curTimeStamp

                                            tsProj = hproj
                                            continueOperation = Not IsNothing(tsProj)

                                        Else
                                            ' kann eigentlich nicht mehr Nothing werden ... die Liste an TimeStamps enthält den größten auftretenden kleinsten datumswert aller Projekte ....
                                            continueOperation = Not IsNothing(tsProj)
                                        End If

                                        If continueOperation Then

                                            Dim q1 As String = pptShape.Tags.Item("Q1")
                                            Dim q2 As String = pptShape.Tags.Item("Q2")
                                            'Dim nids As String = pptShape.Tags.Item("NIDS")

                                            ' ist showrangeleft und Right zu setzen ? 
                                            If Not (showRangeLeft > 0 And showRangeRight > showRangeLeft) Then
                                                If Not IsNothing(pptShape.Tags.Item("SRLD")) And Not IsNothing(pptShape.Tags.Item("SRRD")) Then
                                                    If pptShape.Tags.Item("SRLD") <> "" And pptShape.Tags.Item("SRRD") <> "" Then
                                                        showRangeLeft = getColumnOfDate(CDate(pptShape.Tags.Item("SRLD")))
                                                        showRangeRight = getColumnOfDate(CDate(pptShape.Tags.Item("SRRD")))
                                                    End If
                                                End If
                                            End If

                                            'ur:16.01.2019: Call zeichneTableBudgetCostAPVCV(pptShape, tsProj, bProj, lProj,
                                            '                                 toDoCollection, q1, q2)
                                            Call zeichneTableBudgetCostAPVCV(pptShape, tsProj, bProj, lProj, q1, q2)


                                        End If

                                    Catch ex As Exception
                                        Call MsgBox("Budget/Kosten Tabelle konnte nicht aktualisiert werden ...")
                                        bProj = Nothing
                                        lProj = Nothing
                                    End Try

                                ElseIf detailID = PTpptTableTypes.prMilestoneAPVCV Then
                                    Try
                                        'bProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(tsProj.name, vorgabeVariantName)
                                        'bProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).beauftragung
                                        'lProj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(tsProj.name, vorgabeVariantName, curTimeStamp.AddHours(-1))
                                        'lProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).lastBeauftragung(curTimeStamp.AddMinutes(-1))

                                        bProj = timeMachine.getFirstContractedVersion(pvName)
                                        lProj = timeMachine.getLastContractedVersion(pvName, curTimeStamp)

                                        Dim toDoCollection As Collection = convertNidsToColl(pptShape.Tags.Item("NIDS"))

                                        Dim q1 As String = pptShape.Tags.Item("Q1")
                                        Dim q2 As String = pptShape.Tags.Item("Q2")
                                        Dim nids As String = pptShape.Tags.Item("NIDS")

                                        Call zeichneTableMilestoneAPVCV(pptShape, tsProj, bProj, lProj,
                                                                     toDoCollection, q1, q2)

                                    Catch ex As Exception
                                        Call MsgBox("Budget/Kosten Tabelle konnte nicht aktualisiert werden ...")
                                        bProj = Nothing
                                    End Try
                                End If

                            End If

                        End If

                    Else    ' vpid <> ""
                        If vpid <> "" Then
                            Dim pvName As String = calcProjektKey(pName, vName)

                            ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                            Dim tsProj As clsProjekt = timeMachine.getProjectVersion(pvName, curTimeStamp, vpid)

                            If Not IsNothing(tsProj) Then

                                ' bei normalen Projekten wird immer mit der Basis-Variante verglichen, bei Portfolio Projekten mit dem Portfolio Name

                                Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString


                                'ElseIf bigType = ptReportBigTypes.tables Then

                                If detailID = PTpptTableTypes.prZiele Then
                                    Call updatePPTProjektTabelleZiele(pptShape, tsProj)

                                ElseIf detailID = PTpptTableTypes.prBudgetCostAPVCV Then
                                    Try
                                        'bProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(tsProj.name, vorgabeVariantName)
                                        'bProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).beauftragung

                                        'lProj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(tsProj.name, vorgabeVariantName, curTimeStamp.AddMinutes(-1))
                                        'lProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).lastBeauftragung(curTimeStamp.AddMinutes(-1))
                                        bProj = timeMachine.getFirstContractedVersion(pvName, vpid)
                                        lProj = timeMachine.getLastContractedVersion(pvName, curTimeStamp, vpid)


                                        Dim q1 As String = pptShape.Tags.Item("Q1")
                                        Dim q2 As String = pptShape.Tags.Item("Q2")
                                        'Dim nids As String = pptShape.Tags.Item("NIDS")

                                        ' ist showrangeleft und Right zu setzen ? 
                                        If Not (showRangeLeft > 0 And showRangeRight > showRangeLeft) Then
                                            If Not IsNothing(pptShape.Tags.Item("SRLD")) And Not IsNothing(pptShape.Tags.Item("SRRD")) Then
                                                If pptShape.Tags.Item("SRLD") <> "" And pptShape.Tags.Item("SRRD") <> "" Then
                                                    showRangeLeft = getColumnOfDate(CDate(pptShape.Tags.Item("SRLD")))
                                                    showRangeRight = getColumnOfDate(CDate(pptShape.Tags.Item("SRRD")))
                                                End If
                                            End If
                                        End If

                                        'ur:16.01.2019: Call zeichneTableBudgetCostAPVCV(pptShape, tsProj, bProj, lProj,
                                        '                                 toDoCollection, q1, q2)
                                        Call zeichneTableBudgetCostAPVCV(pptShape, tsProj, bProj, lProj, q1, q2)



                                    Catch ex As Exception
                                        Call MsgBox("Budget/Kosten Tabelle konnte nicht aktualisiert werden ...")
                                        bProj = Nothing
                                        lProj = Nothing
                                    End Try

                                ElseIf detailID = PTpptTableTypes.prMilestoneAPVCV Then
                                    Try
                                        'bProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(tsProj.name, vorgabeVariantName)
                                        'bProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).beauftragung
                                        'lProj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(tsProj.name, vorgabeVariantName, curTimeStamp.AddHours(-1))
                                        'lProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).lastBeauftragung(curTimeStamp.AddMinutes(-1))

                                        bProj = timeMachine.getFirstContractedVersion(pvName, vpid)
                                        lProj = timeMachine.getLastContractedVersion(pvName, curTimeStamp, vpid)

                                        Dim toDoCollection As Collection = convertNidsToColl(pptShape.Tags.Item("NIDS"))

                                        Dim q1 As String = pptShape.Tags.Item("Q1")
                                        Dim q2 As String = pptShape.Tags.Item("Q2")
                                        Dim nids As String = pptShape.Tags.Item("NIDS")

                                        Call zeichneTableMilestoneAPVCV(pptShape, tsProj, bProj, lProj,
                                                                     toDoCollection, q1, q2)

                                    Catch ex As Exception
                                        Call MsgBox("Budget/Kosten Tabelle konnte nicht aktualisiert werden ...")
                                        bProj = Nothing
                                    End Try
                                End If

                            End If

                        End If
                    End If       ' end of Tabellen

                End If    ' end of unterschiedliche Report-Componenten



            Else
                ' kein zu aktualisierendes Shape ... 
            End If


        Catch ex As Exception
            Call MsgBox("UpdateVisboComponent: " & ex.Message)
            Dim a As Integer = 1
        End Try
    End Sub


    ''' <summary>
    ''' aktualisiert eine Smart PPT Komponenten, das sind Felder 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="detailID"></param>
    ''' 
    ''' <remarks></remarks>

    Public Sub updatePPTComponent(ByRef pptShape As PowerPoint.Shape,
                                  ByVal detailID As Integer, ByVal curTimeStamp As Date)

        'Public Sub updatePPTComponent(ByVal hproj As clsProjekt, ByRef pptShape As PowerPoint.Shape,
        '                                  ByVal detailID As Integer, ByVal curTimeStamp As Date)
        Try
            Dim hproj As clsProjekt = Nothing
            Dim portfolio As clsConstellation = Nothing
            Dim portfolioTS As Date = Nothing
            Dim scInfo As New clsSmartPPTCompInfo
            Call scInfo.getValuesFromPPTShape(pptShape)

            If scInfo.pName <> "" Or scInfo.vpid <> "" Then

                Dim continueOperation As Boolean = False
                If scInfo.prPF = ptPRPFType.portfolio And Not scInfo.pName.Contains("_last") Then


                    If currentConstellationPvName = calcPortfolioKey(scInfo.pName, scInfo.vName) Then
                        ' nix tun
                        continueOperation = True
                        portfolio = currentSessionConstellation
                    Else


                        Try
                            currentConstellationPvName = calcPortfolioKey(scInfo.pName, scInfo.vName)
                            ShowProjekte.Clear(updateCurrentConstellation:=False)

                            ' lade das Portfolio 
                            Dim err As New clsErrorCodeMsg
                            currentSessionConstellation =
                                CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(scInfo.pName, scInfo.vpid,
                                                                                                      portfolioTS,
                                                                                                     err,
                                                                                                     scInfo.vName, storedAtOrBefore:=curTimeStamp)
                            portfolio = currentSessionConstellation

                            '' bringe alles in ShowProjekte 
                            'For Each kvp As KeyValuePair(Of String, clsProjekt) In pfListe
                            '    ShowProjekte.Add(kvp.Value, updateCurrentConstellation:=False)
                            'Next

                            ' besetzte ggf den Zeitraum
                            If scInfo.hasValidZeitraum Then
                                showRangeLeft = getColumnOfDate(scInfo.zeitRaumLeft)
                                showRangeRight = getColumnOfDate(scInfo.zeitRaumRight)
                            End If

                            continueOperation = Not IsNothing(portfolio)
                        Catch ex As Exception
                            Call MsgBox("Componente kann nicht aktualisiert werden ..")
                        End Try
                    End If
                Else
                    ' ist Projekt
                    Dim pName As String = pptShape.Tags.Item("PNM")
                    Dim vName As String = pptShape.Tags.Item("VNM")
                    Dim vpid As String = pptShape.Tags.Item("VPID")

                    'If showOtherVariant Then
                    '    vName = currentVariantname
                    '    If pptShape.Tags.Item("VNM").Length > 0 Then
                    '        pptShape.Tags.Delete("VNM")
                    '    End If
                    '    pptShape.Tags.Add("VNM", vName)
                    '    Dim chck As String = pptShape.Tags.Item("VNM")
                    'End If
                    If vpid <> "" Then

                        Dim pvName As String = calcProjektKey(pName, vName)

                        ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                        hproj = timeMachine.getProjectVersion(pvName, curTimeStamp, vpid)

                        If Not IsNothing(hproj) Then

                            ' bei normalen Projekten wird immer mit der Basis-Variante verglichen, bei Portfolio Projekten mit dem Portfolio Name

                            Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString

                        End If

                    Else
                        If pName <> "" Then
                            Dim pvName As String = calcProjektKey(pName, vName)

                            ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                            hproj = timeMachine.getProjectVersion(pvName, curTimeStamp, vpid)

                            If Not IsNothing(hproj) Then

                                ' bei normalen Projekten wird immer mit der Basis-Variante verglichen, bei Portfolio Projekten mit dem Portfolio Name

                                Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString
                            Else


                            End If
                        End If
                    End If

                End If



                Select Case detailID

                    Case ptReportComponents.prName

                        If Not IsNothing(hproj) Then
                            pptShape.TextFrame2.TextRange.Text = hproj.getShapeText
                        End If
                    Case ptReportComponents.pfName

                        If Not IsNothing(portfolio) Then
                            If portfolio.variantName <> "" Then
                                pptShape.TextFrame2.TextRange.Text = portfolio.constellationName & "[" & portfolio.variantName & "]"
                            Else
                                pptShape.TextFrame2.TextRange.Text = portfolio.constellationName
                            End If

                        End If

                    Case ptReportComponents.prCustomField
                        Dim qualifier As String = pptShape.Tags.Item("Q1")
                        If Not IsNothing(qualifier) Then
                            If qualifier.Length > 0 Then
                                Dim uid As Integer = customFieldDefinitions.getUid(qualifier)

                                If uid <> -1 Then
                                    Dim cftype As Integer = customFieldDefinitions.getTyp(uid)

                                    Select Case cftype
                                        Case ptCustomFields.Str
                                            Dim wert As String = Nothing
                                            If Not IsNothing(hproj) Then
                                                wert = hproj.getCustomSField(uid)
                                                If Not IsNothing(wert) Then
                                                    pptShape.TextFrame2.TextRange.Text = qualifier & ": " & wert
                                                Else
                                                    pptShape.TextFrame2.TextRange.Text = qualifier & " : n.a"
                                                End If
                                            End If

                                        Case ptCustomFields.Dbl
                                            Dim wert As Double = Nothing
                                            If Not IsNothing(hproj) Then
                                                wert = hproj.getCustomDField(uid)
                                                If Not IsNothing(wert) Then
                                                    pptShape.TextFrame2.TextRange.Text = qualifier & ": " & wert.ToString("#0.##")
                                                Else
                                                    pptShape.TextFrame2.TextRange.Text = qualifier & " : n.a"
                                                End If

                                            End If

                                        Case ptCustomFields.bool
                                            Dim wert As Boolean = Nothing

                                            If Not IsNothing(hproj) Then
                                                wert = hproj.getCustomBField(uid)

                                                If Not IsNothing(wert) Then
                                                    If wert Then
                                                        ' Sprache !
                                                        pptShape.TextFrame2.TextRange.Text = qualifier & ": Yes"
                                                    Else
                                                        ' Sprache !
                                                        pptShape.TextFrame2.TextRange.Text = qualifier & ": No"
                                                    End If

                                                Else
                                                    pptShape.TextFrame2.TextRange.Text = qualifier & " : n.a"
                                                End If

                                            End If

                                    End Select

                                Else
                                    pptShape.TextFrame2.TextRange.Text = qualifier & " : n.a"
                                End If
                            End If
                        End If

                    Case ptReportComponents.prAmpel
                        If scInfo.prPF = ptPRPFType.project Then
                            If Not IsNothing(hproj) Then
                                Select Case hproj.ampelStatus
                                    Case 0
                                        'pptShape.Fill.ForeColor.RGB = System.Drawing.Color.Gray.ToArgb
                                        pptShape.Fill.ForeColor.RGB = PowerPoint.XlRgbColor.rgbGray
                                    Case 1
                                        'pptShape.Fill.ForeColor.RGB = System.Drawing.Color.Green.ToArgb
                                        pptShape.Fill.ForeColor.RGB = PowerPoint.XlRgbColor.rgbGreen
                                    Case 2
                                        'pptShape.Fill.ForeColor.RGB = System.Drawing.Color.Yellow.ToArgb
                                        pptShape.Fill.ForeColor.RGB = PowerPoint.XlRgbColor.rgbYellow
                                    Case 3
                                        'pptShape.Fill.ForeColor.RGB = System.Drawing.Color.Red.ToArgb
                                        pptShape.Fill.ForeColor.RGB = PowerPoint.XlRgbColor.rgbRed
                                    Case Else
                                End Select
                            End If
                        Else
                            ' ist Portfolio
                        End If


                    Case ptReportComponents.prAmpelText

                        If scInfo.prPF = ptPRPFType.project Then
                            If Not IsNothing(hproj) Then
                                'Dim qualifier2 As String = pptShape.Tags.Item("Q2")
                                'pptShape.TextFrame2.TextRange.Text = qualifier2 & ": " & hproj.ampelErlaeuterung
                                ' 23.6.18 nur noch den eigentlichen Ampel-Text schreiben ...
                                pptShape.TextFrame2.TextRange.Text = hproj.ampelErlaeuterung
                            End If
                        Else
                            ' ist Portfolio

                        End If

                    Case ptReportComponents.prBusinessUnit
                        If Not IsNothing(hproj) Then
                            Dim qualifier2 As String = pptShape.Tags.Item("Q2")
                            pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & hproj.businessUnit
                        End If

                    Case ptReportComponents.prStand
                        If scInfo.prPF = ptPRPFType.project Then
                            If Not IsNothing(hproj) Then
                                Dim qualifier2 As String = pptShape.Tags.Item("Q2")
                                'pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & hproj.timeStamp.ToShortDateString
                                'pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & curTimeStamp.ToShortDateString
                                pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & curTimeStamp.ToShortDateString & " (DB: " & hproj.timeStamp.ToString("g", repCult) & ")"
                            End If
                        Else
                            ' ist Portfolio
                            If Not IsNothing(portfolio) Then
                                Dim qualifier2 As String = pptShape.Tags.Item("Q2")
                                'pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & hproj.timeStamp.ToShortDateString
                                'pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & curTimeStamp.ToShortDateString
                                pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & curTimeStamp.ToShortDateString & " (DB: " & portfolioTS.ToString("g", repCult) & ")"
                            End If
                        End If
                    Case ptReportComponents.prDescription
                        If scInfo.prPF = ptPRPFType.project Then
                            If Not IsNothing(hproj) Then
                                Dim qualifier2 As String = pptShape.Tags.Item("Q2")
                                ' tk 23.6.18 nur noch den eigentlichen Text schreiben  
                                Dim initialText As String = hproj.description

                                If hproj.variantDescription.Length > 0 Then

                                    pptShape.TextFrame2.TextRange.Text = initialText & vbLf & vbLf &
                            "Varianten-Beschreibung: " & hproj.variantDescription
                                End If
                                pptShape.TextFrame2.TextRange.Text = initialText
                            End If
                        Else
                            ' ist Portfolio
                        End If


                    Case ptReportComponents.prLaufzeit

                        If scInfo.prPF = ptPRPFType.project Then
                            If Not IsNothing(hproj) Then
                                Dim qualifier2 As String = pptShape.Tags.Item("Q2")
                                pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & textZeitraum(hproj.startDate, hproj.endeDate)

                            End If
                        Else
                            ' ist Portfolio
                            If Not IsNothing(portfolio) Then
                                If scInfo.hasValidZeitraum Then
                                    pptShape.TextFrame2.TextRange.Text = textZeitraum(scInfo.zeitRaumLeft, scInfo.zeitRaumRight)
                                Else
                                    pptShape.TextFrame2.TextRange.Text = "     "
                                End If
                            End If
                        End If


                    Case ptReportComponents.prVerantwortlich
                        If scInfo.prPF = ptPRPFType.project Then
                            If Not IsNothing(hproj) Then
                                Dim qualifier2 As String = pptShape.Tags.Item("Q2")
                                pptShape.TextFrame2.TextRange.Text = qualifier2 & " " & hproj.leadPerson

                            End If
                        Else
                            ' ist Portfolio
                        End If



                    Case Else
                        If detailID = ptReportComponents.prSymDescription Or
                    detailID = ptReportComponents.prSymTrafficLight Or
                    detailID = ptReportComponents.prSymFinance Or
                    detailID = ptReportComponents.prSymProject Or
                    detailID = ptReportComponents.prSymRisks Or
                    detailID = ptReportComponents.prSymSchedules Or
                    detailID = ptReportComponents.prSymTeam Then

                            If Not IsNothing(hproj) Then

                                If detailID = ptReportComponents.prSymTrafficLight Then
                                    Call switchOnTrafficLightColor(pptShape, hproj.ampelStatus)
                                End If

                                Dim qualifier As String = pptShape.Tags.Item("Q1")
                                Dim qualifier2 As String = pptShape.Tags.Item("Q2")

                                ' jetzt müssen an das Shape wieder die Smart-Infos angebunden werden 
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2, ptReportBigTypes.components, detailID)

                            End If

                        End If


                End Select

            End If          ' scInfo.pName <> ""


        Catch ex As Exception
            If awinSettings.visboDebug Then
                Call MsgBox("hier in updatePPTComponent: " & ex.Message)
            End If
            Call MsgBox("hier in updatePPTComponent: " & ex.Message)
        End Try


    End Sub

    ''' <summary>
    ''' neue Methode, um Charts zu aktualisieren
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="curTimeStamp"></param>
    ''' <param name="prevTimeStamp"></param>
    ''' <param name="showOtherVariant"></param>
    Public Sub updateVisboChart(ByRef pptShape As PowerPoint.Shape,
                                ByVal curTimeStamp As Date, ByVal prevTimeStamp As Date,
                                    Optional ByVal showOtherVariant As Boolean = False)

        Dim chtObjName As String
        Dim pvName As String = ""

        If pptShape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue Then
            Dim pptChart As PowerPoint.Chart = pptShape.Chart

            chtObjName = pptChart.Name

            Dim scInfo As New clsSmartPPTChartInfo
            Call scInfo.getValuesFromPPTShape(pptShape)

            '' showRangeLeft und showRangeRight müssen gesetzt werden, damit bei der Bestimmung der Kapas und
            '' Plandaten der Zeitraum bekannt ist.
            'showRangeLeft = getColumnOfDate(scInfo.zeitRaumLeft)
            'showRangeRight = getColumnOfDate(scInfo.zeitRaumRight)

            If scInfo.pName <> "" Or scInfo.vpid <> "" Then


                Dim continueOperation As Boolean = False
                If scInfo.prPF = ptPRPFType.portfolio Then

                    If Not scInfo.pName.Contains("_last") Then


                        If currentConstellationPvName = calcPortfolioKey(scInfo.pName, scInfo.vName) Then

                            If ShowProjekte.Count <> 0 Then
                                ' nix tun
                                continueOperation = True
                            Else
                                ShowProjekte.Clear(updateCurrentConstellation:=False)

                                ' lade das Portfolio 
                                Dim err As New clsErrorCodeMsg
                                Dim pfListe As SortedList(Of String, clsProjekt) = CType(databaseAcc, DBAccLayer.Request).retrieveProjectsOfOneConstellationFromDB(scInfo.pName,
                                                                                                                                                                   scInfo.vpid, scInfo.vName, err, storedAtOrBefore:=curTimeStamp)

                                ' bringe alles in ShowProjekte 
                                For Each kvp As KeyValuePair(Of String, clsProjekt) In pfListe
                                    ShowProjekte.Add(kvp.Value, updateCurrentConstellation:=False)
                                Next

                                '' besetzte ggf den Zeitraum
                                If scInfo.hasValidZeitraum Then
                                    showRangeLeft = getColumnOfDate(scInfo.zeitRaumLeft)
                                    showRangeRight = getColumnOfDate(scInfo.zeitRaumRight)
                                End If

                                continueOperation = Not IsNothing(ShowProjekte)
                            End If

                        Else

                            Try
                                currentConstellationPvName = calcPortfolioKey(scInfo.pName, scInfo.vName)
                                ShowProjekte.Clear(updateCurrentConstellation:=False)

                                ' lade das Portfolio 
                                Dim err As New clsErrorCodeMsg
                                Dim pfListe As SortedList(Of String, clsProjekt) = CType(databaseAcc, DBAccLayer.Request).retrieveProjectsOfOneConstellationFromDB(scInfo.pName, scInfo.vpid, scInfo.vName, err, storedAtOrBefore:=curTimeStamp)

                                ' bringe alles in ShowProjekte 
                                For Each kvp As KeyValuePair(Of String, clsProjekt) In pfListe
                                    ShowProjekte.Add(kvp.Value, updateCurrentConstellation:=False)
                                Next

                                '' besetzte ggf den Zeitraum
                                If scInfo.hasValidZeitraum Then
                                    showRangeLeft = getColumnOfDate(scInfo.zeitRaumLeft)
                                    showRangeRight = getColumnOfDate(scInfo.zeitRaumRight)
                                End If

                                continueOperation = Not IsNothing(ShowProjekte)

                            Catch ex As Exception
                                Call MsgBox("Chart kann nicht aktualisiert werden ..")
                            End Try
                        End If
                    Else
                        If awinSettings.englishLanguage Then
                            Call MsgBox("Portfolio named: " & scInfo.pName & " cannot be updated")
                        Else
                            Call MsgBox("Das Portfolio " & scInfo.pName & " kann nicht aktualisiert werden")
                        End If
                    End If

                Else

                    ' tk 23.4.19
                    pvName = calcProjektKey(scInfo.pName, scInfo.vName)
                    ' damit auch eine andere Variante gezeigt werden kann ... 

                    If showOtherVariant Then
                        Dim tmpPName As String = getPnameFromKey(pvName)
                        pvName = calcProjektKey(tmpPName, currentVariantname)
                        scInfo.vName = currentVariantname
                    End If

                    ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                    'scInfo.hproj = smartSlideLists.getTSProject(pvName, curTimeStamp)
                    scInfo.hproj = timeMachine.getProjectVersion(pvName, curTimeStamp, scInfo.vpid)

                    ' hier unterscheiden, ob Summary-Projekt oder normales
                    If scInfo.hproj.projectType = ptPRPFType.portfolio Then

                        ' lade das Portfolio 
                        Dim err As New clsErrorCodeMsg
                        Dim realTimestamp As Date
                        Dim hproj As clsProjekt

                        Dim aktConst As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(scInfo.pName, scInfo.vpid, realTimestamp, err, scInfo.vName, curTimeStamp)

                        If Not IsNothing(aktConst) Then
                            hproj = calcUnionProject(aktConst, False, curTimeStamp)
                        Else
                            hproj = Nothing
                        End If

                        ' das neu berechnete SummaryProjekt hat den curTimeStamp
                        'hproj.timeStamp = curTimeStamp

                        scInfo.hproj = hproj
                        continueOperation = Not IsNothing(scInfo.hproj)

                    Else
                        ' kann eigentlich nicht mehr Nothing werden ... die Liste an TimeStamps enthält den größten auftretenden kleinsten datumswert aller Projekte ....
                        continueOperation = Not IsNothing(scInfo.hproj)
                    End If


                End If

                If continueOperation Then
                    Try

                        ' -----------------------------
                        ' Alternative 2 und 3: ja, tun
                        'Call createNewHiddenExcel()
                        ' -----------------------------

                        ' jetzt muss das chtobj aktualisiert werden ... 
                        Try

                            If (scInfo.chartTyp = PTChartTypen.Balken) Or
                                (scInfo.chartTyp = PTChartTypen.CurveCumul) Then


                                If scInfo.prPF = ptPRPFType.project Then
                                    Try
                                        Dim a As Integer = scInfo.hproj.dauerInDays

                                        If scInfo.vergleichsTyp = PTVergleichsTyp.erster Then
                                            'scInfo.vglProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).beauftragung
                                            scInfo.vglProj = timeMachine.getFirstContractedVersion(pvName, scInfo.vpid)
                                        ElseIf scInfo.vergleichsTyp = PTVergleichsTyp.letzter Then
                                            'scInfo.vglProj = smartSlideLists.ListOfProjektHistorien.Item(pvName).lastBeauftragung(curTimeStamp.AddMinutes(-1))
                                            scInfo.vglProj = timeMachine.getLastContractedVersion(pvName, curTimeStamp, scInfo.vpid)
                                        End If


                                    Catch ex As Exception

                                        scInfo.vglProj = Nothing

                                    End Try
                                End If


                                ' Alternative 1a - pptApp.activate auskommentiert
                                Call updateProjectChartInPPT(scInfo, pptShape)
                                'pptAPP.Activate()

                                ' -----------------------------------------
                                ' Alternative 2 - funktioniert nicht 
                                'Call updateProjektChartinPPT2(scInfo, pptShape)
                                'pptAPP.Activate()
                                'pptShape.Chart.Refresh()
                                ' --------------------------------------------

                                ' -----------------------------------------
                                ' Alternative 3 - funktioniert etwas unschön , ständiges Update Geflacker 
                                'Call updateProjectChartInPPT3(scInfo, pptShape)
                                'pptAPP.Activate()
                                'pptShape.Chart.Refresh()
                                ' --------------------------------------------


                            ElseIf scInfo.chartTyp = PTChartTypen.Bubble Then



                            ElseIf scInfo.chartTyp = PTChartTypen.Pie Then


                            ElseIf scInfo.chartTyp = PTChartTypen.Waterfall Then


                            ElseIf scInfo.chartTyp = PTChartTypen.ZweiBalken Then

                            Else

                            End If




                        Catch ex As Exception
                            Call MsgBox(ex.Message)
                        End Try



                    Catch ex As Exception
                        Call MsgBox("CreateNewHiddenExcel und chartCopypptPaste:" & ex.Message)
                    End Try

                End If


            End If        'scInfo.pName <> "" or scinfo.vpid <> ""

        End If            ' hasChart


    End Sub


    '''' <summary>
    '''' Breaklink - dann Aufbau der Daten im updateWorkbook - setsourceData - 
    '''' in der übergeordneten Methode ppt.activate, dann refresh chart  
    '''' </summary>
    '''' <param name="scInfo"></param>
    '''' <param name="pptShape"></param>
    'Public Sub updateProjektChartinPPT2(ByVal scInfo As clsSmartPPTChartInfo, ByRef pptShape As PowerPoint.Shape)

    '    Dim pptChart As PowerPoint.Chart = Nothing

    '    If Not (pptShape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue) Then
    '        Exit Sub
    '    End If

    '    pptChart = pptShape.Chart
    '    ' ------ Alternative 2 -------
    '    ' jetzt den Breaklink machen 
    '    pptChart.ChartData.BreakLink()

    '    Dim curWS As Excel.Worksheet = CType(updateWorkbook.Worksheets.Item(1), Excel.Worksheet)
    '    curWS.UsedRange.Clear()
    '    'curWS.Name = "VISBO-Chart"
    '    ' ----------------------------


    'Dim Xdatenreihe() As String
    '    Dim tdatenreihe() As Double
    '    Dim istDatenReihe() As Double
    '    Dim prognoseDatenReihe() As Double
    '    Dim vdatenreihe() As Double
    '    Dim internKapaDatenreihe() As Double = Nothing
    '    Dim vSum As Double = 0.0
    '    Dim tSum As Double


    '    Dim Xdatenreihe() As String
    '    Dim tdatenreihe() As Double
    '    Dim istDatenReihe() As Double
    '    Dim prognoseDatenReihe() As Double
    '    Dim vdatenreihe() As Double
    '    Dim vSum As Double = 0.0
    '    Dim tSum As Double


    '    Dim pkIndex As Integer = CostDefinitions.Count
    '    Dim pstart As Integer

    '    Dim zE As String = awinSettings.kapaEinheit

    '    Dim tmpCollection As New Collection
    '    Dim maxlenTitle1 As Integer = 20

    '    Dim curmaxScale As Double

    '    Dim IstCharttype As Microsoft.Office.Core.XlChartType
    '    Dim PlanChartType As Microsoft.Office.Core.XlChartType
    '    Dim vglChartType As Microsoft.Office.Core.XlChartType

    '    Dim considerIstDaten As Boolean = scInfo.hproj.actualDataUntil > scInfo.hproj.startDate

    '    If scInfo.chartTyp = PTChartTypen.CurveCumul Then
    '        IstCharttype = Microsoft.Office.Core.XlChartType.xlArea

    '        If considerIstDaten Then
    '            PlanChartType = Microsoft.Office.Core.XlChartType.xlArea
    '        Else
    '            PlanChartType = Microsoft.Office.Core.XlChartType.xlLine
    '        End If

    '        vglChartType = Microsoft.Office.Core.XlChartType.xlLine
    '    Else
    '        IstCharttype = Microsoft.Office.Core.XlChartType.xlColumnStacked
    '        PlanChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    '        vglChartType = Microsoft.Office.Core.XlChartType.xlLine
    '    End If


    '    ' die ganzen Vor-Klärungen machen ...
    '    With pptChart

    '        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    '            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    '                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    '                ' hinausgehende Werte hat 
    '                curmaxScale = .MaximumScale
    '                .MaximumScaleIsAuto = False
    '            End With

    '        End If

    '    End With


    '    'Dim pname As String = scInfo.hproj.name

    '    '
    '    ' hole die Projektdauer; berücksichtigen: die können unterschiedlich starten und unterschiedlich lang sein
    '    ' deshalb muss die Zeitspanne bestimmt werden, die beides umfasst  
    '    '

    '    Call bestimmePstartPlen(scInfo, pstart, plen)




    'ReDim Xdatenreihe(plen - 1)
    '    ReDim tdatenreihe(plen - 1)
    '    ReDim istDatenReihe(plen - 1)
    '    ReDim prognoseDatenReihe(plen - 1)
    '    ReDim vdatenreihe(plen - 1)
    '    ReDim internKapaDatenreihe(plen - 1)


    '    ReDim Xdatenreihe(plen - 1)
    '    ReDim tdatenreihe(plen - 1)
    '    ReDim istDatenReihe(plen - 1)
    '    ReDim prognoseDatenReihe(plen - 1)
    '    ReDim vdatenreihe(plen - 1)


    ' hier werden die Istdaten, die Prognosedaten, die Vergleichsdaten sowie die XDaten bestimmt
    'Dim errMsg As String = ""
    '    Call bestimmeXtipvDatenreihen(pstart, plen, scInfo,
    '                                   Xdatenreihe, tdatenreihe, vdatenreihe, istDatenReihe, prognoseDatenReihe, internKapaDatenreihe, errMsg)


    '    ' hier werden die Istdaten, die Prognosedaten, die Vergleichsdaten sowie die XDaten bestimmt
    '    Dim errMsg As String = ""
    '    Call bestimmeXtipvDatenreihen(pstart, plen, scInfo,
    '                                   Xdatenreihe, tdatenreihe, vdatenreihe, istDatenReihe, prognoseDatenReihe, errMsg)

    '    If errMsg <> "" Then
    '        ' es ist ein Fehler aufgetreten
    '        If pptShape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
    '            pptShape.TextFrame2.TextRange.Text = errMsg
    '        End If
    '        Exit Sub
    '    End If


    '    Dim vProjDoesExist As Boolean = Not IsNothing(scInfo.vglProj)

    '    If scInfo.chartTyp = PTChartTypen.CurveCumul Then
    '        tSum = tdatenreihe(tdatenreihe.Length - 1)
    '        vSum = vdatenreihe(vdatenreihe.Length - 1)
    '    Else
    '        tSum = tdatenreihe.Sum
    '        vSum = vdatenreihe.Sum

    '    End If

    '    Dim startRed As Integer = 0
    '    Dim lengthRed As Integer = 0
    '    diagramTitle = bestimmeChartDiagramTitle(scInfo, tSum, vSum, startRed, lengthRed)



    '    With CType(pptChart, PowerPoint.Chart)

    '        ' remove old series
    '        Try
    '            Dim anz As Integer = CInt(CType(.SeriesCollection, PowerPoint.SeriesCollection).Count)
    '            Do While anz > 0
    '                .SeriesCollection(1).Delete()
    '                anz = anz - 1
    '            Loop
    '        Catch ex As Exception

    '        End Try
    '    End With


    '    ' jetzt werden die Collections in dem Chart aufgebaut ...
    '    With CType(pptChart, PowerPoint.Chart)


    '        ' Planung / Forecast
    '        With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

    '            .Name = bestimmeLegendNameIPB("P") & scInfo.hproj.timeStamp.ToShortDateString
    '            .Interior.Color = visboFarbeBlau
    '            .Values = prognoseDatenReihe
    '            .XValues = Xdatenreihe
    '            .ChartType = PlanChartType

    '            If scInfo.chartTyp = PTChartTypen.CurveCumul And Not considerIstDaten Then
    '                ' es handelt sich um eine Line
    '                .Format.Line.Weight = 4
    '                .Format.Line.ForeColor.RGB = visboFarbeBlau
    '                .Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
    '            End If

    '        End With

    '        ' Beauftragung bzw. Vergleichsdaten
    '        If Not IsNothing(scInfo.vglProj) Then

    '            'series
    '            With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
    '                .Name = bestimmeLegendNameIPB("B") & scInfo.vglProj.timeStamp.ToShortDateString
    '                .Values = vdatenreihe
    '                .XValues = Xdatenreihe

    '                .ChartType = vglChartType

    '                If vglChartType = Microsoft.Office.Core.XlChartType.xlLine Then
    '                    With .Format.Line
    '                        .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
    '                        .ForeColor.RGB = visboFarbeOrange
    '                        .Weight = 4
    '                    End With
    '                Else
    '                    ' ggf noch was definieren ..
    '                End If

    '            End With

    '        End If

    '        ' jetzt kommt der Neu-Aufbau der Series-Collections
    '        If considerIstDaten Then

    '            ' jetzt die Istdaten zeichnen 
    '            With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
    '                '.Name = repMessages.getmsg(194) & " " & hproj.timeStamp.ToShortDateString
    '                .Name = bestimmeLegendNameIPB("I")
    '                .Interior.Color = awinSettings.SollIstFarbeArea
    '                .Values = istDatenReihe
    '                .XValues = Xdatenreihe
    '                .ChartType = IstCharttype
    '            End With

    '        End If


    '    End With



    '    ' Skalierung etc anpassen 
    '    With CType(pptChart, PowerPoint.Chart)

    '        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    '            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    '                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    '                ' hinausgehende Werte hat 

    '                If System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) > .MaximumScale - 3 Then
    '                    .MaximumScale = System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) + 3
    '                End If


    '            End With

    '        End If

    '        ' nur wenn es auch einen Titel gibt ... 
    '        If .HasTitle Then
    '            .ChartTitle.Text = diagramTitle
    '        End If


    '    End With

    '    ' -----------------------------------------------
    '    ' 1. Variante : seriesCollections verändern : funktioniert nicht ! Chart wird aktualisiert, aber erst mit interaktiv Bearbeiten-Daten sieht man das auch 
    '    ' 2. Variante : curWS aus HiddenExcel beziehen 
    '    ' 3. Variante : as-is curWS aus 
    '    ' die Frage ist: braucht man das hier wirklich 
    '    ' tk 21.10.18
    '    ' jetzt wird myRange gesetzt und setSourceData gesetzt 
    '    'Dim fZeile As Integer = usedRange.Rows.Count + 1


    '    ' wird in Alternative 2 nicht gebraucht 
    '    'With pptShape.Chart.ChartData
    '    '    .Activate()
    '    '    '.ActivateChartDataWindow()

    '    '    xlApp = CType(CType(.Workbook, Excel.Workbook).Application, Excel.Application)


    '    '    Try

    '    '        If Not CStr(CType(xlApp.ActiveWindow, Excel.Window).Caption) = "VISBO Smart Diagram" Then
    '    '            xlApp.DisplayFormulaBar = False
    '    '            With xlApp.ActiveWindow

    '    '                .Caption = "VISBO Smart Diagram"
    '    '                .DisplayHeadings = False
    '    '                .DisplayWorkbookTabs = False

    '    '                .Width = 500
    '    '                .Height = 150
    '    '                .Top = 100
    '    '                .Left = -1200

    '    '            End With
    '    '        End If

    '    '    Catch ex As Exception

    '    '    End Try

    '    '    curWS = CType(CType(.Workbook, Excel.Workbook).Worksheets.Item(1), Excel.Worksheet)
    '    '    curWS.UsedRange.Clear()

    '    '    If Not smartChartsAreEditable Then
    '    '        With xlApp
    '    '            '.Visible = False
    '    '            '.ActiveWindow.Visible = False
    '    '        End With
    '    '    End If

    '    'End With

    '    Dim fzeile As Integer = 1
    '    Dim anzSpalten As Integer = plen + 1
    '    Dim anzRows As Integer = 0


    '    ' für das SetSourceData 
    '    Dim myRange As Excel.Range = Nothing
    '    'Dim usedRange As Excel.Range = curWS.UsedRange
    '    ' Ende setsource Vorbereitungen 

    '    With curWS
    '        ' neu 

    '        .Cells(fzeile, 1).value = ""
    '        .Range(.Cells(fzeile, 2), .Cells(fzeile, anzSpalten)).Value = Xdatenreihe

    '        If considerIstDaten Then

    '            anzRows = 3

    '            .Cells(fzeile + 1, 1).value = bestimmeLegendNameIPB("I")
    '            .Range(.Cells(fzeile + 1, 2), .Cells(fzeile + 1, anzSpalten)).Value = istDatenReihe

    '            .Cells(fzeile + 2, 1).value = bestimmeLegendNameIPB("P") & scInfo.hproj.timeStamp.ToShortDateString
    '            .Range(.Cells(fzeile + 2, 2), .Cells(fzeile + 2, anzSpalten)).Value = prognoseDatenReihe

    '            If Not IsNothing(scInfo.vglProj) Then

    '                anzRows = 4
    '                .Cells(fzeile + 3, 1).value = bestimmeLegendNameIPB("B") & scInfo.vglProj.timeStamp.ToShortDateString
    '                .Range(.Cells(fzeile + 3, 2), .Cells(fzeile + 3, anzSpalten)).Value = vdatenreihe

    '            End If

    '        Else

    '            anzRows = 2

    '            .Cells(fzeile + 1, 1).value = bestimmeLegendNameIPB("P") & scInfo.hproj.timeStamp.ToShortDateString
    '            .Range(.Cells(fzeile + 1, 2), .Cells(fzeile + 1, anzSpalten)).Value = prognoseDatenReihe

    '            If Not IsNothing(scInfo.vglProj) Then
    '                anzRows = 3

    '                .Cells(fzeile + 2, 1).value = bestimmeLegendNameIPB("B") & scInfo.vglProj.timeStamp.ToShortDateString
    '                .Range(.Cells(fzeile + 2, 2), .Cells(fzeile + 2, anzSpalten)).Value = vdatenreihe

    '            End If

    '        End If

    '        myRange = curWS.Range(.Cells(fzeile, 1), .Cells(fzeile + anzRows - 1, anzSpalten))

    '        ' Ende neu 

    '    End With



    '    Try
    '        ' es ist der Trick, hier die Verbindung zu einem ohnehin bereits non-visible gesetzten Excel herzustellen ...
    '        Dim rangeString As String = "= '" & curWS.Name & "'!" & myRange.Address & ""
    '        pptShape.Chart.SetSourceData(Source:=rangeString)

    '        pptShape.Chart.ChartData.Activate()

    '    Catch ex As Exception

    '    End Try




    'End Sub


    '''' <summary>
    '''' eine hidden ExcelApp ist mit screenupdate = false geöffnet , es wird nur mit seriesCollections gearbeitet
    '''' 
    '''' </summary>
    '''' <param name="scInfo"></param>
    '''' <param name="pptShape"></param>
    'Public Sub updateProjectChartInPPT3(ByVal scInfo As clsSmartPPTChartInfo, ByRef pptShape As PowerPoint.Shape)

    '    Dim pptChart As PowerPoint.Chart = Nothing

    '    If Not (pptShape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue) Then
    '        Exit Sub
    '    End If

    '    pptChart = pptShape.Chart


    'Dim Xdatenreihe() As String
    '    Dim tdatenreihe() As Double
    '    Dim istDatenReihe() As Double
    '    Dim prognoseDatenReihe() As Double
    '    Dim vdatenreihe() As Double
    '    Dim internKapaDatenreihe() As Double
    '    Dim vSum As Double = 0.0
    '    Dim tSum As Double

    '    Dim Xdatenreihe() As String
    '    Dim tdatenreihe() As Double
    '    Dim istDatenReihe() As Double
    '    Dim prognoseDatenReihe() As Double
    '    Dim vdatenreihe() As Double
    '    Dim vSum As Double = 0.0
    '    Dim tSum As Double


    '    Dim pkIndex As Integer = CostDefinitions.Count
    '    Dim pstart As Integer

    '    Dim zE As String = awinSettings.kapaEinheit

    '    Dim tmpCollection As New Collection
    '    Dim maxlenTitle1 As Integer = 20

    '    Dim curmaxScale As Double

    '    Dim IstCharttype As Microsoft.Office.Core.XlChartType
    '    Dim PlanChartType As Microsoft.Office.Core.XlChartType
    '    Dim vglChartType As Microsoft.Office.Core.XlChartType

    '    Dim considerIstDaten As Boolean = scInfo.hproj.actualDataUntil > scInfo.hproj.startDate

    '    If scInfo.chartTyp = PTChartTypen.CurveCumul Then
    '        IstCharttype = Microsoft.Office.Core.XlChartType.xlArea

    '        If considerIstDaten Then
    '            PlanChartType = Microsoft.Office.Core.XlChartType.xlArea
    '        Else
    '            PlanChartType = Microsoft.Office.Core.XlChartType.xlLine
    '        End If

    '        vglChartType = Microsoft.Office.Core.XlChartType.xlLine
    '    Else
    '        IstCharttype = Microsoft.Office.Core.XlChartType.xlColumnStacked
    '        PlanChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    '        vglChartType = Microsoft.Office.Core.XlChartType.xlLine
    '    End If


    '    ' die ganzen Vor-Klärungen machen ...
    '    With pptChart

    '        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    '            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    '                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    '                ' hinausgehende Werte hat 
    '                curmaxScale = .MaximumScale
    '                .MaximumScaleIsAuto = False
    '            End With

    '        End If

    '    End With


    '    'Dim pname As String = scInfo.hproj.name

    '    '
    '    ' hole die Projektdauer; berücksichtigen: die können unterschiedlich starten und unterschiedlich lang sein
    '    ' deshalb muss die Zeitspanne bestimmt werden, die beides umfasst  
    '    '

    '    Call bestimmePstartPlen(scInfo, pstart, plen)




    'ReDim Xdatenreihe(plen - 1)
    '    ReDim tdatenreihe(plen - 1)
    '    ReDim istDatenReihe(plen - 1)
    '    ReDim prognoseDatenReihe(plen - 1)
    '    ReDim vdatenreihe(plen - 1)
    '    ReDim internKapaDatenreihe(plen - 1)


    '    ReDim Xdatenreihe(plen - 1)
    '    ReDim tdatenreihe(plen - 1)
    '    ReDim istDatenReihe(plen - 1)
    '    ReDim prognoseDatenReihe(plen - 1)
    '    ReDim vdatenreihe(plen - 1)


    ' hier werden die Istdaten, die Prognosedaten, die Vergleichsdaten sowie die XDaten bestimmt
    'Dim errMsg As String = ""
    '    Call bestimmeXtipvDatenreihen(pstart, plen, scInfo,
    '                                   Xdatenreihe, tdatenreihe, vdatenreihe, istDatenReihe, prognoseDatenReihe, internKapaDatenreihe, errMsg)


    '    ' hier werden die Istdaten, die Prognosedaten, die Vergleichsdaten sowie die XDaten bestimmt
    '    Dim errMsg As String = ""
    '    Call bestimmeXtipvDatenreihen(pstart, plen, scInfo,
    '                                   Xdatenreihe, tdatenreihe, vdatenreihe, istDatenReihe, prognoseDatenReihe, errMsg)

    '    If errMsg <> "" Then
    '        ' es ist ein Fehler aufgetreten
    '        If pptShape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
    '            pptShape.TextFrame2.TextRange.Text = errMsg
    '        End If
    '        Exit Sub
    '    End If


    '    Dim vProjDoesExist As Boolean = Not IsNothing(scInfo.vglProj)

    '    If scInfo.chartTyp = PTChartTypen.CurveCumul Then
    '        tSum = tdatenreihe(tdatenreihe.Length - 1)
    '        vSum = vdatenreihe(vdatenreihe.Length - 1)
    '    Else
    '        tSum = tdatenreihe.Sum
    '        vSum = vdatenreihe.Sum

    '    End If

    '    Dim startRed As Integer = 0
    '    Dim lengthRed As Integer = 0
    '    diagramTitle = bestimmeChartDiagramTitle(scInfo, tSum, vSum, startRed, lengthRed)



    '    With CType(pptChart, PowerPoint.Chart)

    '        ' remove old series
    '        Try
    '            Dim anz As Integer = CInt(CType(.SeriesCollection, PowerPoint.SeriesCollection).Count)
    '            Do While anz > 0
    '                .SeriesCollection(1).Delete()
    '                anz = anz - 1
    '            Loop
    '        Catch ex As Exception

    '        End Try
    '    End With


    '    ' jetzt werden die Collections in dem Chart aufgebaut ...
    '    With CType(pptChart, PowerPoint.Chart)


    '        ' Planung / Forecast
    '        With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

    '            .Name = bestimmeLegendNameIPB("P") & scInfo.hproj.timeStamp.ToShortDateString
    '            .Interior.Color = visboFarbeBlau
    '            .Values = prognoseDatenReihe
    '            .XValues = Xdatenreihe
    '            .ChartType = PlanChartType

    '            If scInfo.chartTyp = PTChartTypen.CurveCumul And Not considerIstDaten Then
    '                ' es handelt sich um eine Line
    '                .Format.Line.Weight = 4
    '                .Format.Line.ForeColor.RGB = visboFarbeBlau
    '                .Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
    '            End If

    '        End With

    '        ' Beauftragung bzw. Vergleichsdaten
    '        If Not IsNothing(scInfo.vglProj) Then

    '            'series
    '            With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
    '                .Name = bestimmeLegendNameIPB("B") & scInfo.vglProj.timeStamp.ToShortDateString
    '                .Values = vdatenreihe
    '                .XValues = Xdatenreihe

    '                .ChartType = vglChartType

    '                If vglChartType = Microsoft.Office.Core.XlChartType.xlLine Then
    '                    With .Format.Line
    '                        .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
    '                        .ForeColor.RGB = visboFarbeOrange
    '                        .Weight = 4
    '                    End With
    '                Else
    '                    ' ggf noch was definieren ..
    '                End If

    '            End With

    '        End If

    '        ' jetzt kommt der Neu-Aufbau der Series-Collections
    '        If considerIstDaten Then

    '            ' jetzt die Istdaten zeichnen 
    '            With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
    '                '.Name = repMessages.getmsg(194) & " " & hproj.timeStamp.ToShortDateString
    '                .Name = bestimmeLegendNameIPB("I")
    '                .Interior.Color = awinSettings.SollIstFarbeArea
    '                .Values = istDatenReihe
    '                .XValues = Xdatenreihe
    '                .ChartType = IstCharttype
    '            End With

    '        End If


    '    End With

    '    ' Skalierung etc anpassen 
    '    With CType(pptChart, PowerPoint.Chart)

    '        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    '            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    '                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    '                ' hinausgehende Werte hat 

    '                If System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) > .MaximumScale - 3 Then
    '                    .MaximumScale = System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) + 3
    '                End If


    '            End With

    '        End If



    '    End With


    '    ' jetzt wird das Aktivieren gemacht 
    '    With pptShape.Chart.ChartData

    '        Try
    '            '.ActivateChartDataWindow()
    '            .Activate()
    '        Catch ex As Exception
    '            ' in Office 10 und 13 scheint es den Befehl Data Window nicht zu geben ..
    '            .Activate()
    '        End Try


    '        If IsNothing(xlApp) Then
    '            xlApp = CType(CType(.Workbook, Excel.Workbook).Application, Excel.Application)
    '        End If

    '        Try
    '            If Not IsNothing(xlApp) Then
    '                With xlApp
    '                    .Visible = smartChartsAreEditable
    '                    xlApp.DisplayFormulaBar = False
    '                    Try
    '                        If Not IsNothing(.ActiveWindow) Then
    '                            .ActiveWindow.Visible = smartChartsAreEditable
    '                            .ActiveWindow.Caption = "VISBO Smart Diagram"
    '                            .ActiveWindow.DisplayHeadings = False
    '                            .ActiveWindow.DisplayWorkbookTabs = False

    '                            .ActiveWindow.Width = 500
    '                            .ActiveWindow.Height = 150
    '                            .ActiveWindow.Top = 100
    '                            .ActiveWindow.Left = -1200

    '                        End If

    '                    Catch ex As Exception

    '                    End Try

    '                End With

    '            End If

    '        Catch ex As Exception

    '        End Try


    '    End With



    '    ' ---- hier dann final den Titel setzen 
    '    With pptShape.Chart
    '        If .HasTitle Then
    '            .ChartTitle.Text = diagramTitle
    '            .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbBlack

    '            If startRed > 0 And lengthRed > 0 Then
    '                ' die aktuelle Summe muss rot eingefärbt werden 
    '                .ChartTitle.Format.TextFrame2.TextRange.Characters(startRed,
    '                    lengthRed).Font.Fill.ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbRed
    '            End If
    '        End If

    '    End With

    '    pptShape.Chart.Refresh()

    'End Sub

    ''' <summary>
    ''' neue Aktualisierungs Methode von Balken und Curce-cumulated Charts 
    ''' </summary>
    ''' <param name="scInfo"></param>
    ''' <param name="pptShape"></param>
    Public Sub updateProjectChartInPPT(ByVal scInfo As clsSmartPPTChartInfo, ByRef pptShape As PowerPoint.Shape)

        Dim xx As Date = currentTimestamp

        Dim pptChart As PowerPoint.Chart = Nothing
        Dim pptChartData As PowerPoint.ChartData = Nothing
        Dim pptChartDataWB As Excel.Workbook = Nothing


        ' ur:2019-09-19 TestDim xlApp As xlNS.Application

        If Not (pptShape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue) Then
            Exit Sub
        End If

        pptChart = pptShape.Chart
        pptChartData = pptChart.ChartData

        'ur:2019-09-19 Test: funktionsfähig
        If Not IsNothing(pptChartData.Workbook) Then

            If Not pptChartData.IsLinked Then
                With pptChartData
                    .Activate()
                    '.ActivateChartDataWindow()
                    .Workbook.Application.Visible = smartChartsAreEditable
                    .Workbook.Application.Width = 50
                    .Workbook.Application.Height = 15
                    .Workbook.Application.Top = 10
                    .Workbook.Application.Left = -120
                    ' .Workbook.Application.WindowState = -4140 '## Minimize Excel
                End With
            End If

        End If



        Dim diagramTitle As String = " "
        Dim plen As Integer

        Dim Xdatenreihe() As String = Nothing
        Dim tdatenreihe() As Double = Nothing
        Dim istDatenReihe() As Double = Nothing
        Dim prognoseDatenReihe() As Double = Nothing
        Dim vdatenreihe() As Double = Nothing
        Dim internKapaDatenreihe() As Double = Nothing
        ' für Rechnungs-Stellung 
        Dim invoiceDatenreihe() As Double = Nothing
        Dim formerInvoiceDatenreihe() As Double = Nothing

        Dim vSum As Double = 0.0
        Dim tSum As Double


        Dim pkIndex As Integer = CostDefinitions.Count
        Dim pstart As Integer

        Dim zE As String = awinSettings.kapaEinheit

        Dim tmpCollection As New Collection
        Dim maxlenTitle1 As Integer = 20

        Dim curmaxScale As Double

        Dim IstCharttype As Microsoft.Office.Core.XlChartType
        Dim PlanChartType As Microsoft.Office.Core.XlChartType
        Dim vglChartType As Microsoft.Office.Core.XlChartType

        Dim considerIstDaten As Boolean = False
        Dim actualDataIX As Integer = -1


        ' tk 19.4.19 wenn es sich um ein Portfolio handelt, dann muss rausgefunden werden, was der kleinste Ist-Daten-Value ist 
        If scInfo.prPF = ptPRPFType.portfolio Then
            considerIstDaten = (ShowProjekte.actualDataUntil > StartofCalendar.AddMonths(showRangeLeft - 1))
            If considerIstDaten Then
                actualDataIX = getColumnOfDate(ShowProjekte.actualDataUntil) - getColumnOfDate(StartofCalendar.AddMonths(showRangeLeft))
            End If
        ElseIf scInfo.prPF = ptPRPFType.project Then
            considerIstDaten = scInfo.hproj.actualDataUntil > scInfo.hproj.startDate
            If considerIstDaten Then
                actualDataIX = getColumnOfDate(scInfo.hproj.actualDataUntil) - getColumnOfDate(scInfo.hproj.startDate)
            End If
        End If


        If scInfo.chartTyp = PTChartTypen.CurveCumul Then
            IstCharttype = Microsoft.Office.Core.XlChartType.xlLine

            If considerIstDaten Then
                'PlanChartType = Microsoft.Office.Core.XlChartType.xlArea
                PlanChartType = Microsoft.Office.Core.XlChartType.xlLine
            Else
                PlanChartType = Microsoft.Office.Core.XlChartType.xlLine
            End If

            vglChartType = Microsoft.Office.Core.XlChartType.xlLine
        Else
            IstCharttype = Microsoft.Office.Core.XlChartType.xlColumnStacked
            PlanChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
            vglChartType = Microsoft.Office.Core.XlChartType.xlLine
        End If



        ' die ganzen Vor-Klärungen machen ...
        With pptChart

            If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

                With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
                    ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
                    ' hinausgehende Werte hat 
                    curmaxScale = .MaximumScale
                    .MaximumScaleIsAuto = False
                End With

            End If

        End With


        'Dim pname As String = scInfo.hproj.name

        '
        ' hole die Projektdauer; berücksichtigen: die können unterschiedlich starten und unterschiedlich lang sein
        ' deshalb muss die Zeitspanne bestimmt werden, die beides umfasst  
        '

        Call bestimmePstartPlen(scInfo, pstart, plen)

        ReDim Xdatenreihe(plen - 1)
        ReDim tdatenreihe(plen - 1)
        ReDim istDatenReihe(plen - 1)
        ReDim prognoseDatenReihe(plen - 1)
        ReDim vdatenreihe(plen - 1)
        ReDim internKapaDatenreihe(plen - 1)
        ReDim invoiceDatenreihe(plen - 1)
        ReDim formerInvoiceDatenreihe(plen - 1)


        ' hier werden die Istdaten, die Prognosedaten, die Vergleichsdaten sowie die XDaten bestimmt
        Dim errMsg As String = ""
        Call bestimmeXtipvDatenreihen(pstart, plen, scInfo,
                                       Xdatenreihe, tdatenreihe, vdatenreihe, istDatenReihe, prognoseDatenReihe, internKapaDatenreihe, invoiceDatenreihe, formerInvoiceDatenreihe, errMsg)
        If errMsg <> "" Then
            ' es ist ein Fehler aufgetreten
            If pptShape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                pptShape.TextFrame2.TextRange.Text = errMsg
            End If
            Exit Sub
        End If


        Dim vProjDoesExist As Boolean = Not IsNothing(scInfo.vglProj)

        If scInfo.chartTyp = PTChartTypen.CurveCumul Then
            tSum = tdatenreihe(tdatenreihe.Length - 1)
            vSum = vdatenreihe(vdatenreihe.Length - 1)
        Else
            tSum = tdatenreihe.Sum
            vSum = vdatenreihe.Sum

        End If

        Dim startRedGreen As Integer = 0
        Dim lengthRedGreen As Integer = 0
        diagramTitle = bestimmeChartDiagramTitle(scInfo, tSum, vSum, startRedGreen, lengthRedGreen)



        With CType(pptChart, PowerPoint.Chart)

            ' remove old series
            ''Try
            Dim anz As Integer = CInt(CType(.SeriesCollection, PowerPoint.SeriesCollection).Count)
            Do While anz > 0
                .SeriesCollection(1).Delete()
                anz = anz - 1
            Loop
            ''Catch ex As Exception

            ''End Try
        End With


        ' jetzt die Farbe bestimme
        Dim balkenFarbe As Integer = bestimmeBalkenFarbe(scInfo)


        ' jetzt werden die Collections in dem Chart aufgebaut ...
        With CType(pptChart, PowerPoint.Chart)


            Dim dontShowPlanung As Boolean = False

            If scInfo.prPF = ptPRPFType.portfolio Then

                If ShowProjekte.actualDataUntil >= StartofCalendar Then
                    dontShowPlanung = getColumnOfDate(ShowProjekte.actualDataUntil) >= showRangeRight
                End If

            Else
                If scInfo.hproj.hasActualValues Then
                    dontShowPlanung = getColumnOfDate(scInfo.hproj.actualDataUntil) >= getColumnOfDate(scInfo.hproj.endeDate)
                End If
            End If

            If scInfo.chartTyp = PTChartTypen.CurveCumul Then

                ' here Actual Data as well as Forecast Data is shown in one Line 
                ' draw Actual and Plan-Data Line

                ' here the budget / Auftragswert is being drawn 
                Try
                    Dim budgetReihe() As Double = Nothing
                    ReDim budgetReihe(tdatenreihe.Length - 1)
                    Dim mybudgetValue = scInfo.hproj.Erloes
                    If mybudgetValue > 0 Then

                        For ix As Integer = 0 To tdatenreihe.Length - 1
                            budgetReihe(ix) = mybudgetValue
                        Next

                        With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                            .Name = bestimmeLegendNameIPB("BG") & scInfo.hproj.timeStamp.ToShortDateString
                            .Interior.Color = visboFarbeNone
                            .Values = budgetReihe
                            .XValues = Xdatenreihe
                            .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                            .Format.Line.Weight = 2.5
                            .Format.Line.ForeColor.RGB = visboFarbeNone

                            .Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid

                        End With

                    End If
                Catch ex As Exception

                End Try



                With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                    .Name = bestimmeLegendNameIPB("PA") & scInfo.hproj.timeStamp.ToShortDateString
                    .Interior.Color = visboFarbeBlau
                    .Values = tdatenreihe
                    .XValues = Xdatenreihe
                    .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                    .Format.Line.Weight = 4
                    If dontShowPlanung Then
                        .Format.Line.ForeColor.RGB = visboFarbeNone
                    Else
                        .Format.Line.ForeColor.RGB = visboFarbeBlau
                    End If

                    .Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid

                    If considerIstDaten And Not dontShowPlanung Then
                        Try
                            For ix As Integer = 0 To actualDataIX
                                .Points(ix + 1).Format.Line.ForeColor.RGB = visboFarbeNone
                            Next
                        Catch ex As Exception

                        End Try


                    End If

                End With

                ' draw Baseline Line 
                If Not IsNothing(scInfo.vglProj) Then
                    With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                        .Name = bestimmeLegendNameIPB("B") & scInfo.vglProj.timeStamp.ToShortDateString
                        .Interior.Color = visboFarbeOrange
                        .Values = vdatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                        .Format.Line.Weight = 1.5
                        .Format.Line.ForeColor.RGB = visboFarbeOrange
                        .Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash

                    End With
                End If



                If scInfo.elementTyp = ptElementTypen.roleCostInvoices Then

                    ' draw invoice Line 
                    With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                        .Name = bestimmeLegendNameIPB("PIV") & scInfo.hproj.timeStamp.ToShortDateString
                        .Interior.Color = visboFarbeGreen
                        .Values = invoiceDatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                        .Format.Line.Weight = 4
                        .Format.Line.ForeColor.RGB = visboFarbeGreen
                        .Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid

                    End With

                    ' draw invoices of Baseline 
                    With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                        .Name = bestimmeLegendNameIPB("BIV") & scInfo.vglProj.timeStamp.ToShortDateString
                        .Interior.Color = visboFarbeGreen
                        .Values = formerInvoiceDatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                        .Format.Line.Weight = 1.5
                        .Format.Line.ForeColor.RGB = visboFarbeGreen
                        .Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash

                    End With

                End If


            Else

                If Not dontShowPlanung Then
                    With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                        If scInfo.prPF = ptPRPFType.portfolio Then
                            .Name = bestimmeLegendNameIPB("PS") & currentTimestamp.ToShortDateString
                            .Interior.Color = balkenFarbe
                        Else
                            .Name = bestimmeLegendNameIPB("P") & scInfo.hproj.timeStamp.ToShortDateString
                            .Interior.Color = visboFarbeBlau
                        End If


                        .Values = prognoseDatenReihe
                        .XValues = Xdatenreihe
                        .ChartType = PlanChartType

                    End With
                End If

                ' Beauftragung bzw. Vergleichsdaten
                If scInfo.prPF = ptPRPFType.portfolio Then
                    'series
                    With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                        .Name = bestimmeLegendNameIPB("C")
                        .Values = vdatenreihe
                        .XValues = Xdatenreihe

                        .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                        With .Format.Line
                            .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
                            .ForeColor.RGB = visboFarbeRed
                            .Weight = 2
                        End With


                    End With

                    Dim tmpSum As Double = internKapaDatenreihe.Sum
                    If vdatenreihe.Sum > tmpSum And tmpSum > 0 Then
                        ' es gibt geplante externe Ressourcen ... 
                        With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
                            .HasDataLabels = False
                            '.name = "Kapazität incl. Externe"
                            .Name = bestimmeLegendNameIPB("CI")
                            '.Name = repMessages.getmsg(118)

                            .Values = internKapaDatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                            With .Format.Line
                                .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSysDot
                                .ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbFuchsia
                                .Weight = 2
                            End With

                        End With
                    End If
                Else
                    If Not IsNothing(scInfo.vglProj) Then

                        'series
                        With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                            .Name = bestimmeLegendNameIPB("B") & scInfo.vglProj.timeStamp.ToShortDateString
                            .Values = vdatenreihe
                            .XValues = Xdatenreihe

                            .ChartType = vglChartType

                            If vglChartType = Microsoft.Office.Core.XlChartType.xlLine Then
                                With .Format.Line
                                    .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
                                    .ForeColor.RGB = visboFarbeOrange
                                    .Weight = 4
                                End With
                            Else
                                ' ggf noch was definieren ..
                            End If

                        End With

                    End If
                End If


                ' jetzt kommt der Neu-Aufbau der Series-Collections
                If considerIstDaten Then

                    ' jetzt die Istdaten zeichnen 
                    With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
                        If scInfo.prPF = ptPRPFType.portfolio Then
                            .Name = bestimmeLegendNameIPB("IS")
                        Else
                            .Name = bestimmeLegendNameIPB("I")
                        End If
                        .Interior.Color = awinSettings.SollIstFarbeArea
                        .Values = istDatenReihe
                        .XValues = Xdatenreihe
                        .ChartType = IstCharttype
                    End With

                End If

            End If

        End With


        ' tk das wurde am 7.2 auskommentiert, weil das komischerweise zu einer Skalierung der x-Achse geführt hat 
        ' Skalierung etc anpassen 
        With CType(pptChart, PowerPoint.Chart)

            If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

                With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
                    ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
                    ' hinausgehende Werte hat 
                    Dim newMaxScale As Double = CInt(System.Math.Max(tdatenreihe.Max, vdatenreihe.Max)) + 1
                    If newMaxScale > curmaxScale Then
                        .MaximumScale = newMaxScale + 10
                    End If


                End With
            End If
        End With

        ' ---- hier dann final den Titel setzen 
        With pptChart
            If .HasTitle Then
                .ChartTitle.Text = diagramTitle
                .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbBlack

                If startRedGreen > 0 And lengthRedGreen > 0 Then
                    If tSum < vSum Then
                        ' die aktuelle Summe muss grün eingefärbt werden 
                        .ChartTitle.Format.TextFrame2.TextRange.Characters(startRedGreen,
                            lengthRedGreen).Font.Fill.ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbGreen
                    Else
                        ' die aktuelle Summe muss rot eingefärbt werden 
                        .ChartTitle.Format.TextFrame2.TextRange.Characters(startRedGreen,
                            lengthRedGreen).Font.Fill.ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbRed
                    End If

                End If
            End If

        End With


        ' ur:2019-09-19 TestxlApp = CType(CType(pptChart.ChartData.Workbook, Excel.Workbook).Application, Excel.Application)


        'xlApp.Visible = smartChartsAreEditable
        'xlApp.ScreenUpdating = False
        'xlApp.DisplayFormulaBar = False


        'Try
        ' ur:2019-09-19 Test
        'If Not IsNothing(xlApp.ActiveWindow) Then

        '    With xlApp.ActiveWindow
        '        .Visible = smartChartsAreEditable
        '        '.Caption = "VISBO Smart Diagram"
        '        '.DisplayHeadings = False
        '        '.DisplayWorkbookTabs = False

        '        .Width = 50
        '        .Height = 15
        '        .Top = 10
        '        .Left = -120

        '    End With
        'End If

        ''Catch ex As Exception

        ''End Try


        pptChart.Refresh()
        ' ur:2019.06.03: anstatt Try CatchEx ohne Aktion versuchsweise
        'On Error Resume Next
        'On Error GoTo 0
        'pptChartData = Nothing
        'pptChart = Nothing
        pptChartData.BreakLink()



        'If xlApp.Workbooks.Count = 1 Then

        '    For Each wb As Excel.Workbook In xlApp.Workbooks
        '        wb.Saved = True
        '        xlApp.Quit()
        '    Next
        'Else
        '    For Each wb As Excel.Workbook In xlApp.Workbooks
        '        wb.Close(SaveChanges:=False)
        '    Next

        'End If


    End Sub


    '''' <summary>
    '''' aktualisiert das übergebene ppt-Chart direkt in PPT
    '''' </summary>
    '''' <param name="hproj"></param>
    '''' <param name="vglProj"></param>
    '''' <param name="pptShape"></param>
    '''' <param name="prcTyp"></param>
    '''' <param name="auswahl"></param>
    '''' <param name="rcName"></param>
    'Public Sub updatePPTBalkenOfProjectInPPT(ByVal hproj As clsProjekt, ByVal vglProj As clsProjekt,
    '                                    ByRef pptShape As PowerPoint.Shape,
    '                                    ByVal prcTyp As Integer, ByVal auswahl As Integer, ByVal rcName As String)



    '    Dim curWS As Excel.Worksheet = Nothing


    '    Dim pptChart As PowerPoint.Chart = Nothing

    '    If Not (pptShape.HasChart = Microsoft.Office.Core.MsoTriState.msoTrue) Then
    '        Exit Sub
    '    End If

    '    pptChart = pptShape.Chart



    '    Dim kennung As String = pptChart.Name
    '    Dim diagramTitle As String = " "
    '    Dim plen As Integer
    '    Dim i As Integer
    '    Dim Xdatenreihe() As String
    '    Dim tdatenreihe() As Double
    '    Dim istDatenReihe() As Double
    '    Dim prognoseDatenReihe() As Double
    '    Dim vdatenreihe() As Double
    '    Dim vSum As Double = 0.0

    '    Dim hsum(1) As Double, gesamt_summe As Double

    '    Dim pkIndex As Integer = CostDefinitions.Count
    '    Dim pstart As Integer

    '    Dim zE As String = awinSettings.kapaEinheit
    '    Dim titelTeile(1) As String
    '    Dim titelTeilLaengen(1) As Integer
    '    Dim tmpCollection As New Collection
    '    Dim maxlenTitle1 As Integer = 20

    '    Dim curmaxScale As Double
    '    Dim considerIstDaten As Boolean = False

    '    ' die Settings herauslesen ...
    '    Dim chartTyp As String = ""
    '    Dim typID As Integer = -1
    '    Dim rcNameChk As String = ""
    '    Dim tmpPName As String = ""
    '    Call getChartKennungen(kennung, chartTyp, typID, auswahl, tmpPName, rcNameChk)

    '    If rcNameChk <> rcName Then
    '        Dim a As Integer = 1
    '    End If

    '    ' solnage die repMessages noch nicht in der Datenbank sind, muss man sich über dieses Konstrukt behelfen ... 
    '    ' (,0) ist deutsch, (,1) ist englisch

    '    Dim repmsg() As String
    '    ReDim repmsg(6)

    '    If awinSettings.englishLanguage Then
    '        repmsg(0) = "Personnel Costs" '164
    '        repmsg(1) = "Forecast" ' 38
    '        repmsg(2) = "other Costs" ' 165
    '        repmsg(3) = "approved version" ' 273, vorher 43
    '        repmsg(4) = "Personnel Needs" '159
    '        repmsg(5) = "Total Costs" ' 166
    '        repmsg(6) = "Actual data"
    '    Else
    '        repmsg(0) = "Personalkosten" '164
    '        repmsg(1) = "Prognose" ' 38
    '        repmsg(2) = "sonstige Kosten" ' 165
    '        repmsg(3) = "Beauftragung" ' 273 ; Beauftragung 43
    '        repmsg(4) = "Personalbedarf" '159
    '        repmsg(5) = "Gesamtkosten" ' 166
    '        repmsg(6) = "Ist-Werte"
    '    End If

    '    Dim series1Name As String = repmsg(1)
    '    Dim series2Name As String = "-"

    '    ' die ganzen Vor-Klärungen machen ...
    '    With pptChart

    '        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then
    '            '??? If CBool(.HasAxis(Microsoft.Office.Interop.Excel.XlAxisType.xlValue)) Then

    '            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    '                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    '                ' hinausgehende Werte hat 
    '                curmaxScale = .MaximumScale
    '                .MaximumScaleIsAuto = False
    '            End With

    '        End If

    '    End With


    '    Dim pname As String = hproj.name

    '    '
    '    ' hole die Projektdauer
    '    '
    '    With hproj
    '        plen = .anzahlRasterElemente
    '        pstart = .Start
    '    End With

    '    If Not IsNothing(vglProj) Then
    '        If plen < vglProj.anzahlRasterElemente Then
    '            plen = vglProj.anzahlRasterElemente
    '        End If
    '    End If




    '    ReDim Xdatenreihe(plen - 1)
    '    ReDim tdatenreihe(plen - 1)
    '    ReDim istDatenReihe(plen - 1)
    '    ReDim prognoseDatenReihe(plen - 1)
    '    ReDim vdatenreihe(plen - 1)


    '    For i = 1 To plen
    '        Xdatenreihe(i - 1) = hproj.startDate.AddDays(-1 * hproj.startDate.Day + 1).AddMonths(i - 1).ToString("MMM yy", repCult)
    '    Next i



    '    With CType(pptChart, PowerPoint.Chart)

    '        ' remove old series
    '        Try
    '            Dim anz As Integer = CInt(CType(.SeriesCollection, PowerPoint.SeriesCollection).Count)
    '            Do While anz > 0
    '                .SeriesCollection(1).Delete()
    '                anz = anz - 1
    '            Loop
    '        Catch ex As Exception

    '        End Try

    '        'Dim series1Name As String = repmsg(1) & " " & hproj.timeStamp.ToShortDateString ' Stand vom 

    '        If Not IsNothing(vglProj) Then
    '            series2Name = repmsg(3) & " " & vglProj.timeStamp.ToShortDateString ' erste Beauftragung vom 
    '        End If

    '        ' roles, auswahl=1: Personalbedarf
    '        ' roles: auswahl=2: Personalkosten
    '        ' costs: auswahl=1: andere Kosten
    '        ' costs: auswahl=2: Gesamtkosten

    '        If prcTyp = ptElementTypen.roles Then
    '            If auswahl = 2 Then
    '                If rcName = "" Then
    '                    tdatenreihe = hproj.getAllPersonalKosten
    '                    If Not IsNothing(vglProj) Then
    '                        vdatenreihe = vglProj.getAllPersonalKosten
    '                    End If
    '                Else
    '                    tdatenreihe = hproj.getRessourcenBedarf(rcName, inclSubRoles:=True, outPutInEuro:=True)
    '                    If Not IsNothing(vglProj) Then
    '                        vdatenreihe = vglProj.getRessourcenBedarf(rcName, inclSubRoles:=True, outPutInEuro:=True)
    '                    End If
    '                End If

    '            Else
    '                If rcName = "" Then
    '                    tdatenreihe = hproj.getAlleRessourcen
    '                    If Not IsNothing(vglProj) Then
    '                        vdatenreihe = vglProj.getAlleRessourcen
    '                    End If
    '                Else
    '                    tdatenreihe = hproj.getRessourcenBedarf(rcName, True)
    '                    If Not IsNothing(vglProj) Then
    '                        vdatenreihe = vglProj.getRessourcenBedarf(rcName, True)
    '                    End If
    '                End If
    '            End If

    '        ElseIf prcTyp = ptElementTypen.costs Then
    '            If auswahl = 2 Then
    '                tdatenreihe = hproj.getGesamtKostenBedarf
    '                If Not IsNothing(vglProj) Then
    '                    vdatenreihe = vglProj.getGesamtKostenBedarf
    '                End If
    '            Else
    '                tdatenreihe = hproj.getGesamtAndereKosten
    '                If Not IsNothing(vglProj) Then
    '                    vdatenreihe = vglProj.getGesamtAndereKosten
    '                End If
    '            End If
    '        Else
    '            ' darf eigentlich gar nicht sein ... 

    '        End If

    '        gesamt_summe = tdatenreihe.Sum
    '        vSum = 0

    '        Call tdatenreihe.CopyTo(prognoseDatenReihe, 0)

    '        considerIstDaten = hproj.actualDataUntil > hproj.startDate
    '        Dim actualdataIndex As Integer = -1

    '        If considerIstDaten Then

    '            Call tdatenreihe.CopyTo(istDatenReihe, 0)

    '            actualdataIndex = getColumnOfDate(hproj.actualDataUntil) - getColumnOfDate(hproj.startDate)
    '            ' die Prognose Daten bereinigen
    '            For ix As Integer = 0 To actualdataIndex
    '                prognoseDatenReihe(ix) = 0
    '            Next

    '            For ix = actualdataIndex + 1 To plen - 1
    '                istDatenReihe(ix) = 0
    '            Next

    '            '' jetzt die Istdaten zeichnen 
    '            With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
    '                .Name = repmsg(6)
    '                .Interior.Color = awinSettings.SollIstFarbeArea
    '                .Values = istDatenReihe
    '                .XValues = Xdatenreihe
    '                .ChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    '            End With


    '        End If


    '        With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

    '            .ChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    '            .Name = series1Name
    '            .Interior.Color = visboFarbeBlau
    '            .Values = prognoseDatenReihe
    '            .XValues = Xdatenreihe

    '        End With

    '        If Not IsNothing(vglProj) Then

    '            vSum = vdatenreihe.Sum

    '            ''series
    '            With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
    '                .ChartType = Microsoft.Office.Core.XlChartType.xlLine
    '                .Name = series2Name

    '                .Values = vdatenreihe
    '                .XValues = Xdatenreihe

    '                With .Format.Line
    '                    .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
    '                    .ForeColor.RGB = visboFarbeOrange
    '                    .Weight = 4
    '                End With
    '            End With
    '        End If


    '        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    '            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    '                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    '                ' hinausgehende Werte hat 

    '                If System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) > .MaximumScale - 3 Then
    '                    .MaximumScale = System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) + 3
    '                End If


    '            End With

    '        End If

    '        ' nur wenn es auch einen Titel gibt ... 
    '        If .HasTitle Then

    '            ' jetzt muss der Header bestimmt werden 
    '            Dim tmpStr() As String = .ChartTitle.Text.Split(New Char() {CType("(", Char)})
    '            titelTeile(0) = tmpStr(0).Trim

    '            If prcTyp = ptElementTypen.roles Then
    '                If auswahl = 1 Then

    '                    Dim anfText As String = repmsg(4)
    '                    If rcName <> "" Then
    '                        anfText = anfText & " " & rcName
    '                    End If

    '                    If Not IsNothing(vglProj) Then
    '                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " " & zE & ")"
    '                    Else
    '                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " " & zE & ")"
    '                    End If
    '                    titelTeile(1) = ""

    '                ElseIf auswahl = 2 Then

    '                    Dim anfText As String = repmsg(0)
    '                    If rcName <> "" Then
    '                        anfText = anfText & " " & rcName
    '                    End If

    '                    If Not IsNothing(vglProj) Then
    '                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    '                    Else
    '                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    '                    End If
    '                    titelTeile(1) = ""
    '                Else
    '                    titelTeile(0) = "--- (T€)"
    '                    titelTeile(1) = ""
    '                End If
    '            Else

    '                ' jetzt muss das aus Kosten übernommen werden 
    '                If auswahl = 1 Then
    '                    If Not IsNothing(vglProj) Then
    '                        titelTeile(0) = repmsg(2) & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    '                    Else
    '                        titelTeile(0) = repmsg(2) & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    '                    End If
    '                ElseIf auswahl = 2 Then
    '                    If Not IsNothing(vglProj) Then
    '                        titelTeile(0) = repmsg(5) & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    '                    Else
    '                        titelTeile(0) = repmsg(5) & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    '                    End If
    '                Else
    '                    titelTeile(0) = "--- (T€)" & vbLf & pname
    '                End If

    '                titelTeile(1) = ""
    '            End If

    '            titelTeilLaengen(1) = titelTeile(1).Length
    '            titelTeilLaengen(0) = titelTeile(0).Length
    '            diagramTitle = titelTeile(0) & titelTeile(1)

    '            .ChartTitle.Text = diagramTitle
    '        End If


    '    End With

    '    ' tk 21.10.18
    '    ' jetzt wird myRange gesetzt und setSourceData gesetzt 
    '    'Dim fZeile As Integer = usedRange.Rows.Count + 1
    '    Dim fzeile As Integer = 1
    '    Dim anzSpalten As Integer = plen + 1
    '    Dim anzRows As Integer = 0


    '    With pptShape.Chart.ChartData
    '        .Activate()
    '        '.ActivateChartDataWindow()

    '        xlApp = CType(CType(.Workbook, Excel.Workbook).Application, Excel.Application)


    '        Try

    '            If Not CStr(CType(xlApp.ActiveWindow, Excel.Window).Caption) = "VISBO Smart Diagram" Then
    '                xlApp.DisplayFormulaBar = False
    '                With xlApp.ActiveWindow

    '                    .Caption = "VISBO Smart Diagram"
    '                    .DisplayHeadings = False
    '                    .DisplayWorkbookTabs = False

    '                    .Width = 500
    '                    .Height = 150
    '                    .Top = 100
    '                    .Left = -1200

    '                End With
    '            End If

    '        Catch ex As Exception

    '        End Try

    '        curWS = CType(CType(.Workbook, Excel.Workbook).Worksheets.Item(1), Excel.Worksheet)
    '        curWS.UsedRange.Clear()

    '        If Not smartChartsAreEditable Then
    '            With xlApp
    '                .Visible = False
    '                .ActiveWindow.Visible = False
    '            End With
    '        End If


    '    End With



    '    ' für das SetSourceData 
    '    Dim myRange As Excel.Range = Nothing
    '    'Dim usedRange As Excel.Range = curWS.UsedRange
    '    ' Ende setsource Vorbereitungen 

    '    With curWS

    '        ' neu 

    '        .Cells(fzeile, 1).value = ""
    '        .Range(.Cells(fzeile, 2), .Cells(fzeile, anzSpalten)).Value = Xdatenreihe

    '        If considerIstDaten Then

    '            anzRows = 3

    '            .Cells(fzeile + 1, 1).value = repmsg(6)
    '            .Range(.Cells(fzeile + 1, 2), .Cells(fzeile + 1, anzSpalten)).Value = istDatenReihe

    '            .Cells(fzeile + 2, 1).value = series1Name
    '            .Range(.Cells(fzeile + 2, 2), .Cells(fzeile + 2, anzSpalten)).Value = prognoseDatenReihe

    '            If Not IsNothing(vglProj) Then

    '                anzRows = 4
    '                .Cells(fzeile + 3, 1).value = series2Name
    '                .Range(.Cells(fzeile + 3, 2), .Cells(fzeile + 3, anzSpalten)).Value = vdatenreihe

    '            End If

    '        Else

    '            anzRows = 2

    '            .Cells(fzeile + 1, 1).value = series1Name
    '            .Range(.Cells(fzeile + 1, 2), .Cells(fzeile + 1, anzSpalten)).Value = prognoseDatenReihe

    '            If Not IsNothing(vglProj) Then
    '                anzRows = 3

    '                .Cells(fzeile + 2, 1).value = series2Name
    '                .Range(.Cells(fzeile + 2, 2), .Cells(fzeile + 2, anzSpalten)).Value = vdatenreihe

    '            End If

    '        End If

    '        myRange = curWS.Range(.Cells(fzeile, 1), .Cells(fzeile + anzRows - 1, anzSpalten))

    '        ' Ende neu 

    '    End With

    '    '
    '    ' Test tk 21.10.18
    '    'Dim chkvalues1() As String
    '    'Dim chkvalues2() As String
    '    'Dim chkvalues3() As String
    '    'Dim chkvalues4() As String
    '    'Try

    '    '    ReDim chkvalues1(plen)
    '    '    For ix As Integer = 0 To plen
    '    '        chkvalues1(ix) = CStr(myRange.Cells(1, ix + 1).value)
    '    '    Next


    '    '    ReDim chkvalues2(plen)
    '    '    For ix As Integer = 0 To plen
    '    '        chkvalues2(ix) = CStr(myRange.Cells(2, ix + 1).value)
    '    '    Next

    '    '    If anzRows > 2 Then

    '    '        ReDim chkvalues3(plen)
    '    '        For ix As Integer = 0 To plen
    '    '            chkvalues3(ix) = CStr(myRange.Cells(3, ix + 1).value)
    '    '        Next

    '    '        If anzRows > 3 Then

    '    '            ReDim chkvalues4(plen)
    '    '            For ix As Integer = 0 To plen
    '    '                chkvalues4(ix) = CStr(myRange.Cells(4, ix + 1).value)
    '    '            Next
    '    '        End If
    '    '    End If
    '    'Catch ex As Exception

    '    'End Try
    '    '
    '    ' Ende Test tk 21.10.18 
    '    '


    '    Try
    '        ' es ist der Trick, hier die Verbindung zu einem ohnehin bereits non-visible gesetzten Excel herzustellen ...
    '        Dim rangeString As String = "= '" & curWS.Name & "'!" & myRange.Address & ""
    '        pptShape.Chart.SetSourceData(Source:=rangeString)

    '    Catch ex As Exception

    '    End Try

    '    pptShape.Chart.Refresh()


    'End Sub

    ''''' <summary>
    ''''' aktualisiert das übergebene ppt-Chart direkt in PPT
    ''''' funktioniert auch mit 2010 erzeugtem Chart 
    ''''' </summary>
    ''''' <param name="hproj"></param>
    ''''' <param name="vglProj"></param>
    ''''' <param name="pptShape"></param>
    ''''' <param name="prcTyp"></param>
    ''''' <param name="auswahl"></param>
    ''''' <param name="rcName"></param>
    ''Public Sub updatePPTBalkenOfProjectInPPT2(ByVal hproj As clsProjekt, ByVal vglProj As clsProjekt,
    ''                                    ByRef pptShape As PowerPoint.Shape,
    ''                                    ByVal prcTyp As Integer, ByVal auswahl As Integer, ByVal rcName As String)


    ''    Dim pptChart As PowerPoint.Chart = Nothing

    ''    If Not pptShape.HasChart Then
    ''        Exit Sub
    ''    End If

    ''    pptChart = pptShape.Chart


    ''    'Try
    ''    '    myWS = pptChart.ChartData.Workbook.Worksheets.item(1)
    ''    'Catch ex As Exception
    ''    '    myWS = curWS
    ''    'End Try


    ''    Dim kennung As String = pptChart.Name
    ''    Dim diagramTitle As String = " "
    ''    Dim plen As Integer
    ''    Dim i As Integer
    ''    Dim Xdatenreihe() As String
    ''    Dim tdatenreihe() As Double
    ''    Dim istDatenReihe() As Double
    ''    Dim prognoseDatenReihe() As Double
    ''    Dim vdatenreihe() As Double
    ''    Dim vSum As Double = 0.0

    ''    Dim hsum(1) As Double, gesamt_summe As Double

    ''    Dim pkIndex As Integer = CostDefinitions.Count
    ''    Dim pstart As Integer

    ''    Dim zE As String = awinSettings.kapaEinheit
    ''    Dim titelTeile(1) As String
    ''    Dim titelTeilLaengen(1) As Integer
    ''    Dim tmpCollection As New Collection
    ''    Dim maxlenTitle1 As Integer = 20

    ''    Dim curmaxScale As Double
    ''    Dim considerIstDaten As Boolean = False

    ''    ' für das SetSourceData 
    ''    Dim myRange As Excel.Range = Nothing
    ''    ' Ende setsource Vorbereitungen 

    ''    ' die Settings herauslesen ...
    ''    Dim chartTyp As String = ""
    ''    Dim typID As Integer = -1
    ''    Dim rcNameChk As String = ""
    ''    Dim tmpPname As String = ""
    ''    Call getChartKennungen(kennung, chartTyp, typID, auswahl, tmpPname, rcNameChk)

    ''    If rcNameChk <> rcName Then
    ''        Dim a As Integer = 1
    ''    End If

    ''    ' solnage die repMessages noch nicht in der Datenbank sind, muss man sich über dieses Konstrukt behelfen ... 
    ''    ' (,0) ist deutsch, (,1) ist englisch

    ''    Dim repmsg() As String
    ''    ReDim repmsg(6)

    ''    If awinSettings.englishLanguage Then
    ''        repmsg(0) = "Personnel Costs" '164
    ''        repmsg(1) = "Forecast" ' 38
    ''        repmsg(2) = "other Costs" ' 165
    ''        repmsg(3) = "version from" ' 273, vorher 43
    ''        repmsg(4) = "Personnel Needs" '159
    ''        repmsg(5) = "Total Costs" ' 166
    ''        repmsg(6) = "Actual data"
    ''    Else
    ''        repmsg(0) = "Personalkosten" '164
    ''        repmsg(1) = "Prognose" ' 38
    ''        repmsg(2) = "sonstige Kosten" ' 165
    ''        repmsg(3) = "Stand vom" ' 273 ; Beauftragung 43
    ''        repmsg(4) = "Personalbedarf" '159
    ''        repmsg(5) = "Gesamtkosten" ' 166
    ''        repmsg(6) = "Ist-Werte"
    ''    End If

    ''    Dim series1Name As String = repmsg(1)
    ''    Dim series2Name As String = "-"

    ''    ' die ganzen Vor-Klärungen machen ...
    ''    With pptChart

    ''        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    ''            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    ''                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    ''                ' hinausgehende Werte hat 
    ''                curmaxScale = .MaximumScale
    ''                .MaximumScaleIsAuto = False
    ''            End With

    ''        End If

    ''    End With


    ''    Dim pname As String = hproj.name

    ''    '
    ''    ' hole die Projektdauer
    ''    '
    ''    With hproj
    ''        plen = .anzahlRasterElemente
    ''        pstart = .Start
    ''    End With

    ''    If Not IsNothing(vglProj) Then
    ''        If plen < vglProj.anzahlRasterElemente Then
    ''            plen = vglProj.anzahlRasterElemente
    ''        End If
    ''    End If

    ''    '
    ''    ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
    ''    '
    ''    '
    ''    ' hole die Anzahl Rollen, die in diesem Projekt vorkommen
    ''    '
    ''    ' tk 9.8.18 braucht man hier nicht 
    ''    ''If prcTyp = ptElementTypen.roles Then
    ''    ''    ErgebnisListeRC = hproj.getRoleNames
    ''    ''Else
    ''    ''    ErgebnisListeRC = hproj.getCostNames
    ''    ''End If

    ''    ''anzElemente = ErgebnisListeRC.Count




    ''    ReDim Xdatenreihe(plen - 1)
    ''    ReDim tdatenreihe(plen - 1)
    ''    ReDim istDatenReihe(plen - 1)
    ''    ReDim prognoseDatenReihe(plen - 1)
    ''    ReDim vdatenreihe(plen - 1)


    ''    For i = 1 To plen
    ''        Xdatenreihe(i - 1) = hproj.startDate.AddDays(-1 * hproj.startDate.Day + 1).AddMonths(i - 1).ToString("MMM yy", repCult)
    ''    Next i



    ''    With CType(pptChart, PowerPoint.Chart)

    ''        ' remove old series
    ''        'Try
    ''        '    Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
    ''        '    Do While anz > 0
    ''        '        .SeriesCollection(1).Delete()
    ''        '        anz = anz - 1
    ''        '    Loop
    ''        'Catch ex As Exception

    ''        'End Try

    ''        'Dim series1Name As String = repmsg(1) & " " & hproj.timeStamp.ToShortDateString ' Stand vom 

    ''        If Not IsNothing(vglProj) Then
    ''            series2Name = repmsg(3) & " " & vglProj.timeStamp.ToShortDateString ' erste Beauftragung vom 
    ''        End If

    ''        ' roles, auswahl=1: Personalbedarf
    ''        ' roles: auswahl=2: Personalkosten
    ''        ' costs: auswahl=1: andere Kosten
    ''        ' costs: auswahl=2: Gesamtkosten

    ''        If prcTyp = ptElementTypen.roles Then
    ''            If auswahl = 2 Then
    ''                If rcName = "" Then
    ''                    tdatenreihe = hproj.getAllPersonalKosten
    ''                    If Not IsNothing(vglProj) Then
    ''                        vdatenreihe = vglProj.getAllPersonalKosten
    ''                    End If
    ''                Else
    ''                    tdatenreihe = hproj.getPersonalKosten(rcName, True)
    ''                    If Not IsNothing(vglProj) Then
    ''                        vdatenreihe = vglProj.getPersonalKosten(rcName, True)
    ''                    End If
    ''                End If

    ''            Else
    ''                If rcName = "" Then
    ''                    tdatenreihe = hproj.getAlleRessourcen
    ''                    If Not IsNothing(vglProj) Then
    ''                        vdatenreihe = vglProj.getAlleRessourcen
    ''                    End If
    ''                Else
    ''                    tdatenreihe = hproj.getRessourcenBedarfNew(rcName, True)
    ''                    If Not IsNothing(vglProj) Then
    ''                        vdatenreihe = vglProj.getRessourcenBedarfNew(rcName, True)
    ''                    End If
    ''                End If
    ''            End If

    ''        ElseIf prcTyp = ptElementTypen.costs Then
    ''            If auswahl = 2 Then
    ''                tdatenreihe = hproj.getGesamtKostenBedarf
    ''                If Not IsNothing(vglProj) Then
    ''                    vdatenreihe = vglProj.getGesamtKostenBedarf
    ''                End If
    ''            Else
    ''                tdatenreihe = hproj.getGesamtAndereKosten
    ''                If Not IsNothing(vglProj) Then
    ''                    vdatenreihe = vglProj.getGesamtAndereKosten
    ''                End If
    ''            End If
    ''        Else
    ''            ' darf eigentlich gar nicht sein ... 

    ''        End If

    ''        gesamt_summe = tdatenreihe.Sum
    ''        vSum = 0

    ''        Call tdatenreihe.CopyTo(prognoseDatenReihe, 0)

    ''        considerIstDaten = hproj.actualDataUntil > hproj.startDate
    ''        Dim actualdataIndex As Integer = -1

    ''        If considerIstDaten Then

    ''            Call tdatenreihe.CopyTo(istDatenReihe, 0)

    ''            actualdataIndex = getColumnOfDate(hproj.actualDataUntil) - getColumnOfDate(hproj.startDate)
    ''            ' die Prognose Daten bereinigen
    ''            For ix As Integer = 0 To actualdataIndex
    ''                prognoseDatenReihe(ix) = 0
    ''            Next

    ''            For ix = actualdataIndex + 1 To plen - 1
    ''                istDatenReihe(ix) = 0
    ''            Next

    ''            '' jetzt die Istdaten zeichnen 
    ''            'With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
    ''            '    '.Name = repmsg(6) & " " & hproj.timeStamp.ToShortDateString
    ''            '    .Name = repmsg(6)
    ''            '    '.Interior.Color = visboFarbeBlau
    ''            '    .Interior.Color = awinSettings.SollIstFarbeArea
    ''            '    .Values = istDatenReihe
    ''            '    .XValues = Xdatenreihe
    ''            '    .ChartType = Excel.XlChartType.xlColumnStacked
    ''            'End With


    ''        End If


    ''        'With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)

    ''        '    .ChartType = Excel.XlChartType.xlColumnStacked
    ''        '    .Name = series1Name
    ''        '    .Interior.Color = visboFarbeBlau
    ''        '    '.Interior.Color = visboFarbeYellow
    ''        '    '.Values = tdatenreihe
    ''        '    .Values = prognoseDatenReihe
    ''        '    .XValues = Xdatenreihe

    ''        'End With

    ''        If Not IsNothing(vglProj) Then

    ''            vSum = vdatenreihe.Sum

    ''            ''series
    ''            'With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
    ''            '    .ChartType = Excel.XlChartType.xlLine
    ''            '    .Name = series2Name

    ''            '    .Values = vdatenreihe
    ''            '    .XValues = Xdatenreihe

    ''            '    With .Format.Line
    ''            '        .DashStyle = core.MsoLineDashStyle.msoLineDash
    ''            '        '.ForeColor.RGB = Excel.XlRgbColor.rgbFireBrick
    ''            '        .ForeColor.RGB = visboFarbeOrange
    ''            '        .Weight = 4
    ''            '    End With
    ''            'End With
    ''        End If


    ''        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    ''            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    ''                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    ''                ' hinausgehende Werte hat 

    ''                If System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) > .MaximumScale - 3 Then
    ''                    .MaximumScale = System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) + 3
    ''                End If


    ''            End With

    ''        End If

    ''        ' nur wenn es auch einen Titel gibt ... 
    ''        If .HasTitle Then
    ''            ' jetzt muss der Header bestimmt werden 
    ''            Dim tmpStr() As String = .ChartTitle.Text.Split(New Char() {CType("(", Char)})
    ''            titelTeile(0) = tmpStr(0).Trim

    ''            If prcTyp = ptElementTypen.roles Then
    ''                If auswahl = 1 Then

    ''                    Dim anfText As String = repmsg(4)
    ''                    If rcName <> "" Then
    ''                        anfText = anfText & " " & rcName
    ''                    End If

    ''                    If Not IsNothing(vglProj) Then
    ''                        'titelTeile(0) = repMessages.getmsg(159) & " (" & gesamt_summe.ToString("####0.") & " / " & vSum.ToString("####0.") & " " & zE & ")"
    ''                        'titelTeile(0) = repmsg(4) & " (" & vSum.ToString("####0.") & " / " & gesamt_summe.ToString("####0.") & " " & zE & ")"
    ''                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " " & zE & ")"
    ''                    Else
    ''                        'titelTeile(0) = repMessages.getmsg(159) & " (" & gesamt_summe.ToString("####0.") & " " & zE & ")"
    ''                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " " & zE & ")"
    ''                    End If
    ''                    titelTeile(1) = ""

    ''                ElseIf auswahl = 2 Then

    ''                    Dim anfText As String = repmsg(0)
    ''                    If rcName <> "" Then
    ''                        anfText = anfText & " " & rcName
    ''                    End If

    ''                    If Not IsNothing(vglProj) Then
    ''                        'titelTeile(0) = repMessages.getmsg(160) & " (" & gesamt_summe.ToString("####0.") & " / " & vSum.ToString("####0.") & " T€" & ")"
    ''                        'titelTeile(0) = repmsg(0) & " (" & vSum.ToString("####0.") & " / " & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    ''                    Else
    ''                        'titelTeile(0) = repMessages.getmsg(160) & " (" & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    ''                    End If
    ''                    titelTeile(1) = ""
    ''                Else
    ''                    titelTeile(0) = "--- (T€)"
    ''                    titelTeile(1) = ""
    ''                End If
    ''            Else
    ''                ' jetzt muss das aus Kosten übernommen werden 
    ''                'titelTeile(0) = repMessages.getmsg(165) & " (" & gesamt_Summe.ToString("####0.") & " T€" & ")"
    ''                If auswahl = 1 Then
    ''                    If Not IsNothing(vglProj) Then
    ''                        'titelTeile(0) = repMessages.getmsg(165) & " (" & gesamt_summe.ToString("####0.") & " / " & vSum.ToString("####0.") & " T€" & ")"
    ''                        'titelTeile(0) = repmsg(2) & " (" & vSum.ToString("####0.") & " / " & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = repmsg(2) & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    ''                    Else
    ''                        'titelTeile(0) = repMessages.getmsg(165) & " (" & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = repmsg(2) & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    ''                    End If
    ''                ElseIf auswahl = 2 Then
    ''                    If Not IsNothing(vglProj) Then
    ''                        'titelTeile(0) = repMessages.getmsg(166) & " (" & gesamt_summe.ToString("####0.") & " / " & vSum.ToString("####0.") & " T€" & ")"
    ''                        'titelTeile(0) = repmsg(5) & " (" & vSum.ToString("####0.") & " / " & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = repmsg(5) & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    ''                    Else
    ''                        'titelTeile(0) = repMessages.getmsg(166) & " (" & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = repmsg(5) & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    ''                    End If
    ''                Else
    ''                    titelTeile(0) = "--- (T€)" & vbLf & pname
    ''                End If

    ''                titelTeile(1) = ""
    ''            End If

    ''            titelTeilLaengen(1) = titelTeile(1).Length
    ''            titelTeilLaengen(0) = titelTeile(0).Length
    ''            diagramTitle = titelTeile(0) & titelTeile(1)

    ''            .ChartTitle.Text = diagramTitle
    ''        End If


    ''    End With


    ''    With pptChart.ChartData
    ''        .Activate()
    ''    End With

    ''    Dim myWb As Excel.Workbook = CType(pptChart.ChartData.Workbook, Excel.Workbook)
    ''    myWb.Application.Visible = False
    ''    Dim myWS As Excel.Worksheet = myWb.ActiveSheet
    ''    myWS.UsedRange.Clear()

    ''    ' tk 21.10.18
    ''    ' jetzt wird myRange gesetzt und setSourceData gesetzt 
    ''    Dim fZeile As Integer = 1
    ''    Dim anzSpalten As Integer = plen + 1
    ''    Dim anzRows As Integer = 0

    ''    With myWS

    ''        ' neu 

    ''        .Cells(fZeile, 1).value = ""
    ''        .Range(.Cells(fZeile, 2), .Cells(fZeile, anzSpalten)).Value = Xdatenreihe

    ''        If considerIstDaten Then

    ''            anzRows = 3

    ''            .Cells(fZeile + 1, 1).value = repmsg(6)
    ''            .Range(.Cells(fZeile + 1, 2), .Cells(fZeile + 1, anzSpalten)).Value = istDatenReihe

    ''            .Cells(fZeile + 2, 1).value = series1Name
    ''            .Range(.Cells(fZeile + 2, 2), .Cells(fZeile + 2, anzSpalten)).Value = prognoseDatenReihe

    ''            If Not IsNothing(vglProj) Then

    ''                anzRows = 4
    ''                .Cells(fZeile + 3, 1).value = series2Name
    ''                .Range(.Cells(fZeile + 3, 2), .Cells(fZeile + 3, anzSpalten)).Value = vdatenreihe

    ''            End If

    ''        Else

    ''            anzRows = 2

    ''            .Cells(fZeile + 1, 1).value = series1Name
    ''            .Range(.Cells(fZeile + 1, 2), .Cells(fZeile + 1, anzSpalten)).Value = prognoseDatenReihe

    ''            If Not IsNothing(vglProj) Then
    ''                anzRows = 3

    ''                .Cells(fZeile + 2, 1).value = series2Name
    ''                .Range(.Cells(fZeile + 2, 2), .Cells(fZeile + 2, anzSpalten)).Value = vdatenreihe

    ''            End If

    ''        End If

    ''        myRange = myWS.Range(.Cells(fZeile, 1), .Cells(fZeile + anzRows - 1, anzSpalten))

    ''        ' Ende neu 

    ''    End With

    ''    ' Test tk 21.10.18
    ''    Dim chkvalues1() As String
    ''    Dim chkvalues2() As String
    ''    Dim chkvalues3() As String
    ''    Dim chkvalues4() As String
    ''    Try

    ''        ReDim chkvalues1(plen)
    ''        For ix As Integer = 0 To plen
    ''            chkvalues1(ix) = CStr(myRange.Cells(1, ix + 1).value)
    ''        Next


    ''        ReDim chkvalues2(plen)
    ''        For ix As Integer = 0 To plen
    ''            chkvalues2(ix) = CStr(myRange.Cells(2, ix + 1).value)
    ''        Next

    ''        If anzRows > 2 Then

    ''            ReDim chkvalues3(plen)
    ''            For ix As Integer = 0 To plen
    ''                chkvalues3(ix) = CStr(myRange.Cells(3, ix + 1).value)
    ''            Next

    ''            If anzRows > 3 Then

    ''                ReDim chkvalues4(plen)
    ''                For ix As Integer = 0 To plen
    ''                    chkvalues4(ix) = CStr(myRange.Cells(4, ix + 1).value)
    ''                Next
    ''            End If
    ''        End If
    ''    Catch ex As Exception

    ''    End Try
    ''    ' Ende Test tk 21.10.18 

    ''    Dim rangeString As String = "= '" & myWS.Name & "'!" & myRange.Address & ""

    ''    pptShape.Chart.SetSourceData(Source:=rangeString)


    ''    '
    ''    ' ---- ab hier die SeriesCollection setzen 
    ''    If considerIstDaten Then
    ''        ' die Istdaten ... 
    ''        With CType(CType(CType(pptShape.Chart, PowerPoint.Chart).SeriesCollection, PowerPoint.SeriesCollection).Item(1), PowerPoint.Series)
    ''            .Interior.Color = awinSettings.SollIstFarbeArea
    ''            .ChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    ''        End With

    ''        ' die Prognose-Daten ... 
    ''        With CType(CType(CType(pptShape.Chart, PowerPoint.Chart).SeriesCollection, PowerPoint.SeriesCollection).Item(2), PowerPoint.Series)
    ''            .Interior.Color = visboFarbeBlau
    ''            .ChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    ''            With .Format.Line
    ''                .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
    ''                .ForeColor.RGB = PowerPoint.XlRgbColor.rgbWhite
    ''                .Weight = 0
    ''            End With
    ''        End With

    ''        If Not IsNothing(vglProj) Then
    ''            With CType(CType(CType(pptShape.Chart, PowerPoint.Chart).SeriesCollection, PowerPoint.SeriesCollection).Item(3), PowerPoint.Series)
    ''                .ChartType = Microsoft.Office.Core.XlChartType.xlLine
    ''                With .Format.Line
    ''                    .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
    ''                    .ForeColor.RGB = visboFarbeOrange
    ''                    .Weight = 4
    ''                End With
    ''            End With
    ''        End If
    ''    Else
    ''        ' nur die Prognose-Daten ... 
    ''        With CType(CType(CType(pptShape.Chart, PowerPoint.Chart).SeriesCollection, PowerPoint.SeriesCollection).Item(1), PowerPoint.Series)
    ''            .Interior.Color = visboFarbeBlau
    ''            .ChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    ''        End With

    ''        If Not IsNothing(vglProj) Then
    ''            With CType(CType(CType(pptShape.Chart, PowerPoint.Chart).SeriesCollection, PowerPoint.SeriesCollection).Item(2), PowerPoint.Series)
    ''                .ChartType = Microsoft.Office.Core.XlChartType.xlLine
    ''                With .Format.Line
    ''                    .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
    ''                    .ForeColor.RGB = visboFarbeOrange
    ''                    .Weight = 4
    ''                End With
    ''            End With
    ''        End If
    ''    End If

    ''    ' jetzt muss der Test kommen 
    ''    'Try
    ''    '    If Not IsNothing(pptShape.Chart.ChartData) Then
    ''    '        With pptShape.Chart.ChartData
    ''    '            Call MsgBox("isLinked: " & .IsLinked.ToString)
    ''    '        End With
    ''    '    End If
    ''    'Catch ex As Exception

    ''    'End Try

    ''    pptShape.Chart.Refresh()


    ''End Sub

    ''''' <summary>
    ''''' funktioniert mit einem in 2016 erzeugten Bericht ebenso wie mit einem in 2010 erzeugten Bericht 
    ''''' </summary>
    ''''' <param name="hproj"></param>
    ''''' <param name="vglProj"></param>
    ''''' <param name="pptShape"></param>
    ''''' <param name="prcTyp"></param>
    ''''' <param name="auswahl"></param>
    ''''' <param name="rcName"></param>
    ''Public Sub updatePPTBalkenOfProjectInPPT3(ByVal hproj As clsProjekt, ByVal vglProj As clsProjekt,
    ''                                    ByRef pptShape As PowerPoint.Shape,
    ''                                    ByVal prcTyp As Integer, ByVal auswahl As Integer, ByVal rcName As String)


    ''    Dim pptChart As PowerPoint.Chart = Nothing

    ''    If Not pptShape.HasChart Then
    ''        Exit Sub
    ''    End If

    ''    pptChart = pptShape.Chart


    ''    'Try
    ''    '    myWS = pptChart.ChartData.Workbook.Worksheets.item(1)
    ''    'Catch ex As Exception
    ''    '    myWS = curWS
    ''    'End Try


    ''    Dim kennung As String = pptChart.Name
    ''    Dim diagramTitle As String = " "
    ''    Dim plen As Integer
    ''    Dim i As Integer
    ''    Dim Xdatenreihe() As String
    ''    Dim tdatenreihe() As Double
    ''    Dim istDatenReihe() As Double
    ''    Dim prognoseDatenReihe() As Double
    ''    Dim vdatenreihe() As Double
    ''    Dim vSum As Double = 0.0

    ''    Dim hsum(1) As Double, gesamt_summe As Double

    ''    Dim pkIndex As Integer = CostDefinitions.Count
    ''    Dim pstart As Integer

    ''    Dim zE As String = awinSettings.kapaEinheit
    ''    Dim titelTeile(1) As String
    ''    Dim titelTeilLaengen(1) As Integer
    ''    Dim tmpCollection As New Collection
    ''    Dim maxlenTitle1 As Integer = 20

    ''    Dim curmaxScale As Double
    ''    Dim considerIstDaten As Boolean = False

    ''    ' für das SetSourceData 
    ''    Dim myRange As Excel.Range = Nothing
    ''    ' Ende setsource Vorbereitungen 

    ''    ' die Settings herauslesen ...
    ''    Dim chartTyp As String = ""
    ''    Dim typID As Integer = -1
    ''    Dim rcNameChk As String = ""
    ''    Dim tmpPname As String = ""
    ''    Call getChartKennungen(kennung, chartTyp, typID, auswahl, tmpPname, rcNameChk)

    ''    If rcNameChk <> rcName Then
    ''        Dim a As Integer = 1
    ''    End If

    ''    ' solnage die repMessages noch nicht in der Datenbank sind, muss man sich über dieses Konstrukt behelfen ... 
    ''    ' (,0) ist deutsch, (,1) ist englisch

    ''    Dim repmsg() As String
    ''    ReDim repmsg(6)

    ''    If awinSettings.englishLanguage Then
    ''        repmsg(0) = "Personnel Costs" '164
    ''        repmsg(1) = "Forecast" ' 38
    ''        repmsg(2) = "other Costs" ' 165
    ''        repmsg(3) = "version from" ' 273, vorher 43
    ''        repmsg(4) = "Personnel Needs" '159
    ''        repmsg(5) = "Total Costs" ' 166
    ''        repmsg(6) = "Actual data"
    ''    Else
    ''        repmsg(0) = "Personalkosten" '164
    ''        repmsg(1) = "Prognose" ' 38
    ''        repmsg(2) = "sonstige Kosten" ' 165
    ''        repmsg(3) = "Stand vom" ' 273 ; Beauftragung 43
    ''        repmsg(4) = "Personalbedarf" '159
    ''        repmsg(5) = "Gesamtkosten" ' 166
    ''        repmsg(6) = "Ist-Werte"
    ''    End If

    ''    Dim series1Name As String = repmsg(1)
    ''    Dim series2Name As String = "-"

    ''    ' die ganzen Vor-Klärungen machen ...
    ''    With pptChart

    ''        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    ''            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    ''                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    ''                ' hinausgehende Werte hat 
    ''                curmaxScale = .MaximumScale
    ''                .MaximumScaleIsAuto = False
    ''            End With

    ''        End If

    ''    End With


    ''    Dim pname As String = hproj.name

    ''    '
    ''    ' hole die Projektdauer
    ''    '
    ''    With hproj
    ''        plen = .anzahlRasterElemente
    ''        pstart = .Start
    ''    End With

    ''    If Not IsNothing(vglProj) Then
    ''        If plen < vglProj.anzahlRasterElemente Then
    ''            plen = vglProj.anzahlRasterElemente
    ''        End If
    ''    End If

    ''    '
    ''    ' hole die Anzahl Kostenarten, die in diesem Projekt vorkommen
    ''    '
    ''    '
    ''    ' hole die Anzahl Rollen, die in diesem Projekt vorkommen
    ''    '
    ''    ' tk 9.8.18 braucht man hier nicht 
    ''    ''If prcTyp = ptElementTypen.roles Then
    ''    ''    ErgebnisListeRC = hproj.getRoleNames
    ''    ''Else
    ''    ''    ErgebnisListeRC = hproj.getCostNames
    ''    ''End If

    ''    ''anzElemente = ErgebnisListeRC.Count




    ''    ReDim Xdatenreihe(plen - 1)
    ''    ReDim tdatenreihe(plen - 1)
    ''    ReDim istDatenReihe(plen - 1)
    ''    ReDim prognoseDatenReihe(plen - 1)
    ''    ReDim vdatenreihe(plen - 1)


    ''    For i = 1 To plen
    ''        Xdatenreihe(i - 1) = hproj.startDate.AddDays(-1 * hproj.startDate.Day + 1).AddMonths(i - 1).ToString("MMM yy", repCult)
    ''    Next i



    ''    With CType(pptChart, PowerPoint.Chart)

    ''        ' remove old series
    ''        Try
    ''            Dim anz As Integer = CInt(CType(.SeriesCollection, PowerPoint.SeriesCollection).Count)
    ''            Do While anz > 0
    ''                .SeriesCollection(1).Delete()
    ''                anz = anz - 1
    ''            Loop
    ''        Catch ex As Exception

    ''        End Try

    ''        'Dim series1Name As String = repmsg(1) & " " & hproj.timeStamp.ToShortDateString ' Stand vom 

    ''        If Not IsNothing(vglProj) Then
    ''            series2Name = repmsg(3) & " " & vglProj.timeStamp.ToShortDateString ' erste Beauftragung vom 
    ''        End If

    ''        If prcTyp = ptElementTypen.roles Then
    ''            If auswahl = 2 Then
    ''                If rcName = "" Then
    ''                    tdatenreihe = hproj.getAllPersonalKosten
    ''                    If Not IsNothing(vglProj) Then
    ''                        vdatenreihe = vglProj.getAllPersonalKosten
    ''                    End If
    ''                Else
    ''                    tdatenreihe = hproj.getPersonalKosten(rcName, True)
    ''                    If Not IsNothing(vglProj) Then
    ''                        vdatenreihe = vglProj.getPersonalKosten(rcName, True)
    ''                    End If
    ''                End If

    ''            Else
    ''                If rcName = "" Then
    ''                    tdatenreihe = hproj.getAlleRessourcen
    ''                    If Not IsNothing(vglProj) Then
    ''                        vdatenreihe = vglProj.getAlleRessourcen
    ''                    End If
    ''                Else
    ''                    tdatenreihe = hproj.getRessourcenBedarfNew(rcName, True)
    ''                    If Not IsNothing(vglProj) Then
    ''                        vdatenreihe = vglProj.getRessourcenBedarfNew(rcName, True)
    ''                    End If
    ''                End If
    ''            End If

    ''        ElseIf prcTyp = ptElementTypen.costs Then
    ''            If auswahl = 2 Then
    ''                tdatenreihe = hproj.getGesamtKostenBedarf
    ''                If Not IsNothing(vglProj) Then
    ''                    vdatenreihe = vglProj.getGesamtKostenBedarf
    ''                End If
    ''            Else
    ''                tdatenreihe = hproj.getGesamtAndereKosten
    ''                If Not IsNothing(vglProj) Then
    ''                    vdatenreihe = vglProj.getGesamtAndereKosten
    ''                End If
    ''            End If
    ''        Else
    ''            ' darf eigentlich gar nicht sein ... 

    ''        End If

    ''        gesamt_summe = tdatenreihe.Sum
    ''        vSum = 0

    ''        Call tdatenreihe.CopyTo(prognoseDatenReihe, 0)

    ''        considerIstDaten = hproj.actualDataUntil > hproj.startDate
    ''        Dim actualdataIndex As Integer = -1

    ''        If considerIstDaten Then

    ''            Call tdatenreihe.CopyTo(istDatenReihe, 0)

    ''            actualdataIndex = getColumnOfDate(hproj.actualDataUntil) - getColumnOfDate(hproj.startDate)
    ''            ' die Prognose Daten bereinigen
    ''            For ix As Integer = 0 To actualdataIndex
    ''                prognoseDatenReihe(ix) = 0
    ''            Next

    ''            For ix = actualdataIndex + 1 To plen - 1
    ''                istDatenReihe(ix) = 0
    ''            Next

    ''            '' jetzt die Istdaten zeichnen 
    ''            With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
    ''                '.Name = repmsg(6) & " " & hproj.timeStamp.ToShortDateString
    ''                .Name = repmsg(6)
    ''                '.Interior.Color = visboFarbeBlau
    ''                .Interior.Color = awinSettings.SollIstFarbeArea
    ''                .Values = istDatenReihe
    ''                .XValues = Xdatenreihe
    ''                .ChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    ''            End With


    ''        End If


    ''        With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

    ''            .ChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
    ''            .Name = series1Name
    ''            .Interior.Color = visboFarbeBlau
    ''            '.Interior.Color = visboFarbeYellow
    ''            '.Values = tdatenreihe
    ''            .Values = prognoseDatenReihe
    ''            .XValues = Xdatenreihe

    ''        End With

    ''        If Not IsNothing(vglProj) Then

    ''            vSum = vdatenreihe.Sum

    ''            ''series
    ''            With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
    ''                .ChartType = Microsoft.Office.Core.XlChartType.xlLine
    ''                .Name = series2Name

    ''                .Values = vdatenreihe
    ''                .XValues = Xdatenreihe

    ''                With .Format.Line
    ''                    .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
    ''                    '.ForeColor.RGB = Excel.XlRgbColor.rgbFireBrick
    ''                    .ForeColor.RGB = visboFarbeOrange
    ''                    .Weight = 4
    ''                End With
    ''            End With
    ''        End If


    ''        If CBool(.HasAxis(PowerPoint.XlAxisType.xlValue)) Then

    ''            With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)
    ''                ' das ist dann relevant, wenn ein anderes Projekt selektiert wird, das über die aktuelle Skalierung 
    ''                ' hinausgehende Werte hat 

    ''                If System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) > .MaximumScale - 3 Then
    ''                    .MaximumScale = System.Math.Max(tdatenreihe.Max, vdatenreihe.Max) + 3
    ''                End If


    ''            End With

    ''        End If

    ''        ' nur wenn es auch einen Titel gibt ... 
    ''        If .HasTitle Then
    ''            ' jetzt muss der Header bestimmt werden 
    ''            Dim tmpStr() As String = .ChartTitle.Text.Split(New Char() {CType("(", Char)})
    ''            titelTeile(0) = tmpStr(0).Trim

    ''            If prcTyp = ptElementTypen.roles Then
    ''                If auswahl = 1 Then

    ''                    Dim anfText As String = repmsg(4)
    ''                    If rcName <> "" Then
    ''                        anfText = anfText & " " & rcName
    ''                    End If

    ''                    If Not IsNothing(vglProj) Then
    ''                        'titelTeile(0) = repMessages.getmsg(159) & " (" & gesamt_summe.ToString("####0.") & " / " & vSum.ToString("####0.") & " " & zE & ")"
    ''                        'titelTeile(0) = repmsg(4) & " (" & vSum.ToString("####0.") & " / " & gesamt_summe.ToString("####0.") & " " & zE & ")"
    ''                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " " & zE & ")"
    ''                    Else
    ''                        'titelTeile(0) = repMessages.getmsg(159) & " (" & gesamt_summe.ToString("####0.") & " " & zE & ")"
    ''                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " " & zE & ")"
    ''                    End If
    ''                    titelTeile(1) = ""

    ''                ElseIf auswahl = 2 Then

    ''                    Dim anfText As String = repmsg(0)
    ''                    If rcName <> "" Then
    ''                        anfText = anfText & " " & rcName
    ''                    End If

    ''                    If Not IsNothing(vglProj) Then
    ''                        'titelTeile(0) = repMessages.getmsg(160) & " (" & gesamt_summe.ToString("####0.") & " / " & vSum.ToString("####0.") & " T€" & ")"
    ''                        'titelTeile(0) = repmsg(0) & " (" & vSum.ToString("####0.") & " / " & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    ''                    Else
    ''                        'titelTeile(0) = repMessages.getmsg(160) & " (" & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = anfText & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    ''                    End If
    ''                    titelTeile(1) = ""
    ''                Else
    ''                    titelTeile(0) = "--- (T€)"
    ''                    titelTeile(1) = ""
    ''                End If
    ''            Else
    ''                ' jetzt muss das aus Kosten übernommen werden 
    ''                'titelTeile(0) = repMessages.getmsg(165) & " (" & gesamt_Summe.ToString("####0.") & " T€" & ")"
    ''                If auswahl = 1 Then
    ''                    If Not IsNothing(vglProj) Then
    ''                        'titelTeile(0) = repMessages.getmsg(165) & " (" & gesamt_summe.ToString("####0.") & " / " & vSum.ToString("####0.") & " T€" & ")"
    ''                        'titelTeile(0) = repmsg(2) & " (" & vSum.ToString("####0.") & " / " & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = repmsg(2) & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    ''                    Else
    ''                        'titelTeile(0) = repMessages.getmsg(165) & " (" & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = repmsg(2) & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    ''                    End If
    ''                ElseIf auswahl = 2 Then
    ''                    If Not IsNothing(vglProj) Then
    ''                        'titelTeile(0) = repMessages.getmsg(166) & " (" & gesamt_summe.ToString("####0.") & " / " & vSum.ToString("####0.") & " T€" & ")"
    ''                        'titelTeile(0) = repmsg(5) & " (" & vSum.ToString("####0.") & " / " & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = repmsg(5) & " (" & gesamt_summe.ToString("##,##0.") & " / " & vSum.ToString("##,##0.") & " T€" & ")"
    ''                    Else
    ''                        'titelTeile(0) = repMessages.getmsg(166) & " (" & gesamt_summe.ToString("####0.") & " T€" & ")"
    ''                        titelTeile(0) = repmsg(5) & " (" & gesamt_summe.ToString("##,##0.") & " T€" & ")"
    ''                    End If
    ''                Else
    ''                    titelTeile(0) = "--- (T€)" & vbLf & pname
    ''                End If

    ''                titelTeile(1) = ""
    ''            End If

    ''            titelTeilLaengen(1) = titelTeile(1).Length
    ''            titelTeilLaengen(0) = titelTeile(0).Length
    ''            diagramTitle = titelTeile(0) & titelTeile(1)

    ''            .ChartTitle.Text = diagramTitle
    ''        End If


    ''    End With



    ''    'pptShape.Chart.ChartData.ActivateChartDataWindow()
    ''    With pptShape.Chart.ChartData
    ''        .Activate()
    ''        CType(.Workbook, Excel.Workbook).Application.Visible = False
    ''    End With


    ''    'pptShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront)


    ''End Sub

    '''' <summary>    ''' 
    '''' im AddIn-Shutdown beendet ...
    '''' danach ist xlAPP gesetzt und es gibt das updateWorkbook 
    '''' wenn das HiddenExcel bereits existiert wird nichts gemacht ... 
    '''' </summary>
    '''' <remarks></remarks>
    'Friend Sub createNewHiddenExcel()

    '    Try
    '        If IsNothing(xlApp) Then
    '            xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
    '            xlApp.Visible = False
    '            xlApp.ScreenUpdating = False
    '        End If

    '    Catch ex As Exception
    '        xlApp = Nothing
    '        updateWorkbook = Nothing
    '        Exit Sub
    '    End Try

    '    'Dim updWS As Excel.Worksheet = Nothing
    '    'Dim creationNeeded As Boolean = False

    '    '' wenn 
    '    'Try
    '    '    If Not IsNothing(xlApp) Then
    '    '        ' lediglich ein Test auf Zugreifbarkeit ... fällt auf die Nase, wenn User das Excel geschlossen hat 
    '    '        Dim testCount As Integer = CType(xlApp.Workbooks, Excel.Workbooks).Count
    '    '    End If

    '    'Catch ex As Exception
    '    '    ' in diesem Fall wurde das Excel gelöscht ... 
    '    '    xlApp = Nothing
    '    'End Try

    '    'If Not IsNothing(xlApp) Then

    '    '    If IsNothing(updateWorkbook) Then
    '    '        If xlApp.Workbooks.Count > 0 Then
    '    '            ' fertig  - es gibt bereits ein Workbook 
    '    '            updateWorkbook = xlApp.Workbooks.Item(1)
    '    '        Else
    '    '            updateWorkbook = xlApp.Workbooks.Add()
    '    '        End If
    '    '    Else
    '    '        ' andernfalsl gibt es das ja schon 
    '    '    End If

    '    'Else

    '    '    Try

    '    '        xlApp = CType(CreateObject("Excel.Application"), Excel.Application)
    '    '        xlApp.Visible = False

    '    '        updateWorkbook = xlApp.Workbooks.Add()

    '    '    Catch ex As Exception
    '    '        xlApp = Nothing
    '    '        updateWorkbook = Nothing
    '    '        Exit Sub
    '    '    End Try

    '    'End If

    'End Sub
    ''' <summary>
    ''' ändert den Kommentar Ampel-Text, Lieferumfang
    ''' je nachdem, ob es sich um eine Ampel-Erläuterung oder einen Lieferumfang handelt ...  
    ''' </summary>
    ''' <param name="cmtShape"></param>
    ''' <param name="timestamp"></param>
    ''' <remarks></remarks>
    Friend Sub modifyComment(ByRef cmtShape As PowerPoint.Shape, ByVal timestamp As Date, ByVal showOtherVariant As Boolean)

        Dim newCmtText As String = ""
        Dim newCmtColor As Integer = 0
        Dim cmtType As Integer
        Dim tmpText As String = ""

        If IsNothing(cmtShape) Then
            Exit Sub
        End If


        Try
            ' jetzt kann die eigentliche Behandlung losgehen 
            ' aber nur, wenn es sich um Ampel-Text oder Lieferumfang Shape handelt ...
            cmtType = GetCmtTypeFromShapeName(cmtShape.Name)

            If cmtType = pptAnnotationType.ampelText Or
                cmtType = pptAnnotationType.lieferumfang Then

                If Not IsNothing(timestamp) Then
                    ' der Text und die Farbe müssen von einem TimeStamp Projekt kommen 

                    ' überprüfe, ob es zu dem angegebenen Shape bereits ein TS Projekt gibt 
                    Dim pvName As String = getPVnameFromShpName(cmtShape.Name)
                    Dim vpid As String = getVPIDFromTags(cmtShape)

                    ' damit auch eine andere Variante gezeigt werden kann ... 
                    Dim tmpPName As String = getPnameFromKey(pvName)
                    If showOtherVariant Then
                        pvName = calcProjektKey(tmpPName, currentVariantname)
                    End If


                    If pvName <> "" Or vpid <> "" Then
                        ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                        'Dim tsProj As clsProjekt = smartSlideLists.getTSProject(pvName, timestamp)
                        Dim tsProj As clsProjekt = timeMachine.getProjectVersion(pvName, timestamp, vpid)

                        If Not IsNothing(tsProj) Then

                            Dim refName As String = cmtShape.Name.Substring(0, cmtShape.Name.Length - 1)
                            Dim refShape As PowerPoint.Shape = Nothing
                            Try
                                refShape = currentSlide.Shapes.Item(refName)
                            Catch ex As Exception
                                Try
                                    If showOtherVariant Then
                                        refName = calcPPTShapeNameOVariant(tmpPName, currentVariantname, refName)
                                        refShape = currentSlide.Shapes.Item(refName)
                                    Else
                                        Exit Sub
                                    End If
                                Catch ex1 As Exception
                                    Exit Sub
                                End Try
                            End Try

                            If Not IsNothing(refShape) Then
                                Dim elemName As String = refShape.Tags.Item("CN")
                                Dim elemBC As String = refShape.Tags.Item("BC")

                                If pptShapeIsMilestone(refShape) Then

                                    Dim ms As clsMeilenstein = tsProj.getMilestone(msName:=elemName, breadcrumb:=elemBC)
                                    If IsNothing(ms) Then
                                        cmtShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                                    Else
                                        If Not cmtShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                            cmtShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                                        End If

                                        If cmtType = pptAnnotationType.ampelText Then
                                            ' Text und Farbe bestimmen 
                                            If englishLanguage Then
                                                tmpText = ms.name & " traffic light text:" & vbLf
                                            Else
                                                tmpText = ms.name & " Ampel-Text:" & vbLf
                                            End If
                                            newCmtText = tmpText & ms.getBewertung(1).description
                                            newCmtColor = ms.getBewertung(1).colorIndex

                                        ElseIf cmtType = pptAnnotationType.lieferumfang Then
                                            ' Text und Farbe bestimmen 
                                            If englishLanguage Then
                                                tmpText = ms.name & " Deliverables:" & vbLf
                                            Else
                                                tmpText = ms.name & " Lieferumfänge:" & vbLf
                                            End If
                                            newCmtText = tmpText & ms.getAllDeliverables(vbLf)
                                            newCmtColor = ms.getBewertung(1).colorIndex
                                        End If

                                    End If


                                ElseIf pptShapeIsPhase(refShape) Then

                                    Dim ph As clsPhase = tsProj.getPhase(name:=elemName, breadcrumb:=elemBC)
                                    If IsNothing(ph) Then
                                        cmtShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                                    Else
                                        If Not cmtShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                            cmtShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                                        End If
                                        If cmtType = pptAnnotationType.ampelText Then
                                            ' Text und Farbe bestimmen 
                                            newCmtText = ph.getBewertung(1).description
                                            newCmtColor = ph.getBewertung(1).colorIndex
                                        ElseIf cmtType = pptAnnotationType.lieferumfang Then
                                            ' Text und Farbe bestimmen 
                                            ' bei Phasen gibt es keine Deliverables
                                            newCmtText = ""
                                            newCmtColor = ph.getBewertung(1).colorIndex
                                        End If
                                    End If

                                End If

                                ' jetzt muss das Shape noch entsprechen modifiziert werden ... 
                                With cmtShape

                                    ' Text 
                                    .TextFrame2.TextRange.Text = newCmtText
                                    ' Farbe
                                    If newCmtColor < 1 Or newCmtColor > 4 Then
                                        .Shadow.ForeColor.RGB = Excel.XlRgbColor.rgbGrey
                                    Else
                                        .Shadow.ForeColor.RGB = CInt(trafficLightColors(newCmtColor))
                                    End If
                                End With
                            End If


                        End If
                    End If

                End If

            End If
        Catch ex As Exception

        End Try



    End Sub

    ''' <summary>
    ''' wird aufgerufen von ShowVariant oder TimeStamp Formular
    ''' bewegt alle Shaes, die in der Variante bzw. im TimeStamp ein anderes Datum haben, auf das neue Datum
    '''   
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub moveAllShapes(Optional ByVal showOtherVariant As Boolean = False)

        Dim changeliste As New clsChangeListe

        Dim namesToBeRenamed As New Collection
        'Dim ix As Integer = 0

        ' alle Shapes zur Time-Stamp Position schicken ...
        ' in diffMvList wird gemerkt, um wieviel sich ein Shape verändert hat und ob überhaupt ...  
        Dim diffMvList As New SortedList(Of String, Double)
        Dim oldProgressValue = 0


        ' nimmt die Shape-Namen auf, um darüber dann die Schleife laufen zu lassen. 
        ' also kein in currentSlide.Shapes mehr !!

        Dim bigToDoList As New Collection
        ' Aufbauen der Liste 
        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            bigToDoList.Add(tmpShape.Name)
        Next

        ' ur: 03.07.2017: lösche alle Ampelfarben
        Call faerbeShapes(PTfarbe.none, False)
        Call faerbeShapes(PTfarbe.green, False)
        Call faerbeShapes(PTfarbe.yellow, False)
        Call faerbeShapes(PTfarbe.red, False)


        Dim toDoList As New Collection

        For Each tmpShpName As String In bigToDoList

            Try

                Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                If Not IsNothing(tmpShape) Then
                    'ix = ix + 1

                    If isRelevantMSPHShape(tmpShape) Then
                        ' es ist ein Meilenstein oder eine Phase

                        If showOtherVariant Then
                            ' wenn es eine Variante gibt, wird currentTimeStamp dort auf den entsprechenden Wert der Variante gelegt 
                            namesToBeRenamed.Add(tmpShape.Name)
                            Call sendToNewPosition(tmpShape.Name, Date.Now, diffMvList, showOtherVariant, changeliste)
                        Else
                            Call sendToNewPosition(tmpShape.Name, currentTimestamp, diffMvList, showOtherVariant, changeliste)
                        End If
                        Try
                            ' PropertiesPane (sofern sichtbar) mit selektiertem Shape aktualisieren
                            If propertiesPane.Visible Then

                                If Not IsNothing(selectedPlanShapes) Then
                                    For Each shp As PowerPoint.Shape In selectedPlanShapes
                                        If shp.Id = tmpShape.Id Then
                                            Call aktualisiereInfoPane(tmpShape)
                                        End If
                                    Next
                                End If

                            End If
                        Catch ex As Exception

                        End Try



                    ElseIf isCommentShape(tmpShape) Then

                        If showOtherVariant Then
                            namesToBeRenamed.Add(tmpShape.Name)
                            ' wenn es eine Variante gibt, wird currentTimeStamp dort auf den entsprechenden Wert der Variante gelegt 
                            Call modifyComment(tmpShape, Date.Now, showOtherVariant)
                        Else
                            Call modifyComment(tmpShape, currentTimestamp, showOtherVariant)
                        End If


                    ElseIf isOtherVisboComponent(tmpShape) Then

                        toDoList.Add(tmpShape.Name)
                        'Call updateVisboComponent(tmpShape, currentTimestamp, previousTimeStamp)

                    End If

                    'If CInt(10 * ix / anzahlShapesOnSlide) > oldProgressValue Then
                    '    oldProgressValue = CInt(10 * ix / anzahlShapesOnSlide)
                    '    ProgressBarNavigate.Value = oldProgressValue
                    'End If
                End If

            Catch ex As Exception

            End Try


        Next

        ' jetzt muss die todolist noch extra abgearbeitet werden , wenn Charts drin waren, dürfen die nicht in der oberen Schleife behandelt werden, weil 
        ' bei der chart Behandlung Charts gelöscht und kopiert werden 
        For Each tmpShpName As String In toDoList
            Try
                Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                If Not IsNothing(tmpShape) Then
                    If showOtherVariant Then
                        ' wenn es eine Variante gibt, wird currentTimeStamp dort auf den entsprechenden Wert der Variante gelegt 
                        Call updateVisboComponent(tmpShape, Date.Now, previousTimeStamp, True)
                    Else
                        Call updateVisboComponent(tmpShape, currentTimestamp, previousTimeStamp, False)
                    End If

                Else
                    Call MsgBox("Error in Update ...")
                End If
            Catch ex As Exception
                Call MsgBox("Error in Update ...")
            End Try

        Next

        ' Behandlung der Namens- und Datumsbeschriftungen 
        For Each tmpShpName As String In bigToDoList

            Try
                Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                If Not IsNothing(tmpShape) Then

                    If isAnnotationShape(tmpShape) Then
                        ' hier müssen alle Annotations entsprechend verschoben werden, wie ihr Meilenstein / Phase verschoben wurde 
                        If tmpShape.Name.Substring(tmpShape.Name.Length - 1, 1) = CStr(CInt(pptAnnotationType.text)) Then

                            namesToBeRenamed.Add(tmpShape.Name)
                            ' es handelt sich um den Text, also nur verschieben 
                            Dim refName As String = tmpShape.Name.Substring(0, tmpShape.Name.Length - 1)

                            If diffMvList.ContainsKey(refName) Then
                                Dim diff As Double = diffMvList.Item(refName)
                                With tmpShape
                                    .Left = CSng(.Left + diff)
                                End With
                            End If


                        ElseIf tmpShape.Name.Substring(tmpShape.Name.Length - 1, 1) = CStr(CInt(pptAnnotationType.datum)) Then

                            namesToBeRenamed.Add(tmpShape.Name)
                            ' es handelt sich um das Datum, also verschieben und Text ändern 
                            Dim refName As String = tmpShape.Name.Substring(0, tmpShape.Name.Length - 1)
                            Dim refShape As PowerPoint.Shape = currentSlide.Shapes.Item(refName)
                            Dim tmpShort As Boolean = (tmpShape.TextFrame2.TextRange.Text.Length < 8)
                            ' showWeek = true: datumsText muss mit KW sein
                            Dim showWeek As Boolean = (tmpShape.TextFrame2.TextRange.Text.Contains("w") Or tmpShape.TextFrame2.TextRange.Text.Contains("W"))
                            Dim descriptionText As String = bestimmeElemDateText(refShape, tmpShort)

                            If diffMvList.ContainsKey(refName) Then
                                Dim diff As Double = diffMvList.Item(refName)
                                With tmpShape
                                    .Left = CSng(.Left + diff)
                                    .TextFrame2.TextRange.Text = descriptionText
                                End With
                            End If
                            ' showWeek wieder zurücksetzen
                            showWeek = False
                        End If

                    End If
                End If

            Catch ex As Exception
                'Call MsgBox("Fehler : " & ex.Message)
            End Try

        Next

        ' und schließlich muss noch nachgesehen werden, ob es eine todayLine gibt 
        Try
            Dim todayLineShape As PowerPoint.Shape = currentSlide.Shapes.Item("todayLine")
            ' ur:2019-05-29: TryCatch vermeiden
            'Dim todayLineShape As PowerPoint.Shape
            'todayLineShape = Nothing
            'For i = 1 To currentSlide.Shapes.Count
            '    If currentSlide.Shapes.Item(i).Name = "todayLine" Then
            '        todayLineShape = currentSlide.Shapes.Item("todayLine")
            '        Exit For
            '    End If
            'Next
            If Not IsNothing(todayLineShape) Then
                Call sendTodayLinetoNewPosition(todayLineShape)
            End If
        Catch ex As Exception

        End Try

        ' jetzt müssen die Shape-Namen neu gesetzt werden, wenn es sich um eine Variante handelte 
        If showOtherVariant Then

            For Each tmpShpName As String In namesToBeRenamed

                Dim pvName As String = getPVnameFromShpName(tmpShpName)
                Dim tmpPName As String = getPnameFromKey(pvName)
                Try
                    Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                    If Not IsNothing(tmpShape) Then
                        Dim newShapeName As String = calcPPTShapeNameOVariant(tmpPName, currentVariantname, tmpShape.Name)
                        tmpShape.Name = newShapeName
                    End If
                Catch ex As Exception

                End Try

            Next
        End If

        Call buildSmartSlideLists()

        ' soll auf alle Fälle angezeigt werden ...
        'Call faerbeShapes(PTfarbe.none, showTrafficLights(PTfarbe.none))
        'Call faerbeShapes(PTfarbe.green, showTrafficLights(PTfarbe.green))
        'Call faerbeShapes(PTfarbe.yellow, showTrafficLights(PTfarbe.yellow))
        'Call faerbeShapes(PTfarbe.red, showTrafficLights(PTfarbe.red))

        ' die gelben und roten sollten auf alle Fälle gezeigt werden , die grünen und nicht-bewerteten nur, wenn entsprechend eingestellt 
        Call faerbeShapes(PTfarbe.none, showTrafficLights(PTfarbe.none))
        Call faerbeShapes(PTfarbe.green, showTrafficLights(PTfarbe.green))
        Call faerbeShapes(PTfarbe.yellow, True)
        Call faerbeShapes(PTfarbe.red, True)

        Dim presChgListe As SortedList(Of Integer, clsChangeListe)
        'Dim hwind As Integer = pptAPP.ActiveWindow.HWND
        Dim key As String = CType(currentSlide.Parent, PowerPoint.Presentation).Name

        If chgeLstListe.ContainsKey(key) Then
            presChgListe = chgeLstListe.Item(key)
        Else
            presChgListe = New SortedList(Of Integer, clsChangeListe)
        End If

        If presChgListe.ContainsKey(currentSlide.SlideID) Then
            presChgListe.Remove(currentSlide.SlideID)
            presChgListe.Add(currentSlide.SlideID, changeliste)
            'chgeLstListe(currentSlide.SlideID) = changeliste
        Else
            presChgListe.Add(currentSlide.SlideID, changeliste)
        End If

    End Sub

    ''' <summary>
    ''' aktualisiert das Shape mit den Daten aus dem entsprechenden TimeStamp Projekt; 
    ''' es wird eine Aktion mit Moved Information gemacht ... denn wenn was manuell verändert wurde, muss es jetzt wieder auf die DB Position gebracht werden, sonst macht es überhaupt keinen Sinn
    ''' wenn es aufgerufen wird mit ShowOtherVariant werden die Werte der anderen Variante gezeigt, sonst einfach der andere TimeStamp derselben Projekt-Variante 
    ''' </summary>
    ''' <param name="tmpShapeName"></param>
    ''' <remarks></remarks>
    Friend Sub sendToNewPosition(ByVal tmpShapeName As String,
                                 ByVal timestamp As Date,
                                 ByRef diffMvList As SortedList(Of String, Double),
                                 ByVal showOtherVariant As Boolean,
                                 ByRef changeliste As clsChangeListe)

        Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShapeName)

        If Not IsNothing(tmpShape) Then
            ' Voraussetzung: es handelt sich um ein relevantes Shapes, also einen Meilenstein, eine Phase, einen Swimlane- oder Segment Bezeichner ... eine Phase oder einen Meilenstein ... 

            Dim pvName As String = getPVnameFromShpName(tmpShape.Name)
            Dim vpid As String = getVPIDFromTags(tmpShape)

            ' damit auch eine andere Variante gezeigt werden kann ... 
            If showOtherVariant Then
                Dim tmpPName As String = getPnameFromKey(pvName)
                pvName = calcProjektKey(tmpPName, currentVariantname)
            End If


            If pvName <> "" Then
                ' wenn das Projekt noch nicht geladen wurde, wird es aus der DB geholt und angelegt  ... 
                'Dim tsProj As clsProjekt = smartSlideLists.getTSProject(pvName, timestamp)
                Dim tsProj As clsProjekt = timeMachine.getProjectVersion(pvName, timestamp, vpid)
                ' kann dann nothing werden, wenn es zu diesem Zeitpunkt noch nicht existiert hat
                If Not IsNothing(tsProj) Then
                    Dim elemName As String = tmpShape.Tags.Item("CN")
                    Dim elemBC As String = tmpShape.Tags.Item("BC")

                    If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                        ' es handelt sich um einen Swimlane Namen oder Segment Name: kein Verschieben , aber das Setzen der Tags ist notwendig  
                        '
                        Dim ph As clsPhase = tsProj.getPhase(name:=elemName, breadcrumb:=elemBC)
                        If IsNothing(ph) Then
                            tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                        Else

                            If Not tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                            End If

                            Dim bsn As String = tmpShape.Tags.Item("BSN")
                            Dim bln As String = tmpShape.Tags.Item("BLN")
                            ' jetzt müssen die Tags-Informationen des Meilensteines gesetzt werden 
                            Call addSmartPPTMsPhInfo(tmpShape, tsProj, elemBC, elemName, ph.shortName, ph.originalName, bsn, bln,
                                                      ph.getStartDate, ph.getEndDate, ph.ampelStatus, ph.ampelErlaeuterung,
                                                      ph.getAllDeliverables("#"), ph.verantwortlich, ph.percentDone, ph.DocURL)

                        End If


                    Else
                        ' es handelt sich um einen echten Meilenstein oder Phase 
                        If pptShapeIsMilestone(tmpShape) Then

                            ' hier wird in der SmartList für dieses Element der Eintrag gelöscht, dass es verschoben wurde .. 
                            Call resetMVInfo(tmpShape)

                            Dim ms As clsMeilenstein = tsProj.getMilestone(msName:=elemName, breadcrumb:=elemBC)
                            If IsNothing(ms) Then
                                ' wenn es diesen Meilenstein in der Varianten oder TimeStamp Version gar nicht gibt, wird er auf invisible gesetzt 
                                tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                            Else
                                ' falls der in einer anderen TimeStamp- / Varianten Versin existierte, wird er wieder auf visible gesetzt 
                                If Not tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                                End If

                                Dim mvDiff As Double = mvMilestoneToTimestampPosition(tmpShape, ms.getDate, showOtherVariant, changeliste)
                                If Not diffMvList.ContainsKey(tmpShape.Name) And mvDiff * mvDiff > 0.01 Then
                                    diffMvList.Add(tmpShape.Name, mvDiff)
                                End If
                                '

                                Dim bsn As String = tmpShape.Tags.Item("BSN")
                                Dim bln As String = tmpShape.Tags.Item("BLN")
                                ' jetzt müssen die Tags-Informationen des Meilensteines gesetzt werden 
                                Call addSmartPPTMsPhInfo(tmpShape, tsProj, elemBC, elemName, ms.shortName, ms.originalName, bsn, bln, Nothing,
                                                          ms.getDate, ms.getBewertung(1).colorIndex, ms.getBewertung(1).description,
                                                          ms.getAllDeliverables("#"), ms.verantwortlich, ms.percentDone, ms.DocURL)

                            End If



                        ElseIf pptShapeIsPhase(tmpShape) Then

                            Call resetMVInfo(tmpShape)

                            Dim ph As clsPhase = tsProj.getPhase(name:=elemName, breadcrumb:=elemBC)
                            If IsNothing(ph) Then
                                tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                            Else
                                If Not tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue Then
                                    tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                                End If

                                Dim mvDiff As Double = mvPhaseToTimestampPosition(tmpShape, ph.getStartDate, ph.getEndDate, showOtherVariant, changeliste)
                                If Not diffMvList.ContainsKey(tmpShape.Name) And mvDiff * mvDiff > 0.01 Then
                                    diffMvList.Add(tmpShape.Name, mvDiff)
                                End If
                                '

                                Dim bsn As String = tmpShape.Tags.Item("BSN")
                                Dim bln As String = tmpShape.Tags.Item("BLN")
                                ' jetzt müssen die Tags-Informationen des Meilensteines gesetzt werden 
                                Call addSmartPPTMsPhInfo(tmpShape, tsProj, elemBC, elemName, ph.shortName, ph.originalName, bsn, bln, ph.getStartDate,
                                                             ph.getEndDate, ph.ampelStatus, ph.ampelErlaeuterung,
                                                             ph.getAllDeliverables("#"), ph.verantwortlich, ph.percentDone, ph.DocURL)

                            End If

                        End If

                    End If
                Else
                    ' es hat zu diesem Zeitpunkt noch nicht existiert und muss unsichtbar gemacht werden 
                    tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse

                    Dim explanation As New clsChangeItem
                    Dim projVarName As String = getPVnameFromShpName(tmpShape.Name)
                    explanation.pName = getPnameFromKey(projVarName)
                    explanation.vName = getVariantnameFromKey(projVarName)
                    explanation.bestElemName = "nicht aktualisierbar"

                    changeliste.addToChangeList(projVarName, explanation)
                End If

            End If


        End If



    End Sub

    ''' <summary>
    ''' schreibt bzw. aktualisiert den Time-Stamp auf die Folie ... 
    ''' </summary>
    ''' <param name="currentTimestamp"></param>
    ''' <remarks></remarks>
    Public Sub showTSMessage(ByVal currentTimestamp As Date)

        Dim tsMsgBox As PowerPoint.Shape
        Dim left As Single = 75, top As Single = 7, width As Single = 70, height As Single = 20

        ' handelt es sich um 23:59 Uhr , dann soll nämlich ohne Time gezeigt werden ... 
        Dim showTimeAndDate = (DateDiff(DateInterval.Minute, currentTimestamp.Date.AddHours(23).AddMinutes(59), currentTimestamp) <> 0)

        ' ' gibt es eine todayLine? 
        Dim todayLineShape As PowerPoint.Shape = Nothing
        Try
            todayLineShape = importantShapes(ptImportantShapes.todayline)
        Catch ex As Exception
            todayLineShape = Nothing
        End Try

        Dim tsMsgDidAlreadyExist As Boolean = False

        'ur: 2019-05-29: Try Catch entfernt (hatte nicht funktioniert)
        Try
            tsMsgBox = currentSlide.Shapes.Item("TimeStampInfo")
        Catch ex As Exception
            tsMsgBox = Nothing
        End Try



        If IsNothing(tsMsgBox) And (Not IsNothing(importantShapes(ptImportantShapes.todayline)) Or IsNothing(importantShapes(ptImportantShapes.version))) Then
            ' erstellen ...
            'tsMsgBox = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
            '                          200, 5, 70, 20)
            tsMsgBox = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                      75, 5, 70, 20)
            With tsMsgBox

                If showTimeAndDate Then
                    .TextFrame2.TextRange.Text = currentTimestamp.ToString
                Else
                    .TextFrame2.TextRange.Text = currentTimestamp.Date.ToShortDateString
                End If

                .TextFrame2.TextRange.Font.Size = CSng(schriftGroesse + 6)
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CInt(trafficLightColors(3))
                .TextFrame2.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue
                .TextFrame2.MarginBottom = 0
                .TextFrame2.MarginLeft = 0
                .Rotation = 0
                .TextFrame2.MarginRight = 0
                .TextFrame2.MarginTop = 0
                .Name = "TimeStampInfo"
                .TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
                .Fill.ForeColor.RGB = RGB(255, 255, 255)

            End With
        Else
            If IsNothing(tsMsgBox) Then
                ' nichts tun ...
            Else
                tsMsgDidAlreadyExist = True
                With tsMsgBox
                    If englishLanguage Then
                        If showTimeAndDate Then
                            .TextFrame2.TextRange.Text = currentTimestamp.ToString
                        Else
                            .TextFrame2.TextRange.Text = currentTimestamp.Date.ToShortDateString
                        End If


                    Else
                        If showTimeAndDate Then
                            .TextFrame2.TextRange.Text = currentTimestamp.ToString
                        Else
                            .TextFrame2.TextRange.Text = currentTimestamp.Date.ToShortDateString
                        End If
                    End If


                End With
            End If

        End If

        ' jetzt muss positioniert werden
        If Not IsNothing(todayLineShape) Then

            ' am unteren Ende der todayLineShape zentriert positionieren 
            If Not IsNothing(tsMsgBox) Then
                With tsMsgBox
                    ' die Farbe angleichen

                    If Not tsMsgDidAlreadyExist Then
                        .Top = todayLineShape.Top + todayLineShape.Height
                        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = todayLineShape.Line.ForeColor.RGB
                    Else
                        ' andernfalls bleibt das auf der Höhe wo der User es hingeschoben  hat 
                    End If

                    .Left = CSng(todayLineShape.Left - 0.5 * tsMsgBox.Width)
                End With
            End If

        End If

    End Sub


    ''' <summary>
    ''' setzt die MV-Info bei dem Element zurück 
    ''' andernfalls behalten ein paar die manual movement Info , die anderen dagegen nicht 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <remarks></remarks>
    Friend Sub resetMVInfo(ByRef tmpShape As PowerPoint.Shape)

        With tmpShape
            If .Tags.Item("MVD").Length > 0 Then
                .Tags.Delete("MVD")
            End If
            If .Tags.Item("MVE").Length > 0 Then
                .Tags.Delete("MVE")
            End If
        End With

        With smartSlideLists
            Call .removeSMLmvInfo(tmpShape.Name)
        End With

    End Sub


    ''' <summary>
    ''' diese Methode verschiebt das Shadow-Shape; es erfolgt keinerlei Setzen von Tag-Information
    ''' Der Wert von Diff wird dafür verwendet, um den zugehörigen Datums- oder Annotation Text des Elements zu verschieben ...er orientiert sich deshalb immer am left des Elements 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="tsStartdate"></param>
    ''' <param name="tsEndDate"></param>
    ''' <remarks></remarks>
    Friend Function mvPhaseShadowToNewPosition(ByRef tmpShape As PowerPoint.Shape, ByVal tsStartdate As Date, ByVal tsEndDate As Date,
                                               ByVal showOtherVariant As Boolean) As Double

        Dim x1Pos As Double, x2Pos As Double
        Dim diffEnde As Double = 0.0
        Dim diffDuration As Double = 0.0
        Dim diff As Double = 0.0

        Dim oldLeft As Double, oldWidth As Double
        'Dim expla As String = "Version: " & timeStamp.ToShortDateString


        Dim oldStartdate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left)
        Dim oldEndDate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left + tmpShape.Width)

        If tsStartdate <> oldStartdate Or tsEndDate <> oldEndDate Then
            '
            ' es hat sich was geändert ... 
            diffEnde = DateDiff(DateInterval.Day, oldEndDate, tsEndDate)
            diffDuration = DateDiff(DateInterval.Day, tsStartdate, tsEndDate) + 1 - (DateDiff(DateInterval.Day, oldStartdate, oldEndDate) + 1)

            If previousTimeStamp > currentTimestamp Then
                diffEnde = -1 * diffEnde
                diffDuration = -1 * diffDuration
            End If

            'homeButtonRelevance = True
            Call slideCoordInfo.calculatePPTx1x2(tsStartdate, tsEndDate, x1Pos, x2Pos)

            With tmpShape
                oldLeft = .Left
                oldWidth = .Width

                .Left = CSng(x1Pos)
                .Width = CSng(x2Pos - x1Pos)

            End With

            ' diff dient dazu , um das ggf angezeigte Annotation-Feld Name / Datum zu verschieben 
            diff = tmpShape.Left + tmpShape.Width / 2 - (oldLeft + oldWidth / 2)

        Else
            ' nichts tun ...
        End If

        mvPhaseShadowToNewPosition = diff
    End Function

    ''' <summary>
    ''' diese Methode verschiebt nur das Shape; es erfolgt keinerlei Setzen von Tag-Information
    ''' auch eine HomeButtonRelevance besetht nicht mehr; das neue Home ist mit dem TimeStamp erreicht 
    ''' Der Wert von Diff wird dafür verwendet, um den zugehörigen Datums- oder Annotation Text des Elements zu verschieben ...er orientiert sich deshalb immer am left des Elements 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="tsStartdate"></param>
    ''' <param name="tsEndDate"></param>
    ''' <remarks></remarks>
    Friend Function mvPhaseToTimestampPosition(ByRef tmpShape As PowerPoint.Shape,
                                               ByVal tsStartdate As Date,
                                               ByVal tsEndDate As Date,
                                               ByVal showOtherVariant As Boolean,
                                               ByRef changeliste As clsChangeListe) As Double

        Dim x1Pos As Double, x2Pos As Double
        Dim diffEnde As Double = 0.0
        Dim diffDuration As Double = 0.0
        Dim diff As Double = 0.0

        Dim oldLeft As Double, oldWidth As Double
        'Dim expla As String = "Version: " & timeStamp.ToShortDateString

        ' wenn der Phasen start oder das Phasen-Ende vor bzw. hinter dem pptStart bzw. EndOfCalendar liegt ...
        If DateDiff(DateInterval.Day, tsStartdate, slideCoordInfo.PPTStartOFCalendar) > 0 Then
            tsStartdate = slideCoordInfo.PPTStartOFCalendar
        End If

        If DateDiff(DateInterval.Day, slideCoordInfo.PPTEndOFCalendar, tsEndDate) > 0 Then
            tsEndDate = slideCoordInfo.PPTEndOFCalendar
        End If

        Dim oldStartdate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left)
        Dim oldEndDate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left + tmpShape.Width)

        If tsStartdate <> oldStartdate Or tsEndDate <> oldEndDate Then
            '
            ' es hat sich was geändert ... 
            diffEnde = DateDiff(DateInterval.Day, oldEndDate, tsEndDate)
            diffDuration = DateDiff(DateInterval.Day, tsStartdate, tsEndDate) + 1 - (DateDiff(DateInterval.Day, oldStartdate, oldEndDate) + 1)

            If previousTimeStamp > currentTimestamp Then
                diffEnde = -1 * diffEnde
                diffDuration = -1 * diffDuration
            End If

            'homeButtonRelevance = True
            Call slideCoordInfo.calculatePPTx1x2(tsStartdate, tsEndDate, x1Pos, x2Pos)

            With tmpShape
                oldLeft = .Left
                oldWidth = .Width

                .Left = CSng(x1Pos)
                .Width = CSng(x2Pos - x1Pos)

                Dim expPvName As String = getPVnameFromShpName(tmpShape.Name)
                Dim newShapeName As String = tmpShape.Name
                If showOtherVariant Then
                    Dim pName As String = getPnameFromKey(expPvName)
                    Dim vName As String = currentVariantname
                    expPvName = calcProjektKey(pName, vName)
                    newShapeName = calcPPTShapeNameOVariant(pName, vName, tmpShape.Name)
                End If
                Dim expElemName As String = tmpShape.Tags.Item("CN")
                Dim oldValue As String = bestimmeChangeDateOfPh(oldStartdate, oldEndDate, False)
                Dim newValue As String = bestimmeChangeDateOfPh(tsStartdate, tsEndDate, False)

                Dim chgExplanation As clsChangeItem = buildChangeExplanation(expPvName, expElemName, oldValue, newValue, CInt(diffEnde))

                If showOtherVariant Then
                    Call changeliste.addToChangeList(newShapeName, chgExplanation)
                Else
                    Call changeliste.addToChangeList(tmpShape.Name, chgExplanation)
                End If


            End With

            ' diff dient dazu , um das ggf angezeigte Annotation-Feld Name / Datum zu verschieben 
            diff = tmpShape.Left + tmpShape.Width / 2 - (oldLeft + oldWidth / 2)

        Else

            ' Änderung tk 13.8.17 keine Markierung mehr, wird zu unübersichtlich ... 
            'With tmpShape.Glow
            '    .Radius = 0
            '    '.Color.RGB = .Color.RGB = PowerPoint.XlRgbColor.rgbWhite
            'End With
        End If


        mvPhaseToTimestampPosition = diff
    End Function

    ''' <summary>
    ''' diese Methode verschiebt nur das Shape; es erfolgt keinerlei Setzen von Tag-Information
    ''' auch eine HomeButtonRelevance besetht nicht mehr; das neue Home ist mit dem TimeStamp erreicht  
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="msDate"></param>
    ''' <param name="showOtherVariant"></param>
    ''' <remarks></remarks>
    Friend Function mvMilestoneToTimestampPosition(ByRef tmpShape As PowerPoint.Shape,
                                                   ByVal msDate As Date,
                                                   ByVal showOtherVariant As Boolean,
                                                   ByRef changeliste As clsChangeListe) As Double
        Dim x1Pos As Double, x2Pos As Double
        Dim diff As Double = 0.0
        Dim diffInDays As Integer


        Dim chgExplanation As clsChangeItem
        Dim oldLeft As Double = tmpShape.Left
        Dim oldDate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left + tmpShape.Width / 2).Date
        'Dim expla As String = "Version: " & timeStamp.ToShortDateString

        msDate = msDate.Date

        If msDate <> oldDate Then
            ' es hat sich was geändert ... 

            Try
                If hasKwInMs(tmpShape) Then
                    Call updateKwInMs(tmpShape, msDate, False)
                End If
            Catch ex As Exception

            End Try

            Call slideCoordInfo.calculatePPTx1x2(msDate, msDate, x1Pos, x2Pos)


            ' jetzt die Shape-Info 
            With tmpShape
                .Left = CSng(x1Pos - tmpShape.Width / 2)
                diff = .Left - oldLeft
                diffInDays = CInt(DateDiff(DateInterval.Day, oldDate, msDate))
                If previousTimeStamp > currentTimestamp Then
                    diffInDays = -1 * diffInDays
                End If
                ' jetzt wird ggf die smartlists.changeList aufgebaut ... 

                Dim expPvName As String = getPVnameFromShpName(tmpShape.Name)
                Dim newShapeName As String = tmpShape.Name
                If showOtherVariant Then
                    Dim pName As String = getPnameFromKey(expPvName)
                    Dim vName As String = currentVariantname
                    expPvName = calcProjektKey(pName, vName)
                    newShapeName = calcPPTShapeNameOVariant(pName, vName, tmpShape.Name)
                End If
                Dim expElemName As String = tmpShape.Tags.Item("CN")
                Dim oldValue As String = bestimmeChangeDateOfMs(oldDate, False)
                Dim newValue As String = bestimmeChangeDateOfMs(msDate, False)

                chgExplanation = buildChangeExplanation(expPvName, expElemName, oldValue, newValue, diffInDays)

                If showOtherVariant Then
                    Call changeliste.addToChangeList(newShapeName, chgExplanation)
                Else
                    Call changeliste.addToChangeList(tmpShape.Name, chgExplanation)
                End If


            End With
        Else
            ' Änderung tk 13.8.17: das Element soll unverändert bleiben ... 
            'With tmpShape.Glow
            '    .Radius = 0
            '    '.Color.RGB = PowerPoint.XlRgbColor.rgbWhite

            'End With
        End If

        mvMilestoneToTimestampPosition = diff

    End Function

    ''' <summary>
    ''' gibt zurück, on dieser Meilenstein die seinem Datum entsprechende KW als Text stehen hat ...
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <returns></returns>
    Friend Function hasKwInMs(ByVal tmpShape As PowerPoint.Shape) As Boolean

        Dim tmpResult As Boolean = False
        ' es muss festgelegt werden, ob es eine KW_in_milestone gibt 
        Try
            If (tmpShape.TextFrame2.HasText = Microsoft.Office.Core.MsoTriState.msoTrue) Then
                Dim refKW As Integer = calcKW(CDate(tmpShape.Tags.Item("ED")))
                Dim vglKWinMs As Integer = CInt(tmpShape.TextFrame.TextRange.Text)
                If refKW = vglKWinMs Then
                    tmpResult = True
                End If
            End If

        Catch ex As Exception

        End Try

        hasKwInMs = tmpResult
    End Function

    ''' <summary>
    ''' prüft, ob irgendeine VISBO Slide enthalten ist; kann auch frozen sein
    ''' </summary>
    ''' <param name="pres"></param>
    ''' <returns></returns>
    Friend Function presentationHasAnySmartSlides(ByVal pres As PowerPoint.Presentation) As Boolean

        Dim tmpResult As Boolean = False
        Try
            For Each sld As PowerPoint.Slide In pres.Slides
                If isVisboSlide(sld) Then
                    tmpResult = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            tmpResult = False
        End Try

        presentationHasAnySmartSlides = tmpResult

    End Function

    ''' <summary>
    ''' schreibt die dem Datum msDate entsprechende KW in den Meilenstein 
    ''' whiteFont gibt an , dass die Schrift weiss sein soll (wenn es für den ShadowMeilenstein gezeichnet werden soll 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="msDate"></param>
    ''' <param name="whiteFont"></param>
    Friend Sub updateKwInMs(ByRef tmpShape As PowerPoint.Shape, ByVal msDate As Date, ByVal whiteFont As Boolean)

        Dim newkw As String = calcKW(msDate).ToString("0#")
        tmpShape.TextFrame.TextRange.Text = newkw

        If whiteFont Then
            tmpShape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        End If

    End Sub

    ''' <summary>
    ''' diese Methode verschiebt das Shadow-Shape an seine Position; es erfolgt keinerlei Setzen von Tag-Information
    ''' vor allem erfolgt hier kein Eintrag in der Changeliste 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="msDate"></param>
    ''' <param name="showOtherVariant"></param>
    ''' <remarks></remarks>
    Friend Function mvMilestoneShadowToNewPosition(ByRef tmpShape As PowerPoint.Shape, ByVal msDate As Date, ByVal showOtherVariant As Boolean) As Double
        Dim x1Pos As Double, x2Pos As Double
        Dim diff As Double = 0.0
        Dim diffInDays As Integer

        Dim oldLeft As Double = tmpShape.Left
        Dim oldDate As Date = slideCoordInfo.calcXtoDate(tmpShape.Left + tmpShape.Width / 2).Date
        'Dim expla As String = "Version: " & timeStamp.ToShortDateString

        msDate = msDate.Date

        If msDate <> oldDate Then
            ' es hat sich was geändert ... 

            Call slideCoordInfo.calculatePPTx1x2(msDate, msDate, x1Pos, x2Pos)


            ' jetzt die Shape-Info 
            With tmpShape
                .Left = CSng(x1Pos - tmpShape.Width / 2)
                diff = .Left - oldLeft
                diffInDays = CInt(DateDiff(DateInterval.Day, oldDate, msDate))
                If previousTimeStamp > currentTimestamp Then
                    diffInDays = -1 * diffInDays
                End If
                ' jetzt wird ggf die smartlists.changeList aufgebaut ... 

            End With
        Else
            ' Änderung tk 13.8.17: das Element soll unverändert bleiben ... 
            'With tmpShape.Glow
            '    .Radius = 0
            '    '.Color.RGB = PowerPoint.XlRgbColor.rgbWhite

            'End With
        End If

        mvMilestoneShadowToNewPosition = diff

    End Function

    ''' <summary>
    ''' baut den Explanation String auf, der erklärt, was sich am Element geändert hat 
    ''' </summary>
    ''' <param name="expPvName"></param>
    ''' <param name="expElemName"></param>
    ''' <param name="oldValue"></param>
    ''' <param name="newValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function buildChangeExplanation(ByVal expPvName As String, ByVal expElemName As String,
                                           ByVal oldValue As String, ByVal newValue As String,
                                           ByVal diffInDays As Integer) As clsChangeItem

        Dim tmpChangeItem As New clsChangeItem

        With tmpChangeItem
            .pName = getPnameFromKey(expPvName)
            .vName = getVariantnameFromKey(expPvName)
            .bestElemName = expElemName
            .oldValue = oldValue
            .newValue = newValue
            .diffInDays = diffInDays
        End With

        buildChangeExplanation = tmpChangeItem

    End Function


    ''' <summary>
    ''' aktualisiert die Info Form mit den Feldern ElemName, ElemDate, BreadCrumb und aLuTv-Text 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="isMovedShape"></param>
    ''' <remarks></remarks>
    Friend Sub aktualisiereInfoFrm(ByVal tmpShape As PowerPoint.Shape, Optional ByVal isMovedShape As Boolean = False)

        If Not IsNothing(infoFrm) Then

            With infoFrm
                '.btnSendToHome.Enabled = homeButtonRelevance
                '.btnSentToChange.Enabled = changedButtonRelevance
            End With

            If Not IsNothing(tmpShape) Then

                If Not IsNothing(selectedPlanShapes) Then

                    If selectedPlanShapes.Count = 1 Then

                        With infoFrm

                            Call .setDTPicture(pptShapeIsMilestone(tmpShape))

                            .elemName.Text = bestimmeElemText(tmpShape, .showAbbrev.Checked, .showOrginalName.Checked, .uniqueNameRequired.Checked)
                            'If showBreadCrumbField Then
                            '    .fullBreadCrumb.Text = bestimmeElemBC(tmpShape)
                            'End If
                            .elemDate.Text = bestimmeElemDateText(tmpShape, False)

                            Dim rdbCode As Integer = calcRDB()

                            Dim tmpStr() As String
                            tmpStr = bestimmeElemALuTvText(tmpShape, rdbCode).Split(New Char() {CType(vbLf, Char), CType(vbCr, Char)})
                            '.aLuTvText.Lines = tmpStr

                            ' Änderungen bei Datum und Erläuterung erlauben 
                            If isMovedShape Then
                                .elemDate.Enabled = True
                                'If .rdbMV.Checked Then
                                '    .aLuTvText.ReadOnly = False
                                'Else
                                '    .aLuTvText.ReadOnly = True
                                'End If
                            Else
                                .elemDate.Enabled = False
                                '.aLuTvText.ReadOnly = True
                            End If


                        End With
                    ElseIf selectedPlanShapes.Count > 1 Then

                        Dim rdbCode As Integer = calcRDB()

                        With infoFrm

                            Call .setDTPicture(Nothing)

                            If .elemName.Text <> bestimmeElemText(tmpShape, .showAbbrev.Checked, .showOrginalName.Checked, .uniqueNameRequired.Checked) Then
                                .elemName.Text = " ... "
                            End If
                            If .elemDate.Text <> bestimmeElemDateText(tmpShape, False) Then
                                .elemDate.Text = " ... "
                            End If

                            '.aLuTvText.Text = " ... "

                            '.aLuTvText.ReadOnly = True
                            .elemDate.Enabled = False


                        End With

                    End If
                Else
                    ' Info Formular Inhalte zurücksetzen ... 
                    With infoFrm
                        .elemName.Text = ""
                        '.fullBreadCrumb.Text = ""
                        .elemDate.Text = ""
                        '.aLuTvText.Text = ""
                    End With

                End If

            Else
                ' es wurde eine Selektion aufgehoben ..
                ' erstmal nichts tun .. 
                ' Info Formular Inhalte zurücksetzen ... 
                With infoFrm
                    .elemName.Text = ""
                    '.fullBreadCrumb.Text = ""
                    .elemDate.Text = ""
                    '.aLuTvText.Text = ""
                End With

            End If

        End If
    End Sub

    ''' <summary>
    ''' gibt den Projekt-/Varianten-Namen in der Form pname#vname zurück 
    ''' bildet ihn aus dem Tag PNM und VNM
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    Friend Function getPVnameFromTags(ByVal curShape As PowerPoint.Shape) As String

        Dim tmpResult As String = ""
        Dim pname As String = curShape.Tags.Item("PNM")
        Dim vname As String = curShape.Tags.Item("VNM")

        If pname.Length > 0 Then
            tmpResult = calcProjektKey(pname, vname)
        End If

        getPVnameFromTags = tmpResult

    End Function

    ''' <summary>
    ''' gibt die vpid aus den Tags zurück 
    ''' 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    Friend Function getVPIDFromTags(ByVal curShape As PowerPoint.Shape) As String

        Dim tmpResult As String = ""

        If curShape.Tags.Item("VPID").Length > 0 Then
            tmpResult = curShape.Tags.Item("VPID")
        End If

        getVPIDFromTags = tmpResult

    End Function

    ''' <summary>
    ''' gibt die vpid und VariantNaame aus String zurück 
    ''' 
    ''' </summary>
    ''' <param name="vpidVN"></param>
    ''' <returns></returns>
    Friend Function getVPIDVNfromString(ByVal vpidVN As String) As clsvpidVN

        Const MongoDBIDLength As Integer = 24
        Dim VNlength As Integer = vpidVN.Length - MongoDBIDLength

        Dim vpid As String = LSet(vpidVN, MongoDBIDLength)
        Dim VN As String = RSet(vpidVN, vpidVN.Length - MongoDBIDLength)


        getVPIDVNfromString = New clsvpidVN(vpid, VN)

    End Function


    ''' <summary>
    ''' gibt den Projekt-/Varianten Namen zurück
    ''' ShapeName ist aufgebaut (pName#variantName)ElemID  
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function getPVnameFromShpName(ByVal shapeName As String) As String

        Dim tmpName As String = ""
        If shapeName.StartsWith("(") Then

            Dim ixEnde As Integer = 0

            If shapeName.Contains(")0§") Then
                ixEnde = shapeName.IndexOf(")0§")

            ElseIf shapeName.Contains(")1§") Then
                ixEnde = shapeName.IndexOf(")1§")

            Else
                ' kein gültiger VISBO Shape-Name, bleibt leer 
            End If

            If ixEnde > 1 And ixEnde < shapeName.Length - 2 Then
                tmpName = shapeName.Substring(1, ixEnde - 1)
            End If
        End If

        getPVnameFromShpName = tmpName

    End Function

    ''' <summary>
    ''' bringt zu dem gegebenen ShapeNamen den Namen des zugrundeliegenden Referenz-Shapes zurück
    ''' also zum Comment das zugehörige Shape , dass dann in Folge mit einem Marker markiert werden kann 
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Friend Sub markReferenceShape(ByVal shapeName As String)
        Dim tmpText As String = ""

        If shapeName.EndsWith(CStr(pptAnnotationType.ampelText)) Or
            shapeName.EndsWith(CStr(pptAnnotationType.lieferumfang)) Or
            shapeName.EndsWith(CStr(pptAnnotationType.movedExplanation)) Then
            Dim strLength As Integer = shapeName.Length
            If strLength > 1 Then
                tmpText = shapeName.Substring(0, strLength - 1)

                Try
                    Dim refShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpText)
                    Call createMarkerShapes(refShape)
                    If formIsShown Then
                        Call aktualisiereInfoFrm(refShape)
                    End If
                Catch ex As Exception

                End Try
            End If

        End If

    End Sub

    ''' <summary>
    ''' erzeugt für jedes Shape in der angegebenen ShapeRange ein Marker Shape 
    ''' </summary>
    ''' <param name="pptShapes"></param>
    ''' <remarks></remarks>
    Friend Sub createMarkerShapes(Optional ByVal pptShape As PowerPoint.Shape = Nothing,
                                  Optional ByVal pptShapes As PowerPoint.ShapeRange = Nothing)


        Dim tmpShapeRange As PowerPoint.ShapeRange

        If Not IsNothing(pptShapes) Then
            tmpShapeRange = pptShapes
            For Each refShape As PowerPoint.Shape In tmpShapeRange
                Call zeichneMarkerShape(refShape)
            Next

        ElseIf Not IsNothing(pptShape) Then
            Call zeichneMarkerShape(pptShape)

        Else
            Exit Sub
        End If

    End Sub

    ''' <summary>
    ''' zeichnet für das übergebene Shape ein MarkerShape 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <remarks></remarks>
    Friend Sub zeichneMarkerShape(ByVal tmpShape As PowerPoint.Shape)

        Dim newHeight As Single
        Dim newWidth As Single
        Dim newLeft As Single
        Dim newTop As Single

        Try
            If Not IsNothing(tmpShape) Then

                If Not markerShpNames.ContainsKey(tmpShape.Name) Then
                    ' dann gibt es noch keinen Marker für dieses Shape ...  
                    With tmpShape
                        newHeight = CSng(markerHeight)
                        newWidth = CSng(markerWidth)
                        newLeft = CSng(.Left + 0.5 * (tmpShape.Width - newWidth))
                        newTop = CSng(.Top - (newHeight + 2))
                    End With

                    Dim markerShape As PowerPoint.Shape =
                                currentSlide.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeDownArrow, newLeft, newTop, newWidth, newHeight)

                    With markerShape
                        '.Fill.ForeColor.RGB = PowerPoint.XlRgbColor.rgbCornflowerBlue
                        .Fill.ForeColor.RGB = visboFarbeBlau
                        .Fill.Transparency = 0.0
                        .Line.Weight = 3
                        .Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
                        .Line.ForeColor.RGB = visboFarbeBlau
                    End With


                    If Not markerShpNames.ContainsKey(tmpShape.Name) Then
                        markerShpNames.Add(tmpShape.Name, markerShape.Name)
                    Else
                        Try
                            Dim oldMarker As PowerPoint.Shape = currentSlide.Shapes(markerShpNames.Item(tmpShape.Name))
                            oldMarker.Delete()
                            markerShpNames.Remove(tmpShape.Name)
                            markerShpNames.Add(tmpShape.Name, markerShape.Name)
                        Catch ex As Exception

                        End Try

                    End If


                End If



            End If
        Catch ex As Exception

        End Try

    End Sub
    ''' <summary>
    ''' löscht das Marker Shape ( Laserpointer 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub deleteMarkerShapes(Optional ByVal exceptShpName As String = "")
        Try
            Dim exceptionKey As String = ""
            Dim exceptionValue As String = ""

            If markerShpNames.Count > 1 Or
                (markerShpNames.Count = 1 And exceptShpName.Length > 0 And
                Not markerShpNames.ContainsKey(exceptShpName)) Then

                Dim arrayOfShpNames() As String

                ' ist eine Ausnahme definiert ? 
                If exceptShpName.Length > 0 Then
                    If markerShpNames.ContainsKey(exceptShpName) Then
                        exceptionKey = exceptShpName
                        exceptionValue = markerShpNames.Item(exceptionKey)
                        markerShpNames.Remove(exceptionKey)
                    End If
                End If

                ReDim arrayOfShpNames(markerShpNames.Count - 1)

                markerShpNames.Values.CopyTo(arrayOfShpNames, 0)


                'Dim markerShape As PowerPoint.Shape = currentSlide.Shapes.Item(markerName)
                Dim markerShapes As PowerPoint.ShapeRange

                Try
                    markerShapes = currentSlide.Shapes.Range(arrayOfShpNames)
                    If Not IsNothing(markerShapes) Then
                        markerShapes.Delete()
                    End If
                Catch ex As Exception
                    ' es ist mindestens ein Shape-Name im Array, der nicht mehr existiert 
                    ' deshalb muss hier einfach eine Schleife gefahren werden 
                    For ti As Integer = 0 To arrayOfShpNames.Length - 1

                        Try
                            Dim tshp As PowerPoint.Shape = currentSlide.Shapes.Item(arrayOfShpNames(ti))
                            If Not IsNothing(tshp) Then
                                tshp.Delete()
                            End If
                        Catch ex1 As Exception

                        End Try

                    Next

                End Try

                ' die Liste komplett bzw. bis auf die Ausnahme löschen
                markerShpNames.Clear()
                If exceptionKey.Length > 0 Then
                    markerShpNames.Add(exceptionKey, exceptionValue)
                End If

            ElseIf markerShpNames.Count = 1 And exceptShpName.Length = 0 Then
                Dim tmpName As String = markerShpNames.First.Value
                Dim markerShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpName)
                markerShpNames.Clear()
                markerShape.Delete()
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' gibt die ElemID eines Elements zurück 
    ''' 
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function getElemIDFromShpName(ByVal shapeName As String) As String

        Dim tmpName As String = ""
        If shapeName.StartsWith("(") Then

            Dim ixEnde As Integer = 0

            If shapeName.Contains(")0§") Then
                ixEnde = shapeName.IndexOf(")0§")

            ElseIf shapeName.Contains(")1§") Then
                ixEnde = shapeName.IndexOf(")1§")

            Else
                ' kein gültiger VISBO Shape-Name, bleibt leer 
            End If

            If ixEnde > 1 And ixEnde < shapeName.Length - 2 Then
                tmpName = shapeName.Substring(ixEnde + 1)
            End If
        End If

        getElemIDFromShpName = tmpName

    End Function


    ''' <summary>
    ''' gibt den Typ des Comments zurück 1: Ampel, 2: Lieferumfänge, 3: Terminverschiebungen
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function GetCmtTypeFromShapeName(ByVal shapeName As String) As Integer
        Try
            GetCmtTypeFromShapeName = CInt(shapeName.Substring(shapeName.Length - 1, 1))
        Catch ex As Exception
            GetCmtTypeFromShapeName = -1
        End Try
    End Function

    ''' <summary>
    ''' gibt den Elem-Namen zurück 
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function getElemNameFromShpName(ByVal shapeName As String) As String

        Dim tmpName As String = ""
        Dim elemName As String = ""
        If shapeName.StartsWith("(") Then

            Dim ixEnde As Integer = 0

            If shapeName.Contains(")0§") Then
                ixEnde = shapeName.IndexOf(")0§")

            ElseIf shapeName.Contains(")1§") Then
                ixEnde = shapeName.IndexOf(")1§")

            Else
                ' kein gültiger VISBO Shape-Name, bleibt leer 
            End If

            If ixEnde > 1 And ixEnde < shapeName.Length - 2 Then
                tmpName = shapeName.Substring(ixEnde + 3)
            End If
        End If

        ' jetzt Elem-Name bestimmen 
        If tmpName.Contains("§") Then
            elemName = tmpName.Substring(0, tmpName.IndexOf("§"))
        Else
            elemName = tmpName
        End If

        getElemNameFromShpName = elemName

    End Function

    ''' <summary>
    ''' entscheidet, ob es sich um einen Meilenstein handelt
    ''' Kriterium ist: Anzahl Tags > 0 und Startdate = Nothing, Enddate nicht gleich Nothing
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function pptShapeIsMilestone(ByVal curShape As PowerPoint.Shape) As Boolean


        If curShape.Tags.Count > 0 Then
            Dim anfang As String = curShape.Tags.Item("SD")
            Dim ende As String = curShape.Tags.Item("ED")


            If curShape.Tags.Item("SD").Length = 0 And curShape.Tags.Item("ED").Length > 0 Then
                ' ----------------------
                ' Test: 
                'If Not curShape.Name.Contains(")1§") Then
                '    Call MsgBox("Test-Fehler: Meilenstein?")
                'End If
                ' --------------------- Ende Test 

                pptShapeIsMilestone = True
            Else
                pptShapeIsMilestone = False
            End If
        Else
            pptShapeIsMilestone = False
        End If

    End Function

    ''' <summary>
    ''' entscheidet, ob es sich um einen Meilenstein handelt
    ''' Kriterium ist: Anzahl Tags > 0 und Startdate, EndDate ungleich Nothing
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function pptShapeIsPhase(ByVal curShape As PowerPoint.Shape) As Boolean


        If curShape.Tags.Count > 0 Then
            Dim anfang As String = curShape.Tags.Item("SD")
            Dim ende As String = curShape.Tags.Item("ED")


            If anfang.Length > 0 And ende.Length > 0 Then
                pptShapeIsPhase = True
            Else
                pptShapeIsPhase = False
            End If
        Else
            pptShapeIsPhase = False
        End If

    End Function

    ''' <summary>
    ''' gibt den Ampeltext / Lieferumfang / Terminveränderungs-Erläuterung des Shapes zurück 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemALuTvText(ByVal curShape As PowerPoint.Shape,
                                          Optional ByVal type As Integer = pptInfoType.aExpl,
                                          Optional ByVal shortForm As Boolean = True) As String

        Dim tmpText As String = ""
        ' tk, 22.11.17 wird später durch globale Variable ersetzt, die in den Settings interaktiv gesetzt werden kann 
        ' jetzt ist Voreinstellung einfach mal true ...
        Dim combined As Boolean = True

        If Not shortForm Then
            If pptShapeIsMilestone(curShape) Then
                tmpText = "(M): "
            ElseIf pptShapeIsPhase(curShape) Then
                tmpText = "(P): "
            End If
            tmpText = tmpText & curShape.Tags.Item("CN") & vbLf
        End If

        Try

            ' unterscheiden: einzeln ausweisen, oder in einem Text ? 
            If combined And Not (type = pptInfoType.resources Or type = pptInfoType.costs Or type = pptInfoType.mvElement) Then
                ' verantwortlich 
                If curShape.Tags.Item("VE").Length > 0 Then
                    If englishLanguage Then
                        tmpText = tmpText & "(Responsible): " & curShape.Tags.Item("VE") & vbLf
                    Else
                        tmpText = tmpText & "(Verantwortl.): " & curShape.Tags.Item("VE") & vbLf
                    End If
                End If

                ' Ampel-Text 
                If curShape.Tags.Item("AE").Length > 0 Then
                    If englishLanguage Then
                        tmpText = tmpText & "(Explanation):" & vbLf
                    Else
                        tmpText = tmpText & "(Erläuterung):" & vbLf
                    End If

                    tmpText = tmpText & curShape.Tags.Item("AE") & vbLf
                End If

                ' Deliverables 
                If curShape.Tags.Item("LU").Length > 0 Then
                    If englishLanguage Then
                        tmpText = tmpText & "(Deliverables):" & vbLf
                    Else
                        tmpText = tmpText & "(Lieferumfänge):" & vbLf
                    End If

                    Dim tmpStr() As String
                    tmpStr = curShape.Tags.Item("LU").Split(New Char() {CType("#", Char)})
                    For i As Integer = 0 To tmpStr.Length - 1
                        tmpText = tmpText & tmpStr(i) & vbLf
                    Next
                End If


            Else

                If type = pptInfoType.lUmfang Then
                    ' bestimmen der ersten Zeile:
                    If englishLanguage Then
                        tmpText = tmpText & "(Deliverables):" & vbLf
                    Else
                        tmpText = tmpText & "(Lieferumfänge):" & vbLf
                    End If


                    Dim tmpStr() As String
                    If curShape.Tags.Item("LU").Length > 0 Then
                        tmpStr = curShape.Tags.Item("LU").Split(New Char() {CType("#", Char)})
                        For i As Integer = 0 To tmpStr.Length - 1
                            tmpText = tmpText & tmpStr(i) & vbLf
                        Next
                    End If

                ElseIf type = pptInfoType.responsible Then
                    If englishLanguage Then
                        tmpText = tmpText & "(Responsible): "
                    Else
                        tmpText = tmpText & "(Verantwortlich): "
                    End If

                    If curShape.Tags.Item("VE").Length > 0 Then
                        tmpText = tmpText & curShape.Tags.Item("VE")
                    End If

                ElseIf type = pptInfoType.mvElement Then
                    If englishLanguage Then
                        tmpText = tmpText & "(moved):" & vbLf
                    Else
                        tmpText = tmpText & "(verschoben):" & vbLf
                    End If

                    If curShape.Tags.Item("MVE").Length > 0 Then
                        tmpText = tmpText & curShape.Tags.Item("MVE")
                    End If

                ElseIf type = pptInfoType.resources Or type = pptInfoType.costs Then
                    If Not noDBAccessInPPT And pptShapeIsPhase(curShape) Then
                        Try
                            Dim pvName As String = getPVnameFromShpName(curShape.Name)
                            Dim vpid As String = getVPIDFromTags(curShape)

                            'Dim hproj As clsProjekt = smartSlideLists.getTSProject(pvName, currentTimestamp)
                            Dim hproj As clsProjekt = timeMachine.getProjectVersion(pvName, currentTimestamp, vpid)
                            Dim phNameID As String = getElemIDFromShpName(curShape.Name)
                            Dim cPhase As clsPhase = hproj.getPhaseByID(phNameID)
                            Dim roleInformations As SortedList(Of String, Double) = cPhase.getRoleNamesAndValues
                            Dim costInformations As SortedList(Of String, Double) = cPhase.getCostNamesAndValues

                            If Not shortForm Then

                                If englishLanguage Then
                                    'tmpText = getElemNameFromShpName(curShape.Name) & " Resource/Costs :" & vbLf
                                    tmpText = tmpText & "(Resource/Costs):" & vbLf
                                Else
                                    'tmpText = getElemNameFromShpName(curShape.Name) & " Ressourcen/Kosten:" & vbLf
                                    tmpText = tmpText & "(Ressourcen/Kosten):" & vbLf
                                End If

                            Else
                                If englishLanguage Then
                                    tmpText = "(Resource/Costs):" & vbLf
                                Else
                                    tmpText = "(Ressourcen/Kosten):" & vbLf
                                End If
                            End If


                            Dim unit As String
                            If englishLanguage Then
                                unit = " PD"
                            Else
                                unit = " PT"
                            End If

                            For i As Integer = 1 To roleInformations.Count
                                tmpText = tmpText &
                                    roleInformations.ElementAt(i - 1).Key & ": " & CInt(roleInformations.ElementAt(i - 1).Value).ToString & unit & vbLf
                            Next

                            If costInformations.Count > 0 And roleInformations.Count > 0 Then
                                tmpText = tmpText & vbLf
                            End If

                            unit = " TE"
                            For i As Integer = 1 To costInformations.Count
                                tmpText = tmpText &
                                    costInformations.ElementAt(i - 1).Key & ": " & CInt(costInformations.ElementAt(i - 1).Value).ToString & unit & vbLf
                            Next

                        Catch ex As Exception
                            tmpText = "Phase " & getElemNameFromShpName(curShape.Name)
                        End Try



                    ElseIf noDBAccessInPPT And pptShapeIsPhase(curShape) Then
                        If Not shortForm Then
                            If englishLanguage Then
                                tmpText = "Resource/Costs " & getElemNameFromShpName(curShape.Name) & ":" & vbLf &
                                "no DB access ..."
                            Else
                                tmpText = "Ressourcen / Kosten " & getElemNameFromShpName(curShape.Name) & ":" & vbLf &
                                    "kein DB Zugriff ..."
                            End If

                        Else
                            If englishLanguage Then
                                tmpText = "no DB access"
                            Else
                                tmpText = "kein DB Zugriff"
                            End If

                        End If
                    Else
                        tmpText = ""
                    End If


                Else
                    ' in allen anderen Fällen den Ampel-Text wählen 
                    If curShape.Tags.Item("AE").Length > 0 Then
                        If englishLanguage Then
                            tmpText = tmpText & "(Explanation):" & vbLf
                        Else
                            tmpText = tmpText & "(Erläuterung):" & vbLf
                        End If

                        tmpText = tmpText & curShape.Tags.Item("AE")
                    End If
                End If

            End If




        Catch ex As Exception

        End Try

        bestimmeElemALuTvText = tmpText

    End Function

    ''' <summary>
    ''' betimmt die Beschriftung, den Namen des Symbols 
    ''' </summary>
    ''' <param name="curshape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeSymbolName(ByVal curshape As PowerPoint.Shape) As String
        Dim tmpText As String = ""

        With curshape

            If .Tags.Item("DID") = CStr(ptReportComponents.prSymTrafficLight) Then
                If englishLanguage Then
                    tmpText = "Explanation for project traffic Light"
                Else
                    tmpText = "Erläuterung zur Projekt-Ampel"
                End If

            ElseIf .Tags.Item("DID") = CStr(ptReportComponents.prSymRisks) Then
                If englishLanguage Then
                    tmpText = "Current Project Risks"
                Else
                    tmpText = "Aktuelle Projekt-Risiken"
                End If

            ElseIf .Tags.Item("DID") = CStr(ptReportComponents.prSymDescription) Then
                If englishLanguage Then
                    tmpText = "Project Goals"
                Else
                    tmpText = "Projekt-Ziele"
                End If

            ElseIf .Tags.Item("DID") = CStr(ptReportComponents.prSymFinance) Then
                If englishLanguage Then
                    tmpText = "Finance Overview"
                Else
                    tmpText = "Finanz-Überblick"
                End If

            ElseIf .Tags.Item("DID") = CStr(ptReportComponents.prSymProject) Then
                If englishLanguage Then
                    tmpText = "Project-Overview"
                Else
                    tmpText = "Projekt-Überblick"
                End If

            ElseIf .Tags.Item("DID") = CStr(ptReportComponents.prSymSchedules) Then
                If englishLanguage Then
                    tmpText = "Schedules-Overview"
                Else
                    tmpText = "Termin-Überblick"
                End If

            ElseIf .Tags.Item("DID") = CStr(ptReportComponents.prSymTeam) Then
                If englishLanguage Then
                    tmpText = "Team"
                Else
                    tmpText = "Team"
                End If
            End If

        End With

        bestimmeSymbolName = tmpText

    End Function

    ''' <summary>
    ''' bestimmt den Text , der dem Symbol zugeordnet ist 
    ''' </summary>
    ''' <param name="curshape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeSymbolText(ByVal curshape As PowerPoint.Shape) As String

        bestimmeSymbolText = curshape.Tags.Item("TXT")


    End Function

    ''' <summary>
    ''' bestimmt den Text in Abhängigkeit, ob classified name, ShortName oder OriginalName gezeigt werden soll 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <param name="showShortName"></param>
    ''' <param name="showOriginalName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemText(ByVal curShape As PowerPoint.Shape,
                                          ByVal showShortName As Boolean, ByVal showOriginalName As Boolean, ByVal useUniqueNames As Boolean) As String

        Dim tmpText As String = ""
        Dim translationNecessary As Boolean = False
        ' im Falle eines Namens , der öfter vorkommt und zu Zwecken der Eindeutigkeit durch den Bestname erstezt werden muss 
        Dim isCombinedName As Boolean = False
        Dim elemName As String = ""
        Dim bestShortName As String = curShape.Tags.Item("BSN")
        Dim bestLongName As String = curShape.Tags.Item("BLN")

        If isProjectCard(curShape) Then
            tmpText = curShape.Tags.Item("CN")

        ElseIf isRelevantMSPHShape(curShape) Then
            If showOriginalName Then
                If curShape.Tags.Item("ON").Length = 0 Then
                    tmpText = curShape.Tags.Item("CN")
                    translationNecessary = (selectedLanguage <> defaultSprache)
                Else
                    tmpText = curShape.Tags.Item("ON")
                End If

            ElseIf showShortName Then
                If curShape.Tags.Item("SN").Length = 0 Then
                    ' 26.3.21 wenn keine Abbrev definiert ist, soll nicht! der Lang-NAme verwendet werden ! 
                    tmpText = ""
                    'If curShape.Tags.Item("CN").Length > 0 Then
                    '    tmpText = curShape.Tags.Item("CN")
                    '    translationNecessary = (selectedLanguage <> defaultSprache)
                    'End If
                Else
                    tmpText = curShape.Tags.Item("SN")
                    If bestShortName.Length > 0 And tmpText <> bestShortName And useUniqueNames Then
                        tmpText = bestShortName
                    End If

                End If

            ElseIf curShape.Tags.Item("CN").Length > 0 Then
                tmpText = curShape.Tags.Item("CN")

                If bestLongName.Length > 0 And bestLongName <> tmpText And useUniqueNames Then
                    elemName = tmpText
                    tmpText = bestLongName
                    isCombinedName = True
                End If
                translationNecessary = (selectedLanguage <> defaultSprache)
            End If
        End If

        If translationNecessary Then
            ' jetzt den Text ersetzen 
            If isCombinedName Then
                tmpText = languages.translate(tmpText, selectedLanguage, elemName, isCombinedName)
            Else
                tmpText = languages.translate(tmpText, selectedLanguage)
            End If

        End If

        bestimmeElemText = tmpText
    End Function

    ''' <summary>
    ''' bestimmt den ChangeDate String für das Meilenstein-Element
    ''' </summary>
    ''' <param name="msDate"></param>
    ''' <param name="showShort"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeChangeDateOfMs(ByVal msDate As Date, ByVal showShort As Boolean) As String
        Dim tmpText As String = ""
        If Not showShort Then
            'tmpText = msDate.ToString("d")
            tmpText = msDate.ToShortDateString
        Else
            tmpText = msDate.Day.ToString & "." & msDate.Month.ToString
        End If
        bestimmeChangeDateOfMs = tmpText
    End Function

    ''' <summary>
    ''' bestimmt den ChangeDate String für das Phase-Element
    ''' </summary>
    ''' <param name="startDate"></param>
    ''' <param name="endDate"></param>
    ''' <param name="showShort"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeChangeDateOfPh(ByVal startDate As Date, ByVal endDate As Date, ByVal showShort As Boolean) As String
        Dim tmpText As String = ""

        If Not showShort Then
            tmpText = startDate.ToShortDateString & "-" & endDate.ToShortDateString
        Else
            Try

                tmpText = startDate.Day.ToString & "." & startDate.Month.ToString & "-" &
                            endDate.Day.ToString & "." & endDate.Month.ToString
            Catch ex As Exception
                tmpText = "? - ?"
            End Try

        End If
        bestimmeChangeDateOfPh = tmpText
    End Function

    ''' <summary>
    ''' bestimmt den Datums-String, für einen MEilenstein nur das Ende-Datum; 
    ''' 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemDateText(ByVal curShape As PowerPoint.Shape,
                                         ByVal showShort As Boolean,
                                         Optional considershowWeek As Boolean = True) As String

        Dim tmpText As String = ""

        If isProjectCard(curShape) Then

            tmpText = curShape.Tags.Item("SD") & "-" & curShape.Tags.Item("ED")

        ElseIf pptShapeIsMilestone(curShape) Then

            Dim msDate As Date = slideCoordInfo.calcXtoDate(curShape.Left + 0.5 * curShape.Width)
            If Not showShort Then
                tmpText = msDate.ToShortDateString
            Else
                tmpText = msDate.Day.ToString & "." & msDate.Month.ToString
            End If
            If considershowWeek And showWeek Then
                If englishLanguage Then
                    tmpText = "W " & calcKW(msDate).ToString
                Else
                    tmpText = "KW " & calcKW(msDate).ToString
                End If

            End If

        ElseIf pptShapeIsPhase(curShape) Then

            Dim startDate As Date = slideCoordInfo.calcXtoDate(curShape.Left)
            Dim endDate As Date = slideCoordInfo.calcXtoDate(curShape.Left + curShape.Width)

            If Not showShort Then
                tmpText = startDate.ToShortDateString & "-" & endDate.ToShortDateString
            Else
                Try

                    tmpText = startDate.Day.ToString & "." & startDate.Month.ToString & "-" &
                                endDate.Day.ToString & "." & endDate.Month.ToString
                Catch ex As Exception
                    tmpText = curShape.Tags.Item("SD") & "-" & curShape.Tags.Item("ED")
                End Try

            End If
            If considershowWeek And showWeek Then
                If englishLanguage Then
                    tmpText = "W " & calcKW(startDate).ToString & "-" & calcKW(endDate).ToString
                Else
                    tmpText = "KW " & calcKW(startDate).ToString & "-" & calcKW(endDate).ToString
                End If
            End If

        End If

        bestimmeElemDateText = tmpText
    End Function


    ''' <summary>
    ''' gibt an, ob ein Element manuell verändert wurde oder nicht 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isMovedElement(ByVal curShape As PowerPoint.Shape,
                                   Optional ByVal testWithMVPosition As Boolean = False) As Boolean

        Dim tmpResult As Boolean = False
        Dim tolerance As Integer = 0
        Dim tstDate As Date

        Try
            If pptShapeIsMilestone(curShape) Then

                Dim msDate As Date = slideCoordInfo.calcXtoDate(curShape.Left + 0.5 * curShape.Width)

                If testWithMVPosition Then
                    If curShape.Tags.Item("MVD").Length > 0 Then
                        Try
                            tstDate = CDate(curShape.Tags.Item("MVD"))
                        Catch ex As Exception
                            tstDate = msDate
                        End Try
                    Else
                        tstDate = msDate
                    End If
                Else
                    tstDate = CDate(curShape.Tags.Item("ED"))
                End If

                Dim diffDays As Integer = CInt(DateDiff(DateInterval.Day, msDate, tstDate))

                If diffDays <> 0 Then
                    tmpResult = True
                End If


            ElseIf pptShapeIsPhase(curShape) Then


                Dim pptSDate As Date = slideCoordInfo.calcXtoDate(curShape.Left)
                Dim pptEDate As Date = slideCoordInfo.calcXtoDate(curShape.Left + curShape.Width)
                Dim planSDate As Date = CDate(curShape.Tags.Item("SD"))
                Dim planEDate As Date = CDate(curShape.Tags.Item("ED"))


                If testWithMVPosition Then
                    Dim mvdString As String = curShape.Tags.Item("MVD")
                    If mvdString.Length > 0 Then

                        Try
                            Dim tmpStr() As String = mvdString.Split(New Char() {CType("#", Char)})
                            planSDate = CDate(tmpStr(0))
                            planEDate = CDate(tmpStr(1))
                        Catch ex As Exception
                            planSDate = CDate(curShape.Tags.Item("SD"))
                            planEDate = CDate(curShape.Tags.Item("ED"))
                        End Try

                    Else
                        planSDate = CDate(curShape.Tags.Item("SD"))
                        planEDate = CDate(curShape.Tags.Item("ED"))
                    End If


                Else
                    planSDate = CDate(curShape.Tags.Item("SD"))
                    planEDate = CDate(curShape.Tags.Item("ED"))
                End If


                ' prüfen, ob es beim Erzeugen abgeschnitten wurde ...
                Dim pptStartOfCalendar As Date = slideCoordInfo.PPTStartOFCalendar
                Dim pptEndOfCalendar As Date = slideCoordInfo.PPTEndOFCalendar

                If DateDiff(DateInterval.Day, pptStartOfCalendar, planSDate) < 0 Then
                    planSDate = pptStartOfCalendar
                End If

                If DateDiff(DateInterval.Day, pptEndOfCalendar, planEDate) > 0 Then
                    planEDate = pptEndOfCalendar
                End If

                Dim diffSD As Long = DateDiff(DateInterval.Day, pptSDate, planSDate)
                Dim diffED As Long = DateDiff(DateInterval.Day, pptEDate, planEDate)


                If diffSD <> 0 Or diffED <> 0 Then
                    tmpResult = True
                End If


            End If
        Catch ex As Exception

        End Try

        isMovedElement = tmpResult

    End Function
    ''' <summary>
    ''' gibt den Breadcrumb des Elements zurück 
    ''' </summary>
    ''' <param name="curshape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemBC(ByVal curshape As PowerPoint.Shape) As String

        Dim tmpText As String = ""

        If curshape.Tags.Item("BC").Length > 0 Then
            tmpText = curshape.Tags.Item("BC")
        End If

        bestimmeElemBC = tmpText

    End Function

    ''' <summary>
    ''' prüft, ob es sich um eine andere VISBO Komponente handelt ... (Chart, Tabelle, Platzhalter, ..) 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isOtherVisboComponent(ByVal curShape As PowerPoint.Shape) As Boolean

        Try
            isOtherVisboComponent = (curShape.Tags.Item("CHON").Length > 0) Or
                (curShape.Tags.Item("BID").Length > 0 And curShape.Tags.Item("DID").Length > 0)
        Catch ex As Exception
            isOtherVisboComponent = False
        End Try

    End Function

    ''' <summary>
    ''' true, wenn das Shape ein VISBO Meilenstein oder eine VISBO Phase ist 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isRelevantMSPHShape(ByVal curShape As PowerPoint.Shape) As Boolean

        Dim tmpResult As Boolean = False
        Dim tmpsTr As String = ""
        Dim pvName As String = getPVnameFromShpName(curShape.Name)
        If pvName <> "" Then
            tmpResult = isRelevantShape(curShape)
        End If

        isRelevantMSPHShape = tmpResult
    End Function



    ''' <summary>
    ''' true, wenn es sich um eine sogenannte Projekt-Karte handelt ...
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    Public Function isProjectCard(ByVal curShape As PowerPoint.Shape) As Boolean

        Dim tmpResult As Boolean = False
        Dim BID As String = curShape.Tags.Item("BID")
        Dim DID As String = curShape.Tags.Item("DID")

        If BID = CStr(ptReportBigTypes.planelements) And
            (DID = CStr(ptReportComponents.prCard) Or DID = CStr(ptReportComponents.prCardinvisible)) Then
            tmpResult = True
        End If

        isProjectCard = tmpResult
    End Function

    ''' <summary>
    ''' true, wenn es sich um eine Non-Prio Projektkarte handelt, die ja by default unsichtbar ist
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    Public Function isProjectCardInvisible(ByVal curShape As PowerPoint.Shape) As Boolean
        Dim tmpResult As Boolean = False
        Dim BID As String = curShape.Tags.Item("BID")
        Dim DID As String = curShape.Tags.Item("DID")

        If BID = CStr(ptReportBigTypes.planelements) And DID = CStr(ptReportComponents.prCardinvisible) Then
            tmpResult = True
        End If

        isProjectCardInvisible = tmpResult
    End Function

    ''' <summary>
    ''' 
    ''' true, wenn es einen Wert für Tag CN enthält
    ''' false , sonst
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isRelevantShape(ByVal curShape As PowerPoint.Shape) As Boolean

        Dim tmpStr As String = ""
        Try
            tmpStr = curShape.Tags.Item("CN")
        Catch ex As Exception

        End Try

        If tmpStr.Length > 0 Then
            isRelevantShape = True
        Else
            isRelevantShape = False
        End If

    End Function

    ''' <summary>
    ''' gibt true zurück, wenn es sich bei dem Shape um ein Symbol Shape handelt 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isSymbolShape(ByVal curShape As PowerPoint.Shape) As Boolean
        Dim bigType As String = curShape.Tags.Item("BID")
        Dim detailID As String = curShape.Tags.Item("DID")

        isSymbolShape = ((bigType = CStr(ptReportBigTypes.components)) And
                         (detailID = CStr(ptReportComponents.prSymDescription) Or
                          detailID = CStr(ptReportComponents.prSymRisks) Or
                          detailID = CStr(ptReportComponents.prSymTrafficLight) Or
                          detailID = CStr(ptReportComponents.prSymFinance) Or
                          detailID = CStr(ptReportComponents.prSymProject) Or
                          detailID = CStr(ptReportComponents.prSymSchedules) Or
                          detailID = CStr(ptReportComponents.prSymTeam)))

    End Function

    Public Sub sendTodayLinetoNewPosition(ByRef curShape As PowerPoint.Shape)

        Dim x1Pos As Double, x2Pos As Double

        With curShape

            Call slideCoordInfo.calculatePPTx1x2(currentTimestamp, currentTimestamp, x1Pos, x2Pos)

            ' Positionieren auf Home Position und aktualisieren des Info-Formulars..
            If .Left <> CSng(x1Pos) - .Width / 2 Then
                .Left = CSng(x1Pos) - .Width / 2
            End If

        End With

    End Sub

    ''' <summary>
    ''' gibt true zurück wenn es sich um ein Visbo Shape handelt, also entweder ein Plan-Element ist, ein Chart oder eine Komponente
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isVisboShape(ByVal curShape As PowerPoint.Shape) As Boolean
        If isRelevantMSPHShape(curShape) Or isCommentShape(curShape) Or isOtherVisboComponent(curShape) Or isAnnotationShape(curShape) Then
            isVisboShape = True
        Else
            isVisboShape = False
        End If
    End Function

    ''' <summary>
    ''' true, wenn es ein VISBO Chart, später dann auch ganz allgemein Reporting Element ist ..
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isVISBOChartElement(ByVal curShape As PowerPoint.Shape) As Boolean
        Dim tmpStr As String = ""
        Try
            tmpStr = curShape.Tags.Item("CHON")
        Catch ex As Exception

        End Try

        If tmpStr.Length > 0 Then
            isVISBOChartElement = True
        Else
            isVISBOChartElement = False
        End If
    End Function

    ''' <summary>
    ''' gibt zurück, ob es sich bei dem Shape um ein Comment-Shape handelt ... 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isCommentShape(ByVal curShape As PowerPoint.Shape) As Boolean
        Dim tmpResult As Boolean = False
        ' ggf noch ergänzen mit : curShape.Name.Contains("§")
        With curShape
            If curShape.Name.Contains("§") And .Tags.Item("CMT").Length > 0 Then
                tmpResult = True
            End If
        End With

        isCommentShape = tmpResult

    End Function

    ''' <summary>
    ''' determines whether or not the shape is a reporting component template
    ''' i.e contains certain keywords in .AlternativeText 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    Public Function isProjectReportingComponent(ByVal curShape As PowerPoint.Shape) As Boolean
        Dim tmpResult As Boolean = False
        Dim keyWords As String() = {"Swimlanes", "Swimlanes2", "AlleProjektElemente", "Einzelprojektsicht"}
        Dim searchtext = curShape.AlternativeText

        If IsNothing(searchtext) Then
            isProjectReportingComponent = False
        Else
            isProjectReportingComponent = keyWords.Contains(curShape.AlternativeText)
        End If

    End Function

    Public Function isMultiProjectReportingComponent(ByVal curShape As PowerPoint.Shape) As Boolean
        Dim tmpResult As Boolean = False
        Dim keyWords As String() = {"Multiprojektsicht"}
        Dim searchtext = curShape.AlternativeText

        If IsNothing(searchtext) Then
            isMultiProjectReportingComponent = False
        Else
            isMultiProjectReportingComponent = keyWords.Contains(curShape.AlternativeText)
        End If

    End Function

    ''' <summary>
    ''' liefert den Enumeration Typ des Comments zurück 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCommentType(ByVal curShape As PowerPoint.Shape) As Integer
        Dim tmpResult As Integer = -1

        With curShape
            Try
                If .Tags.Item("CMT").Length > 0 Then
                    If IsNumeric(.Tags.Item("CMT")) Then
                        tmpResult = CInt(.Tags.Item("CMT"))
                        If tmpResult < 0 Or tmpResult > 4 Then
                            ' ungültiger Wert
                            tmpResult = -1
                        End If

                    End If
                End If

            Catch ex As Exception

            End Try

        End With
        getCommentType = tmpResult

    End Function

    ''' <summary>
    ''' prüft, ob eine Slide geupdated werden sollte. 
    ''' ja, wenn gilt:  Slide enthält VISBO smart Elements And Not Frozen and LastUpdateDate ist ungleich  Heute 
    ''' </summary>
    ''' <param name="curSlide"></param>
    ''' <returns></returns>
    Public Function isSlideWithNeedToBeUpdated(ByVal curSlide As PowerPoint.Slide) As Boolean
        Dim tmpResult As Boolean = False

        With curSlide
            If .Tags.Item("SMART") = "visbo" Then
                If .Tags.Item("FROZEN").Length = 0 Then
                    If .Tags.Item("CRD").Length > 0 Then
                        Dim slideDate As Date = CDate(.Tags.Item("CRD"))
                        If DateDiff(DateInterval.Day, slideDate.Date, Date.Now.Date) <> 0 Then
                            tmpResult = True
                        End If
                    End If
                End If

            End If

        End With

        isSlideWithNeedToBeUpdated = tmpResult

    End Function

    ''' <summary>
    ''' prüft ob es sich um eine VISBO Slide handelt - dafür muss sie das Tag "SMART" enthalten 
    ''' Frozen nicht, da ja auch eine Frozen Slide interaktiv Auskunft geben können soll 
    ''' </summary>
    ''' <param name="sld"></param>
    ''' <returns></returns>
    Public Function isVisboSlide(ByVal sld As PowerPoint.Slide) As Boolean
        Dim tmpResult As Boolean = False

        Try
            With sld

                If .Tags.Item("SMART") = "visbo" Then
                    tmpResult = True
                End If

            End With
        Catch ex As Exception
            tmpResult = False
        End Try

        isVisboSlide = tmpResult

    End Function

    ''' <summary>
    ''' prüft, ob ein Shape ein Text oder Datums-Annotation-Shape ist 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isAnnotationShape(ByVal curShape As PowerPoint.Shape) As Boolean

        Dim criteria1 As Boolean = (curShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox)
        Dim criteria2 As Boolean = (curShape.Name.Contains(")1§") Or curShape.Name.Contains(")0§"))

        isAnnotationShape = criteria1 And criteria2
    End Function

    ''' <summary>
    ''' prüft, ob ein Shape für Schutz relevant ist oder nicht 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isRelevantForProtection(ByVal curShape As PowerPoint.Shape) As Boolean
        Dim criteria1 As String = ""
        'Dim criteria2 As Boolean


        isRelevantForProtection = isVisboShape(curShape) And Not isProjectCardInvisible(curShape)
        'Try
        '    criteria1 = curShape.Tags.Item("CN")
        'Catch ex As Exception

        'End Try

        'Try
        '    ' alle VISBO Beschriftungen oder Kommentare enthalten das im Namen ... 
        '    criteria2 = (curShape.Name.Contains(")1§") Or curShape.Name.Contains(")0§"))
        'Catch ex As Exception

        'End Try

        'If criteria1.Length > 0 Or criteria2 Then
        '    isRelevantForProtection = True
        'Else
        '    isRelevantForProtection = False
        'End If
    End Function

    ''' <summary>
    ''' löscht von einem Powerpoint Shape die entsprechenden Tags
    ''' das wird z.B dann benötigt, wenn auf einer Folie ein relevantes Shape kopiert wurde ... 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <remarks></remarks>
    Private Sub deleteShpTags(ByVal curShape As PowerPoint.Shape)

        If curShape.Tags.Item("CN").Length > 0 Then
            curShape.Tags.Delete("CN")
        End If

        If curShape.Tags.Item("ON").Length > 0 Then
            curShape.Tags.Delete("ON")
        End If

        If curShape.Tags.Item("SN").Length > 0 Then
            curShape.Tags.Delete("SN")
        End If

        If curShape.Tags.Item("BC").Length > 0 Then
            curShape.Tags.Delete("BC")
        End If

        If curShape.Tags.Item("SD").Length > 0 Then
            curShape.Tags.Delete("SD")
        End If

        If curShape.Tags.Item("ED").Length > 0 Then
            curShape.Tags.Delete("ED")
        End If

        If curShape.Tags.Item("AC").Length > 0 Then
            curShape.Tags.Delete("AC")
        End If

        If curShape.Tags.Item("AE").Length > 0 Then
            curShape.Tags.Delete("AE")
        End If

        curShape.Name = "copied_from" & curShape.Name


    End Sub

    ''' <summary>
    ''' fügt in der Powerpoint an das selektierte Plan-Element Lang-Name, Original-Name, Kurz-Name bzw Datum an 
    ''' wenn das Element bereits existiert, so wird es mit dem betreffenden Text beschriftet   
    ''' globale Variable, die im Zugriff sind: 
    ''' currentSlide: die aktuelle PPT-Slide
    ''' selectedplanShape: das aktuell selektierte Plan-Shape 
    ''' </summary>
    ''' <param name="descriptionType"></param>
    ''' <param name="positionIndex"></param>
    ''' <remarks></remarks>
    Public Sub annotatePlanShape(ByVal selectedPlanShape As PowerPoint.Shape,
                                  ByVal descriptionType As Integer, ByVal positionIndex As Integer)

        Dim newShape As PowerPoint.Shape
        Dim txtShpLeft As Single = selectedPlanShape.Left - 4
        Dim txtShpTop As Single = selectedPlanShape.Top - 5
        Dim txtShpWidth As Single = 5
        Dim txtShpHeight As Single = 5
        Dim normalFarbe As Integer = RGB(10, 10, 10)
        Dim ampelFarbe As Integer = 0

        Dim descriptionText As String = ""

        Dim shapeName As String = ""
        Dim ok As Boolean = False

        ' bestimme den Info Type ..
        ' handelt es sich um den Lang-/Kurz-Namen oder um das Datum ? 

        If descriptionType = pptAnnotationType.text Then
            descriptionText = bestimmeElemText(selectedPlanShape, showShortName, showOrigName, showBestName)

        ElseIf descriptionType = pptAnnotationType.datum Then
            descriptionText = bestimmeElemDateText(selectedPlanShape, showShortName)

        ElseIf descriptionType = pptAnnotationType.ampelText Or
                descriptionType = pptAnnotationType.lieferumfang Or
                descriptionType = pptAnnotationType.responsible Or
                descriptionType = pptAnnotationType.movedExplanation Then

            If IsNumeric(selectedPlanShape.Tags.Item("AC")) Then
                ampelFarbe = CInt(selectedPlanShape.Tags.Item("AC"))
            End If

            If descriptionType = pptAnnotationType.movedExplanation Then
                descriptionText = bestimmeElemALuTvText(selectedPlanShape, pptInfoType.mvElement, False)
                ampelFarbe = 4

            ElseIf descriptionType = pptAnnotationType.lieferumfang Then
                descriptionText = bestimmeElemALuTvText(selectedPlanShape, pptInfoType.lUmfang, False)

            ElseIf descriptionType = pptAnnotationType.responsible Then
                descriptionText = bestimmeElemALuTvText(selectedPlanShape, pptInfoType.responsible, False)

            Else
                descriptionText = bestimmeElemALuTvText(selectedPlanShape, pptInfoType.aExpl, False)
            End If

            txtShpLeft = CSng(selectedPlanShape.Left + 1.5 * selectedPlanShape.Width + 5)
            txtShpTop = selectedPlanShape.Top - 75
            txtShpWidth = 70
            txtShpHeight = 70

        ElseIf descriptionType = pptAnnotationType.resourceCost Then
            descriptionText = bestimmeElemALuTvText(selectedPlanShape, pptInfoType.resources, False)
        End If

        Try
            If Not IsNothing(descriptionType) Then
                If descriptionType >= 0 Then
                    shapeName = selectedPlanShape.Name & descriptionType.ToString
                    ok = True
                End If
            End If

        Catch ex As Exception
            ok = False
        End Try

        If Not ok Then
            Exit Sub
        End If

        Try
            newShape = currentSlide.Shapes(shapeName)
            If descriptionType = pptAnnotationType.ampelText Or
                    descriptionType = pptAnnotationType.movedExplanation Or
                    descriptionType = pptAnnotationType.lieferumfang Or
                    descriptionType = pptAnnotationType.resourceCost Then
                newShape.Delete()
                newShape = Nothing
            End If
        Catch ex As Exception
            newShape = Nothing
        End Try


        If IsNothing(newShape) Then

            If descriptionType = pptAnnotationType.ampelText Or
                    descriptionType = pptAnnotationType.movedExplanation Or
                    descriptionType = pptAnnotationType.lieferumfang Or
                    descriptionType = pptAnnotationType.resourceCost Then

                'newShape = currentSlide.Shapes.AddComment()
                newShape = currentSlide.Shapes.AddCallout(Microsoft.Office.Core.MsoCalloutType.msoCalloutOne,
                                      txtShpLeft, txtShpTop, txtShpWidth, txtShpHeight)
                With newShape
                    ' das Shape als Comment Shape kennzeichnen ... 
                    .Tags.Add("CMT", descriptionType.ToString)

                    .Fill.ForeColor.RGB = RGB(240, 240, 240)


                    .Shadow.Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow
                    .Shadow.Blur = 4
                    .Shadow.Size = 100
                    .Shadow.Transparency = 0.66
                    .Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    .Shadow.OffsetX = 2
                    .Shadow.OffsetY = 3.4641016151
                    .Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse

                    If ampelFarbe = 1 Then
                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbGreen
                    ElseIf ampelFarbe = 2 Then
                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbYellow
                    ElseIf ampelFarbe = 3 Then
                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbRed
                    ElseIf ampelFarbe = 4 Then
                        .Shadow.ForeColor.RGB = changeColor
                    Else
                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbGrey
                    End If

                    .TextFrame2.TextRange.Text = descriptionText
                    '.TextFrame2.TextRange.Font.Size = CDbl(schriftGroesse)
                    .TextFrame2.TextRange.Font.Size = 12
                    .TextFrame2.MarginBottom = 3
                    .TextFrame2.MarginLeft = 3
                    .TextFrame2.MarginRight = 3
                    .TextFrame2.MarginTop = 3
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = normalFarbe
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft
                    .Name = shapeName
                    .TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
                    .TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeShapeToFitText
                End With
            Else
                newShape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                      txtShpLeft, txtShpTop, 50, txtShpHeight)
                With newShape
                    .TextFrame2.TextRange.Text = descriptionText
                    .TextFrame2.TextRange.Font.Size = CSng(schriftGroesse)
                    .TextFrame2.MarginBottom = 0
                    .TextFrame2.MarginLeft = 0
                    .TextFrame2.MarginRight = 0
                    .TextFrame2.MarginTop = 0
                    .Name = shapeName
                    .TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
                End With

            End If

        Else
            With newShape
                .TextFrame2.TextRange.Text = descriptionText
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = normalFarbe
            End With
        End If


        ' jetzt wird das TextShape noch positioniert - in Abhängigkeit vom Position Index, 
        ' aber nur wenn es sich nicht um einen Comment handelt ...

        If ((Not descriptionType = pptAnnotationType.ampelText) And
             (Not descriptionType = pptAnnotationType.movedExplanation) And
             (Not descriptionType = pptAnnotationType.lieferumfang) And
             (Not descriptionType = pptAnnotationType.resourceCost)) Then

            Select Case positionIndex

                Case pptPositionType.center

                    If newShape.Width > 1.5 * selectedPlanShape.Width Then
                        ' keine Farbänderung 
                    Else
                        ' wenn die Beschriftung von der Ausdehnung kleiner als die Phase/der Meilenstein ist
                        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB =
                            selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End If
                    txtShpLeft = CSng(selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width))
                    txtShpTop = CSng(selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height))

                Case pptPositionType.aboveCenter

                    txtShpLeft = CSng(selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width))
                    txtShpTop = selectedPlanShape.Top - newShape.Height

                Case pptPositionType.aboveRight

                    If newShape.Width > selectedPlanShape.Width Then
                        txtShpLeft = selectedPlanShape.Left
                    Else
                        txtShpLeft = selectedPlanShape.Left + selectedPlanShape.Width - newShape.Width
                        'If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        '    newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        '    selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        'End If
                    End If

                    txtShpTop = selectedPlanShape.Top - newShape.Height

                Case pptPositionType.centerRight

                    txtShpLeft = selectedPlanShape.Left + selectedPlanShape.Width + 2
                    ' es wird jetzt rechts davon positioniert 
                    'If newShape.Width > selectedPlanShape.Width Then
                    '    txtShpLeft = selectedPlanShape.Left
                    'Else
                    '    txtShpLeft = selectedPlanShape.Left + selectedPlanShape.Width - newShape.Width
                    '    If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                    '        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                    '        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    '    End If

                    'End If

                    txtShpTop = CSng(selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height))

                Case pptPositionType.belowRight

                    If newShape.Width > selectedPlanShape.Width Then
                        txtShpLeft = selectedPlanShape.Left
                    Else
                        txtShpLeft = selectedPlanShape.Left + selectedPlanShape.Width - newShape.Width
                        'If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        '    newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        '    selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        'End If
                    End If

                    txtShpTop = selectedPlanShape.Top + selectedPlanShape.Height

                Case pptPositionType.belowCenter
                    txtShpLeft = CSng(selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width))
                    txtShpTop = selectedPlanShape.Top + selectedPlanShape.Height

                Case pptPositionType.belowLeft

                    If newShape.Width > selectedPlanShape.Width Then
                        txtShpLeft = selectedPlanShape.Left - (newShape.Width - selectedPlanShape.Width)
                    Else
                        txtShpLeft = selectedPlanShape.Left
                        'If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        '    newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        '    selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        'End If
                    End If

                    txtShpTop = selectedPlanShape.Top + selectedPlanShape.Height

                Case pptPositionType.centerLeft
                    txtShpLeft = selectedPlanShape.Left - (newShape.Width + 2)
                    'If newShape.Width > selectedPlanShape.Width Then
                    '    txtShpLeft = selectedPlanShape.Left - (newShape.Width - selectedPlanShape.Width)
                    'Else
                    '    txtShpLeft = selectedPlanShape.Left
                    '    If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                    '        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                    '        selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    '    End If
                    'End If
                    txtShpTop = CSng(selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height))

                Case pptPositionType.aboveLeft
                    If newShape.Width > selectedPlanShape.Width Then
                        txtShpLeft = selectedPlanShape.Left - (newShape.Width - selectedPlanShape.Width)
                    Else
                        txtShpLeft = selectedPlanShape.Left
                        'If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                        '    newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                        '    selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        'End If
                    End If
                    txtShpTop = selectedPlanShape.Top - newShape.Height

                Case pptPositionType.asis
                    txtShpLeft = newShape.Left
                    txtShpTop = newShape.Top

                Case Else
                    txtShpLeft = selectedPlanShape.Left - 5
                    txtShpTop = selectedPlanShape.Top - 10
            End Select

            ' jetzt die Position zuweisen

            With newShape
                .Top = txtShpTop
                .Left = txtShpLeft
            End With
        Else
            With newShape
                .Top = selectedPlanShape.Top - .Height - selectedPlanShape.Height / 2
                .Left = selectedPlanShape.Left
            End With
        End If






    End Sub

    ''' <summary>
    ''' wechselt die Sprache in der Annotation; tut dies für alle bereits dargestellten Beschriftungen 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub changeLanguageInAnnotations()


        ' andernfalls jetzt für alle Shapes ... 
        Dim bigToList As New Collection

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            bigToList.Add(tmpShape.Name)
        Next

        For Each tmpShpName As String In bigToList
            Try
                Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                If Not IsNothing(tmpShape) Then
                    If isRelevantMSPHShape(tmpShape) Then

                        ' hat es einen Text ? 
                        Dim searchName As String = tmpShape.Name & CInt(pptAnnotationType.text).ToString
                        Try
                            Dim txtShape As PowerPoint.Shape = currentSlide.Shapes(searchName)
                            If Not IsNothing(txtShape) Then
                                Dim curText As String = txtShape.TextFrame2.TextRange.Text
                                ' wenn der Text jetzt weder dem ShortName noch dem Original Name entspricht, dann soll er ersetzt werden ... 

                                Dim shortText As String = bestimmeElemText(tmpShape, True, False, showBestName)
                                Dim origText As String = bestimmeElemText(tmpShape, False, True, showBestName)

                                If ((curText <> shortText) And (curText <> origText)) Then
                                    ' dann ist es kein ShortName oder ein Original-Name , eine Unterscheidung in Meilenstein / Phase ist hier nicht notwendig, da asis gewählt wurde
                                    Call annotatePlanShape(tmpShape, pptAnnotationType.text, pptPositionType.asis)
                                End If
                            End If
                        Catch ex As Exception

                        End Try

                    End If
                End If
            Catch ex As Exception

            End Try
        Next




    End Sub
    ''' <summary>
    ''' Das Objekt vom Typ clsLanguages wird umgewandelt in einen String
    ''' über einen MemoryStream, der dann in String gewandelt wird
    ''' </summary>
    ''' <param name="obj">Objekt vom Typ clsLanguages</param>
    ''' <returns>XML String</returns>
    ''' <remarks></remarks>
    Public Function xml_serialize(ByVal obj As clsLanguages) As String

        Dim serializer As New DataContractSerializer(GetType(clsLanguages))
        Dim s As String

        ' --- Serialisieren in MemoryStream
        Dim ms As New MemoryStream()
        serializer.WriteObject(ms, obj)
        'Call MsgBox("Objekt wurde serialisiert!")

        ' --- Stream in String umwandeln
        Dim r As StreamReader = New StreamReader(ms)
        r.BaseStream.Seek(0, SeekOrigin.Begin)
        s = r.ReadToEnd

        Return s
    End Function
    ''' <summary>
    ''' Es wird ein String in die Struktur clsLanguages eingelesen
    ''' </summary>
    ''' <param name="langXMLstring">String in XML-Format</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function xml_deserialize(ByVal langXMLstring As String) As clsLanguages

        Dim languages As New clsLanguages

        ' --- Objekt  in Stream kopieren
        Dim ms As New MemoryStream()
        Dim w As StreamWriter = New StreamWriter(ms)
        w.BaseStream.Seek(0, SeekOrigin.Begin)
        w.WriteLine(langXMLstring)
        w.Close()

        '
        ' zu folgenden Befehlen: siehe Beschreibung unter Link 
        'https://books.google.de/books?id=zoBPnnGcASEC&pg=PA418&lpg=PA418&dq=xmlstring+erzeugen+mit+serializer&source=bl&ots=oMaIaszAh2&sig=l3E0WzuSsQ2IjvPIz50VahjJaNw&hl=de&sa=X&ved=0ahUKEwipiaa0yfXPAhVF7xQKHfHeDHEQ6AEIRjAG#v=onepage&q=xmlstring%20erzeugen%20mit%20serializer&f=false
        '
        ' --- MemoryStream umwandeln in Struktur clsLanguages
        Dim serializer As New DataContractSerializer(GetType(clsLanguages))
        ms = New MemoryStream(ms.ToArray)
        languages = CType(serializer.ReadObject(ms), clsLanguages)
        'Call MsgBox("Objekt wurde deserialisiert!")
        Return languages
    End Function

    ''' <summary>
    ''' macht die Visbo Shapes sichtbar bzw. unsichtbar .... 
    ''' </summary>
    ''' <param name="visible"></param>
    ''' <remarks></remarks>
    Public Sub makeVisboShapesVisible(ByVal visible As Microsoft.Office.Core.MsoTriState)

        For Each pptSlide As PowerPoint.Slide In pptAPP.ActivePresentation.Slides

            For Each pptShape As PowerPoint.Shape In pptSlide.Shapes
                If isRelevantForProtection(pptShape) Then
                    pptShape.Visible = visible
                End If
            Next

        Next

    End Sub


    ''' <summary>
    ''' sets the global parameters currentSldhasprojecttemplates, 
    ''' </summary>
    ''' <param name="sld"></param>
    ''' <returns></returns>
    Public Function slideHasReportComponents(ByVal sld As PowerPoint.Slide) As Boolean

        Dim found As Boolean = False
        currentSldHasProjectTemplates = False
        currentSldHasMultiProjectTemplates = False
        currentSldHasPortfolioTemplates = False

        Try
            If smartSlideLists.getElementNamen.Count > 0 Then
                ' nix weitermachen, dann sollen keine weiteren hier erstellt werden können 
            Else
                For Each pptShape As PowerPoint.Shape In sld.Shapes

                    Try
                        Dim visboKeyWord As String = ""
                        Dim tmpStr() As String = Nothing

                        If pptShape.Title <> "" Then
                            tmpStr = pptShape.Title.Split(New Char() {CChar("("), CChar(")")}, 3)
                        ElseIf pptShape.AlternativeText <> "" Then
                            tmpStr = pptShape.AlternativeText.Split(New Char() {CChar("("), CChar(")")}, 3)
                        ElseIf pptShape.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                            tmpStr = pptShape.TextFrame2.TextRange.Text.Split(New Char() {CChar("("), CChar(")")}, 3)
                        End If

                        If Not IsNothing(tmpStr) Then
                            visboKeyWord = tmpStr(0)
                            If Not currentSldHasProjectTemplates Then
                                currentSldHasProjectTemplates = projectComponentNames.Contains(visboKeyWord)
                            End If

                            If Not currentSldHasMultiProjectTemplates Then
                                currentSldHasMultiProjectTemplates = multiprojectComponentNames.Contains(visboKeyWord)
                            End If

                            If Not currentSldHasPortfolioTemplates Then
                                currentSldHasPortfolioTemplates = portfolioComponentNames.Contains(visboKeyWord)
                            End If

                        End If
                    Catch ex As Exception

                    End Try

                Next
            End If

        Catch ex As Exception
            found = False
        End Try

        slideHasReportComponents = currentSldHasProjectTemplates Or currentSldHasMultiProjectTemplates Or currentSldHasPortfolioTemplates
    End Function

    'Public Function getProjektHistory(ByVal pvName) As clsProjektHistorie

    '    Dim tmpResult As clsProjektHistorie = Nothing
    '    Dim pName As String
    '    Dim variantName As String = ""
    '    Dim pHistory As New clsProjektHistorie

    '    If IsNothing(pvName) Then
    '        ' nichts tun 
    '    ElseIf pvName.trim.length = 0 Then
    '        ' auch nichts tun ...
    '    Else

    '        Dim tmpstr() As String = pvName.Split(New Char() {CType("#", Char)})
    '        pName = tmpstr(0).Trim
    '        If tmpstr.Length > 1 Then
    '            variantName = tmpstr(1).Trim
    '        Else
    '            variantName = ""
    '        End If

    '        If Not noDBAccessInPPT Then

    '            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

    '            If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
    '                Try
    '                    pHistory.liste = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName,
    '                                                                    storedEarliest:=Date.MinValue, storedLatest:=Date.Now)
    '                Catch ex As Exception
    '                    pHistory = Nothing
    '                End Try
    '            Else
    '                If englishLanguage Then
    '                    Call MsgBox("database connection lost !")
    '                Else
    '                    Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
    '                End If

    '            End If




    '        End If


    '    End If

    '    getProjektHistory = tmpResult
    'End Function

    ''' <summary>
    ''' prüft, ob Home bzw Changed Button enabled werden muss 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub checkHomeChangeBtnEnablement()

        Dim atleastOneHomey As Boolean = False
        Dim atleastOneChanged As Boolean = False

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes

            If Not tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                If isRelevantMSPHShape(tmpShape) Then
                    If tmpShape.Tags.Item("MVD").Length > 0 Then
                        If isMovedElement(tmpShape) Then
                            atleastOneHomey = True
                        Else
                            atleastOneChanged = True
                        End If
                    End If
                End If
            End If

            If atleastOneChanged And atleastOneHomey Then
                Exit For
            End If
        Next

        homeButtonRelevance = atleastOneHomey
        changedButtonRelevance = atleastOneChanged

    End Sub

    ''' <summary>
    ''' gibt für ein existierendes Shape und ein entsprechendes Varianten-Projekt den neuen Shape-Namen zurück ... 
    ''' wird nur in SmartInfo benutzt, wenn die Shapes einer anderen Variante angezeigt werden sollen ...
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="newVariantName"></param>
    ''' <param name="oldShapeName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function calcPPTShapeNameOVariant(ByVal pName As String, ByVal newVariantName As String, ByVal oldShapeName As String) As String

        Dim tmpStr() As String = oldShapeName.Split(New Char() {CChar("("), CChar(")")}, 3)
        Dim tmpResult As String = oldShapeName

        Try
            If tmpStr.Length = 3 Then
                If tmpStr(2).Length > 0 Then
                    tmpResult = "(" & pName & "#" & newVariantName & ")" & tmpStr(2)
                End If
            End If
        Catch ex As Exception

        End Try

        calcPPTShapeNameOVariant = tmpResult

    End Function

    ''' <summary>
    ''' bringt wieder alle No-Prio Projekt-Karten ins No-Show 
    ''' </summary>
    Friend Sub putAllNoPrioShapesInNoshow()

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes

            If isProjectCardInvisible(tmpShape) And tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue Then
                tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
            End If

        Next

    End Sub
    ''' <summary>
    ''' wird aufgerufen, um die Elemente aus der ChangeListe (TimeMachine) hervorheben zu können, die sich verändert haben. 
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub dimAllShapesExceptThese(ByVal exceptionArray() As String)

        Dim oldValue As Single = 0.0

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes

            If isRelevantMSPHShape(tmpShape) Then
                If Not exceptionArray.Contains(tmpShape.Name) And Not tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                    ' Shape abdimmen , aber vorher den Wert merken .. 

                    If tmpShape.Tags.Item("DIMF").Length > 0 Then
                        tmpShape.Tags.Delete("DIMF")
                    End If

                    oldValue = tmpShape.Fill.Transparency
                    tmpShape.Tags.Add("DIMF", oldValue.ToString("#0.#"))

                    tmpShape.Glow.Radius = 0

                    tmpShape.Fill.Transparency = 0.8
                    tmpShape.Fill.Solid()

                    tmpShape.Line.Transparency = 0.8



                End If

            ElseIf isAnnotationShape(tmpShape) Then
                tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
            End If

        Next

    End Sub

    ''' <summary>
    ''' zeichnet die Shadows, das heisst die Elemente des letzten Timestamps zu dem angegebenen Shapes 
    ''' </summary>
    ''' <param name="nameArrayO"></param>
    ''' <remarks></remarks>
    Friend Sub zeichneShadows(ByVal nameArrayO() As String, ByVal isOtherVariant As Boolean)

        Dim anzElemente As Integer = nameArrayO.Length
        Dim vTextShadowShape As PowerPoint.Shape = Nothing
        Dim vTextOrigShape As PowerPoint.Shape = Nothing

        ' wird verwendet , um nur ein einziges Mal die Beschriftung der Versionen anzubringen, aber bei Mehrfach-Selektion kein zweites Mal 
        Dim versionAlreadyNotedAtMS As Boolean = False
        Dim versionAlreadyNotedAtPH As Boolean = False

        For i As Integer = 1 To anzElemente

            Dim mvDiff As Double = 0.0
            Dim elemID As String = getElemIDFromShpName(nameArrayO(i - 1))
            Dim pvName As String = getPVnameFromShpName(nameArrayO(i - 1))
            Dim origShape As PowerPoint.Shape = currentSlide.Shapes(nameArrayO(i - 1))
            Dim vpid As String = getVPIDFromTags(origShape)

            'Dim tsProj As clsProjekt = smartSlideLists.getTSProject(pvName:=pvName, tsDate:=previousTimeStamp)
            Dim tsProj As clsProjekt = timeMachine.getProjectVersion(pvName, previousTimeStamp, vpid)
            Dim isMilestone As Boolean = elemIDIstMeilenstein(elemID)

            If Not IsNothing(origShape) Then

                origShape.Copy()
                Dim newShape As PowerPoint.ShapeRange = currentSlide.Shapes.Paste()
                Dim shadowShape As PowerPoint.Shape = newShape(1)

                With shadowShape
                    ' das shadow soll nicht den Schatten aus dem Original Shape übernehmen 
                    .Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    .Name = origShape.Name & shadowName

                    If Not isMilestone Then
                        ' damit der Unterschied bei den Phasen besser erkennbar, d.h überlappungsfrei erkennbar ist ...
                        .Top = origShape.Top - (origShape.Height + 3)
                    Else
                        .Top = origShape.Top
                    End If

                    .Left = origShape.Left

                    ' das Shadow Shape soll immer als dunkles Schatten-Shape gezeichnet werden ...
                    .Fill.ForeColor.RGB = RGB(20, 20, 20)
                    .Line.ForeColor.RGB = RGB(20, 20, 20)


                    'If isMilestone Then
                    '    .Shadow.Type = Microsoft.Office.Core.MsoShadowType.msoShadow25
                    '    .Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    '    .Shadow.Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow
                    '    .Shadow.OffsetX = 0
                    '    .Shadow.OffsetY = 0
                    '    .Shadow.Blur = 15.0
                    '    .Shadow.Size = 180.0
                    '    .Shadow.ForeColor.RGB = RGB(220, 220, 220)
                    'End If

                End With

                If isMilestone Then
                    ' Meilenstein
                    Dim cMilestone As clsMeilenstein = tsProj.getMilestoneByID(elemID)

                    If IsNothing(cMilestone) Then
                        ' wenn es diesen Meilenstein in der Variante bzw. Timestamp nicht gibt wird das new'Shape wieder gelöscht ...  
                        newShape.Delete()
                    Else

                        Dim shadowDate As Date = cMilestone.getDate
                        If hasKwInMs(origShape) Then
                            Call updateKwInMs(shadowShape, shadowDate, True)
                        End If

                        ' jetzt bewegen 
                        mvDiff = mvMilestoneShadowToNewPosition(shadowShape, shadowDate, isOtherVariant)
                        Dim bsn As String = origShape.Tags.Item("BSN")
                        Dim bln As String = origShape.Tags.Item("BLN")
                        Dim elemName As String = origShape.Tags.Item("CN")
                        Dim elemBC As String = origShape.Tags.Item("BC")
                        ' jetzt müssen die Tags-Informationen des Meilensteines gesetzt werden 
                        Call addSmartPPTMsPhInfo(shadowShape, tsProj, elemBC, elemName, cMilestone.shortName, cMilestone.originalName, bsn, bln, Nothing,
                                                  cMilestone.getDate, cMilestone.getBewertung(1).colorIndex, cMilestone.getBewertung(1).description,
                                                  cMilestone.getAllDeliverables("#"), cMilestone.verantwortlich, cMilestone.percentDone, cMilestone.DocURL)

                        If Not versionAlreadyNotedAtMS Then
                            Call beschrifteOrigAndShadow(shadowShape.Name, origShape.Name, True)
                            versionAlreadyNotedAtMS = True
                        End If


                    End If
                Else
                    ' es handelt sich um eine Phase
                    ' wichtig: die Project Linie soll aber nicht betrachtet werden  
                    Dim ph As clsPhase = tsProj.getPhaseByID(elemID)

                    If IsNothing(ph) Then
                        ' wenn es diese Phase in der Variante bzw. Timestamp nicht gibt wird das new'Shape wieder gelöscht ...  
                        newShape.Delete()
                    Else
                        mvDiff = mvPhaseShadowToNewPosition(newShape(1), ph.getStartDate, ph.getEndDate, isOtherVariant)

                        Dim bsn As String = origShape.Tags.Item("BSN")
                        Dim bln As String = origShape.Tags.Item("BLN")
                        Dim elemName As String = origShape.Tags.Item("CN")
                        Dim elemBC As String = origShape.Tags.Item("BC")
                        ' jetzt müssen die Tags-Informationen der Phase gesetzt werden 
                        Call addSmartPPTMsPhInfo(shadowShape, tsProj, elemBC, elemName, ph.shortName, ph.originalName, bsn, bln,
                                                  ph.getStartDate, ph.getEndDate, ph.ampelStatus, ph.ampelErlaeuterung,
                                                  ph.getAllDeliverables("#"), ph.verantwortlich, ph.percentDone, ph.DocURL)

                    End If

                    If Not versionAlreadyNotedAtPH Then
                        Call beschrifteOrigAndShadow(shadowShape.Name, origShape.Name, False)
                        versionAlreadyNotedAtPH = True
                    End If

                End If


                ' jetzt wird entscheiden , ob eine Verbindungslinie gezeichnet wird 
                ' bei Phasen wird überhaupt keine Verbindungslinie gezeichnet , hier wird der Unterschied durch oben / unten klar 

                If isMilestone Then

                    If System.Math.Abs(mvDiff) > 1.5 * shadowShape.Width Then
                        Dim verbindungsShape As PowerPoint.Shape = Nothing

                        If previousTimeStamp < currentTimestamp Then
                            'If currentTimestamp > previousTimeStamp Then

                            If mvDiff < 0 Then
                                verbindungsShape = currentSlide.Shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
                                                                                                        shadowShape.Left + shadowShape.Width, shadowShape.Top + shadowShape.Height / 2,
                                                                                                        origShape.Left, origShape.Top + origShape.Height / 2)
                            Else
                                verbindungsShape = currentSlide.Shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
                                                                                                        shadowShape.Left, shadowShape.Top + shadowShape.Height / 2,
                                                                                                        origShape.Left + origShape.Width, origShape.Top + origShape.Height / 2)
                            End If

                        Else
                            ' currentTimeStamp < previoustimestamp
                            If mvDiff > 0 Then

                                verbindungsShape = currentSlide.Shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
                                                                                                        origShape.Left + origShape.Width, origShape.Top + origShape.Height / 2,
                                                                                                        shadowShape.Left, shadowShape.Top + shadowShape.Height / 2)


                            Else
                                verbindungsShape = currentSlide.Shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorStraight,
                                                                                                        origShape.Left, origShape.Top + origShape.Height / 2,
                                                                                                        shadowShape.Left + shadowShape.Width, shadowShape.Top + shadowShape.Height / 2)
                            End If

                        End If

                        With verbindungsShape

                            .Name = .Name & shadowName
                            .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadNone
                            .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle
                            .Line.Weight = 5.0

                        End With
                    End If

                Else
                    Dim verbindungsShape As PowerPoint.Shape = Nothing

                    ' tk, für Phasen soll keine Verbindunslinie gezeichnet werden  
                    'If previousTimeStamp < currentTimestamp Then

                    '    If shadowShape.Left + shadowShape.Width / 2 < origShape.Left Then

                    '        verbindungsShape = currentSlide.Shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorElbow, _
                    '                                                                                        shadowShape.Left + shadowShape.Width / 2, shadowShape.Top + shadowShape.Height, _
                    '                                                                                        origShape.Left, origShape.Top + shadowShape.Height / 2)
                    '    ElseIf shadowShape.Left + shadowShape.Width / 2 > origShape.Left + origShape.Width Then

                    '        verbindungsShape = currentSlide.Shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorElbow, _
                    '                                                                                        shadowShape.Left + shadowShape.Width / 2, shadowShape.Top + shadowShape.Height, _
                    '                                                                                        origShape.Left + origShape.Width, origShape.Top + shadowShape.Height / 2)

                    '    End If

                    'Else
                    '    If shadowShape.Left + shadowShape.Width / 2 < origShape.Left Then

                    '        verbindungsShape = currentSlide.Shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorElbow, _
                    '                                                            origShape.Left, origShape.Top + shadowShape.Height / 2, _
                    '                                                            shadowShape.Left + shadowShape.Width / 2, shadowShape.Top + shadowShape.Height)

                    '    ElseIf shadowShape.Left + shadowShape.Width / 2 > origShape.Left + origShape.Width Then

                    '        verbindungsShape = currentSlide.Shapes.AddConnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorElbow, _
                    '                                                            origShape.Left + origShape.Width, origShape.Top + shadowShape.Height / 2, _
                    '                                                            shadowShape.Left + shadowShape.Width / 2, shadowShape.Top + shadowShape.Height)
                    '    End If

                    '    If Not IsNothing(verbindungsShape) Then
                    '        With verbindungsShape
                    '            .Name = .Name & shadowName
                    '            .Line.BeginArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadNone
                    '            .Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle
                    '            .Line.Weight = 5.0
                    '        End With
                    '    End If

                    'End If


                End If
            End If

        Next



    End Sub

    Friend Sub beschrifteOrigAndShadow(ByVal shadowShapeName As String, ByVal origShapeName As String, ByVal ismilestone As Boolean)

        Try
            Dim shadowShape As PowerPoint.Shape = currentSlide.Shapes(shadowShapeName)
            Dim origShape As PowerPoint.Shape = currentSlide.Shapes(origShapeName)

            Dim shadowIsLeft As Boolean = (shadowShape.Left < origShape.Left)

            ' jetzt die Beschriftung vornehmen
            Dim vTextShadow As String = "Version" & vbLf & previousTimeStamp.ToShortDateString
            Dim vTextOrig As String = "Version" & vbLf & currentTimestamp.ToShortDateString

            Dim vTextOrigShape As PowerPoint.Shape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                                origShape.Left + origShape.Width + 2,
                                                                origShape.Top, 50, 50)
            With vTextOrigShape
                .TextFrame2.TextRange.Text = vTextOrig
                .TextFrame2.TextRange.Font.Size = 10
                .TextFrame2.MarginLeft = 0
                .TextFrame2.MarginRight = 0
                .TextFrame2.MarginTop = 0
                .Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Fill.Solid()
                .Name = .Name & shadowName ' durch den Zusatz shadowName wird sichergestellt, dass die hinterher gelöscht werden
            End With


            Dim vTextShadowShape As PowerPoint.Shape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                            shadowShape.Left - 50,
                                                            shadowShape.Top, 50, 50)

            With vTextShadowShape
                .TextFrame2.TextRange.Text = vTextShadow
                .TextFrame2.TextRange.Font.Size = 10
                .TextFrame2.MarginLeft = 0
                .TextFrame2.MarginRight = 0
                .TextFrame2.MarginTop = 0
                .Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Fill.Solid()
                .Name = .Name & shadowName ' durch den Zusatz shadowName wird sichergestellt, dass die hinterher gelöscht werden 
            End With


            If ismilestone Then
                If shadowIsLeft Then
                    ' Shadow links beschriften , Orig rechts beschriften

                    With vTextShadowShape
                        .Left = shadowShape.Left - (.Width + 3)
                        .Top = shadowShape.Top - (.Height - shadowShape.Height) / 2
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    End With

                    With vTextOrigShape
                        .Left = origShape.Left + origShape.Width + 3
                        .Top = origShape.Top - (.Height - origShape.Height) / 2
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    End With

                Else
                    ' Shadow rechts beschriften, orig links 
                    With vTextOrigShape
                        .Left = origShape.Left - (.Width + 3)
                        .Top = origShape.Top
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    End With

                    With vTextShadowShape
                        .Left = shadowShape.Left + shadowShape.Width + 3
                        .Top = shadowShape.Top
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    End With
                End If
            Else
                ' bei Phasen: Shadow oben beschriften , Original unten
                ' Shadow links beschriften , Orig rechts beschriften

                With vTextShadowShape
                    .Left = shadowShape.Left - (.Width - shadowShape.Width) / 2
                    .Top = shadowShape.Top - (.Height + 3)
                    .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                End With

                With vTextOrigShape
                    .Left = origShape.Left - (.Width - origShape.Width) / 2
                    .Top = origShape.Top + (origShape.Height + 3)
                    .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                End With
            End If


        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' hebt das Abdimmen wieder auf, das eingesetzt wurde, um die Phasen und Meilenstein Shapes, die sich verändert haben hervorzuheben 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub undimAllShapes()



        Dim tValue As Single = 1.0

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes

            If isRelevantMSPHShape(tmpShape) Then
                tValue = 0.0

                If tmpShape.Tags.Item("DIMF").Length > 0 Then
                    Try
                        tValue = CSng(tmpShape.Tags.Item("DIMF"))
                        If tValue < 0 Or tValue > 1.0 Then
                            tValue = 0.0
                        End If
                    Catch ex As Exception

                    End Try
                    tmpShape.Tags.Delete("DIMF")
                    ' innerhalb der if .. Clause aufrufen
                    tmpShape.Fill.Transparency = tValue


                    tmpShape.Line.Transparency = tValue
                End If

            ElseIf isAnnotationShape(tmpShape) Then
                tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
            End If

        Next

        ' jetzt noch alle Shadow-Elemente löschen
        Call deleteShadows()

    End Sub

    ''' <summary>
    ''' wird benötigt, um eine Smart Powerpoint Slide von allen smart Tags zu befreien ... 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub stripOffAllSmartInfo()

        If MsgBox("Wirklich alle Smart-Info löschen ? ", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then

            Dim tagNames(35) As String
            tagNames(0) = "BC"
            tagNames(1) = "CN"
            tagNames(2) = "SN"
            tagNames(3) = "ON"
            tagNames(4) = "BSN"
            tagNames(5) = "BLN"
            tagNames(6) = "SD"
            tagNames(7) = "ED"
            tagNames(8) = "AC"
            tagNames(9) = "AE"
            tagNames(10) = "LU"
            tagNames(11) = "MVD"
            tagNames(21) = "MVE"
            tagNames(22) = "CMT"
            tagNames(23) = "VE"
            tagNames(24) = "PD"
            tagNames(25) = "CHON"
            tagNames(26) = "PRPF"
            tagNames(27) = "PNM"
            tagNames(28) = "VNM"
            tagNames(29) = "CHT"
            tagNames(30) = "ASW"
            tagNames(31) = "COL"
            tagNames(31) = "Q1"
            tagNames(32) = "Q2"
            tagNames(33) = "BID"
            tagNames(34) = "DID"
            tagNames(35) = "NIDS"

            ' Smartslide Info löschen ..

            With currentSlide
                If .Tags.Item("DBURL").Length > 0 Then
                    .Tags.Delete("DBURL")
                End If

                If .Tags.Item("DBNAME").Length > 0 Then
                    .Tags.Delete("DBNAME")
                End If

                If .Tags.Item("SMART") = "visbo" Then
                    .Tags.Delete("SMART")
                End If

                If .Tags.Item("SOC").Length > 0 Then
                    .Tags.Delete("SOC")
                End If

                If .Tags.Item("CRD").Length > 0 Then
                    .Tags.Delete("CRD")
                End If

                If .Tags.Item("FROZEN").Length > 0 Then
                    .Tags.Delete("FROZEN")
                End If

                If .Tags.Item("PREV").Length > 0 Then
                    .Tags.Delete("PREV")
                End If
            End With

            Try
                For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes

                    If isRelevantShape(tmpShape) Then

                        Dim anzTags As Integer = tmpShape.Tags.Count

                        For i As Integer = 0 To 35

                            If tmpShape.Tags.Item(tagNames(i)).Length > 0 Then
                                tmpShape.Tags.Delete(tagNames(i))
                            End If

                        Next


                    End If

                Next

                Call MsgBox("ok, alle SmartInfo gelöscht ...")

            Catch ex As Exception

            End Try


        Else ' nichts tun 
        End If


    End Sub

    ''' <summary>
    ''' löscht alle Shadow Shapes: ein Shadow Element ist das zu einem bestimmten Element gehörende  timestamp Element 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub deleteShadows()
        Dim bigTodoList As New Collection

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            bigTodoList.Add(tmpShape.Name)
        Next

        For Each tmpShpName As String In bigTodoList
            Try

                If tmpShpName.EndsWith(shadowName) Then
                    Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                    tmpShape.Delete()
                End If
            Catch ex As Exception

            End Try
        Next
    End Sub



    '''' <summary>
    '''' löscht beim Beenden von Powerpoint die Hidden App wieder 
    '''' </summary>
    Friend Sub closeExcelAPP(ByVal xlApp As xlNS.Application)
        Try
            'ur:2019-06-04: Test
            If xlApp.Workbooks.Count = 1 Then

                For Each wb As Excel.Workbook In xlApp.Workbooks
                    wb.Saved = True
                    xlApp.Quit()
                Next
            Else
                For Each wb As Excel.Workbook In xlApp.Workbooks
                    wb.Close(SaveChanges:=False)
                Next

            End If

            'If Not IsNothing(xlApp) Then
            '    For Each tmpWB As Excel.Workbook In CType(xlApp.Workbooks, Excel.Workbooks)
            '        tmpWB.Close(SaveChanges:=False)
            '    Next
            '    xlApp.Quit()
            'End If

            'xlApp = Nothing
        Catch ex As Exception

        End Try

    End Sub

    ''
    ''' <summary>
    ''' löscht das Search-Pane mit den Feldern  
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="isMovedShape"></param>
    ''' <remarks></remarks>
    Friend Sub clearSearchPane(ByVal tmpShape As PowerPoint.Shape, Optional ByVal isMovedShape As Boolean = False)

        If IsNothing(tmpShape) Then
            With ucSearchView
                ' eigentlich soll doch nur selListboxNames zurückgesetzt werden und die Auswahlen daraus ...
                '.cathegoryList.SelectedItem = Nothing
                '.shwOhneLight.Checked = False
                '.shwGreenLight.Checked = False
                '.shwYellowLight.Checked = False
                '.shwRedLight.Checked = False
                '.filterText.Text = ""
                '.listboxNames.Items.Clear()
                '.selListboxNames.Items.Clear()
                '.fülltListbox()

                ' tk 11.1.18
                .listboxNames.SelectedItems.Clear()
                .selListboxNames.Items.Clear()
            End With
        End If

    End Sub


    ''
    ''' <summary>
    ''' aktualisiert das Info-Pane mit den Feldern ElemName, ElemDate, BreadCrumb und aLuTv-Text 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="isMovedShape"></param>
    ''' <remarks></remarks>
    Friend Sub aktualisiereInfoPane(ByVal tmpShape As PowerPoint.Shape, Optional ByVal isMovedShape As Boolean = False)

        If Not IsNothing(ucPropertiesView) Then

            ' ''With infoFrm
            ' ''    .btnSendToHome.Enabled = homeButtonRelevance
            ' ''    .btnSentToChange.Enabled = changedButtonRelevance
            ' ''End With

            If Not IsNothing(tmpShape) And Not IsNothing(selectedPlanShapes) Then

                If selectedPlanShapes.Count = 1 Then

                    If isSymbolShape(tmpShape) Then
                        With ucPropertiesView

                            ' positioniert die Darstellungs-Elemente entsprechend
                            .symbolMode(True)
                            .eleName.Text = bestimmeSymbolName(tmpShape)
                            .eleAmpelText.Text = bestimmeSymbolText(tmpShape)

                            ' Dokumenten Links ausblenden 
                            .setLinksToVisible(False)
                            .setLinkValues(tmpShape)

                        End With
                    Else
                        With ucPropertiesView

                            ' positioniert die Darstellungs-Elemente entsprechend
                            .symbolMode(False)

                            .eleName.Text = "                                                                   "
                            .eleName.Text = bestimmeElemText(tmpShape, False, False, showBestName)

                            .eleDatum.Text = bestimmeElemDateText(tmpShape, False, False)

                            'Dim rgbFarbe As Drawing.Color = Drawing.Color.FromArgb(CType(trafficLightColors(CInt(tmpShape.Tags.Item("AC"))), Integer))

                            Dim ampelfarbe As Integer = CInt(tmpShape.Tags.Item("AC"))

                            Select Case CInt(tmpShape.Tags.Item("AC"))

                                Case PTfarbe.none
                                    .eleAmpel.BackColor = Drawing.Color.Silver
                                Case PTfarbe.green
                                    .eleAmpel.BackColor = Drawing.Color.Green
                                Case PTfarbe.yellow
                                    .eleAmpel.BackColor = Drawing.Color.Yellow
                                Case PTfarbe.red
                                    .eleAmpel.BackColor = Drawing.Color.Firebrick

                            End Select

                            .percentDone.Text = bestimmeElemPD(tmpShape)

                            .eleAmpelText.Text = bestimmeElemAE(tmpShape)

                            .eleDeliverables.Text = bestimmeElemLU(tmpShape)

                            .eleRespons.Text = bestimmeElemVE(tmpShape)

                            ' die Link Buttons grundsätzlich einblenden 
                            .setLinksToVisible(False)
                            .setLinkValues(tmpShape)


                        End With
                    End If


                ElseIf selectedPlanShapes.Count > 1 Then

                    'Dim rdbCode As Integer = calcRDB()

                    With ucPropertiesView
                        ' leeren ...
                        Call .emptyPane()

                        If .eleName.Text <> bestimmeElemText(tmpShape, False, True, showBestName) Then
                            .eleName.Text = " ... "
                        End If
                        If .eleDatum.Text <> bestimmeElemDateText(tmpShape, False) Then
                            .eleDatum.Text = " ... "
                        End If

                        .eleRespons.Text = ""
                        .eleDatum.Enabled = False
                        .eleDeliverables.Text = ""
                        .eleAmpelText.Text = ""
                        .eleAmpel.BackColor = Drawing.Color.Silver
                        .percentDone.Text = ""

                        .documentsLink = ""
                        .myDocumentsLink = ""

                        ' Dokumenten Links ausblenden 
                        .setLinksToVisible(False)
                        .setLinkValues(Nothing)

                    End With

                End If
            Else
                ' Info Pane Inhalte zurücksetzen ... 
                With ucPropertiesView
                    ' leeren 
                    Call .emptyPane()

                    .eleName.Text = ""
                    .eleDatum.Text = ""
                    .eleDeliverables.Text = ""
                    .eleAmpel.BackColor = Drawing.Color.Silver
                    .eleAmpelText.Text = ""
                    .eleRespons.Text = ""
                    .percentDone.Text = ""

                    .documentsLink = ""
                    .myDocumentsLink = ""

                End With

            End If

        End If
    End Sub

    ''' <summary>
    ''' ''' gibt den  Lieferumfang zurück (Tag= LU)
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemLU(ByVal curShape As PowerPoint.Shape) As String

        Dim tmpText As String = ""

        Dim tmpStr() As String
        If curShape.Tags.Item("LU").Length > 0 Then
            tmpStr = curShape.Tags.Item("LU").Split(New Char() {CType("#", Char)})
            For i As Integer = 0 To tmpStr.Length - 1
                tmpText = tmpText & tmpStr(i) & vbLf
            Next
        End If

        bestimmeElemLU = tmpText

    End Function

    ''' <summary>
    ''' bestimme die Ampel-Erläuterung ( Tag= AE)
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemAE(ByVal curShape As PowerPoint.Shape) As String

        Dim tmpText As String = ""

        If curShape.Tags.Item("AE").Length > 0 Then
            tmpText = tmpText & curShape.Tags.Item("AE")
        End If

        bestimmeElemAE = tmpText

    End Function

    ''' <summary>
    ''' bestimme den Verantwortlichen (Tag= VE)
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemVE(ByVal curShape As PowerPoint.Shape) As String

        Dim tmpText As String = ""

        If curShape.Tags.Item("VE").Length > 0 Then
            tmpText = tmpText & curShape.Tags.Item("VE")
        End If

        bestimmeElemVE = tmpText

    End Function

    Public Function bestimmeElemMVE(ByVal curShape As PowerPoint.Shape) As String

        Dim tmpText As String = ""

        If curShape.Tags.Item("MVE").Length > 0 Then
            tmpText = tmpText & curShape.Tags.Item("MVE")
        End If
        bestimmeElemMVE = tmpText

    End Function
    ''' <summary>
    ''' bestimme die percentDone einer Phase ( Tag= PD)
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemPD(ByVal curShape As PowerPoint.Shape) As String

        Dim tmpText As String = ""

        If curShape.Tags.Item("PD").Length > 0 Then

            ' Änderung tk, wenn PercentDone als 50, etc.  eingetragen worden ist ..
            If CDbl(curShape.Tags.Item("PD")) > 1.0 Then
                tmpText = tmpText & (CDbl(curShape.Tags.Item("PD"))).ToString & " %"
            Else
                tmpText = tmpText & (CDbl(curShape.Tags.Item("PD")) * 100).ToString & " %"
            End If


        End If

        bestimmeElemPD = tmpText

    End Function


    Public Function bestimmeElemResCosts(ByVal curshape As PowerPoint.Shape) As String

        Dim tmpText As String = ""

        If Not noDBAccessInPPT And pptShapeIsPhase(curshape) Then
            Try
                Dim pvName As String = getPVnameFromShpName(curshape.Name)
                Dim vpid As String = getVPIDFromTags(curshape)

                'Dim hproj As clsProjekt = smartSlideLists.getTSProject(pvName, currentTimestamp)
                Dim hproj As clsProjekt = timeMachine.getProjectVersion(pvName, currentTimestamp, vpid)
                Dim phNameID As String = getElemIDFromShpName(curshape.Name)
                Dim cPhase As clsPhase = hproj.getPhaseByID(phNameID)
                Dim roleInformations As SortedList(Of String, Double) = cPhase.getRoleNamesAndValues
                Dim costInformations As SortedList(Of String, Double) = cPhase.getCostNamesAndValues

                ' ''If Not shortForm Then

                ' ''    If englishLanguage Then
                ' ''        'tmpText = getElemNameFromShpName(curShape.Name) & " Resource/Costs :" & vbLf
                ' ''        tmpText = tmpText & "Resource/Costs :" & vbLf
                ' ''    Else
                ' ''        'tmpText = getElemNameFromShpName(curShape.Name) & " Ressourcen/Kosten:" & vbLf
                ' ''        tmpText = tmpText & "Ressourcen/Kosten:" & vbLf
                ' ''    End If

                ' ''Else
                ' ''    If englishLanguage Then
                ' ''        tmpText = "Resource/Costs :" & vbLf
                ' ''    Else
                ' ''        tmpText = "Ressourcen/Kosten:" & vbLf
                ' ''    End If
                ' ''End If


                Dim unit As String
                If englishLanguage Then
                    unit = " PD"
                Else
                    unit = " PT"
                End If

                For i As Integer = 1 To roleInformations.Count
                    tmpText = tmpText &
                        roleInformations.ElementAt(i - 1).Key & ": " & CInt(roleInformations.ElementAt(i - 1).Value).ToString & unit & vbLf
                Next

                If costInformations.Count > 0 And roleInformations.Count > 0 Then
                    tmpText = tmpText & vbLf
                End If

                unit = " TE"
                For i As Integer = 1 To costInformations.Count
                    tmpText = tmpText &
                        costInformations.ElementAt(i - 1).Key & ": " & CInt(costInformations.ElementAt(i - 1).Value).ToString & unit & vbLf
                Next

            Catch ex As Exception
                tmpText = "Phase " & getElemNameFromShpName(curshape.Name)
            End Try



        ElseIf noDBAccessInPPT And pptShapeIsPhase(curshape) Then
            ' ''If Not shortForm Then
            ' ''    If englishLanguage Then
            ' ''        tmpText = "Resource/Costs " & getElemNameFromShpName(curShape.Name) & ":" & vbLf & _
            ' ''        "no DB access ..."
            ' ''    Else
            ' ''        tmpText = "Ressourcen / Kosten " & getElemNameFromShpName(curShape.Name) & ":" & vbLf & _
            ' ''            "kein DB Zugriff ..."
            ' ''    End If

            ' ''Else
            ' ''    If englishLanguage Then
            ' ''        tmpText = "no DB access"
            ' ''    Else
            ' ''        tmpText = "kein DB Zugriff"
            ' ''    End If

            ' ''End If
        Else
            tmpText = ""
        End If

        bestimmeElemResCosts = tmpText

    End Function


    ''' <summary>
    ''' führt die Time-Machine Action aus, übergeben wird lediglich die Kennzeichnung um welchen Time-Machine Button es sich handelt 
    ''' wird aufgerufen direkt aus den Buttons des Ribbon1
    ''' </summary>
    ''' <param name="ptNavType"></param>
    Public Sub updateSelectedSlide(ByVal ptNavType As Integer, ByVal specDate As Date)

        Try
            Dim errmsg As String = ""
            'Call closeExcelAPP()

            Dim pres As PowerPoint.Presentation = CType(currentSlide.Parent, PowerPoint.Presentation)
            Dim formerSlide As PowerPoint.Slide = currentSlide
            'Dim saveCurrentTimeStamp As Date = currentTimestamp

            'ur:2019-06-04

            Dim slideIDList As New SortedList(Of Integer, Integer)

            For Each sl As PowerPoint.Slide In pres.Slides
                slideIDList.Add(sl.SlideID, sl.SlideID)
            Next


            Dim sld As PowerPoint.Slide = currentSlide


            ' neue Slide , also leer machen ... 
            smartSlideLists = New clsSmartSlideListen

            If Not IsNothing(sld) Then
                If Not (sld.Tags.Item("FROZEN").Length > 0) And (sld.Tags.Item("SMART") = "visbo") Then

                    If userIsEntitled(errmsg, sld) Then    ' User ist bereits eingeloggt 

                        currentTimestamp = getCurrentTimeStampFromSlide(sld)
                        Call pptAPP_AufbauSmartSlideLists(sld)
                        Call prepareAndPerformBtnAction(ptNavType, specDate, False)

                    Else
                        ' hier ggf auf invisible setzen, wenn erforderlich 
                        Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoFalse)
                    End If



                End If
            End If
            'Next

            ' ur:2019-06-04: wird nicht benötigt, wenn nur jede selektierte Slide einzeln upgedated wird

            'If currentSlide.SlideID <> formerSlide.SlideID Then

            '    currentSlide = formerSlide
            '    currentTimestamp = getCurrentTimeStampFromSlide(currentSlide)

            '    smartSlideLists = New clsSmartSlideListen

            '    If Not IsNothing(currentSlide) Then
            '        If Not (currentSlide.Tags.Item("FROZEN").Length > 0) And (currentSlide.Tags.Item("SMART") = "visbo") Then

            '            If userIsEntitled(errmsg, currentSlide) Then

            '                Call pptAPP_AufbauSmartSlideLists(currentSlide)

            '            Else
            '                ' hier ggf auf invisible setzen, wenn erforderlich
            '                Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoFalse)
            '            End If

            '        End If
            '    End If

            'End If

            ' das Formular ggf, also wenn aktiv,  updaten 
            If Not IsNothing(changeFrm) Then
                changeFrm.neuAufbau()
            End If

            'Dim ticsEnd As Integer = My.Computer.Clock.TickCount
            'Call MsgBox("zeit benötigt: " & CStr(ticsEnd - ticsStart))

            pres.Application.Activate()


        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try


    End Sub


    ''' <summary>
    ''' führt die Button Action der Time-Machine aus in der currentSlide aus; setzt in der Slide den previousTimestamp und den currentTimestamp
    ''' </summary>
    ''' <param name="newdate"></param>
    ''' <remarks></remarks>
    Public Sub performBtnAction(ByVal newdate As Date)

        ' tk 28.10.18 braucht man doch nicht ... 
        '' Versuch den Undo-Stack zu löschen
        '' pptAPP.StartNewUndoEntry()

        Dim ddiff As Long = DateDiff(DateInterval.Second, newdate, currentTimestamp)

        If ddiff <> 0 Then


            previousVariantName = currentVariantname
            previousTimeStamp = currentTimestamp
            currentTimestamp = newdate

            ' jetzt muss auch die currentConstellationName wieder zurück gesetzt werden 
            currentConstellationPvName = ""

            Call moveAllShapes()

            'Call setBtnEnablements()

            Call setCurrentTimestampInSlide(currentTimestamp)
            Call setPreviousTimestampInSlide(previousTimeStamp)

            Call showTSMessage(currentTimestamp)

            Try
                If Not IsNothing(selectedPlanShapes) Then

                    If selectedPlanShapes.Count = 1 Then
                        Dim curShape As PowerPoint.Shape = selectedPlanShapes.Item(1)

                        Call aktualisiereInfoPane(curShape)
                        Call aktualisiereInfoFrm(curShape)

                    End If
                End If
            Catch ex As Exception

            End Try

            ' jetzt noch das InfoPane aktualisieren
            If Not IsNothing(selectedPlanShapes) Then
                If selectedPlanShapes.Count >= 1 Then
                    Dim tmpShape As PowerPoint.Shape = selectedPlanShapes.Item(1)
                    If isRelevantMSPHShape(tmpShape) Then
                        Call aktualisiereInfoPane(tmpShape)
                    End If
                End If
            End If


        End If

    End Sub


    ''' <summary>
    ''' wird aufgerufen, sobald der User eine spezielle Slide updaten will
    ''' der currentTimestamp wird hier nicht mehr gesetzt ... der wird nur in den Time-Machine Routinen und bei Window_Activate geholt bzw. bei De-Activate gespeichert  geändert  
    ''' </summary>
    ''' <param name="specSlide"></param>
    ''' <remarks></remarks>
    Public Sub pptAPP_AufbauSmartSlideLists(specSlide As PowerPoint.Slide)

        ' die aktuelle Slide setzen 

        Try

            smartSlideLists = New clsSmartSlideListen
            ' tk 22.8.18, SlideCoordInfo steuert die Berechnung des Datums anhand der "versteckten" Kalenderlinie ..
            slideCoordInfo = Nothing
            Try

                If Not IsNothing(searchPane) Then
                    If searchPane.Visible Then
                        Call clearSearchPane(Nothing)
                    End If
                End If

            Catch ex As Exception

            End Try


            ' jetzt ggf gesetzte Glow MArker zurücksetzen ... 
            currentSlide = specSlide

            ' tk 29.10.18 nicht nötig , wird an andere Stelle besser gemacht 
            'Try
            '    If Not IsNothing(currentSlide) Then
            '        If currentSlide.Tags.Item("SMART").Length > 0 Then

            '            Call deleteMarkerShapes()
            '            Call putAllNoPrioShapesInNoshow()

            '        End If
            '    End If


            'Catch ex As Exception

            'End Try


            If Not IsNothing(currentSlide) Then

                If currentSlide.Tags.Count > 0 Then
                    Try
                        If currentSlide.Tags.Item("SMART") = "visbo" Then

                            ' Aufbau SmartSlideLists muss immer ohne DB erfolgen können ! 

                            ' die HomeButtonRelevanz setzen 
                            homeButtonRelevance = False
                            changedButtonRelevance = False

                            slideHasSmartElements = True

                            Try

                                slideCoordInfo = New clsPPTShapes
                                slideCoordInfo.pptSlide = currentSlide


                                With currentSlide

                                    If .Tags.Item("CALL").Length > 0 And .Tags.Item("CALR").Length > 0 Then
                                        Dim tmpSD As String = .Tags.Item("CALL")
                                        Dim tmpED As String = .Tags.Item("CALR")
                                        slideCoordInfo.setCalendarDates(CDate(tmpSD), CDate(tmpED))
                                    End If

                                    If .Tags.Item("SOC").Length > 0 Then
                                        StartofCalendar = CDate(.Tags.Item("SOC"))
                                    End If



                                End With

                            Catch ex As Exception
                                slideCoordInfo = Nothing
                            End Try


                            Call buildSmartSlideLists()

                            ' jetzt merken, wie die Settings für homeButton und changedButton waren ..
                            initialHomeButtonRelevance = homeButtonRelevance
                            initialChangedButtonRelevance = changedButtonRelevance

                            If Not IsNothing(searchPane) Then
                                If searchPane.Visible Then

                                    If slideHasSmartElements Then

                                        ucSearchView.fülltListbox(showTrafficLights)

                                    End If
                                End If
                            End If

                        End If
                    Catch ex As Exception

                    End Try

                Else

                    slideHasSmartElements = False
                    ' Listen löschen
                    smartSlideLists = New clsSmartSlideListen
                    ' tk ergänzt am 22.8.18; slideCoordInfo steuert die Berechnung des Datums aufgrund PPTCalendarStart und PPTCalendarEnde
                    slideCoordInfo = Nothing

                    If Not IsNothing(searchPane) Then
                        If searchPane.Visible Then
                            Call clearSearchPane(Nothing)
                        End If
                    End If

                End If

            End If


        Catch ex As Exception
            Call MsgBox("Fehler in pptAPP_AufbauSmartSlideLists")
        End Try
    End Sub


    ''' <summary>
    ''' führt den Code gehe-zum-letzten bzw Visbo-Update aus 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub prepareAndPerformBtnAction(ByVal updateModus As Integer,
                           ByRef specDate As Date,
                           Optional ByVal showMessage As Boolean = True)

        Dim newDate As Date


        ' jetzt kann das newDate errechnet werden 
        '
        If updateModus = ptNavigationButtons.previous Then

            If currentSlide.Tags.Item("PREV").Length > 0 Then
                'smartSlideLists.prevDate = CDate(currentSlide.Tags.Item("PREV"))
                newDate = CDate(currentSlide.Tags.Item("PREV"))
            End If


        Else
            newDate = timeMachine.getNextNavigationDate(updateModus, specDate)
        End If

        Call performBtnAction(newDate)


        If updateModus = ptNavigationButtons.letzter Then
            specDate = newDate
        End If
    End Sub

    Public Sub own_SlideSelectionChanged(ByVal Sld As PowerPoint.Slide)

        Dim beforeSlideTimestamp As Date = Date.MinValue

        ' die aktuelle Slide setzen 
        If Not IsNothing(Sld) Then


            ' re-set parameters necessary for Creating reporting templates
            currentSldHasProjectTemplates = False
            currentSldHasMultiProjectTemplates = False
            currentSldHasPortfolioTemplates = False

            If currentPresHasVISBOElements Then
                ' nur dann muss irgendwas weitergemacht werden ..

                Dim afterSlideID As Integer = Sld.SlideID ' aktuell selektierte SlideID

                ' hier muss nur weitergemacht werden, wenn es sich um eine VISBO slide handelt 
                If isVisboSlide(Sld) Then

                    Dim afterSlideKennung As String = CType(Sld.Parent, PowerPoint.Presentation).Name & afterSlideID.ToString
                    Dim beforeSlideKennung As String = ""

                    Dim key As String = CType(Sld.Parent, PowerPoint.Presentation).Name

                    Dim beforeSlideID As Integer = 0               ' zuvor selektierte SlideID

                    If Not IsNothing(currentSlide) Then
                        Try
                            beforeSlideID = currentSlide.SlideID
                            beforeSlideKennung = CType(currentSlide.Parent, PowerPoint.Presentation).Name & beforeSlideID.ToString
                            ' jetzt die beforeSlideTimestamp setzen 
                            With currentSlide
                                If .Tags.Item("CRD").Length > 0 Then
                                    beforeSlideTimestamp = getCurrentTimeStampFromSlide(currentSlide)
                                End If
                            End With

                        Catch ex As Exception

                        End Try

                    End If


                    '' jetzt die CurrentSlide setzen , denn evtl kommt man ja gar nicht in pptAPP_UpdateOneSlide
                    currentSlide = Sld

                    If beforeSlideKennung <> afterSlideKennung Or smartSlideLists.countProjects = 0 Then
                        Call pptAPP_AufbauSmartSlideLists(Sld)

                    End If

                    ' jetzt die currentTimeStamp setzen 
                    With currentSlide
                        If .Tags.Item("CRD").Length > 0 Then
                            currentTimestamp = getCurrentTimeStampFromSlide(currentSlide)
                        End If
                    End With

                    If beforeSlideTimestamp > Date.MinValue Then

                        Dim diff As Long = DateDiff(DateInterval.Minute, currentTimestamp, beforeSlideTimestamp)

                        If diff <> 0 And Not noDBAccessInPPT Then
                            Call updateSelectedSlide(ptNavigationButtons.individual, beforeSlideTimestamp)
                        End If


                        ' nur wenn die SlideID gewechselt hat, muss agiert werden
                        ' dabei auch berücksichtigen, ob sich Presentation geändert hat 
                        If beforeSlideKennung <> afterSlideKennung Then
                            Try
                                ' das Change-Formular aktualisieren, wenn es gezeigt wird  
                                Dim hwind As Integer = pptAPP.ActiveWindow.HWND
                                If Not IsNothing(changeFrm) Then

                                    changeFrm.changeliste.clearChangeList()

                                    If chgeLstListe.ContainsKey(key) Then
                                        If chgeLstListe.Item(key).ContainsKey(currentSlide.SlideID) Then
                                            changeFrm.changeliste = chgeLstListe.Item(key).Item(currentSlide.SlideID)
                                        Else
                                            ' eine Liste für die neue SlideID einfügen ..
                                        End If
                                    End If

                                    changeFrm.neuAufbau()
                                End If
                            Catch ex As Exception

                            End Try

                        End If       'Ende ob SlideIDs ungleich sind
                    End If
                Else
                    'tk 8.10.19 es muss immer eine Slide geben 
                    currentSlide = Sld
                End If
            Else
                'tk 8.10.19 es muss immer eine Slide geben 
                currentSlide = Sld
            End If ' if currentPresHasVisboElements
        Else
            'nichts tun
        End If

    End Sub


End Module
