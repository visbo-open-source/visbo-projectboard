Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ProjectBoardBasic
Imports xlNS = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core.MsoThemeColorIndex

Module Module1

    Friend WithEvents pptAPP As PowerPoint.Application

    'Friend visboInfoActivated As Boolean = False
    Friend formIsShown As Boolean = False
    Friend Const markerName As String = "VisboMarker"
    Friend Const protectionTag As String = "VisboProtection"
    Friend Const protectionValue As String = "VisboValue"
    Friend Const noVariantName As String = "-9999999"

    Friend myPPTWindow As PowerPoint.DocumentWindow = Nothing

    Friend xlApp As xlNS.Application = Nothing
    Friend updateWorkbook As xlNS.Workbook = Nothing

    Friend Const changeColor As Integer = PowerPoint.XlRgbColor.rgbSteelBlue
    Friend currentSlide As PowerPoint.Slide
    Friend VisboProtected As Boolean = False
    Friend protectionSolved As Boolean = False

    Friend thereIsNoVersionFieldOnSlide As Boolean = True
    ' bestimmt, ob in englisch oder auf deutsch ..
    Friend englishLanguage As Boolean = True

    ' was ist der aktuelle Timestamp der Slide 
    Friend currentTimestamp As Date = Date.MinValue
    Friend previousTimeStamp As Date = Date.MinValue

    Friend currentVariantname As String = ""
    Friend previousVariantName As String = noVariantName

    ' der Key ist der Name des Referenz-Shapes, zu dem der Marker gezeichnet wird , der Value ist der Name des Marker-Shapes 
    Friend markerShpNames As New SortedList(Of String, String)

    ' wird gesetzt in Einstellungen 
    ' steuert, ob extended seach gemacht werden kann; wirkt auf Suchfeld (NAme, Original Name, Abkürzung, ..)  
    Friend extSearch As Boolean = False
    ' wird gesetzt in Einstellungen 
    ' gibt an, mit welcher Schriftgroesse der Text geschrieben wird 
    Friend schriftGroesse As Double = 8.0
    ' wird gesetzt in Einstellungen 
    ' gibt an, ob das Breadcrumb Feld gezeigt werden soll 
    Friend showBreadCrumbField As Boolean = False
    ' gibt die MArker-Höhe und Breite an 
    Friend markerHeight As Double = 19
    Friend markerWidth As Double = 13


    ' gibt an, ob bei der suche die gefundenen Elemente mit AMrker angezeigt werden sollen oder nicht .. 
    Friend showMarker As Boolean = False

    ' globale Variable, die angibt, ob ShortName gezeichnet werden soll 
    Friend showShortName As Boolean = False
    ' globlaela Variable, die anzeigt, ob Orginal Name gezeigt werden soll 
    Friend showOrigName As Boolean = False

    Friend protectType As Integer
    Friend protectFeld1 As String = ""
    Friend protectFeld2 As String = ""

    ' hier sollen die Namen aus projectboardDefinitions übernommen werden 
    'Friend dbURL As String = ""
    'Friend dbName As String = ""
    'Friend userName As String = ""
    'Friend userPWD As String = ""

    Friend noDBAccessInPPT As Boolean = True

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
    Friend Function calcColorCode() As Integer

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
    ''' zeigt bei den Shapes, die die angegebene Ampelfarbe haben, diese Farbe als Hintergrund Schatten an bzw. löscht den Hintergrund Schatten wieder
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

                If show Then
                    ' mit Schatten einfärben 
                    With shapesToBeColored.Shadow
                        .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        .Type = Microsoft.Office.Core.MsoShadowType.msoShadow25
                        .Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow
                        .Blur = 0
                        .Size = 160
                        .OffsetX = 0
                        .OffsetY = 0
                        .Transparency = 0
                        .ForeColor.RGB = trafficLightColors(ampelColor)
                    End With
                Else
                    ' Schatten wieder wegnehmen 
                    With shapesToBeColored.Shadow
                        .Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                    End With
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
    Friend Sub faerbeShape(ByRef tmpShape As PowerPoint.Shape, _
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
                        .ForeColor.RGB = trafficLightColors(ampelColor)
                    End With
                Else
                    ' Schatten wieder wegnehmen 
                    With tmpShape.Shadow
                        .ForeColor.RGB = trafficLightColors(ampelColor)
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
    Friend Function userIsEntitled(ByRef msg As String) As Boolean

        Dim tmpResult As Boolean = False

        If userhasValidLicence Then

            If pptAPP.ActivePresentation.Tags.Item(protectionTag) = "PWD" Or _
            pptAPP.ActivePresentation.Tags.Item(protectionTag) = "COMPUTER" Or _
            pptAPP.ActivePresentation.Tags.Item(protectionTag) = "DATABASE" Then

                VisboProtected = True

                If Not protectionSolved Then
                    If pptAPP.ActivePresentation.Tags.Item(protectionTag) = "PWD" Then

                        Dim pwdFormular As New frmPassword
                        If pwdFormular.ShowDialog() = Windows.Forms.DialogResult.OK Then
                            If pwdFormular.pwdText.Text = pptAPP.ActivePresentation.Tags.Item(protectionValue) Then
                                ' in allen Slides den Sicht Schutz aufheben 
                                protectionSolved = True
                                Call makeVisboShapesVisible(True)
                            End If
                        Else
                            If englishLanguage Then
                                msg = "wrong password ..."
                            Else
                                msg = "Password falsch ..."
                            End If

                            tmpResult = False
                        End If

                    ElseIf pptAPP.ActivePresentation.Tags.Item(protectionTag) = "COMPUTER" Then
                        Dim userName As String = My.Computer.Name
                        If pptAPP.ActivePresentation.Tags.Item(protectionValue) = userName Then
                            ' in allen Slides den Sicht Schutz aufheben 
                            protectionSolved = True
                            Call makeVisboShapesVisible(True)
                        Else
                            tmpResult = False
                            If englishLanguage Then
                                msg = "computer / user not entitled ..."
                            Else
                                msg = "nicht berechtigter Computer bzw. User ..."
                            End If

                        End If

                    ElseIf pptAPP.ActivePresentation.Tags.Item(protectionTag) = "DATABASE" Then
                        ' die Login Maske aufschalten ... 
                        ' muss noch eingeloggt werden ? 
                        If noDBAccessInPPT Then
                            ' jetzt die Login Maske aufrufen ... 

                            If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

                                Call logInToMongoDB()

                                If Not noDBAccessInPPT Then
                                    ' in allen Slides den Sicht Schutz aufheben 
                                    protectionSolved = True
                                    Call makeVisboShapesVisible(True)

                                End If

                            End If

                        End If

                    End If
                End If

                If protectionSolved Then
                    tmpResult = True
                End If
            Else
                tmpResult = True
            End If

        Else
            tmpResult = False
            If englishLanguage Then
                msg = "no valid licence ... please contact your system-administrator"
            Else
                msg = "keine gültige Lizenz ... bitte kontaktieren Sie Ihren System-Administrator"
            End If

        End If

        

        userIsEntitled = tmpResult

    End Function

    Private Function userHasValidLicence()
        userHasValidLicence = True
    End Function

    ''' <summary>
    ''' hier wird bestimmt, ob es sich um eine VisboProtected Präsentation handelt 
    ''' </summary>
    ''' <param name="Pres"></param>
    ''' <remarks></remarks>
    Private Sub pptAPP_AfterPresentationOpen(Pres As PowerPoint.Presentation) Handles pptAPP.AfterPresentationOpen

        ' ein ggf. vorhandener Schutz  muss wieder aktiviert werden ... 
        protectionSolved = False


        ' gibt es eine Sprachen-Tabelle ? 
        Dim langGUID As String = pptAPP.ActivePresentation.Tags.Item("langGUID")
        If langGUID.Length > 0 Then

            Dim langXMLpart As Office.CustomXMLPart = pptAPP.ActivePresentation.CustomXMLParts.SelectByID(langGUID)

            Dim langXMLstring = langXMLpart.XML
            languages = xml_deserialize(langXMLstring)

        End If


    End Sub

    Private Sub pptAPP_NewPresentation(Pres As Microsoft.Office.Interop.PowerPoint.Presentation) Handles pptAPP.NewPresentation

    End Sub

    Private Sub pptAPP_PresentationBeforeClose(Pres As PowerPoint.Presentation, ByRef Cancel As Boolean) Handles pptAPP.PresentationBeforeClose


        If Not IsNothing(currentSlide) Then
            If currentSlide.Tags.Item("SMART").Length > 0 Then
                Call resetMovedGlowOfShapes()
            End If
        End If

        Try
            Call closeExcelAPP()
        Catch ex As Exception

        End Try

        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If
    End Sub

    Private Sub pptAPP_PresentationBeforeSave(Pres As PowerPoint.Presentation, ByRef Cancel As Boolean) Handles pptAPP.PresentationBeforeSave
        ' wenn VisboProtected, dann müssen jetzt alle relevanten Shapes auf invisible gesetzt werden ...

        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If

    End Sub

    Private Sub pptAPP_PresentationCloseFinal(Pres As PowerPoint.Presentation) Handles pptAPP.PresentationCloseFinal



    End Sub

    ''' <summary>
    ''' ein VISBO Protected File kann nur als pptx gespeichert werden ...
    ''' </summary>
    ''' <param name="Pres"></param>
    ''' <remarks></remarks>
    Private Sub pptAPP_PresentationSave(Pres As PowerPoint.Presentation) Handles pptAPP.PresentationSave
        If VisboProtected And Not Pres.Name.EndsWith(".pptx") Then
            If englishLanguage Then
                Call MsgBox("Store only possible with file extension .pptx !")
            Else
                Call MsgBox("Speichern nur als .pptx möglich!")
            End If

            Dim vollerName As String = Pres.FullName
            Dim correctName As String = Pres.Name & ".pptx"

            Pres.SaveAs(correctName)
            My.Computer.FileSystem.DeleteFile(vollerName)
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
            ' jetzt ggf gesetzte Glow MArker zurücksetzen ... 

            currentSlide = SldRange.Item(1)

            Try
                If Not IsNothing(currentSlide) Then
                    If currentSlide.Tags.Item("SMART").Length > 0 Then
                        Call resetMovedGlowOfShapes()
                    End If
                End If

                Call deleteMarkerShapes()


            Catch ex As Exception

            End Try
            
            thereIsNoVersionFieldOnSlide = True

            If currentSlide.Tags.Count > 0 Then
                Try
                    If currentSlide.Tags.Item("SMART").Length > 0 Then

                        ' wird benötigt, um jetzt die Infos zu der Datenbank rauszulesen ...
                        Call getDBsettings()

                        Dim msg As String = ""
                        If userIsEntitled(msg) Then

                            ' die HomeButtonRelevanz setzen 
                            homeButtonRelevance = False
                            changedButtonRelevance = False

                            slideHasSmartElements = True

                            Try

                                slideCoordInfo = New clsPPTShapes
                                slideCoordInfo.pptSlide = currentSlide

                                With currentSlide

                                    ' currentTimeStamp setzen 
                                    If .Tags.Item("CRD").Length > 0 Then
                                        currentTimestamp = CDate(.Tags.Item("CRD"))
                                    End If

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

                            ' jetzt merken, wie die Settings für homeButton und chengedButton waren ..
                            initialHomeButtonRelevance = homeButtonRelevance
                            initialChangedButtonRelevance = changedButtonRelevance

                        Else
                            Call MsgBox(msg)
                        End If

                    End If
                Catch ex As Exception

                End Try
            Else
                slideHasSmartElements = False
            End If


        Else
            ' nichts tun, das heisst auch nichts verändern ...
        End If

    End Sub

    ''' <summary>
    ''' bestimmt die Settings der Datenbank, sofern welche da sind 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub getDBsettings()
        With currentSlide
            
            If .Tags.Item("DBURL").Length > 0 And _
                .Tags.Item("DBNAME").Length > 0 Then

                If .Tags.Item("DBURL") = awinSettings.databaseURL And _
                    .Tags.Item("DBNAME") = awinSettings.databaseName And Not noDBAccessInPPT Then
                    ' nichts machen, user ist schon berechtigt ...
                Else
                    noDBAccessInPPT = True
                    awinSettings.databaseURL = .Tags.Item("DBURL")
                    awinSettings.databaseName = .Tags.Item("DBNAME")
                End If
                

            End If
        End With
    End Sub

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
    ''' erstellt die SmartSlideListen neu ... 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub buildSmartSlideLists()

        ' zurücksetzen der SmartSlideLists
        smartSlideLists = New clsSmartSlideListen
        bekannteIDs = New SortedList(Of Integer, String)


        With currentSlide
            If .Tags.Item("CRD").Length > 0 Then
                smartSlideLists.creationDate = CDate(.Tags.Item("CRD"))
            End If

            If .Tags.Item("DBURL").Length > 0 And _
                .Tags.Item("DBNAME").Length > 0 Then

                smartSlideLists.slideDBName = .Tags.Item("DBNAME")
                smartSlideLists.slideDBUrl = .Tags.Item("DBURL")

                If awinSettings.databaseURL <> smartSlideLists.slideDBUrl Or _
                    awinSettings.databaseName <> smartSlideLists.slideDBName Then

                    noDBAccessInPPT = True
                    awinSettings.databaseURL = smartSlideLists.slideDBUrl
                    awinSettings.databaseName = smartSlideLists.slideDBName

                End If
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

                        Dim bigID As Integer = CInt(tmpShape.Tags.Item("BID"))
                        Dim detailID As Integer = CInt(tmpShape.Tags.Item("DID"))
                        If Not (bigID = ptReportBigTypes.components And (detailID = ptReportComponents.prStand Or detailID = ptReportComponents.pfStand)) Then
                            thereIsNoVersionFieldOnSlide = False
                        End If

                        Dim pvName As String = ""
                        If tmpShape.Tags.Item("PNM").Length > 0 Then
                            Dim pName As String = tmpShape.Tags.Item("PNM")
                            Dim vName As String = tmpShape.Tags.Item("VNM")
                            pvName = calcProjektKey(pName, vName)
                        End If
                        ' um zu berücksichtigen, dass auch Slides ohne Meilensteine / Phasen als Smart-Slides aufgefasst werden ...

                        If pvName <> "" Then
                            If smartSlideLists.containsProject(pvName) Then
                                ' nichts tun, ist schon drin ..
                            Else
                                smartSlideLists.addProject(pvName)
                            End If
                        End If

                    End If

                    If tmpShape.Tags.Count > 0 Then
                        If isRelevantMSPHShape(tmpShape) Then

                            bekannteIDs.Add(tmpShape.Id, tmpShape.Name)

                            Call aktualisiereSortedLists(tmpShape)

                            If protectionSolved And tmpShape.Visible = False Then
                                tmpShape.Visible = True
                            End If

                        ElseIf isVISBOChartElement(tmpShape) Then
                            If protectionSolved And tmpShape.Visible = False Then
                                tmpShape.Visible = True
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
                Dim tmpName As String = smartSlideLists.getPVName(i)
                Dim pName As String = getPnameFromKey(tmpName)
                Dim vName As String = getVariantnameFromKey(tmpName)
                Dim pvName As String = calcProjektKeyDB(pName, vName)
                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                Dim tsCollection As Collection = request.retrieveZeitstempelFromDB(pvName)
                smartSlideLists.addToListOfTS(tsCollection)
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
        myPPTWindow = Wn
    End Sub

    Private Sub pptAPP_WindowDeactivate(Pres As PowerPoint.Presentation, Wn As PowerPoint.DocumentWindow) Handles pptAPP.WindowDeactivate

        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If
    End Sub

    Private Sub pptAPP_WindowSelectionChange(Sel As PowerPoint.Selection) Handles pptAPP.WindowSelectionChange

        'Dim relevantShape As PowerPoint.Shape
        Dim arrayOfNames() As String
        Dim relevantShapeNames As New Collection


        selectedPlanShapes = Nothing

        Try
            Dim shpRange As PowerPoint.ShapeRange = Sel.ShapeRange

            If Not IsNothing(shpRange) And slideHasSmartElements Then


                ' es sind ein oder mehrere Shapes selektiert worden 
                Dim i As Integer = 0
                If shpRange.Count = 1 Then

                    ' prüfen, ob inzwischen was selektiert wurde, was nicht zu der Selektion in der 
                    ' Listbox passt 

                    ' prüfen, ob das Info Fenster offen ist und der Search bereich sichtbar - 
                    ' dann muss der Klarheit wegen die Listbox neu aufgebaut werden 
                    If Not IsNothing(infoFrm) And formIsShown Then
                        If infoFrm.rdbName.Visible Then
                            If infoFrm.listboxNames.SelectedItems.Count > 0 Then
                                'Call infoFrm.listboxNames.SelectedItems.Clear()
                            End If
                        End If
                    End If

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
                            If bekannteIDs.ContainsKey(tmpShape.Id) Then

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

                    ' hier muss geprüft werden, ob das Info - Fenster angezeigt wird ... 
                    If IsNothing(infoFrm) And Not formIsShown Then
                        infoFrm = New frmInfo
                        formIsShown = True
                        infoFrm.Show()
                    End If

                    ReDim arrayOfNames(relevantShapeNames.Count - 1)

                    For ix As Integer = 1 To relevantShapeNames.Count
                        arrayOfNames(ix - 1) = CStr(relevantShapeNames(ix))
                    Next

                    selectedPlanShapes = currentSlide.Shapes.Range(arrayOfNames)
                Else
                    ' in diesem Fall wurden nur nicht-relevante Shapes selektiert 
                    Call checkHomeChangeBtnEnablement()
                    If formIsShown Then
                        Call aktualisiereInfoFrm(Nothing)
                    End If
                End If
                '' Ende ...


                If Not IsNothing(selectedPlanShapes) Then

                    Dim tmpShape As PowerPoint.Shape = Nothing
                    Dim elemWasMoved As Boolean = False
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

                    If formIsShown Then
                        Call aktualisiereInfoFrm(tmpShape, elemWasMoved)
                    End If


                    ' jetzt den Window Ausschnitt kontrollieren: ist das oder die selectedPlanShapes überhaupt sichtbar ? 
                    ' wenn nein, dann sicherstellen, dass sie sichtbar werden 
                    Call ensureVisibilityOfSelection(selectedPlanShapes)
                Else

                    Call checkHomeChangeBtnEnablement()
                    If formIsShown Then
                        Call aktualisiereInfoFrm(Nothing)
                    End If

                End If

            End If


        Catch ex As Exception

        End Try

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
            Dim selectionLeft As Single = slideCoordInfo.drawingAreaRight + 1000
            Dim selectionTop As Single = slideCoordInfo.drawingAreaBottom + 1000
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
            selectionTop = selectionTop - markerTol
            selectionWidth = selectionRight - selectionLeft
            selectionHeight = selectionBottom - selectionTop

            With slideCoordInfo
                If selectionLeft >= .drawingAreaLeft And _
                    selectionTop >= .drawingAreaTop - markerTol And _
                    selectionWidth <= .drawingAreaWidth And _
                    selectionHeight <= .drawingAreaBottom - .drawingAreaTop Then
                    ' zulässig ... 

                    pptAPP.ActiveWindow.ScrollIntoView(selectionLeft, selectionTop, _
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
        If formIsShown And Not IsNothing(infoFrm) Then
            With infoFrm
                If .rdbName.Checked Then
                    tmpResult = pptInfoType.cName
                ElseIf .rdbOriginalName.Checked Then
                    tmpResult = pptInfoType.oName
                ElseIf .rdbAbbrev.Checked Then
                    tmpResult = pptInfoType.sName
                ElseIf .rdbBreadcrumb.Checked Then
                    tmpResult = pptInfoType.bCrumb
                ElseIf .rdbLU.Checked Then
                    tmpResult = pptInfoType.lUmfang
                ElseIf .rdbMV.Checked Then
                    tmpResult = pptInfoType.mvElement
                ElseIf .rdbResources.Checked Then
                    tmpResult = pptInfoType.resources
                ElseIf .rdbCosts.Checked Then
                    tmpResult = pptInfoType.costs
                Else
                    tmpResult = pptInfoType.cName
                End If
            End With
        Else
            tmpResult = pptInfoType.cName
        End If

        calcRDB = tmpResult

    End Function

    ''' <summary>
    ''' ruft Formular zum Login auf und holt die RoleDefinitions, CostDefinitions aus der Datenbank 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub logInToMongoDB()
        ' jetzt die Login Maske aufrufen, aber nur wenn nicht schon ein Login erfolgt ist .. ... 

        If noDBAccessInPPT Then
            Dim msg As String
            If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

                ' tk: 17.11.16: Einloggen in Datenbank 
                noDBAccessInPPT = Not loginProzedur()

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
                    RoleDefinitions = request.retrieveRolesFromDB(currentTimestamp)
                    CostDefinitions = request.retrieveCostsFromDB(currentTimestamp)
                End If
            Else
                If englishLanguage Then
                    If englishLanguage Then
                        msg = "no database URL information available ... "
                    Else
                        msg = "keine Datenbank URL verfügbar ... "
                    End If
                    Call MsgBox(msg)
                End If
            End If
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
    Private Sub aktualisiereSortedLists(ByVal tmpShape As PowerPoint.Shape)
        Dim shapeName As String = tmpShape.Name
        Dim checkIT As Boolean = False
        Dim isMilestone As Boolean

        Dim pvName As String = getPVnameFromShpName(tmpShape.Name)
        If pvName <> "" Then
            If smartSlideLists.containsProject(pvName) Then
                ' nichts tun, ist schon drin ..
            Else
                smartSlideLists.addProject(pvName)
            End If
        End If


        If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Or _
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

            ' den classified Name behandeln ...
            Dim tmpName As String = tmpShape.Tags.Item("CN")
            If tmpName.Trim.Length = 0 Then
                Exit Sub
            End If

            Call smartSlideLists.addCN(tmpName, shapeName, isMilestone)

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

            ' AmpelColor behandeln
            Dim ampelColor As Integer = 0
            tmpName = tmpShape.Tags.Item("AC")
            If tmpName.Trim.Length > 0 And pptShapeIsMilestone(tmpShape) Then
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

            ' wurde das Element verschoben ? 
            ' SmartslideLists werden auch gleich mit aktualisiert ... 
            Call checkShpOnManualMovement(tmpShape.Name)

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


        If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Or _
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

                If pvName <> "" Then
                    Dim hproj As clsProjekt = smartSlideLists.getTSProject(pvName, currentTimestamp)
                    Dim phNameID As String = getElemIDFromShpName(tmpShape.Name)
                    Dim cPhase As clsPhase = hproj.getPhaseByID(phNameID)
                    Dim roleInformations As SortedList(Of String, Double) = cPhase.getRoleNamesAndValues
                    Dim costInformations As SortedList(Of String, Double) = cPhase.getCostNamesAndValues

                    Try
                        Call smartSlideLists.addRoleAndCostInfos(roleInformations, _
                                                                 costInformations, _
                                                                 shapeName, _
                                                                 isMilestone)
                    Catch ex As Exception

                    End Try
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
    Friend Sub updateVisboComponent(ByRef pptShape As PowerPoint.Shape, ByVal curTimeStamp As Date, ByVal prevTimeStamp As Date, _
                                    Optional ByVal showOtherVariant As Boolean = False)
        Dim chtObjName As String = ""
        Dim bigType As Integer = -1
        Dim detailID As Integer = -1
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

                            chtObjName = pptChart.Name

                            Dim auswahl As Integer = -1
                            Dim prpfTyp As Integer = -1
                            Dim pName As String = ""
                            Dim vName As String = ""
                            Dim chartTyp As Integer = -1
                            Dim prcTyp As Integer = ptElementTypen.roles
                            Dim ws As xlNS.Worksheet

                            ' der Chart-ObjectName enthält sehr viel ..
                            'pr#ptprdk#projekt-Name/Varianten-Name#Auswahl 
                            Call bestimmeChartInfosFromName(chtObjName, prpfTyp, prcTyp, pName, vName, chartTyp, auswahl)

                            If pName <> "" Then
                                Dim pvName As String = calcProjektKey(pName, vName)

                                ' damit auch eine andere Variante gezeigt werden kann ... 
                                If showOtherVariant Then
                                    Dim tmpPName As String = getPnameFromKey(pvName)
                                    pvName = calcProjektKey(tmpPName, currentVariantname)
                                    vName = currentVariantname
                                End If

                                ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                                Dim tsProj As clsProjekt = smartSlideLists.getTSProject(pvName, curTimeStamp)
                                ' kann eigentlich nicht mehr Nothing werden ... die Liste an TimeStamps enthält den größten auftretenden kleinsten datumswert aller Projekte ....
                                If Not IsNothing(tsProj) Then

                                    '' '' jetzt muss , falls nicht schon geschehen, Excel versteckt in einer neuen Instanz geöffnet werden und das Chart dorthin kopiert werden und wieder 
                                    '' '' zurückgeholt werden; damit wird der Link aufgebrochen 

                                    ' das neue Chart ..
                                    Dim newchtobj As xlNS.ChartObject = Nothing


                                    Try
                                        Call createNewHiddenExcel()

                                        If Not IsNothing(updateWorkbook) Then

                                            ws = CType(updateWorkbook.Worksheets.Item(1), xlNS.Worksheet)
                                            ' das Workbook wird aktiviert ... 

                                            ' dann muss das Shape in Excel kopiert werden 
                                            pptShape.Copy()
                                            ws.Paste()
                                            Dim anzCharts As Integer = CType(ws.ChartObjects, Excel.ChartObjects).Count

                                            If anzCharts > 0 Then
                                                newchtobj = CType(ws.ChartObjects(anzCharts), Excel.ChartObject)

                                                If Not IsNothing(newchtobj) Then

                                                    ' jetzt muss das chtobj aktualisiert werden ... 
                                                    Try

                                                        If prpfTyp = ptPRPFType.project Then

                                                            If chartTyp = PTprdk.PersonalBalken Or chartTyp = PTprdk.KostenBalken Then
                                                                Call updatePPTBalkenOfProject(tsProj, newchtobj, prcTyp, auswahl)

                                                            ElseIf chartTyp = PTprdk.PersonalPie Or chartTyp = PTprdk.KostenPie Then
                                                                ' Aktualisieren der Personal- bzw. Kosten-Pies ...

                                                            ElseIf chartTyp = PTprdk.Ergebnis Then
                                                                ' Aktualisieren des Ergebnis Charts 
                                                                Call updatePPTProjektErgebnis(tsProj, newchtobj)

                                                            ElseIf chartTyp = PTprdk.StrategieRisiko Or _
                                                                chartTyp = PTprdk.ZeitRisiko Or _
                                                                chartTyp = PTprdk.FitRisikoVol Or _
                                                                chartTyp = PTprdk.ComplexRisiko Then
                                                                ' Aktualisieren der Strategie-Charts

                                                                Call updatePPTProjectPfDiagram(tsProj, newchtobj, chartTyp, 0)

                                                            End If

                                                        ElseIf prpfTyp = ptPRPFType.portfolio Then

                                                        End If

                                                    Catch ex As Exception

                                                    End Try


                                                End If

                                                ' jetzt wird das aktualisierte Excel-Chart kopiert
                                                newchtobj.Copy()

                                                ' dann muss das Excel-Shape wieder zurück in PPT kopiert werden 
                                                Dim newShapeRange As PowerPoint.ShapeRange = currentSlide.Shapes.Paste()
                                                Dim newPPTShape As PowerPoint.Shape = newShapeRange.Item(1)

                                                ' dann mus das Powerpoint Shape aktualisiert werden ...
                                                With newPPTShape
                                                    .Top = pptShape.Top
                                                    .Left = pptShape.Left
                                                    .Height = pptShape.Height
                                                    .Width = pptShape.Width
                                                    .Name = pptShape.Name
                                                    .Tags.Add("CHON", pptShape.Tags("CHON"))
                                                    .Tags.Add("PNM", pptShape.Tags("PNM"))
                                                    If showOtherVariant Then
                                                        .Tags.Add("VNM", vName)
                                                    Else
                                                        .Tags.Add("VNM", pptShape.Tags("VNM"))
                                                    End If

                                                    .Tags.Add("CHT", pptShape.Tags("CHT"))
                                                    .Tags.Add("ASW", pptShape.Tags("ASW"))
                                                    .Tags.Add("COL", pptShape.Tags("COL"))
                                                    .Tags.Add("UPDT", "TRUE")
                                                    .Tags.Add("BID", pptShape.Tags("BID"))
                                                    .Tags.Add("DID", pptShape.Tags("DID"))
                                                    .Tags.Add("Q1", pptShape.Tags("Q1"))
                                                    .Tags.Add("Q2", pptShape.Tags("Q2"))
                                                End With

                                                ' das Original Shape wird gelöscht und das neue tritt an seine Stelle ... 
                                                ' sowohl newChtobj als auch das late Powerpoint Shape ... 
                                                If Not IsNothing(newchtobj) Then
                                                    newchtobj.Delete()
                                                End If
                                                pptShape.Delete()


                                            End If
                                        Else

                                        End If
                                    Catch ex As Exception

                                    End Try

                                End If

                            End If

                        End If

                    End If




                ElseIf bigType = ptReportBigTypes.components Or _
                       bigType = ptReportBigTypes.tables Then

                    Dim pName As String = pptShape.Tags.Item("PNM")
                    Dim vName As String = pptShape.Tags.Item("VNM")

                    If showOtherVariant Then
                        vName = currentVariantname
                        If pptShape.Tags.Item("VNM").Length > 0 Then
                            pptShape.Tags.Delete("VNM")
                        End If
                        pptShape.Tags.Add("VNM", vName)
                        Dim chck As String = pptShape.Tags.Item("VNM")
                    End If

                    If pName <> "" Then
                        Dim pvName As String = calcProjektKey(pName, vName)

                        ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                        Dim tsProj As clsProjekt = smartSlideLists.getTSProject(pvName, curTimeStamp)

                        If Not IsNothing(tsProj) Then

                            If bigType = ptReportBigTypes.components Then
                                Call updatePPTComponent(tsProj, pptShape, detailID)

                            ElseIf bigType = ptReportBigTypes.tables Then

                                If detailID = ptReportTables.prMilestones Then
                                    Call updatePPTProjektTabelleZiele(pptShape, tsProj)
                                End If

                            End If

                        End If


                    Else
                        ' kein zu aktualisierendes Shape ... 
                    End If


                End If
            End If

        Catch ex As Exception
            Dim a As Integer = 1
        End Try
    End Sub
    ''' <summary>
    ''' erzeugt eine verborgene Excel-Instanz, die verwendet werden kann, um PPT charts hin und her zu kopieren und damit die Referenz zu löschen, 
    ''' die verhindert, dass ein PPT Chart geupdated werden kann;
    ''' wenn das HiddenExcel bereits existiert wird nichts gemacht ... 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub createNewHiddenExcel()

        If IsNothing(updateWorkbook) Then
            ' es wird auf jeden Fall eine neue, verborgene Excel-Instanz aufgemacht 
            ' die wird dann beim Schliessen einer Presentation wieder beendet bzw. zugemacht 
            Try
                xlApp = CreateObject("Excel.Application")
                xlApp.Visible = False
                xlApp.ScreenUpdating = False
                '' prüft, ob bereits Powerpoint geöffnet ist 
                'xlApp = GetObject(, "Excel.Application")
            Catch ex As Exception
                xlApp = Nothing
                updateWorkbook = Nothing
                Exit Sub
            End Try

            If My.Computer.FileSystem.FileExists("visboupdate.xlsx") Then
                ' öffnen
                xlApp.Workbooks.Open("visboupdate.xlsx")
            Else
                xlApp.Workbooks.Add()
                xlApp.ActiveWorkbook.SaveAs("visboupdate.xlsx")
            End If
            updateWorkbook = xlApp.ActiveWorkbook
        Else
            ' existiert schon, also existiert auch xlApp bereits ...
        End If

    End Sub
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

            If cmtType = pptAnnotationType.ampelText Or _
                cmtType = pptAnnotationType.lieferumfang Then

                If Not IsNothing(timestamp) Then
                    ' der Text und die Farbe müssen von einem TimeStamp Projekt kommen 

                    ' überprüfe, ob es zu dem angegebenen Shape bereits ein TS Projekt gibt 
                    Dim pvName As String = getPVnameFromShpName(cmtShape.Name)

                    ' damit auch eine andere Variante gezeigt werden kann ... 
                    Dim tmpPName As String = getPnameFromKey(pvName)
                    If showOtherVariant Then
                        pvName = calcProjektKey(tmpPName, currentVariantname)
                    End If

                    If pvName <> "" Then
                        ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                        Dim tsProj As clsProjekt = smartSlideLists.getTSProject(pvName, timestamp)

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
                                        cmtShape.Visible = False
                                    Else
                                        If Not cmtShape.Visible Then
                                            cmtShape.Visible = True
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
                                            newCmtText = tmpText & ms.getAllDeliverables
                                            newCmtColor = ms.getBewertung(1).colorIndex
                                        End If

                                    End If


                                ElseIf pptShapeIsPhase(refShape) Then

                                    Dim ph As clsPhase = tsProj.getPhase(name:=elemName, breadcrumb:=elemBC)
                                    If IsNothing(ph) Then
                                        cmtShape.Visible = False
                                    Else
                                        If Not cmtShape.Visible Then
                                            cmtShape.Visible = True
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
                                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbGrey
                                    Else
                                        .Shadow.ForeColor.RGB = trafficLightColors(newCmtColor)
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
    ''' bewegt alle Shapes an 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub moveAllShapes(Optional ByVal showOtherVariant As Boolean = False)

        Dim namesToBeRenamed As New Collection
        Dim ix As Integer = 0

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
                    ix = ix + 1

                    If isRelevantMSPHShape(tmpShape) Then


                        If showOtherVariant Then
                            ' wenn es eine Variante gibt, wird currentTimeStamp dort auf den entsprechenden Wert der Variante gelegt 
                            namesToBeRenamed.Add(tmpShape.Name)
                            Call sendToNewPosition(tmpShape.Name, Date.Now, diffMvList, showOtherVariant)
                        Else
                            Call sendToNewPosition(tmpShape.Name, currentTimestamp, diffMvList, showOtherVariant)
                        End If

                  
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
        ' bei de rchart Behandlung Charts gelöscht udn kopiert werden 
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


        For Each tmpShpName As String In bigToDoList

            Try
                Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                If Not IsNothing(tmpShape) Then
                    If isAnnotationShape(tmpShape) Then

                        If tmpShape.Name.Substring(tmpShape.Name.Length - 1, 1) = pptAnnotationType.text Then

                            namesToBeRenamed.Add(tmpShape.Name)
                            ' es handelt sich um den Text, also nur verschieben 
                            Dim refName As String = tmpShape.Name.Substring(0, tmpShape.Name.Length - 1)

                            If diffMvList.ContainsKey(refName) Then
                                Dim diff As Double = diffMvList.Item(refName)
                                With tmpShape
                                    .Left = .Left + diff
                                End With
                            End If


                        ElseIf tmpShape.Name.Substring(tmpShape.Name.Length - 1, 1) = pptAnnotationType.datum Then

                            namesToBeRenamed.Add(tmpShape.Name)
                            ' es handelt sich um das Datum, also verschieben und Text ändern 
                            Dim refName As String = tmpShape.Name.Substring(0, tmpShape.Name.Length - 1)
                            Dim refShape As PowerPoint.Shape = currentSlide.Shapes.Item(refName)
                            Dim tmpShort As Boolean = (tmpShape.TextFrame2.TextRange.Text.Length < 8)
                            Dim descriptionText As String = bestimmeElemDateText(refShape, tmpShort)

                            If diffMvList.ContainsKey(refName) Then
                                Dim diff As Double = diffMvList.Item(refName)
                                With tmpShape
                                    .Left = .Left + diff
                                    .TextFrame2.TextRange.Text = descriptionText
                                End With
                            End If

                        End If

                    End If
                End If

            Catch ex As Exception
                Call MsgBox("Fehler : " & ex.Message)
            End Try

        Next

        ' und schließlich muss noch nachgesehen werden, ob es eine todayLine gibt 
        Try
            Dim todayLineShape As PowerPoint.Shape = currentSlide.Shapes.Item("todayLine")
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

        ' ur: 03.07.2017: setze alle Ampelfarben
        Call faerbeShapes(PTfarbe.none, showTrafficLights(PTfarbe.none))
        Call faerbeShapes(PTfarbe.green, showTrafficLights(PTfarbe.green))
        Call faerbeShapes(PTfarbe.yellow, showTrafficLights(PTfarbe.yellow))
        Call faerbeShapes(PTfarbe.red, showTrafficLights(PTfarbe.red))


    End Sub

    ''' <summary>
    ''' aktualisiert das Shape mit den Daten aus dem entsprechenden TimeStamp Projekt; 
    ''' es wird keine Aktion mit MV gemacht, das ist manually moved Information
    ''' wenn es aufgerufen wird mit ShowOtherVariant werden die Werte der anderen Variante gezeigt, sonst einfach der andere TimeStamp derselben Projekt-Variante 
    ''' </summary>
    ''' <param name="tmpShapeName"></param>
    ''' <remarks></remarks>
    Friend Sub sendToNewPosition(ByVal tmpShapeName As String, ByVal timestamp As Date, ByRef diffMvList As SortedList(Of String, Double), _
                                       ByVal showOtherVariant As Boolean)

        Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShapeName)

        If Not IsNothing(tmpShape) Then
            ' Voraussetzung: es handelt sich um ein relevantes Shapes, also einen Meilenstein, eine Phase, einen Swimlane- oder Segment Bezeichner ... eine Phase oder einen Meilenstein ... 

            Dim pvName As String = getPVnameFromShpName(tmpShape.Name)

            ' damit auch eine andere Variante gezeigt werden kann ... 
            If showOtherVariant Then
                Dim tmpPName As String = getPnameFromKey(pvName)
                pvName = calcProjektKey(tmpPName, currentVariantname)
            End If

            If pvName <> "" Then
                ' wenn das noch nicht existiert, wird es aus der DB geholt und angelegt  ... 
                Dim tsProj As clsProjekt = smartSlideLists.getTSProject(pvName, timestamp)
                ' kann eigentlich nicht mehr Nothing werden ... die Liste an TimeStamps enthält den größten auftretenden kleinsten datumswert aller Projekte ....
                If Not IsNothing(tsProj) Then
                    Dim elemName As String = tmpShape.Tags.Item("CN")
                    Dim elemBC As String = tmpShape.Tags.Item("BC")

                    If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                        ' Swimlane Name oder Segment Name: kein Verschieben , aber das das Setzen der Tags ist notwendig  
                        '
                        Dim ph As clsPhase = tsProj.getPhase(name:=elemName, breadcrumb:=elemBC)
                        If IsNothing(ph) Then
                            tmpShape.Visible = False
                        Else

                            If Not tmpShape.Visible Then
                                tmpShape.Visible = True
                            End If

                            Dim bsn As String = tmpShape.Tags.Item("BSN")
                            Dim bln As String = tmpShape.Tags.Item("BLN")
                            ' jetzt müssen die Tags-Informationen des Meilensteines gesetzt werden 
                            Call addSmartPPTShapeInfo(tmpShape, elemBC, elemName, ph.shortName, ph.originalName, bsn, bln, _
                                                      ph.getStartDate, ph.getEndDate, ph.getBewertung(1).colorIndex, ph.getBewertung(1).description, _
                                                      Nothing)
                        End If


                    Else

                        If pptShapeIsMilestone(tmpShape) Then

                            'Call resetMVInfo(tmpShape)

                            Dim ms As clsMeilenstein = tsProj.getMilestone(msName:=elemName, breadcrumb:=elemBC)
                            If IsNothing(ms) Then
                                tmpShape.Visible = False
                            Else

                                If Not tmpShape.Visible Then
                                    tmpShape.Visible = True
                                End If

                                Dim mvDiff As Double = mvMilestoneToTimestampPosition(tmpShape, ms.getDate, timestamp)
                                If Not diffMvList.ContainsKey(tmpShape.Name) And mvDiff * mvDiff > 0.01 Then
                                    diffMvList.Add(tmpShape.Name, mvDiff)
                                End If
                                '
                                'ur:3.7.2017: soll nun nach MoveAllShapes erfolgen für alle Elemente gemäß gemerkten ShowTrafficligths
                                ' jetzt muss ggf die Farbe gesetzt werden 
                                ''Dim ampelFarbe As Integer = ms.getBewertung(1).colorIndex
                                ''Call faerbeShape(tmpShape, ampelFarbe, showTrafficLights(ampelFarbe))

                                Dim bsn As String = tmpShape.Tags.Item("BSN")
                                Dim bln As String = tmpShape.Tags.Item("BLN")
                                ' jetzt müssen die Tags-Informationen des Meilensteines gesetzt werden 
                                Call addSmartPPTShapeInfo(tmpShape, elemBC, elemName, ms.shortName, ms.originalName, bsn, bln, Nothing, _
                                                          ms.getDate, ms.getBewertung(1).colorIndex, ms.getBewertung(1).description, _
                                                          ms.getAllDeliverables("#"))
                            End If



                        ElseIf pptShapeIsPhase(tmpShape) Then

                            'Call resetMVInfo(tmpShape)

                            Dim ph As clsPhase = tsProj.getPhase(name:=elemName, breadcrumb:=elemBC)
                            If IsNothing(ph) Then
                                tmpShape.Visible = False
                            Else
                                If Not tmpShape.Visible Then
                                    tmpShape.Visible = True
                                End If

                                Dim mvDiff As Double = mvPhaseToTimestampPosition(tmpShape, ph.getStartDate, ph.getEndDate, timestamp)
                                If Not diffMvList.ContainsKey(tmpShape.Name) And mvDiff * mvDiff > 0.01 Then
                                    diffMvList.Add(tmpShape.Name, mvDiff)
                                End If
                                '
                                'ur:3.7.2017: soll nun nach MoveAllShapes erfolgen für alle Elemente gemäß gemerkten ShowTrafficligths
                                ' '' jetzt muss ggf die Farbe gesetzt werden 
                                ''Dim ampelFarbe As Integer = ph.getBewertung(1).colorIndex
                                ''Call faerbeShape(tmpShape, ampelFarbe, showTrafficLights(ampelFarbe))

                                Dim bsn As String = tmpShape.Tags.Item("BSN")
                                Dim bln As String = tmpShape.Tags.Item("BLN")
                                ' jetzt müssen die Tags-Informationen des Meilensteines gesetzt werden 
                                Call addSmartPPTShapeInfo(tmpShape, elemBC, elemName, ph.shortName, ph.originalName, bsn, bln, ph.getStartDate, _
                                                             ph.getEndDate, ph.getBewertung(1).colorIndex, ph.getBewertung(1).description, _
                                                             Nothing)
                            End If

                        End If

                    End If
                End If

            End If


        End If



    End Sub

    ''' <summary>
    ''' schreibt bzw. aktualisiert den Time-Stamp auf die Folie ... 
    ''' </summary>
    ''' <param name="currentTimestamp"></param>
    ''' <remarks></remarks>
    Friend Sub showTSMessage(ByVal currentTimestamp As Date)

        Dim tsMsgBox As PowerPoint.Shape
        Try
            tsMsgBox = currentSlide.Shapes.Item("TimeStampInfo")
        Catch ex As Exception
            tsMsgBox = Nothing
        End Try

        If IsNothing(tsMsgBox) Then
            ' erstellen ...
            tsMsgBox = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, _
                                      200, 5, 70, 20)
            With tsMsgBox
                .TextFrame2.TextRange.Text = "Stand: " & currentTimestamp.ToString
                .TextFrame2.TextRange.Font.Size = CDbl(schriftGroesse + 6)
                .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = trafficLightColors(3)
                .TextFrame2.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue
                .TextFrame2.MarginBottom = 0
                .TextFrame2.MarginLeft = 0
                .Rotation = 0
                .TextFrame2.MarginRight = 0
                .TextFrame2.MarginTop = 0
                .Name = "TimeStampInfo"
                .TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse
            End With
        Else
            With tsMsgBox
                If englishLanguage Then
                    .TextFrame2.TextRange.Text = "Version: " & currentTimestamp.ToString
                Else
                    .TextFrame2.TextRange.Text = "Stand: " & currentTimestamp.ToString
                End If

            End With
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
    ''' diese MEthode verschiebt nur das Shape; es erfolgt keinerlei Setzen von Tag-Information
    ''' auch eine HomeButtonRelevance besetht nicht mehr; das neue Home ist mit dem TimeStamp erreicht 
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="tsStartdate"></param>
    ''' <param name="tsEndDate"></param>
    ''' <param name="timeStamp"></param>
    ''' <remarks></remarks>
    Friend Function mvPhaseToTimestampPosition(ByRef tmpShape As PowerPoint.Shape, ByVal tsStartdate As Date, ByVal tsEndDate As Date, _
                                              ByVal timeStamp As Date) As Double

        Dim x1Pos As Double, x2Pos As Double
        Dim diff As Double = 0.0
        Dim oldLeft As Double
        'Dim expla As String = "Version: " & timeStamp.ToShortDateString

        ' wenn der Phasen start oder das Phasen-Ende vor bzw. hinter dem pptStart bzw. EndOfCalendar liegt ...
        If DateDiff(DateInterval.Day, tsStartdate, slideCoordInfo.PPTStartOFCalendar) > 0 Then
            tsStartdate = slideCoordInfo.PPTStartOFCalendar
        End If

        If DateDiff(DateInterval.Day, slideCoordInfo.PPTEndOFCalendar, tsEndDate) > 0 Then
            tsEndDate = slideCoordInfo.PPTEndOFCalendar
        End If


        If tsStartdate <> slideCoordInfo.calcXtoDate(tmpShape.Left) Or _
            tsEndDate <> slideCoordInfo.calcXtoDate(tmpShape.Left + tmpShape.Width) Then
            ' es hat sich was geändert ... 

            'homeButtonRelevance = True
            Call slideCoordInfo.calculatePPTx1x2(tsStartdate, tsEndDate, x1Pos, x2Pos)

            With tmpShape
                oldLeft = .Left

                .Left = x1Pos
                .Width = x2Pos - tmpShape.Left

                With .Glow
                    .Radius = 5
                    .Color.RGB = changeColor
                End With

                diff = .Left - oldLeft
                'Dim mvdString As String = tsStartdate.ToString & "#" & tsEndDate.ToString
                '.Tags.Add("MVD", mvdString)
                '.Tags.Add("MVE", expla)

            End With
        Else
            With tmpShape.Glow
                .Radius = 0
                '.Color.RGB = .Color.RGB = PowerPoint.XlRgbColor.rgbWhite
            End With
        End If

        mvPhaseToTimestampPosition = diff
    End Function

    ''' <summary>
    ''' diese MEthode verschiebt nur das Shape; es erfolgt keinerlei Setzen von Tag-Information
    ''' auch eine HomeButtonRelevance besetht nicht mehr; das neue Home ist mit dem TimeStamp erreicht  
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="msDate"></param>
    ''' <param name="timeStamp"></param>
    ''' <remarks></remarks>
    Friend Function mvMilestoneToTimestampPosition(ByRef tmpShape As PowerPoint.Shape, ByVal msDate As Date, ByVal timeStamp As Date) As Double
        Dim x1Pos As Double, x2Pos As Double
        Dim diff As Double = 0.0
        Dim oldLeft As Double = 0.0
        'Dim expla As String = "Version: " & timeStamp.ToShortDateString

        If msDate <> slideCoordInfo.calcXtoDate(tmpShape.Left + tmpShape.Width / 2) Then
            ' es hat sich was geändert ... 
            'homeButtonRelevance = True

            Call slideCoordInfo.calculatePPTx1x2(msDate, msDate, x1Pos, x2Pos)

            ' jetzt die Shape-Info 
            With tmpShape
                oldLeft = .Left
                .Left = x1Pos - tmpShape.Width / 2
                diff = .Left - oldLeft

                With .Glow
                    .Radius = 5
                    .Color.RGB = changeColor
                End With

            End With
        Else
            With tmpShape.Glow
                .Radius = 0
                '.Color.RGB = PowerPoint.XlRgbColor.rgbWhite

            End With
        End If

        mvMilestoneToTimestampPosition = diff

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
                .btnSendToHome.Enabled = homeButtonRelevance
                .btnSentToChange.Enabled = changedButtonRelevance
            End With

            If Not IsNothing(tmpShape) Then

                If Not IsNothing(selectedPlanShapes) Then

                    If selectedPlanShapes.Count = 1 Then

                        With infoFrm

                            Call .setDTPicture(pptShapeIsMilestone(tmpShape))

                            .elemName.Text = bestimmeElemText(tmpShape, .showAbbrev.Checked, .showOrginalName.Checked)
                            If showBreadCrumbField Then
                                .fullBreadCrumb.Text = bestimmeElemBC(tmpShape)
                            End If
                            .elemDate.Text = bestimmeElemDateText(tmpShape, False)

                            Dim rdbCode As Integer = calcRDB()

                            Dim tmpStr() As String
                            tmpStr = bestimmeElemALuTvText(tmpShape, rdbCode).Split(New Char() {CType(vbLf, Char), CType(vbCr, Char)})
                            .aLuTvText.Lines = tmpStr

                            ' Änderungen bei Datum und Erläuterung erlauben 
                            If isMovedShape Then
                                .elemDate.Enabled = True
                                If .rdbMV.Checked Then
                                    .aLuTvText.Enabled = True
                                Else
                                    .aLuTvText.Enabled = False
                                End If
                            Else
                                .elemDate.Enabled = False
                                .aLuTvText.Enabled = False
                            End If


                        End With
                    ElseIf selectedPlanShapes.Count > 1 Then

                        Dim rdbCode As Integer = calcRDB()

                        With infoFrm

                            Call .setDTPicture(Nothing)

                            If .elemName.Text <> bestimmeElemText(tmpShape, .showAbbrev.Checked, .showOrginalName.Checked) Then
                                .elemName.Text = " ... "
                            End If
                            If .elemDate.Text <> bestimmeElemDateText(tmpShape, False) Then
                                .elemDate.Text = " ... "
                            End If

                            .aLuTvText.Text = " ... "

                            .aLuTvText.Enabled = False
                            .elemDate.Enabled = False


                        End With

                    End If
                Else
                    ' Info Formular Inhalte zurücksetzen ... 
                    With infoFrm
                        .elemName.Text = ""
                        .fullBreadCrumb.Text = ""
                        .elemDate.Text = ""
                        .aLuTvText.Text = ""
                    End With

                End If

            Else
                ' es wurde eine Selektion aufgehoben ..
                ' erstmal nichts tun .. 
                ' Info Formular Inhalte zurücksetzen ... 
                With infoFrm
                    .elemName.Text = ""
                    .fullBreadCrumb.Text = ""
                    .elemDate.Text = ""
                    .aLuTvText.Text = ""
                End With

            End If

        End If
    End Sub

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

        If shapeName.EndsWith(CStr(pptAnnotationType.ampelText)) Or _
            shapeName.EndsWith(CStr(pptAnnotationType.lieferumfang)) Or _
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
    Friend Sub createMarkerShapes(Optional ByVal pptShape As PowerPoint.Shape = Nothing, _
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
                        newLeft = .Left + 0.5 * (tmpShape.Width - newWidth)
                        newTop = .Top - (newHeight + 2)
                    End With

                    Dim markerShape As PowerPoint.Shape = _
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

            If markerShpNames.Count > 1 Or _
                (markerShpNames.Count = 1 And exceptShpName.Length > 0 And _
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
    Public Function bestimmeElemALuTvText(ByVal curShape As PowerPoint.Shape, _
                                          Optional ByVal type As Integer = pptInfoType.aExpl, _
                                          Optional ByVal shortForm As Boolean = True) As String

        Dim tmpText As String = ""

        If Not shortForm Then
            tmpText = curShape.Tags.Item("CN") & " "
        End If

        Try

            If type = pptInfoType.lUmfang Then
                ' bestimmen der ersten Zeile:
                If englishLanguage Then
                    tmpText = tmpText & "Deliverables:" & vbLf
                Else
                    tmpText = tmpText & "Lieferumfänge:" & vbLf
                End If


                Dim tmpStr() As String
                If curShape.Tags.Item("LU").Length > 0 Then
                    tmpStr = curShape.Tags.Item("LU").Split(New Char() {CType("#", Char)})
                    For i As Integer = 0 To tmpStr.Length - 1
                        tmpText = tmpText & tmpStr(i) & vbLf
                    Next
                End If

            ElseIf type = pptInfoType.mvElement Then
                If englishLanguage Then
                    tmpText = tmpText & "moved:" & vbLf
                Else
                    tmpText = tmpText & "verschoben:" & vbLf
                End If

                If curShape.Tags.Item("MVE").Length > 0 Then
                    tmpText = tmpText & curShape.Tags.Item("MVE")
                End If

            ElseIf type = pptInfoType.resources Or type = pptInfoType.costs Then
                If Not noDBAccessInPPT And pptShapeIsPhase(curShape) Then
                    Try
                        Dim pvName As String = getPVnameFromShpName(curShape.Name)
                        Dim hproj As clsProjekt = smartSlideLists.getTSProject(pvName, currentTimestamp)
                        Dim phNameID As String = getElemIDFromShpName(curShape.Name)
                        Dim cPhase As clsPhase = hproj.getPhaseByID(phNameID)
                        Dim roleInformations As SortedList(Of String, Double) = cPhase.getRoleNamesAndValues
                        Dim costInformations As SortedList(Of String, Double) = cPhase.getCostNamesAndValues

                        If Not shortForm Then

                            If englishLanguage Then
                                'tmpText = getElemNameFromShpName(curShape.Name) & " Resource/Costs :" & vbLf
                                tmpText = tmpText & "Resource/Costs :" & vbLf
                            Else
                                'tmpText = getElemNameFromShpName(curShape.Name) & " Ressourcen/Kosten:" & vbLf
                                tmpText = tmpText & "Ressourcen/Kosten:" & vbLf
                            End If

                        Else
                            If englishLanguage Then
                                tmpText = "Resource/Costs :" & vbLf
                            Else
                                tmpText = "Ressourcen/Kosten:" & vbLf
                            End If
                        End If


                        Dim unit As String
                        If englishLanguage Then
                            unit = " PD"
                        Else
                            unit = " PT"
                        End If

                        For i As Integer = 1 To roleInformations.Count
                            tmpText = tmpText & _
                                roleInformations.ElementAt(i - 1).Key & ": " & CInt(roleInformations.ElementAt(i - 1).Value).ToString & unit & vbLf
                        Next

                        If costInformations.Count > 0 And roleInformations.Count > 0 Then
                            tmpText = tmpText & vbLf
                        End If

                        unit = " TE"
                        For i As Integer = 1 To costInformations.Count
                            tmpText = tmpText & _
                                costInformations.ElementAt(i - 1).Key & ": " & CInt(costInformations.ElementAt(i - 1).Value).ToString & unit & vbLf
                        Next

                    Catch ex As Exception
                        tmpText = "Phase " & getElemNameFromShpName(curShape.Name)
                    End Try



                ElseIf noDBAccessInPPT And pptShapeIsPhase(curShape) Then
                    If Not shortForm Then
                        If englishLanguage Then
                            tmpText = "Resource/Costs " & getElemNameFromShpName(curShape.Name) & ":" & vbLf & _
                            "no DB access ..."
                        Else
                            tmpText = "Ressourcen / Kosten " & getElemNameFromShpName(curShape.Name) & ":" & vbLf & _
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
                        tmpText = tmpText & "traffic light text:" & vbLf
                    Else
                        tmpText = tmpText & "Ampel-Text:" & vbLf
                    End If

                    tmpText = tmpText & curShape.Tags.Item("AE")
                End If
            End If

        Catch ex As Exception

        End Try

        bestimmeElemALuTvText = tmpText

    End Function

    ''' <summary>
    ''' bestimmt den Text in Abhängigkeit, ob classified name, ShortName oder OriginalName gezeigt werden soll 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <param name="showShortName"></param>
    ''' <param name="showOriginalName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemText(ByVal curShape As PowerPoint.Shape, _
                                          ByVal showShortName As Boolean, ByVal showOriginalName As Boolean) As String

        Dim tmpText As String = ""
        Dim translationNecessary As Boolean = False
        ' im Falle eines Namens , der öfter vorkommt und zu Zwecken der Eindeutigkeit durch den Bestname erstezt werden muss 
        Dim isCombinedName As Boolean = False
        Dim elemName As String = ""
        Dim bestShortName As String = curShape.Tags.Item("BSN")
        Dim bestLongName As String = curShape.Tags.Item("BLN")

        If isRelevantMSPHShape(curShape) Then
            If showOriginalName Then
                If curShape.Tags.Item("ON").Length = 0 Then
                    tmpText = curShape.Tags.Item("CN")
                    translationNecessary = (selectedLanguage <> defaultSprache)
                Else
                    tmpText = curShape.Tags.Item("ON")
                End If

            ElseIf showShortName Then
                If curShape.Tags.Item("SN").Length = 0 Then
                    If curShape.Tags.Item("CN").Length > 0 Then
                        tmpText = curShape.Tags.Item("CN")
                        translationNecessary = (selectedLanguage <> defaultSprache)
                    End If
                Else
                    tmpText = curShape.Tags.Item("SN")
                    If bestShortName.Length > 0 And tmpText <> bestShortName Then
                        tmpText = bestShortName
                    End If

                End If

            ElseIf curShape.Tags.Item("CN").Length > 0 Then
                tmpText = curShape.Tags.Item("CN")

                If bestLongName.Length > 0 And bestLongName <> tmpText Then
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
    ''' bestimmt den Datums-String, für einen MEilenstein nur das Ende-Datum; 
    ''' 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemDateText(ByVal curShape As PowerPoint.Shape, ByVal showShort As Boolean) As String

        Dim tmpText As String = ""

        If pptShapeIsMilestone(curShape) Then

            Dim msDate As Date = slideCoordInfo.calcXtoDate(curShape.Left + 0.5 * curShape.Width)
            If Not showShort Then
                'tmpText = msDate.ToString("d")
                tmpText = msDate.ToShortDateString
            Else
                tmpText = msDate.Day.ToString & "." & msDate.Month.ToString
            End If

            'Dim tstDate As Date = CDate(curShape.Tags.Item("ED"))
            'If DateDiff(DateInterval.Day, msDate, tstDate) <> 0 Then
            '    tmpText = tmpText & " (M)"
            'End If

        ElseIf pptShapeIsPhase(curShape) Then

            Dim startDate As Date = slideCoordInfo.calcXtoDate(curShape.Left)
            Dim endDate As Date = slideCoordInfo.calcXtoDate(curShape.Left + curShape.Width)

            If Not showShort Then
                'tmpText = startDate.ToString("d") & "-" & endDate.ToString("d")
                tmpText = startDate.ToShortDateString & "-" & endDate.ToShortDateString
            Else
                Try

                    tmpText = startDate.Day.ToString & "." & startDate.Month.ToString & "-" & _
                                endDate.Day.ToString & "." & endDate.Month.ToString
                Catch ex As Exception
                    tmpText = curShape.Tags.Item("SD") & "-" & curShape.Tags.Item("ED")
                End Try

            End If

            'Dim tstDate1 As Date = CDate(curShape.Tags.Item("SD"))
            'Dim tstDate2 As Date = CDate(curShape.Tags.Item("ED"))

            'If DateDiff(DateInterval.Day, startDate, tstDate1) <> 0 Or _
            '    DateDiff(DateInterval.Day, startDate, tstDate1) <> 0 Then
            '    tmpText = tmpText & " (M)"
            'End If


        End If


        bestimmeElemDateText = tmpText
    End Function


    ''' <summary>
    ''' gibt an, ob ein Element manuell verändert wurde oder nicht 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isMovedElement(ByVal curShape As PowerPoint.Shape, _
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

                Dim diffDays As Integer = DateDiff(DateInterval.Day, msDate, tstDate)

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

                Dim diffSD As Integer = DateDiff(DateInterval.Day, pptSDate, planSDate)
                Dim diffED As Integer = DateDiff(DateInterval.Day, pptEDate, planEDate)


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
            isOtherVisboComponent = (curShape.Tags.Item("CHON").Length > 0) Or _
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
        If isRelevantMSPHShape(curShape) Or isCommentShape(curShape) Or isOtherVisboComponent(curShape) Then
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


        isRelevantForProtection = isVisboShape(curShape)
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
    Public Sub annotatePlanShape(ByVal selectedPlanShape As PowerPoint.Shape, _
                                  ByVal descriptionType As Integer, ByVal positionIndex As Integer)

        Dim newShape As PowerPoint.Shape
        Dim txtShpLeft As Double = selectedPlanShape.Left - 4
        Dim txtShpTop As Double = selectedPlanShape.Top - 5
        Dim txtShpWidth As Double = 5
        Dim txtShpHeight As Double = 5
        Dim normalFarbe As Integer = RGB(10, 10, 10)
        Dim ampelFarbe As Integer = 0

        Dim descriptionText As String = ""

        Dim shapeName As String = ""
        Dim ok As Boolean = False

        ' bestimme den Info Type ..
        ' handelt es sich um den Lang-/Kurz-Namen oder um das Datum ? 

        If descriptionType = pptAnnotationType.text Then
            descriptionText = bestimmeElemText(selectedPlanShape, showShortName, showOrigName)

        ElseIf descriptionType = pptAnnotationType.datum Then
            descriptionText = bestimmeElemDateText(selectedPlanShape, showShortName)

        ElseIf descriptionType = pptAnnotationType.ampelText Or _
                descriptionType = pptAnnotationType.lieferumfang Or _
                descriptionType = pptAnnotationType.movedExplanation Then

            If IsNumeric(selectedPlanShape.Tags.Item("AC")) Then
                ampelFarbe = CInt(selectedPlanShape.Tags.Item("AC"))
            End If

            If descriptionType = pptAnnotationType.movedExplanation Then
                descriptionText = bestimmeElemALuTvText(selectedPlanShape, pptInfoType.mvElement, False)
                ampelFarbe = 4

            ElseIf descriptionType = pptAnnotationType.lieferumfang Then
                descriptionText = bestimmeElemALuTvText(selectedPlanShape, pptInfoType.lUmfang, False)
            Else

                descriptionText = bestimmeElemALuTvText(selectedPlanShape, pptInfoType.aExpl, False)
            End If

            txtShpLeft = selectedPlanShape.Left + 1.5 * selectedPlanShape.Width + 5
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
            If descriptionType = pptAnnotationType.ampelText Or _
                    descriptionType = pptAnnotationType.movedExplanation Or _
                    descriptionType = pptAnnotationType.lieferumfang Or _
                    descriptionType = pptAnnotationType.resourceCost Then
                newShape.Delete()
                newShape = Nothing
            End If
        Catch ex As Exception
            newShape = Nothing
        End Try


        If IsNothing(newShape) Then

            If descriptionType = pptAnnotationType.ampelText Or _
                    descriptionType = pptAnnotationType.movedExplanation Or _
                    descriptionType = pptAnnotationType.lieferumfang Or _
                    descriptionType = pptAnnotationType.resourceCost Then

                'newShape = currentSlide.Shapes.AddComment()
                newShape = currentSlide.Shapes.AddCallout(Microsoft.Office.Core.MsoCalloutType.msoCalloutOne, _
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
                newShape = currentSlide.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, _
                                      txtShpLeft, txtShpTop, 50, txtShpHeight)
                With newShape
                    .TextFrame2.TextRange.Text = descriptionText
                    .TextFrame2.TextRange.Font.Size = CDbl(schriftGroesse)
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

        If ((Not descriptionType = pptAnnotationType.ampelText) And _
             (Not descriptionType = pptAnnotationType.movedExplanation) And _
             (Not descriptionType = pptAnnotationType.lieferumfang) And _
             (Not descriptionType = pptAnnotationType.resourceCost)) Then

            Select Case positionIndex

                Case pptPositionType.center

                    If newShape.Width > 1.5 * selectedPlanShape.Width Then
                        ' keine Farbänderung 
                    Else
                        ' wenn die Beschriftung von der Ausdehnung kleiner als die Phase/der Meilenstein ist
                        newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                            selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                    End If
                    txtShpLeft = selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width)
                    txtShpTop = selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height)

                Case pptPositionType.aboveCenter

                    txtShpLeft = selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width)
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

                    txtShpTop = selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height)

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
                    txtShpLeft = selectedPlanShape.Left + 0.5 * (selectedPlanShape.Width - newShape.Width)
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
                    txtShpTop = selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height)

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

                                Dim shortText As String = bestimmeElemText(tmpShape, True, False)
                                Dim origText As String = bestimmeElemText(tmpShape, False, True)

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
    Public Sub makeVisboShapesVisible(ByVal visible As Boolean)

        For Each pptSlide As PowerPoint.Slide In pptAPP.ActivePresentation.Slides

            For Each pptShape As PowerPoint.Shape In pptSlide.Shapes
                If isRelevantForProtection(pptShape) Then
                    pptShape.Visible = visible
                End If
            Next

        Next

    End Sub

    Public Function getProjektHistory(ByVal pvName) As clsProjektHistorie

        Dim tmpResult As clsProjektHistorie = Nothing
        Dim pName As String
        Dim variantName As String = ""
        Dim pHistory As New clsProjektHistorie

        If IsNothing(pvName) Then
            ' nichts tun 
        ElseIf pvName.trim.length = 0 Then
            ' auch nichts tun ...
        Else

            Dim tmpstr() As String = pvName.Split(New Char() {CType("#", Char)})
            pName = tmpstr(0).Trim
            If tmpstr.Length > 1 Then
                variantName = tmpstr(1).Trim
            Else
                variantName = ""
            End If

            If Not noDBAccessInPPT Then

                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

                If request.pingMongoDb() Then
                    Try

                        pHistory.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                        storedEarliest:=Date.MinValue, storedLatest:=Date.Now)
                    Catch ex As Exception
                        pHistory = Nothing
                    End Try
                Else
                    If englishLanguage Then
                        Call MsgBox("database connection lost !")
                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
                    End If

                End If




            End If


        End If

        getProjektHistory = tmpResult
    End Function

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
    ''' setzt die Markierung zurück, dass Elemente über Time-Machine / andere Variante  verschoben wurden 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub resetMovedGlowOfShapes()

        Dim bigTodoList As New Collection

        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
            bigTodoList.Add(tmpShape.Name)
        Next

        For Each tmpShpName As String In bigTodoList
            Try
                Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpShpName)
                If Not IsNothing(tmpShape) Then
                    If isVisboShape(tmpShape) Then

                        With tmpShape
                            If .Glow.Radius > 0 Then
                                tmpShape.Glow.Radius = 0.0
                                'tmpShape.Glow.Color.RGB = PowerPoint.XlRgbColor.rgbWhite

                                If .Tags.Item("MVD").Length > 0 Then
                                    ' nichts machen 
                                Else
                                    If .Tags.Item("MVE").Length > 0 Then
                                        .Tags.Delete("MVE")
                                    End If
                                End If

                            End If

                        End With

                    End If
                End If
            Catch ex As Exception

            End Try
        Next

        ' ur: 03.07.2017: setze alle Ampelfarben
        Call faerbeShapes(PTfarbe.none, showTrafficLights(PTfarbe.none))
        Call faerbeShapes(PTfarbe.green, showTrafficLights(PTfarbe.green))
        Call faerbeShapes(PTfarbe.yellow, showTrafficLights(PTfarbe.yellow))
        Call faerbeShapes(PTfarbe.red, showTrafficLights(PTfarbe.red))


    End Sub

    Friend Sub closeExcelAPP()
        Try
            If Not IsNothing(xlApp) Then
                For Each tmpWB As Excel.Workbook In CType(xlApp.Workbooks, Excel.Workbooks)
                    tmpWB.Close(SaveChanges:=False)
                Next
                xlApp.Quit()
            End If

            updateWorkbook = Nothing
            Call Sleep(300)
            xlApp = Nothing
        Catch ex As Exception

        End Try

    End Sub


End Module
