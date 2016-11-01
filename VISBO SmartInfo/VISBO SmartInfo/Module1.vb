Imports ProjectBoardDefinitions
Module Module1

    Friend WithEvents pptAPP As PowerPoint.Application

    Friend visboInfoActivated As Boolean = False
    Friend formIsShown As Boolean = False
    Friend Const markerName As String = "VisboMarker"
    Friend Const protectionTag As String = "VisboProtection"
    Friend Const protectionValue As String = "VisboValue"

    Friend currentSlide As PowerPoint.Slide
    Friend VisboProtected As Boolean = False

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

    ' gibt an, ob bei der suche die gefundenen Elemente mit AMrker angezeigt werden sollen oder nicht .. 
    Friend showMarker As Boolean = False

    ' globale Variable, die angibt, ob ShortName gezeichnet werden soll 
    Friend showShortName As Boolean = False
    ' globlaela Variable, die anzeigt, ob Orginal Name gezeigt werden soll 
    Friend showOrigName As Boolean = False

    Friend protectType As Integer
    Friend protectFeld1 As String = ""
    Friend protectFeld2 As String = ""

    Friend dbURL As String = ""
    Friend dbName As String = ""
    Friend userName As String = ""
    Friend userPWD As String = ""

    Friend noDBAccess As Boolean = True

    Friend defaultSprache As String = "Original"
    Friend selectedLanguage As String = defaultSprache

    Friend absEinheit As Integer = 0

    ' gibt an, ob irgendwelche Ampeln gesetzt sind 
    Friend ampelnExistieren As Boolean = False

    Friend selectedPlanShapes As PowerPoint.ShapeRange = Nothing

    ' hier werden PPTClander, linker Rand etc gehalten
    ' mit dieser Klasse können auch die Berechnungen Koord->Datum und umgekehrt durchgeführt werden 
    Friend slideCoordInfo As New clsPPTShapes

    Friend infoFrm As New frmInfo

    ' diese Listen enthalten die Infos welche Shapes Ampel grün, gelb etc haben
    ' welche welchen Namen tragen, ...
    Friend smartSlideLists As New clsSmartSlideListen
    Friend languages As New clsLanguages

    Friend bekannteIDs As SortedList(Of Integer, String)

    Friend Enum pptAbsUnit
        tage = 0
        wochen = 1
        monate = 2
    End Enum

    Friend Enum pptAnnotationType
        text = 0
        datum = 1
        ampelText = 2
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
    ''' hier wird bestimmt, ob es sich um eine VisboProtected Präsentation handelt 
    ''' </summary>
    ''' <param name="Pres"></param>
    ''' <remarks></remarks>
    Private Sub pptAPP_AfterPresentationOpen(Pres As PowerPoint.Presentation) Handles pptAPP.AfterPresentationOpen

        ' gibt es eine Sprachen-Tabelle ? 
        Dim langGUID As String = pptAPP.ActivePresentation.Tags.Item("langGUID")
        If langGUID.Length > 0 Then

            Dim langXMLpart As Office.CustomXMLPart = pptAPP.ActivePresentation.CustomXMLParts.SelectByID(langGUID)

            Dim langXMLstring = langXMLpart.XML
            languages = xml_deserialize(langXMLstring)

        End If

        ' Abrufen von Datenbank URL und Datenbank-Name 
        Try
            dbName = Pres.Tags.Item("DBNAME")
            dbURL = Pres.Tags.Item("DBURL")
        Catch ex As Exception
            dbName = ""
            dbURL = ""
        End Try
        

    End Sub

    Private Sub pptAPP_PresentationBeforeClose(Pres As PowerPoint.Presentation, ByRef Cancel As Boolean) Handles pptAPP.PresentationBeforeClose
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
        'If Not Pres.Name.EndsWith(".pptx") Then
        '    Call MsgBox("Alarm! unter falschem Namen gespeichert ... ")
        'End If
    End Sub

    ''' <summary>
    ''' ein VISBO Protected File kann nur als pptx gespeichert werden ...
    ''' </summary>
    ''' <param name="Pres"></param>
    ''' <remarks></remarks>
    Private Sub pptAPP_PresentationSave(Pres As PowerPoint.Presentation) Handles pptAPP.PresentationSave
        If VisboProtected And Not Pres.Name.EndsWith(".pptx") Then
            Call MsgBox("Speichern nur als .pptx möglich!")
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

        ' jetzt müssen die sortedLists 

        ' die aktuelle Slide setzen 
        If SldRange.Count = 1 Then
            currentSlide = SldRange.Item(1)

            If currentSlide.Tags.Count > 0 Then
                Try
                    If currentSlide.Tags.Item("SMART").Length > 0 Then


                        Try

                            slideCoordInfo = New clsPPTShapes
                            slideCoordInfo.pptSlide = currentSlide

                            With currentSlide
                                Dim tmpSD As String = .Tags.Item("CALL")
                                Dim tmpED As String = .Tags.Item("CALR")
                                slideCoordInfo.setCalendarDates(CDate(tmpSD), CDate(tmpED))
                            End With

                        Catch ex As Exception
                            slideCoordInfo = Nothing
                        End Try
                        
                        ' zurücksetzen der SmartSlideLists
                        smartSlideLists = New clsSmartSlideListen
                        bekannteIDs = New SortedList(Of Integer, String)

                        Dim anzShapes As Integer = currentSlide.Shapes.Count
                        ' jetzt werden die ganzen Listen aufgebaut 
                        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes
                            If tmpShape.Tags.Count > 0 Then
                                If isRelevantShape(tmpShape) Then
                                    ' invisible setzen ....
                                    'tmpShape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                                    bekannteIDs.Add(tmpShape.Id, tmpShape.Name)

                                    Call aktualisiereSortedLists(tmpShape, smartSlideLists)

                                    If visboInfoActivated And tmpShape.Visible = False Then
                                        tmpShape.Visible = True
                                    End If
                                End If
                            End If
                        Next
                    End If
                Catch ex As Exception

                End Try

            End If


        Else
            ' nichts tun, das heisst auch nichts verändern ...
        End If


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

            If Not IsNothing(shpRange) And visboInfoActivated Then


                ' es sind ein oder mehrere Shapes selektiert worden 
                Dim i As Integer = 0
                If shpRange.Count = 1 Then

                    If Not markerShpNames.ContainsKey(shpRange(1).Name) Then
                        Call deleteMarkerShapes()
                    ElseIf markerShpNames.Count > 1 Then
                        Call deleteMarkerShapes(shpRange(1).Name)
                    End If

                    ' prüfen, ob es ein Kommentar ist 
                    Dim tmpShape As PowerPoint.Shape = shpRange(1)
                    If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoComment Or _
                        (tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoAutoShape And tmpShape.Name.Contains("§")) Then
                        Call markReferenceShape(tmpShape.Name)
                    End If
                ElseIf shpRange.Count > 1 Then
                    ' für jedes Shape prüfen, ob es ein Comment Shape ist .. 
                    For Each tmpShape As PowerPoint.Shape In shpRange
                        If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoComment Or _
                        (tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoAutoShape And tmpShape.Name.Contains("§")) Then
                            Call markReferenceShape(tmpShape.Name)
                        End If
                    Next
                Else
                    If Not markerShpNames.ContainsKey(shpRange(1).Name) Then
                        Call deleteMarkerShapes()
                    End If

                End If

                For Each tmpShape As PowerPoint.Shape In shpRange


                    'If Not tmpShape.HasChart And Not tmpShape.HasTable Then
                    If tmpShape.Tags.Count > 0 Then

                        'If tmpShape.AlternativeText <> "" And tmpShape.Title <> "" Then

                        If isRelevantShape(tmpShape) Then
                            If bekannteIDs.ContainsKey(tmpShape.Id) Then

                                If Not relevantShapeNames.Contains(tmpShape.Name) Then
                                    relevantShapeNames.Add(tmpShape.Name, tmpShape.Name)
                                End If

                                If relevantShapeNames.Count = 1 Then

                                    If IsNothing(infoFrm) Then
                                        infoFrm = New frmInfo
                                        formIsShown = False
                                    End If

                                    With infoFrm
                                        .elemName.Text = bestimmeElemText(tmpShape, .showAbbrev.Checked, .showOrginalName.Checked)
                                        .elemDate.Text = bestimmeElemDateText(tmpShape, .showAbbrev.Checked)
                                        .fullBreadCrumb.Text = bestimmeElemBC(tmpShape)
                                        .ampelText.Text = bestimmeElemAmpelText(tmpShape)
                                        ' Festlegen der Beschriftungs-Position für Name und Text
                                        Call .setDTPicture(pptShapeIsMilestone(tmpShape))

                                    End With




                                    If Not formIsShown Then
                                        infoFrm.Show()
                                        formIsShown = True
                                    End If

                                Else
                                    With infoFrm
                                        If .elemName.Text <> bestimmeElemText(tmpShape, .showAbbrev.Checked, .showOrginalName.Checked) Then
                                            .elemName.Text = " ... "
                                        End If
                                        If .elemDate.Text <> bestimmeElemDateText(tmpShape, .showAbbrev.Checked) Then
                                            .elemDate.Text = " ... "
                                        End If

                                        If .ampelText.Text <> bestimmeElemAmpelText(tmpShape) Then
                                            .ampelText.Text = " ... "
                                        End If

                                        .positionTextButton.Image = Nothing
                                        .positionDateButton.Image = Nothing

                                    End With
                                End If

                            Else
                                ' die vorhandenen Tags löschen ... und den Namen ändern 
                                Call deleteShpTags(tmpShape)
                            End If

                        End If

                    End If

                Next

                'End If

                ' jetzt muss geprüft werden, ob relevantShapeNames mindestens ein Element enthält ..
                If relevantShapeNames.Count >= 1 Then

                    ReDim arrayOfNames(relevantShapeNames.Count - 1)

                    For ix As Integer = 1 To relevantShapeNames.Count
                        arrayOfNames(ix - 1) = CStr(relevantShapeNames(ix))
                    Next

                    selectedPlanShapes = currentSlide.Shapes.Range(arrayOfNames)
                Else
                    ' in diesem Fall wurden nur nicht-relevante Shapes selektiert 
                    If Not IsNothing(infoFrm) And formIsShown Then
                        With infoFrm
                            .elemName.Text = ""
                            .elemDate.Text = ""
                            .ampelText.Text = ""
                        End With
                    End If
                End If

            End If


        Catch ex As Exception

        End Try

    End Sub

    
    ''' <summary>
    ''' wird nur für relevante Shapes aufgerufen
    ''' baut die intelligenten Listen für das Slide auf 
    ''' wenn das Shape keine Abkürzung hat, so wird eine aus der laufenden Nummer erzeugt ...
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="smartSlideLists"></param>
    ''' <remarks></remarks>
    Private Sub aktualisiereSortedLists(ByVal tmpShape As PowerPoint.Shape, _
                                            ByRef smartSlideLists As clsSmartSlideListen)
        Dim shapeName As String = tmpShape.Name

        ' den classified Name behandeln ...
        Dim tmpName As String = tmpShape.Tags.Item("CN")
        If tmpName.Trim.Length = 0 Then
            Exit Sub
        End If

        Call smartSlideLists.addCN(tmpName, shapeName)

        ' den original Name behandeln ...
        tmpName = tmpShape.Tags.Item("ON")
        If tmpName.Trim.Length > 0 Then
            Call smartSlideLists.addON(tmpName, shapeName)
        End If

        ' den Short Name behandeln ...
        tmpName = tmpShape.Tags.Item("SN")
        If tmpName.Trim.Length = 0 Then
            ' es gibt keinen Short-Name, also soll einer aufgrund der laufenden Nummer erzeugt werden ...
            tmpName = smartSlideLists.getUID(shapeName).ToString
        End If
        Call smartSlideLists.addSN(tmpName, shapeName)

        ' den BreadCrumb behandeln 
        tmpName = tmpShape.Tags.Item("BC")
        If tmpName.Trim.Length > 0 Then
            Call smartSlideLists.addBC(tmpName, shapeName)
        End If

        ' AmpelColor behandeln
        Dim ampelColor As Integer = 0
        tmpName = tmpShape.Tags.Item("AC")
        If tmpName.Trim.Length > 0 Then
            Try
                If IsNumeric(tmpName) Then
                    ampelColor = CInt(tmpName)
                    Call smartSlideLists.addAC(ampelColor, shapeName)
                End If

            Catch ex As Exception

            End Try

        End If

        ' Lieferumfänge behandeln
        tmpName = tmpShape.Tags.Item("LU")
        If tmpName.Trim.Length > 0 Then
            Try
                Call smartSlideLists.addLU(tmpName, shapeName)
            Catch ex As Exception

            End Try
        End If

        ' wurde das Element verschoben ? 
        ' SmartslideLists werden auch gleich mit aktualisiert ... 
        Call checkShpOnManualMovement(tmpShape.Name)


    End Sub

    ''' <summary>
    ''' prüft, ob ein Shape manuell verschoben wurde; 
    ''' wenn ja, wird dem Shape die Movement Info gleich in Tags mitgegeben und die SmartSlideLists werden aktualisiert  
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Private Sub checkShpOnManualMovement(ByVal shapeName As String)

        Dim tmpShape As PowerPoint.Shape = currentSlide.Shapes(shapeName)
        Dim defaultExplanation As String = "manuell verschoben durch " & My.User.Name

        If IsNothing(tmpShape) Then
            Exit Sub
        Else

            If tmpShape.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                ' die Swimlane Texte sollen nicht berücksichtigt werden ...
            Else
                If pptShapeIsMilestone(tmpShape) Then
                    Dim pptDate As Date = slideCoordInfo.calcStartDate(tmpShape.Left + 0.5 * tmpShape.Width)
                    Dim planDate As Date = CDate(tmpShape.Tags.Item("ED"))

                    If DateDiff(DateInterval.Day, pptDate, planDate) = 0 Then
                        ' keine Änderung 
                    Else

                        With tmpShape
                            If .Tags.Item("MVD").Length > 0 Then
                                .Tags.Delete("MVD")
                            End If

                            .Tags.Add("MVD", pptDate.ToString)
                            If .Tags.Item("MVE").Length > 0 Then
                                ' nichts tun, der alte Wert soll erhalten bleiben 
                            Else
                                .Tags.Add("MVE", defaultExplanation)
                            End If
                        End With

                        Call smartSlideLists.addMV(tmpShape.Name)
                    End If
                Else
                    Dim pptSDate As Date = slideCoordInfo.calcStartDate(tmpShape.Left)
                    Dim pptEDate As Date = slideCoordInfo.calcStartDate(tmpShape.Left + tmpShape.Width)
                    Dim planSDate As Date = CDate(tmpShape.Tags.Item("SD"))
                    Dim planEDate As Date = CDate(tmpShape.Tags.Item("ED"))

                    If ((DateDiff(DateInterval.Day, pptSDate, planSDate) = 0) And _
                        (DateDiff(DateInterval.Day, pptEDate, planEDate) = 0)) Then
                        ' keine Änderung 
                    Else

                        With tmpShape

                            If .Tags.Item("MVD").Length > 0 Then
                                .Tags.Delete("MVD")
                            End If

                            .Tags.Add("MVD", pptSDate.ToString & "#" & pptEDate.ToString)
                            ' wenn bereits eine Explanation existiert, soll die erhalten bleiben 
                            If .Tags.Item("MVE").Length > 0 Then
                                ' nichts tun, der alte Wert soll erhalten bleiben 
                            Else
                                .Tags.Add("MVE", defaultExplanation)
                            End If

                        End With

                        Call smartSlideLists.addMV(tmpShape.Name)
                    End If

                End If
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
    Friend Function getPnameFromShpName(ByVal shapeName As String) As String

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

        getPnameFromShpName = tmpName

    End Function

    ''' <summary>
    ''' bringt zu dem gegebenen ShapeNamen den Namen des zugrundeliegenden Referenz-Shapes zurück
    ''' also zum Comment das zugehörige Shape , dass dann in Folge mit einem Marker markiert werden kann 
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <remarks></remarks>
    Friend Sub markReferenceShape(ByVal shapeName As String)
        Dim tmpText As String = ""

        If shapeName.EndsWith(CStr(pptAnnotationType.ampelText)) Then
            Dim strLength As Integer = shapeName.Length
            If strLength > 1 Then
                tmpText = shapeName.Substring(0, strLength - 1)

                Try
                    Dim refShape As PowerPoint.Shape = currentSlide.Shapes.Item(tmpText)
                    Call createMarkerShapes(refShape)
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
                        newHeight = 19
                        newWidth = 13
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

                    markerShpNames.Add(tmpShape.Name, markerShape.Name)

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
                Dim markerShapes As PowerPoint.ShapeRange = currentSlide.Shapes.Range(arrayOfShpNames)
                If Not IsNothing(markerShapes) Then
                    markerShapes.Delete()
                End If
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
    ''' entscheidet, ob es sich um einen Meilenstein oder eine Phase handelt
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
    ''' gibt den Ampeltext des Shapes zurück 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function bestimmeElemAmpelText(ByVal curShape As PowerPoint.Shape) As String
        Dim tmpText As String = ""

        Try
            If curShape.Tags.Item("AE").Length > 0 Then
                tmpText = curShape.Tags.Item("AE")
            End If
        Catch ex As Exception

        End Try

        bestimmeElemAmpelText = tmpText
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

        If isRelevantShape(curShape) Then
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
                End If

            ElseIf curShape.Tags.Item("CN").Length > 0 Then
                tmpText = curShape.Tags.Item("CN")
                translationNecessary = (selectedLanguage <> defaultSprache)
            End If
        End If

        If translationNecessary Then
            ' jetzt den Text ersetzen 
            tmpText = languages.translate(tmpText, selectedLanguage)
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
            If curShape.Tags.Item("ED").Length > 0 Then
                If Not showShort Then
                    tmpText = curShape.Tags.Item("ED")
                Else
                    Try
                        Dim msDate As Date = CDate(curShape.Tags.Item("ED"))
                        tmpText = msDate.Day.ToString & "." & msDate.Month.ToString
                    Catch ex As Exception
                        tmpText = curShape.Tags.Item("ED")
                    End Try
                End If
            End If
        Else
            If curShape.Tags.Item("SD").Length > 0 And curShape.Tags.Item("ED").Length > 0 Then
                If Not showShort Then
                    tmpText = curShape.Tags.Item("SD") & "-" & curShape.Tags.Item("ED")
                Else
                    Try
                        Dim startDate As Date = CDate(curShape.Tags.Item("SD"))
                        Dim endDate As Date = CDate(curShape.Tags.Item("ED"))
                        tmpText = startDate.Day.ToString & "." & startDate.Month.ToString & "-" & _
                                    endDate.Day.ToString & "." & endDate.Month.ToString
                    Catch ex As Exception
                        tmpText = curShape.Tags.Item("SD") & "-" & curShape.Tags.Item("ED")
                    End Try

                End If

            End If
        End If

        bestimmeElemDateText = tmpText
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
    ''' true, wenn das Shape wenigstens einen Wert für Tag CN enthält
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
    ''' prüft, ob ein Shape für Schutz relevant ist oder nicht 
    ''' </summary>
    ''' <param name="curShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isRelevantForProtection(ByVal curShape As PowerPoint.Shape) As Boolean
        Dim criteria1 As String = ""
        Dim criteria2 As Boolean

        Try
            criteria1 = curShape.Tags.Item("CN")
        Catch ex As Exception

        End Try

        Try
            ' alle VISBO Beschriftungen oder Kommentare enthalten das im Namen ... 
            criteria2 = (curShape.Name.Contains(")1§") Or curShape.Name.Contains(")0§"))
        Catch ex As Exception

        End Try

        If criteria1.Length > 0 Or criteria2 Then
            isRelevantForProtection = True
        Else
            isRelevantForProtection = False
        End If
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

        ' handelt es sich um den Lang-/Kurz-Namen oder um das Datum ? 

        If descriptionType = pptAnnotationType.text Then
            descriptionText = bestimmeElemText(selectedPlanShape, showShortName, showOrigName)
        ElseIf descriptionType = pptAnnotationType.datum Then
            descriptionText = bestimmeElemDateText(selectedPlanShape, showShortName)
        ElseIf descriptionType = pptAnnotationType.ampelText Then
            If IsNumeric(selectedPlanShape.Tags.Item("AC")) Then
                ampelFarbe = CInt(selectedPlanShape.Tags.Item("AC"))
            End If
            descriptionText = bestimmeElemAmpelText(selectedPlanShape)
            txtShpLeft = selectedPlanShape.Left + 1.5 * selectedPlanShape.Width + 5
            txtShpTop = selectedPlanShape.Top - 75
            txtShpWidth = 70
            txtShpHeight = 70
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
            If descriptionType = pptAnnotationType.ampelText Then
                newShape.Delete()
                newShape = Nothing
            End If
        Catch ex As Exception
            newShape = Nothing
        End Try


        If IsNothing(newShape) Then

            If descriptionType = pptAnnotationType.ampelText Then
                newShape = currentSlide.Shapes.AddComment()
                'newShape = currentSlide.Shapes.AddCallout(Microsoft.Office.Core.MsoCalloutType.msoCalloutOne, _
                '                      txtShpLeft, txtShpTop, txtShpWidth, txtShpHeight)
                With newShape
                    .Fill.ForeColor.RGB = RGB(240, 240, 240)
                    If ampelFarbe = 1 Then
                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbGreen
                    ElseIf ampelFarbe = 2 Then
                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbYellow
                    ElseIf ampelFarbe = 3 Then
                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbRed
                    Else
                        .Shadow.ForeColor.RGB = PowerPoint.XlRgbColor.rgbGrey
                    End If
                    '.Line.Weight = 3
                    .TextFrame2.TextRange.Text = descriptionText
                    .TextFrame2.TextRange.Font.Size = CDbl(schriftGroesse)
                    .TextFrame2.MarginBottom = 3
                    .TextFrame2.MarginLeft = 3
                    .TextFrame2.MarginRight = 3
                    .TextFrame2.MarginTop = 3
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = normalFarbe
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = Microsoft.Office.Core.MsoParagraphAlignment.msoAlignLeft
                    .Name = shapeName
                    .TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoTrue
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
        ' aber nur wenn es sich nicht um die Ampel handelt ...

        If Not descriptionType = pptAnnotationType.ampelText Then
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
                        If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                            newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                            selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        End If
                    End If

                    txtShpTop = selectedPlanShape.Top - newShape.Height

                Case pptPositionType.centerRight

                    If newShape.Width > selectedPlanShape.Width Then
                        txtShpLeft = selectedPlanShape.Left
                    Else
                        txtShpLeft = selectedPlanShape.Left + selectedPlanShape.Width - newShape.Width
                        If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                            newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                            selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        End If

                    End If

                    txtShpTop = selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height)

                Case pptPositionType.belowRight

                    If newShape.Width > selectedPlanShape.Width Then
                        txtShpLeft = selectedPlanShape.Left
                    Else
                        txtShpLeft = selectedPlanShape.Left + selectedPlanShape.Width - newShape.Width
                        If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                            newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                            selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        End If
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
                        If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                            newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                            selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        End If
                    End If

                    txtShpTop = selectedPlanShape.Top + selectedPlanShape.Height

                Case pptPositionType.centerLeft
                    If newShape.Width > selectedPlanShape.Width Then
                        txtShpLeft = selectedPlanShape.Left - (newShape.Width - selectedPlanShape.Width)
                    Else
                        txtShpLeft = selectedPlanShape.Left
                        If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                            newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                            selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        End If
                    End If
                    txtShpTop = selectedPlanShape.Top + 0.5 * (selectedPlanShape.Height - newShape.Height)

                Case pptPositionType.aboveLeft
                    If newShape.Width > selectedPlanShape.Width Then
                        txtShpLeft = selectedPlanShape.Left - (newShape.Width - selectedPlanShape.Width)
                    Else
                        txtShpLeft = selectedPlanShape.Left
                        If pptShapeIsMilestone(selectedPlanShape) And newShape.Width < 2 * selectedPlanShape.Width Then
                            newShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                            selectedPlanShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                        End If
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
                .Left = selectedPlanShape.Left + 2 * selectedPlanShape.Width
            End With
        End If






    End Sub

    ''' <summary>
    ''' wechselt die Sprache in der Annotation; tut dies für alle bereits dargestellten Beschriftungen 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub changeLanguageInAnnotations()


        ' andernfalls jetzt für alle Shapes ... 
        For Each tmpShape As PowerPoint.Shape In currentSlide.Shapes

            If isRelevantShape(tmpShape) Then

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
End Module
