Module Module1

    Friend WithEvents pptAPP As PowerPoint.Application

    Friend visboInfoActivated As Boolean = False
    Friend formIsShown As Boolean = False

    Friend currentSlide As PowerPoint.Slide
    Friend VisboProtected As Boolean = False

    ' wird gesetzt in Einstellungen 
    ' steuert, ob extended seach gemacht werden kann; wirkt auf Suchfeld (NAme, Original Name, Abkürzung, ..)  
    Friend extSearch As Boolean = False
    ' wird gesetzt in Einstellungen 
    ' gibt an, mit welcher Schriftgroesse der Text geschrieben wird 
    Friend schriftGroesse As Double = 8.0
    ' wird gesetzt in Einstellungen 
    ' gibt an, ob das Breadcrumb Feld gezeigt werden soll 
    Friend showBreadCrumbField As Boolean = False

    Friend absEinheit As Integer = 0

    ' gibt an, ob irgendwelche Ampeln gesetzt sind 
    Friend ampelnExistieren As Boolean = False

    Friend selectedPlanShapes As PowerPoint.ShapeRange = Nothing

    Friend infoFrm As New frmInfo

    ' diese Listen enthalten die Infos welche Shapes Ampel grün, gelb etc haben
    ' welche welchen Namen tragen, ...
    Friend smartSlideLists As New clsSmartSlideListen

    Friend bekannteIDs As SortedList(Of Integer, String)

    Friend Enum pptAbsUnit
        tage = 0
        wochen = 1
        monate = 2
    End Enum

    Friend Enum pptAnnotationType
        text = 0
        datum = 1
    End Enum

    Friend Enum pptInfoType
        cName = 0
        oName = 1
        sName = 2
        bCrumb = 3
        aColor = 4
        aExpl = 5
        appClass = 6
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
    End Enum


    ''' <summary>
    ''' hier wird bestimmt, ob es sich um eine VisboProtected Präsentation handelt 
    ''' </summary>
    ''' <param name="Pres"></param>
    ''' <remarks></remarks>
    Private Sub pptAPP_AfterPresentationOpen(Pres As PowerPoint.Presentation) Handles pptAPP.AfterPresentationOpen

    End Sub

    Private Sub pptAPP_PresentationBeforeSave(Pres As PowerPoint.Presentation, ByRef Cancel As Boolean) Handles pptAPP.PresentationBeforeSave
        ' wenn VisboProtected, dann müssen jetzt alle relevanten Shapes auf invisible gesetzt werden ...
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

    Private Sub pptAPP_WindowSelectionChange(Sel As PowerPoint.Selection) Handles pptAPP.WindowSelectionChange

        'Dim relevantShape As PowerPoint.Shape
        Dim arrayOfNames() As String
        Dim relevantShapeNames As New Collection

        selectedPlanShapes = Nothing

        Try
            Dim shpRange As PowerPoint.ShapeRange = Sel.ShapeRange

            If Not IsNothing(shpRange) And visboInfoActivated Then

                'If shpRange.Count = 1 Then
                '    relevantShape = shpRange.Item(1)

                '    'If Not relevantShape.HasChart And Not relevantShape.HasTable Then
                '    If relevantShape.Tags.Count > 0 Then

                '        'If relevantShape.AlternativeText <> "" And relevantShape.Title <> "" Then
                '        If Not IsNothing(relevantShape.Tags.Item("CN")) Then

                '            ' das Shape merken, damit im Formular später die Beschriftung ausgegeben werden kann 
                '            relevantShapeNames.Add(relevantShape.Name, relevantShape.Name)

                '            If IsNothing(infoFrm) Then
                '                infoFrm = New frmInfo
                '                formIsShown = False
                '            End If

                '            With infoFrm
                '                .elemName.Text = relevantShape.Title
                '                .elemDate.Text = relevantShape.AlternativeText
                '            End With

                '            If Not formIsShown Then
                '                infoFrm.Show()
                '                formIsShown = True
                '            End If
                '        Else
                '            With infoFrm
                '                .elemName.Text = ""
                '                .elemDate.Text = ""
                '            End With
                '        End If
                '    Else
                '        With infoFrm
                '            .elemName.Text = ""
                '            .elemDate.Text = ""
                '        End With
                '    End If
                'Else

                ' es sind mehrere Shapes selektiert worden 
                Dim i As Integer = 0
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
                                        .elemDate.Text = bestimmeElemDateText(tmpShape)
                                        .fullBreadCrumb.Text = bestimmeElemBC(tmpShape)
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
                                        If .elemDate.Text <> bestimmeElemDateText(tmpShape) Then
                                            .elemDate.Text = " ... "
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
                        End With
                    End If
                End If

            End If


        Catch ex As Exception

        End Try

    End Sub


    ''' <summary>
    ''' baut die intelligenten Listen für das Slide auf 
    ''' wenn das Shape keine Abkürzung hat, so wird eine aus der laufenden Nummer erzeugt ...
    ''' </summary>
    ''' <param name="tmpShape"></param>
    ''' <param name="smartSlideLists"></param>
    ''' <remarks></remarks>
    Private Sub aktualisiereSortedLists(ByRef tmpShape As PowerPoint.Shape, _
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
        Try
            ampelColor = CInt(tmpShape.Tags.Item("AC"))
        Catch ex As Exception

        End Try

        If tmpName.Trim.Length > 0 Then
            Call smartSlideLists.addAC(ampelColor, shapeName)
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

        If isRelevantShape(curShape) Then
            If showOriginalName Then
                If curShape.Tags.Item("ON").Length = 0 Then
                    tmpText = curShape.Tags.Item("CN")
                Else
                    tmpText = curShape.Tags.Item("ON")
                End If

            ElseIf showShortName Then
                If curShape.Tags.Item("SN").Length = 0 Then
                    If curShape.Tags.Item("CN").Length > 0 Then
                        tmpText = curShape.Tags.Item("CN")
                    End If
                Else
                    tmpText = curShape.Tags.Item("SN")
                End If

            ElseIf curShape.Tags.Item("CN").Length > 0 Then
                tmpText = curShape.Tags.Item("CN")
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
    Public Function bestimmeElemDateText(ByVal curShape As PowerPoint.Shape) As String

        Dim tmpText As String = ""

        If pptShapeIsMilestone(curShape) Then
            If curShape.Tags.Item("ED").Length > 0 Then
                tmpText = curShape.Tags.Item("ED")
            End If
        Else
            If curShape.Tags.Item("SD").Length > 0 And curShape.Tags.Item("ED").Length > 0 Then
                tmpText = curShape.Tags.Item("SD") & "-" & curShape.Tags.Item("ED")
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

        curShape.Name = "copied relevant Shape"


    End Sub

End Module
