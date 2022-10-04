Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports core = Microsoft.Office.Core
Imports xlNS = Microsoft.Office.Interop.Excel
'Imports pptNS = Microsoft.Office.Interop.PowerPoint
Imports System.ComponentModel
Imports Microsoft.VisualBasic.Constants
Imports ProjectBoardDefinitions
Imports System.Globalization

Public Module ProjekteZeichne

    ''' <summary>
    ''' zeichnet die Abhängigkeiten zu dem übergebenen Projekt 
    ''' </summary>
    ''' <param name="hproj">Projekt, dessen Abhängigkeiten dargestellt werden sollen</param>
    ''' <param name="type">welche Art Abhängigkeit soll dargestellt werden</param>
    ''' <param name="auswahl">0: sowohl incoming als outgoing Abhängigkeiten
    ''' 1: nur outgoing Abhängigkeiten
    ''' 2: nur incoming abhängigkeiten</param>
    ''' <remarks></remarks>
    Public Sub zeichneDependenciesOfProject(ByVal hproj As clsProjekt, ByVal type As Integer, ByVal auswahl As Integer)

        Dim listeDep As Collection ' nimmt die Liste der abhängigen Projekte auf
        Dim depListe As Collection ' nimmt die Liste der Projekte auf, von denen hproj abhängig ist 
        Dim pShape As Excel.Shape
        Dim dpShape As Excel.Shape
        Dim newConnector As Excel.Shape
        Dim X1, X2, Y1, Y2 As Single
        Dim dProj As clsProjekt
        Dim curDependency As clsDependency

        Dim pName As String = hproj.name, dpName As String

        Dim tmpshapes As Excel.Shapes
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerEOU As Boolean = enableOnUpdate



        listeDep = allDependencies.activeListe(hproj.name, PTdpndncyType.inhalt)
        depListe = allDependencies.passiveListe(hproj.name, PTdpndncyType.inhalt)

        If listeDep.Count = 0 And depListe.Count = 0 Then
            ' es gibt keine Abhängigkeiten
            Throw New Exception("keine Abhängigkeiten vorhanden")
        Else
            ' jetzt werden die Abhängigkeiten gezeichnet ...


            ' Event Behandlung ausschalten 
            enableOnUpdate = False
            appInstance.EnableEvents = False

            tmpshapes = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes
            pShape = tmpshapes.Item(pName)

            ' outgoing dependencies
            If auswahl = 0 Or auswahl = 1 Then

                For d = 1 To listeDep.Count

                    Try
                        dpName = CStr(listeDep.Item(d))
                        dpShape = tmpshapes.Item(dpName)
                        dProj = ShowProjekte.getProject(dpName, True)
                        Dim curDegree As Integer
                        curDependency = allDependencies.getDependency(PTdpndncyType.inhalt, pName, dpName)
                        If Not IsNothing(curDependency) Then
                            curDegree = curDependency.degree
                        Else
                            curDegree = PTdpndncy.schwach
                        End If

                        'Dim newShapeName As String = pName.Trim & "#" & dpName.Trim
                        Dim newShapeName As String = projectboardShapes.calcDependencyShapeName(pName, dpName)

                        ' prüfen , ob das Shape schon existiert ? 
                        Try

                            newConnector = tmpshapes.Item(newShapeName)

                            With newConnector
                                If curDegree = PTdpndncy.schwach Then
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = core.MsoLineDashStyle.msoLineLongDash
                                Else
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = core.MsoLineDashStyle.msoLineSolid
                                End If
                            End With



                        Catch ex As Exception

                            Call calculateDepCoord(pShape, dpShape, X1, Y1, X2, Y2)
                            newConnector = tmpshapes.AddConnector(core.MsoConnectorType.msoConnectorStraight, X1, Y1, X2, Y2)

                            With newConnector
                                .Line.EndArrowheadStyle = core.MsoArrowheadStyle.msoArrowheadTriangle
                                .ConnectorFormat.BeginConnect(pShape, 3)
                                .ConnectorFormat.EndConnect(dpShape, 1)
                                .Line.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                                If curDegree = PTdpndncy.schwach Then
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = core.MsoLineDashStyle.msoLineLongDash
                                Else
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = core.MsoLineDashStyle.msoLineSolid
                                End If
                                .Name = newShapeName
                                .AlternativeText = CInt(PTshty.dependency).ToString
                            End With

                            Call bringChartsToFront(newConnector)

                        End Try



                    Catch ex As Exception

                    End Try

                Next


            End If

            ' incoming dependencies
            dpShape = tmpshapes.Item(pName)
            dpName = pName
            dProj = hproj
            If auswahl = 0 Or auswahl = 2 Then

                For d = 1 To depListe.Count

                    Try
                        pName = CStr(depListe.Item(d))
                        pShape = tmpshapes.Item(pName)
                        hproj = ShowProjekte.getProject(pName)

                        Dim curDegree As Integer
                        curDependency = allDependencies.getDependency(PTdpndncyType.inhalt, pName, dpName)
                        If Not IsNothing(curDependency) Then
                            curDegree = curDependency.degree
                        Else
                            curDegree = PTdpndncy.schwach
                        End If

                        Dim newShapeName As String = projectboardShapes.calcDependencyShapeName(pName, dpName)
                        'Dim newShapeName As String = pName.Trim & "#" & dpName.Trim

                        ' prüfen , ob das Shape schon existiert ? 
                        Try

                            newConnector = tmpshapes.Item(newShapeName)

                            With newConnector
                                If curDegree = PTdpndncy.schwach Then
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = core.MsoLineDashStyle.msoLineLongDash
                                Else
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = core.MsoLineDashStyle.msoLineSolid
                                End If
                            End With


                        Catch ex As Exception

                            Call calculateDepCoord(pShape, dpShape, X1, Y1, X2, Y2)
                            newConnector = tmpshapes.AddConnector(core.MsoConnectorType.msoConnectorStraight, X1, Y1, X2, Y2)

                            With newConnector
                                .Line.EndArrowheadStyle = core.MsoArrowheadStyle.msoArrowheadTriangle
                                .ConnectorFormat.BeginConnect(pShape, 3)
                                .ConnectorFormat.EndConnect(dpShape, 1)
                                .Line.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                                If curDegree = PTdpndncy.schwach Then
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = core.MsoLineDashStyle.msoLineLongDash
                                Else
                                    .Line.Weight = 4.0
                                    .Line.DashStyle = core.MsoLineDashStyle.msoLineSolid
                                End If
                                .Name = newShapeName
                                .Title = "Dependency"
                            End With

                            Call bringChartsToFront(newConnector)

                        End Try

                    Catch ex As Exception

                    End Try

                Next


            End If

            ' Event Behandlung auf vorherigen Zustand setzen ...
            appInstance.EnableEvents = formerEE
            enableOnUpdate = formerEOU
        End If
    End Sub

    ''' <summary>
    ''' wird von der WPFPIE Vorage aufgerufen !
    ''' </summary>
    ''' <param name="nameList"></param>
    ''' <param name="farbTyp"></param>
    ''' <param name="numberIt"></param>
    ''' <remarks></remarks>
    Public Sub zeichneMilestones(ByVal nameList As Collection, ByVal farbTyp As Integer, ByVal numberIt As Boolean)
        ' tue es für alle Projekte in Showprojekte 


        Dim todoListe As New SortedList(Of Long, clsProjekt)
        Dim key As Long
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formereO As Boolean = enableOnUpdate

        appInstance.EnableEvents = False
        enableOnUpdate = False

        If selectedProjekte.Count > 0 Then
            For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste

                key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                todoListe.Add(key, kvp.Value)

            Next
        Else
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                todoListe.Add(key, kvp.Value)

            Next
        End If


        Dim msNumber As Integer = 1

        For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

            Call zeichneMilestonesInProjekt(kvp.Value, nameList, farbTyp, showRangeLeft, showRangeRight, numberIt, msNumber, False)

        Next

        Call awinSelect()

        appInstance.EnableEvents = formerEE
        enableOnUpdate = formereO

    End Sub


    ''' <summary>
    ''' zeichnet die Meilensteine eines Projektes
    ''' </summary>
    ''' <param name="hproj">
    ''' das Projekt, das die Meilensteine enthält</param>
    ''' <param name="namenListe">
    ''' enthält die Namen,der Meilensteine die gezeichnet werden sollen
    ''' wenn leer, werden alle gezeichnet</param>
    ''' <param name="farbTyp">
    ''' gibt an , welche Farbe gezeichnet werden soll; bei 4 werden alle gezeichnet </param>
    ''' <param name="tmpShowrangeleft">
    ''' gibt den linken Rand des Zeitraums an, sofern einer betrachtet werden soll </param>
    ''' <param name="tmpShowrangeRight">gibt den rechten Rand des Zeitraums an, sofern einer betrachtet werden soll </param>
    ''' <param name="numberIt">
    ''' gibt an, ob der Meilenstein nummeriert werden soll</param>
    ''' <param name="msNumber">
    ''' gibt die Nummer an, aber nummeriert werden soll</param>
    ''' <param name="report">
    ''' gibt an, ob vom Reporting aufgerufen
    ''' </param>
    ''' <remarks></remarks>
    Public Sub zeichneMilestonesInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection, ByVal farbTyp As Integer, ByVal tmpShowRangeLeft As Integer, ByVal tmpShowrangeRight As Integer,
                                                      ByVal numberIt As Boolean, ByRef msNumber As Integer, ByVal report As Boolean)

        Dim top As Double, left As Double, width As Double, height As Double
        Dim resultShape As Excel.Shape
        Dim worksheetShapes As Excel.Shapes
        Dim heute As Date = Date.Now
        Dim alreadyGroup As Boolean = False
        Dim shpElement As Excel.Shape
        Dim shpName As String
        Dim resultColumn As Integer
        Dim onlyFew As Boolean
        Dim projectShape As Excel.Shape
        Dim shapeGruppe As Excel.ShapeRange
        Dim newShape As Excel.ShapeRange = Nothing
        Dim listOFShapes As New Collection
        Dim found As Boolean = True
        Dim showOnlyWithinTimeFrame As Boolean
        'Dim vorlagenShape As xlNS.Shape
        Dim realNameList As New Collection

        Try
            If namenListe.Count > 0 Then
                onlyFew = True
                realNameList = hproj.getElemIdsOf(namenListe, True)
            Else
                onlyFew = False
                realNameList = hproj.getAllElemIDs(True)
            End If
        Catch ex As Exception
            onlyFew = False
        End Try

        ' jetzt wurde aus der Liste von Namen / oder IDs gesichert eine Liste von IDs gemacht 


        ' es muss abgefangen werden, daß nicht alle Meilensteine gezeichnet werden, wenn namenListe.count > 0, aber später 
        ' die neue namenliste.count  = 0 ; dass wird dadurch sichergestellt, dass onlyFew bereits vor Bearbeitung / Ersetzung der Namenliste gesetzt ist  


        If tmpShowRangeLeft <= 0 Or
            tmpShowrangeRight <= 0 Or
            tmpShowRangeLeft > tmpShowrangeRight Then

            showOnlyWithinTimeFrame = False

        Else

            showOnlyWithinTimeFrame = True

        End If



        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

            worksheetShapes = .Shapes
            ' Änderung 12.7.14 Alle Milestone Shapes in ein gruppiertes Shape
            ' jetzt muss das Projekt-Shape gesucht werden
            Try
                projectShape = worksheetShapes.Item(hproj.name)
            Catch ex As Exception
                found = False
                projectShape = Nothing
            End Try


            ' found=true bedeutet, dass das Shape bereits angezeigt wird  
            If found Then

                ' jetzt muss die Liste an Shapes aufgebaut werden 
                If projectShape.AlternativeText = CInt(PTshty.projektL).ToString Or
                    projectShape.AutoShapeType = core.MsoAutoShapeType.msoShapeRoundedRectangle Then

                    listOFShapes.Add(projectShape.Name)

                Else
                    shapeGruppe = projectShape.Ungroup
                    Dim anzElements As Integer = shapeGruppe.Count

                    Dim i As Integer
                    For i = 1 To anzElements
                        listOFShapes.Add(shapeGruppe.Item(i).Name)
                    Next


                End If

                ' hier muss jetzt ausgenutzt werden, dass man bereits die direkten IDs der Meilensteine hat ... 

                ' es muss aber auch berücksichtigt werden, wenn alle gezeigt werden sollen ... 

                For m As Integer = 1 To realNameList.Count

                    Dim cMilestone As clsMeilenstein = hproj.getMilestoneByID(CStr(realNameList.Item(m)))
                    Dim isMissingDefinition As Boolean = Not MilestoneDefinitions.Contains(cMilestone.name)

                    If Not IsNothing(cMilestone) Then
                        Dim cBewertung As clsBewertung

                        cBewertung = cMilestone.getBewertung(1)
                        resultColumn = getColumnOfDate(cMilestone.getDate)

                        If farbTyp = 4 Or farbTyp = cBewertung.colorIndex Then
                            ' es muss nur etwas gemacht werden , wenn entweder alle Farben gezeichnet werden oder eben die übergebene

                            If (showOnlyWithinTimeFrame And (resultColumn < tmpShowRangeLeft Or resultColumn > tmpShowrangeRight)) Then
                                ' nichts machen 
                            Else
                                Dim zeilenoffset As Integer = 0
                                ' hier die übergeordnete Phase holen ...

                                ' Änderung tk 25.11.15: sofern die Definition in definitions.. enthalten ist: auch berücksichtigen
                                'If MilestoneDefinitions.Contains(cMilestone.name) Then
                                '    vorlagenShape = MilestoneDefinitions.getShape(cMilestone.name)

                                'Else
                                '    vorlagenShape = missingMilestoneDefinitions.getShape(cMilestone.name)

                                'End If


                                'Dim factorB2H As Double = vorlagenShape.Width / vorlagenShape.Height
                                'ur:190725
                                Dim appear As clsAppearance = appearanceDefinitions.getMileStoneAppearance(cMilestone)

                                Dim factorB2H As Double = appear.width / appear.height


                                hproj.calculateMilestoneCoord(cMilestone.getDate, zeilenoffset, factorB2H, top, left, width, height)
                                'hproj.calculateResultCoord(cResult.getDate, zeilenoffset, top, left, width, height)

                                shpName = projectboardShapes.calcMilestoneShapeName(hproj.name, cMilestone.nameID)

                                ' existiert das schon ? 
                                Try
                                    shpElement = worksheetShapes.Item(shpName)
                                Catch ex As Exception
                                    shpElement = Nothing
                                End Try

                                If shpElement Is Nothing Then

                                    If report Then
                                        top = top - boxWidth
                                    End If

                                    '' Alt - Start 
                                    'resultShape = .Shapes.AddShape(Type:=vorlagenShape.AutoShapeType,
                                    '                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                    'vorlagenShape.PickUp()
                                    'resultShape.Apply()

                                    'resultShape.Rotation = vorlagenShape.Rotation

                                    'With resultShape
                                    '    .Name = shpName
                                    '    .Title = cMilestone.nameID
                                    '    .AlternativeText = CInt(PTshty.milestoneN).ToString
                                    'End With
                                    '' Alt - Ende
                                    'Neu - Start
                                    Try
                                        resultShape = worksheetShapes.Item(shpName)
                                    Catch ex As Exception
                                        resultShape = Nothing
                                    End Try

                                    If resultShape Is Nothing Then


                                        resultShape = worksheetShapes.AddShape(Type:=appear.shpType,
                                                                        Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                        'ur:190725
                                        'msShape = worksheetShapes.AddShape(Type:=vorlagenShape.AutoShapeType,
                                        '                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                        'vorlagenShape.PickUp()
                                        'msShape.Apply()

                                        resultShape.Rotation = appear.Rotation
                                        With resultShape
                                            .Name = shpName
                                            .Title = cMilestone.nameID
                                            .AlternativeText = CInt(PTshty.milestoneE).ToString
                                        End With

                                        ' Neu - Ende



                                        msNumber = msNumber + 1
                                        If numberIt Then
                                            Call defineResultAppearance(hproj, msNumber, resultShape, cBewertung, isMissingDefinition, cMilestone.farbe)

                                        Else
                                            Call defineResultAppearance(hproj, 0, resultShape, cBewertung, isMissingDefinition, cMilestone.farbe)
                                        End If

                                        ' jetzt der Liste der ProjectboardShapes hinzufügen
                                        projectboardShapes.add(resultShape)

                                        ' jetzt der Liste von Shapes hinzufügen, die dann nachher zum ProjektShape gruppiert werden sollen 
                                        listOFShapes.Add(resultShape.Name)

                                    End If

                                End If
                            End If
                        End If

                    End If

                Next


            End If


            If listOFShapes.Count > 1 Then
                ' hier werden die Shapes gruppiert
                projectShape = projectboardShapes.groupShapes(listOFShapes, hproj.name)

                ' jetzt der Liste der ProjectboardShapes hinzufügen
                projectboardShapes.add(projectShape)
            End If


        End With


        ' jetzt müssen ggf die Charts wieder in den Vordergrund gebracht werden 
        Call bringChartsToFront(projectShape)

    End Sub

    ''' <summary>
    ''' zeichnet die Werte der Rollen und Kosten auf die Projekt-Tafel
    ''' </summary>
    ''' <param name="hproj">das Projekt, das gezeichnet werden soll </param>
    ''' <param name="namenListe">die Liste der Rollen bzw. Kosten</param>
    ''' <param name="tmpShowRangeLeft">linke Spalte des Bereiches, in dem gezeichnet werden soll</param>
    ''' <param name="tmpShowrangeRight">rechte Spalte des Bereiches, in dem gezeichnet werden soll</param>
    ''' <param name="type">gibt an den Type an, damit lässt sich entscheiden, ob Rolle / Kosten in der Namenliste stehen</param>
    ''' <remarks></remarks>
    Public Sub zeichneRollenKostenWerteInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection, ByVal tmpShowRangeLeft As Integer, ByVal tmpShowrangeRight As Integer,
                                                          ByVal type As String)

        ' aktuell wird das nur im Fall nicht-extended Mode angezeigt 
        If awinSettings.drawphases Then
            ' nichts tun 
            Call MsgBox("wird aktuell nur im Einzeilen - Modus unterstützt" & vbLf &
                         "Wählen Sie Extended Mode = Nein")
            Exit Sub
        End If

        ' aktuell wird nur unterstützt, einen Monat anzuzeigen 
        If tmpShowRangeLeft <> tmpShowrangeRight Then
            Call MsgBox("aktuell wird nur ein Monat unterstützt")
            Exit Sub
        End If

        ' bestimme die Zeile und die Spalte 
        Dim currentRow As Integer = hproj.tfZeile + 1
        Dim currentColumn As Integer = tmpShowRangeLeft

        ' bestimme den Wert
        Dim currentValue As Double = hproj.getBedarfeInMonth(namenListe, type, tmpShowRangeLeft, True)

        ' schreibe jetzt den Wert in die Zelle
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)
            If currentValue > 0 Then
                .Cells(currentRow, currentColumn).value = CInt(currentValue)

            End If

        End With

        appInstance.EnableEvents = formerEE


    End Sub

    ''' <summary>
    ''' trägt bei Projektlinie den Namen ein ... 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub zeichneNameInProjekt(ByVal hproj As clsProjekt)

        Dim projectTop As Single, projectLeft As Single, projectHeight As Single, projectWidth As Single
        Dim txtTop As Single, txtLeft As Single, txtwidth As Single, txtHeight As Single
        Dim pNameShape As Excel.Shape
        Dim worksheetShapes As Excel.Shapes

        Dim projectShape As Excel.Shape
        Dim shapeGruppe As Excel.ShapeRange

        Dim listOFShapes As New Collection
        Dim found As Boolean = True

        Dim pvName As String = calcProjektKey(hproj.name, hproj.variantName)

        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

            worksheetShapes = .Shapes
            ' Änderung 12.7.14 Alle Milestone Shapes in ein gruppiertes Shape
            ' jetzt muss das Projekt-Shape gesucht werden
            Try
                projectShape = worksheetShapes.Item(hproj.name)
            Catch ex As Exception
                found = False
                projectShape = Nothing
            End Try


            ' found=true bedeutet, dass das Shape bereits angezeigt wird  
            If found Then

                ' Merken der Koordinaten 
                ' bestimmen der Text Koordinaten 
                With projectShape
                    projectTop = .Top
                    projectLeft = .Left
                    projectWidth = .Width
                    projectHeight = .Height
                End With

                txtTop = projectTop
                txtLeft = projectLeft + 7
                txtwidth = 30
                txtHeight = 30

                ' jetzt muss die Liste an Shapes aufgebaut werden 

                If projectShape.AlternativeText = CInt(PTshty.projektL).ToString Or
                    projectShape.AutoShapeType = core.MsoAutoShapeType.msoShapeRoundedRectangle Then
                    listOFShapes.Add(projectShape.Name)
                Else
                    shapeGruppe = projectShape.Ungroup
                    Dim anzElements As Integer = shapeGruppe.Count
                    ' hier muss der alte Shape Text rausgelöscht werdewn 

                    Dim oldTxtxShape As Excel.Shape = Nothing
                    For Each tmpshape As Excel.Shape In shapeGruppe
                        If tmpshape.AlternativeText = "(Projektname)" Then
                            oldTxtxShape = tmpshape
                        Else
                            listOFShapes.Add(tmpshape.Name)
                        End If
                    Next

                    ' jetzt muss der alte Text gelöscht werden ...
                    If Not IsNothing(oldTxtxShape) Then
                        oldTxtxShape.Delete()
                    End If


                End If



                ' ab jetzt darf auf projectShape nicht mehr zugegriffen werden, da es ggf bereits im Else-Zweig aufgelöst wurde ...


                ' jetzt muss das Textshape erzeugt werden 
                pNameShape = worksheetShapes.AddLabel(core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                        txtLeft, txtTop, txtwidth, txtHeight)

                With pNameShape
                    .AlternativeText = "(Projektname)"
                    .TextFrame2.AutoSize = core.MsoAutoSize.msoAutoSizeShapeToFitText
                    .TextFrame2.WordWrap = core.MsoTriState.msoFalse
                    .TextFrame2.TextRange.Text = hproj.getShapeText
                    .TextFrame2.TextRange.Font.Size = hproj.Schrift
                    .TextFrame2.MarginLeft = 4
                    .TextFrame2.MarginRight = 4
                    .TextFrame2.MarginTop = 0
                    .TextFrame2.MarginBottom = 0
                    .TextFrame2.VerticalAnchor = core.MsoVerticalAnchor.msoAnchorMiddle
                    .TextFrame2.HorizontalAnchor = core.MsoHorizontalAnchor.msoAnchorCenter
                    ' braucht man für die Update Routine 
                    .Name = calcProjectTextShapeName(hproj.name)

                    ' hier werden die Farben und Fonts gemäß dem Protection Status bestimmt 
                    If writeProtections.isProtected(pvName) Then
                        If writeProtections.isPermanentProtected(pvName) Then
                            ' use permanent Font
                            .TextFrame.Characters.Font.FontStyle = awinSettings.protectedPermanentFont

                            If writeProtections.isProtected(pvName, dbUsername) Then
                                ' use byOtherProtectedColor 
                                .TextFrame.Characters.Font.Color = awinSettings.protectedByOtherColor
                            Else
                                ' use byMeProtected Color 
                                .TextFrame.Characters.Font.Color = awinSettings.protectedByMeColor
                            End If
                        Else
                            ' use normal font
                            If writeProtections.isProtected(pvName, dbUsername) Then
                                ' use byOtherProtectedColor 
                                .TextFrame.Characters.Font.Color = awinSettings.protectedByOtherColor
                            Else
                                ' use byMeProtected Color 
                                .TextFrame.Characters.Font.Color = awinSettings.protectedByMeColor
                            End If
                        End If
                    Else
                        ' es ist nicht protected, also muss nichts verändert werden 
                    End If

                    .Fill.Visible = core.MsoTriState.msoTrue
                    .Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .Fill.Transparency = 0
                    .Fill.Solid()

                End With

                ' jetzt muss das Shape noch in der Höhe richtig positioniert werden 
                Dim diff As Single
                If awinSettings.drawphases Or hproj.extendedView Then
                    diff = CSng(0.3 * boxHeight)
                Else
                    diff = (pNameShape.Height - projectHeight) / 2
                End If
                pNameShape.Top = projectTop - diff

                Try
                    If pNameShape.Width > projectWidth Then
                        Dim newName As String = pNameShape.TextFrame2.TextRange.Text
                        Dim anzZeichen As Integer = newName.Length
                        Do Until pNameShape.Width < projectWidth Or anzZeichen < 3
                            pNameShape.TextFrame2.TextRange.Text = newName.Substring(0, anzZeichen - 2)
                            anzZeichen = anzZeichen - 1
                        Loop
                    End If
                Catch ex As Exception
                    pNameShape.TextFrame2.TextRange.Text = ""
                End Try


                ' jetzt wird das Shape aufgenommen 
                listOFShapes.Add(pNameShape.Name)


            End If


            If listOFShapes.Count > 1 Then
                ' hier werden die Shapes gruppiert
                projectShape = projectboardShapes.groupShapes(listOFShapes, hproj.name)

                ' jetzt der Liste der ProjectboardShapes hinzufügen
                projectboardShapes.add(projectShape)
            End If


        End With


        ' jetzt müssen ggf die Charts wieder in den Vordergrund gebracht werden 
        Call bringChartsToFront(projectShape)


    End Sub

    ''' <summary>
    ''' zeichnet den Pfeil, der anzeigt, um wieviel ein Projekt bei Optimierung verschoben werden würde
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <remarks></remarks>
    Public Sub ZeichneMoveLineOfProjekt(ByRef pname As String)

        Dim start As Integer
        Dim laenge As Integer
        Dim pcolor As Integer, schriftfarbe As Object, fillColor As Integer, borderColor As Integer
        Dim schriftgroesse As Integer
        Dim zeilenOffset As Integer = 1
        Dim spaltenOffset As Integer = 0
        Dim hproj As clsProjekt
        Dim leftDrawn As Boolean
        Dim moveLength As Integer
        Dim tfz As Integer, tfs As Integer
        Dim top As Double, left As Double, width As Double, height As Double

        Dim straightLine As core.MsoConnectorType = core.MsoConnectorType.msoConnectorStraight


        hproj = ShowProjekte.getProject(pname)
        With hproj
            laenge = .anzahlRasterElemente
            start = .Start + .StartOffset
            moveLength = .StartOffset
            pcolor = CInt(.farbe)
            schriftfarbe = .Schriftfarbe
            schriftgroesse = .Schrift
            tfz = .tfZeile
            tfs = .tfspalte
        End With

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        If moveLength <> 0 Then
            height = 0.4 * boxHeight

            If moveLength < 0 Then
                leftDrawn = True
                left = (tfs - 1 + 0.5) * boxWidth + moveLength * boxWidth ' movelength ist negativ , deshalb "+"
                width = moveLength * boxWidth * (-1)
                top = topOfMagicBoard + (tfz - 1 + 0.75) * boxHeight
            Else
                leftDrawn = False
                left = (tfs + laenge - 1 - 0.5) * boxWidth
                top = topOfMagicBoard + (tfz - 1 + 0.25) * boxHeight
                width = moveLength * boxWidth
            End If

            Dim shp As Excel.Shape
            With appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT))


                fillColor = RGB(255, 255, 255)
                borderColor = pcolor

                If leftDrawn Then
                    shp = CType(.Shapes, Excel.Shapes).AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeLeftArrow,
                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))

                Else
                    shp = CType(.Shapes, Excel.Shapes).AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRightArrow,
                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                End If

                ' jetzt wird der Pfeil gezeichnet




                With shp
                    With .Fill
                        .ForeColor.RGB = fillColor
                        .Transparency = 0.0
                    End With
                    With .Line
                        '.Visible = True
                        .Weight = 1.5
                        .ForeColor.RGB = borderColor
                        .Transparency = 0
                    End With

                End With



            End With



        End If

        appInstance.EnableEvents = formerEE

    End Sub



    ''' <summary>
    ''' zeichnet für das angegebene Projekt hproj alle in namenliste enthaltenen Phasen, sofern die Phase innerhalb 
    ''' der vonMonth, bisMonth aufgespannten Grenzen liegt 
    ''' wenn namenliste leer ist, werden alle Phasen des Projekts gezeichnet
    ''' numberit steuet, ob die Phase für Reporting Zwecke eine Nummerierung erhalten soll 
    ''' </summary>
    ''' <param name="hproj">das Projekt-Objekt</param>
    ''' <param name="namenListe">Liste der Phasen, die gezeichnet werden sollen</param>
    ''' <param name="vonMonth">linker rand des Kalenderzeitraums, der betrachtet werden soll</param>
    ''' <param name="bisMonth">rechter Rand des Kalenderzeitraums, der betrachtet werden soll</param>
    ''' <param name="numberIt">soll nummeriert werden </param>
    ''' <param name="msNumber">Start der Nummerierung</param>
    ''' <remarks></remarks>
    Public Sub zeichnePhasenInProjekt(ByVal hproj As clsProjekt, ByVal namenListe As Collection,
                                      ByVal numberIt As Boolean, ByRef msNumber As Integer,
                                      Optional ByVal vonMonth As Integer = 0, Optional ByVal bisMonth As Integer = 0)

        'Dim top1 As Double, left1 As Double, top2 As Double, left2 As Double
        Dim top As Double, left As Double, width As Double, height As Double
        Dim nummer As Integer
        Dim phasenShape As xlNS.Shape
        Dim worksheetShapes As xlNS.Shapes
        Dim heute As Date = Date.Now
        Dim alreadyGroup As Boolean = False
        Dim shpElement As xlNS.Shape
        ' vorlagenshape ist durch Ute's Umsetzung der Appearances ohne Excel Shapes unnötig geworden ... 
        'Dim vorlagenshape As xlNS.Shape
        Dim appear As clsAppearance
        Dim shpName As String
        Dim todoListe As New Collection
        Dim realNameList As New Collection
        Dim phasenSchriftgroesse As Double = 5.0

        Dim onlyFew As Boolean
        Dim projectShape As xlNS.Shape
        Dim shapeGruppe As xlNS.ShapeRange
        Dim listOFShapes As New Collection
        Dim found As Boolean = True


        Dim ok As Boolean = True

        ' alle Phasen auslesen , die NameIDs dazu holen 
        Try
            If namenListe.Count > 0 Then
                onlyFew = True
                realNameList = hproj.getElemIdsOf(namenListe, False)
            Else
                onlyFew = False
                realNameList = hproj.getAllElemIDs(False)
            End If
        Catch ex As Exception
            onlyFew = False
        End Try


        ' als wievielte Phase wird das Shape gezeichnet ... 
        nummer = 1




        Try
            If vonMonth = 0 Or bisMonth = 0 Then
                ' alle Phasen betrachten 
                todoListe = realNameList
            Else
                'bringt eine List von Phasen ElemIDs zurück, die den angegebenen Zeitraum berühren / überdecken

                todoListe = hproj.phasesWithinTimeFrame(False, vonMonth, bisMonth, realNameList)

            End If



        Catch ex As Exception

        End Try

        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

            worksheetShapes = .Shapes

            ' Änderung 12.7.14 Alle Phasen Shapes in ein gruppiertes Shape
            ' jetzt muss das Projekt-Shape gesucht werden
            Try
                projectShape = worksheetShapes.Item(hproj.name)
            Catch ex As Exception
                found = False
                projectShape = Nothing
            End Try


            If found Then

                ' jetzt muss die Liste an Shapes aufgebaut werden 
                If projectShape.AlternativeText = CInt(PTshty.projektL).ToString Or
                    projectShape.AutoShapeType = core.MsoAutoShapeType.msoShapeRoundedRectangle Then

                    listOFShapes.Add(projectShape.Name)

                Else
                    shapeGruppe = projectShape.Ungroup
                    Dim anzElements As Integer = shapeGruppe.Count

                    Dim i As Integer
                    For i = 1 To anzElements
                        listOFShapes.Add(shapeGruppe.Item(i).Name)
                    Next
                End If


                Dim cphase As clsPhase

                ' in der todoListe stehen jetzt nur Phasen, die den angegeben Zeitraum betreffen 
                For p = 1 To todoListe.Count

                    Dim phaseNameID As String = CStr(todoListe(p))


                    If realNameList.Contains(phaseNameID) Then

                        cphase = hproj.getPhaseByID(phaseNameID)
                        Dim isMissingDefinition As Boolean = Not PhaseDefinitions.Contains(cphase.name)

                        appear = appearanceDefinitions.getPhaseAppearance(cphase)


                        Try
                            'cphase.calculateLineCoord(hproj.tfZeile, nummer, gesamtZahl, top1, left1, top2, left2, linienDicke)
                            cphase.calculatePhaseShapeCoord(top, left, width, height)
                        Catch ex As Exception
                            ok = False
                        End Try



                        If ok Then
                            nummer = nummer + 1

                            shpName = projectboardShapes.calcPhaseShapeName(hproj.name, cphase.nameID)
                            'shpName = hproj.name & "#" & cphase.name
                            ' existiert das schon ? 
                            Try
                                shpElement = worksheetShapes.Item(shpName)
                            Catch ex As Exception
                                shpElement = Nothing
                            End Try

                            If shpElement Is Nothing Then



                                phasenShape = .Shapes.AddShape(Type:=appear.shpType,
                                                                    Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                With phasenShape
                                    .Name = shpName
                                    .Title = cphase.nameID
                                    .AlternativeText = CInt(PTshty.phaseN).ToString
                                End With

                                msNumber = msNumber + 1
                                If numberIt Then
                                    Call definePhaseAppearance(hproj, cphase, msNumber, phasenShape, isMissingDefinition)

                                Else
                                    Call definePhaseAppearance(hproj, cphase, 0, phasenShape, isMissingDefinition)
                                End If

                                ' jetzt der Liste der ProjectboardShapes hinzufügen
                                projectboardShapes.add(phasenShape)

                                ' jetzt der Liste von Shapes hinzufügen, die dann nachher zum ProjektShape gruppiert werden sollen 
                                listOFShapes.Add(phasenShape.Name)


                            End If

                        End If

                    End If

                    ok = True

                Next

            End If

            If listOFShapes.Count > 1 Then
                ' hier werden die Shapes gruppiert
                projectShape = projectboardShapes.groupShapes(listOFShapes, hproj.name)

                ' jetzt der Liste der ProjectboardShapes hinzufügen
                projectboardShapes.add(projectShape)

            End If

        End With

        ' jetzt müssen die Charts ggf wieder nach vorne gebracht werden 
        Call bringChartsToFront(projectShape)


    End Sub



    ''' <summary>
    ''' zeichnet das Projekt "pname" in die Plantafel; 
    ''' wenn es bereits vorhanden ist: keine Aktion  
    ''' noCollection ist eine Collection von Projekt-Namen, die beim Suchen nach einem Platz 
    ''' auf der Projekt-Tafel nicht berücksichtigt werden soll
    ''' ist insbesondere wichtig, wenn mehrere Projekte selektiert wurden und verschoben werden 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <remarks></remarks>
    Public Sub ZeichneProjektinPlanTafel(ByVal noCollection As Collection, ByVal pname As String, ByVal tryzeile As Integer,
                                         ByVal drawPhaseList As Collection, ByVal drawMilestoneList As Collection,
                                         Optional useTryZeileAnyway As Boolean = True)


        Dim drawphases As Boolean = awinSettings.drawphases
        Dim phasenNameID As String
        Dim phaseShapeName As String
        Dim msShapeName As String

        Dim start As Integer
        Dim laenge As Integer
        Dim status As String
        Dim pMarge As Double
        Dim pcolor As Object, schriftfarbe As Object
        Dim schriftgroesse As Integer
        Dim zeile As Integer
        Dim hproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim projectShape As Excel.Shape
        Dim phaseShape As Excel.Shape, milestoneShape As Excel.Shape
        Dim shpUID As String
        'Dim tmpshapes As Excel.Shapes = appInstance.ActiveSheet.shapes
        Dim worksheetShapes As Excel.Shapes
        Dim heute As Date = Date.Now
        Dim tmpShapeRange As Excel.ShapeRange
        'Dim vorlagenShape As xlNS.Shape
        Dim isMissingPhaseDefinition As Boolean = False
        Dim isMissingMilestoneDefinition As Boolean = False

        Dim shpExists As Boolean
        Dim oldAlternativeText As String = ""
        Dim isSummaryProject As Boolean


        Try

            worksheetShapes = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes

        Catch ex As Exception
            Throw New Exception("in ZeichneProjektinPlanTafel : keine Shapes Zuordnung möglich ")
        End Try

        Try
            hproj = ShowProjekte.getProject(pname)
            With hproj
                laenge = .anzahlRasterElemente
                shpUID = .shpUID
                start = .Start + .StartOffset
                pcolor = .farbe
                schriftfarbe = .Schriftfarbe
                schriftgroesse = .Schrift
                'status = .Status
                status = .vpStatus
                pMarge = .ProjectMarge
                isSummaryProject = (.projectType = ptPRPFType.portfolio)
            End With
        Catch ex As Exception
            Throw New ArgumentException("in zeichneProjektinBoard - Projektname existiert nicht: " & pname)
        End Try

        If isSummaryProject Then
            pcolor = visboFarbeOrange
            drawphases = False
            hproj.extendedView = False
        End If

        ' prüfen, ob das Shape bereits existiert ...
        If shpUID <> "" Then
            Try
                projectShape = worksheetShapes.Item(pname)
                shpExists = True
                ' merken, weil bei Variante erzeugen der Alternative Text nicht geändert werden soll 
                oldAlternativeText = projectShape.AlternativeText
            Catch ex As Exception
                shpExists = False
                projectShape = Nothing
            End Try
        Else
            shpExists = False
            projectShape = Nothing
        End If



        '
        ' ist dort überhaupt Platz ? wenn nicht, dann Zeile mit freiem Platz suchen ...
        If tryzeile < 2 Then
            tryzeile = projectboardShapes.getMaxZeile
        End If

        If useTryZeileAnyway Then
            zeile = tryzeile
        Else
            zeile = findeMagicBoardPosition(noCollection, pname, tryzeile, start, laenge)
        End If



        Dim formerEE As Boolean = appInstance.EnableEvents
        enableOnUpdate = False
        appInstance.EnableEvents = False



        If shpExists Then

            If drawphases Or hproj.extendedView Then

                ' ungroup Shape, damit die einzelnen Phasen- bzw Milestone Shapes im Zugriff sind 
                Try
                    tmpShapeRange = projectShape.Ungroup
                Catch ex As Exception
                    tmpShapeRange = Nothing
                End Try

                Dim cphase As clsPhase

                For i = 1 To hproj.CountPhases
                    cphase = hproj.getPhase(i)
                    phasenNameID = cphase.nameID

                    isMissingPhaseDefinition = Not PhaseDefinitions.Contains(cphase.name)

                    '''' tk/ur: 28.9.15 
                    '''' damit die Phase (1) gefunden werden kann.  muss bei Phase(1) der Name anders zusammengesetzt sein als bei den anderen 
                    If phasenNameID = rootPhaseName Then
                        phaseShapeName = projectboardShapes.calcPhaseShapeName(pname, phasenNameID)
                    Else
                        phaseShapeName = projectboardShapes.calcPhaseShapeName(pname, phasenNameID) & "#" & i.ToString
                    End If

                    'phaseShapeName = pname & "#" & phasenName & "#" & i.ToString

                    Try
                        phaseShape = worksheetShapes.Item(phaseShapeName)
                        Call definePhaseAppearance(hproj, cphase, 0, phaseShape, isMissingPhaseDefinition)
                    Catch ex As Exception

                    End Try

                    For r = 1 To cphase.countMilestones

                        Dim cMilestone As clsMeilenstein
                        Dim cBewertung As clsBewertung

                        cMilestone = cphase.getMilestone(r)
                        cBewertung = cMilestone.getBewertung(1)

                        isMissingMilestoneDefinition = Not MilestoneDefinitions.Contains(cMilestone.name)

                        msShapeName = projectboardShapes.calcMilestoneShapeName(hproj.name, cMilestone.nameID)

                        ' existiert das schon ? 
                        Try
                            milestoneShape = worksheetShapes.Item(msShapeName)
                            Call defineResultAppearance(hproj, 0, milestoneShape, cBewertung, isMissingMilestoneDefinition, cMilestone.farbe)
                        Catch ex As Exception

                        End Try


                    Next


                Next

                ' Gruppieren des Shapes 
                projectShape = tmpShapeRange.Group
                projectShape.Name = hproj.name



            Else

                Call defineShapeAppearance(hproj, projectShape)

            End If


        Else

            ' ///////////////
            ' Shape existiert noch nicht 
            ' ///////////////

            ' hier wird der vorher bestimmte Wert gesetzt, wo das Shape gezeichnet werden kann 
            hproj.tfZeile = zeile

            If (drawphases And (hproj.CountPhases > 1)) Or (hproj.extendedView And (hproj.CountPhases > 1)) Then
                ' stelle das Projekt im Extended Mode dar  
                'Dim shapeGroupListe() As Object
                Dim shapeGroupListe() As String
                Dim arrayOfMSNames() As String
                Dim msShapeNames As New Collection
                Dim anzGroupElemente As Integer = 0
                Dim projectShapesCollection As New Collection



                'oldShape = Nothing
                phaseShape = Nothing

                Dim zeilenOffset As Integer = 0
                Dim lastEndDate As Date = StartofCalendar.AddDays(-1)

                For i = 1 To hproj.CountPhases

                    Dim cphase As clsPhase = hproj.getPhase(i)
                    With cphase

                        phasenNameID = .nameID

                    End With

                    isMissingPhaseDefinition = Not PhaseDefinitions.Contains(cphase.name)

                    Try
                        zeilenOffset = 0
                        Call hproj.calculateShapeCoord(i, zeilenOffset, top, left, width, height)

                        If i = 1 Then

                            If awinSettings.drawProjectLine Then

                                phaseShape = worksheetShapes.AddConnector(core.MsoConnectorType.msoConnectorStraight, CSng(left), CSng(top),
                                                                            CSng(left + width), CSng(top))
                            Else

                                phaseShape = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle,
                                                        Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                            End If

                        Else

                            'ur:190725
                            Dim appear As clsAppearance = appearanceDefinitions.getPhaseAppearance(cphase)

                            Dim tmpName As String = elemNameOfElemID(phasenNameID)

                            phaseShape = worksheetShapes.AddShape(Type:=appear.shpType,
                              Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))


                        End If


                    Catch ex As Exception
                        Throw New Exception("in zeichneProjektinPlantafel2 : keine Shape-Erstellung möglich ...  ")
                    End Try

                    phaseShapeName = projectboardShapes.calcPhaseShapeName(pname, phasenNameID) & "#" & i.ToString
                    'phaseShapeName = pname & "#" & phasenName & "#" & i.ToString
                    With phaseShape
                        .Name = phaseShapeName
                        .Title = phasenNameID
                        .AlternativeText = CInt(PTshty.phaseE).ToString
                    End With



                    If i = 1 Then
                        Call defineShapeAppearance(hproj, phaseShape)
                    Else
                        Call definePhaseAppearance(hproj, cphase, 0, phaseShape, isMissingPhaseDefinition)
                    End If


                    ' jetzt der Liste der ProjectboardShapes hinzufügen
                    projectboardShapes.add(phaseShape)

                    Try
                        projectShapesCollection.Add(phaseShapeName, Key:=phaseShapeName)
                    Catch ex As Exception

                    End Try


                    ' jetzt müssen alle Meilensteine dieser Phase gezeichnet werden 

                    With CType(hproj.getPhase(i), clsPhase)
                        Dim msName As String
                        Dim msShape As Excel.Shape

                        For r = 1 To .countMilestones

                            Dim cMilestone As clsMeilenstein
                            Dim cBewertung As clsBewertung

                            cMilestone = .getMilestone(r)
                            cBewertung = cMilestone.getBewertung(1)

                            isMissingMilestoneDefinition = Not MilestoneDefinitions.Contains(cMilestone.name)

                            ' Änderung tk 26.11.15
                            'If MilestoneDefinitions.Contains(cMilestone.name) Then
                            '    vorlagenShape = MilestoneDefinitions.getShape(cMilestone.name)
                            'Else
                            '    vorlagenShape = missingMilestoneDefinitions.getShape(cMilestone.name)
                            'End If
                            'Dim factorB2H As Double = vorlagenShape.Width / vorlagenShape.Height

                            'ur:190725

                            Dim appear As clsAppearance = appearanceDefinitions.getMileStoneAppearance(cMilestone)

                            Dim factorB2H As Double = appear.width / appear.height


                            hproj.calculateMilestoneCoord(cMilestone.getDate, zeilenOffset, factorB2H, top, left, width, height)

                            msName = projectboardShapes.calcMilestoneShapeName(hproj.name, cMilestone.nameID)
                            'msName = hproj.name & "#" & .name & "#M" & r.ToString
                            ' existiert das schon ? 
                            Try
                                msShape = worksheetShapes.Item(msName)
                            Catch ex As Exception
                                msShape = Nothing
                            End Try

                            If msShape Is Nothing Then


                                msShape = worksheetShapes.AddShape(Type:=appear.shpType,
                                                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                'ur:190725
                                'msShape = worksheetShapes.AddShape(Type:=vorlagenShape.AutoShapeType,
                                '                                Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))
                                'vorlagenShape.PickUp()
                                'msShape.Apply()

                                With msShape
                                    .Name = msName
                                    .Title = cMilestone.nameID
                                    .AlternativeText = CInt(PTshty.milestoneE).ToString
                                End With


                                ' tk 24.3.2015 um nachher die Milestone Shapes nach vorne zu holen
                                If Not msShapeNames.Contains(msName) Then
                                    msShapeNames.Add(msName, msName)
                                End If

                                Call defineResultAppearance(hproj, 0, msShape, cBewertung, isMissingMilestoneDefinition,
                                                            cMilestone.farbe, appear)

                                ' jetzt der Liste der ProjectboardShapes hinzufügen
                                projectboardShapes.add(msShape)

                            Else
                                ' Koordinaten anpassen 
                                msShape.Top = CSng(top)
                            End If

                            Try
                                projectShapesCollection.Add(msName, Key:=msName)
                            Catch ex As Exception

                            End Try

                        Next


                    End With

                Next

                ' Änderung tk 24.3.2015
                ' jetzt müssen ggf die Meilensteine noch nach vorne gebracht werden ...
                Dim anzElements As Integer
                anzElements = msShapeNames.Count

                If anzElements > 0 Then

                    ReDim arrayOfMSNames(anzElements - 1)
                    For ix = 1 To anzElements
                        arrayOfMSNames(ix - 1) = CStr(msShapeNames.Item(ix))
                    Next

                    Try
                        CType(worksheetShapes.Range(arrayOfMSNames), Excel.ShapeRange).ZOrder(core.MsoZOrderCmd.msoBringToFront)
                    Catch ex As Exception

                    End Try

                End If


                ' hier werden die Shapes gruppiert
                anzGroupElemente = projectShapesCollection.Count

                If anzGroupElemente > 1 Then
                    ' es macht nur Sinn zu gruppieren, wenn es mehr als 1 Element ist ....

                    ReDim shapeGroupListe(anzGroupElemente - 1)
                    For i = 1 To anzGroupElemente
                        shapeGroupListe(i - 1) = CStr(projectShapesCollection.Item(i))
                    Next

                    Dim ShapeGroup As Excel.ShapeRange
                    ShapeGroup = worksheetShapes.Range(shapeGroupListe)
                    projectShape = ShapeGroup.Group()

                Else
                    ' in diesem Fall besteht das Projekt nur aus einer einzigen Phase
                    projectShape = phaseShape

                End If
                projectShape.Name = pname


            Else
                ' stelle das Projekt im Einzeilen Modus dar

                With hproj
                    .tfZeile = zeile ' calculateShapeCoord verwendet .tfzeile ! 
                    .CalculateShapeCoord(top, left, width, height)
                End With

                'If awinSettings.drawProjectLine Then
                If awinSettings.drawProjectLine Then

                    projectShape = worksheetShapes.AddConnector(core.MsoConnectorType.msoConnectorStraight, CSng(left), CSng(top),
                                                                CSng(left + width), CSng(top))

                    projectShape.AlternativeText = CInt(PTshty.projektL).ToString

                Else
                    projectShape = worksheetShapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle,
                        Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))

                End If

                projectShape.Name = pname
                Call defineShapeAppearance(hproj, projectShape)

            End If


        End If

        With projectShape
            If shpExists Then
                .AlternativeText = oldAlternativeText
            Else
                If hproj.extendedView Or drawphases Then
                    .AlternativeText = CInt(PTshty.projektE).ToString
                Else
                    If awinSettings.drawProjectLine Then
                        .AlternativeText = CInt(PTshty.projektL).ToString
                    Else
                        .AlternativeText = CInt(PTshty.projektN).ToString
                    End If
                End If
            End If


            hproj.shpUID = .ID.ToString
            hproj.tfZeile = calcYCoordToZeile(projectShape.Top)
        End With

        ' jetzt der Liste der ProjectboardShapes hinzufügen
        projectboardShapes.add(projectShape)

        ' jetzt muss das neue Shape in der ShowProjekte.ShapeListe eingetragen werden ..
        ShowProjekte.AddShape(pname, shpUID:=projectShape.ID.ToString)


        ' zu guter Letzt, 
        ' aber noch vor den Meilensteinen und Phasen  muss der Projekt-Name gezeichnet werden ; 
        If awinSettings.drawProjectLine Then
            Call zeichneNameInProjekt(hproj)
        End If

        ' jetzt müssen ggf die noch zu zeichnenden Meilensteine und Phasen eingezeichnet werden  
        ' das wird jetzt nach zeichneNameInProjekt gemacht, damit die Meilensteine / Phasen nicht vom Projekt-Namen überdeckt werden 

        Dim msNumber As Integer = 0
        If drawPhaseList.Count > 0 And Not (drawphases Or hproj.extendedView) Then
            Call zeichnePhasenInProjekt(hproj:=hproj, namenListe:=drawPhaseList, numberIt:=False, msNumber:=msNumber,
                                        vonMonth:=showRangeLeft, bisMonth:=showRangeRight)
            'Call zeichnePhasenInProjekt(hproj, drawPhaseList, False, msNumber)
        End If

        msNumber = 0
        If drawMilestoneList.Count > 0 And Not (drawphases Or hproj.extendedView) Then
            Call zeichneMilestonesInProjekt(hproj, drawMilestoneList, 4, showRangeLeft, showRangeRight, False, msNumber, False)
            'Call zeichneMilestonesInProjekt(hproj, drawMilestoneList, 4, 0, 0, False, msNumber, False)
        End If


        'If roentgenBlick.isOn Then
        '    With roentgenBlick
        '        Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
        '    End With
        'End If

        If drawPhaseList.Count = 0 And drawMilestoneList.Count = 0 Then
            ' jetzt müssen die Charts, die vom Projekt evtl überdeckt werden in den Vordergrund geholt werden 
            ' das muss jedoch nur gemacht werden, wenn nicht vorher schon zeichnePhasenInProjekt oder zeichneMilestonesInProjekt aufgerufen wurde 
            Call bringChartsToFront(projectShape)
        End If



        appInstance.EnableEvents = formerEE
        enableOnUpdate = True

    End Sub


    ''' <summary>
    ''' zeichnet für das Projekt das Status Shape; wenn es bereits existiert, wird das alte gelöscht, das neue gezeichnet 
    ''' wenn number > 0 , wird diese Zahl in das Symbol geschrieben 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="number"></param>
    ''' <remarks></remarks>
    Public Sub zeichneStatusSymbolInPlantafel(ByVal hproj As clsProjekt, ByVal number As Integer)
        Dim top As Double, left As Double, height As Double, width As Double
        Dim worksheetShapes As Excel.Shapes
        Dim statusShape As Excel.Shape
        Dim shpName As String
        Dim timeAtStatus As Date = hproj.timeStamp
        Dim heuteColumn As Integer = getColumnOfDate(timeAtStatus)

        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)
            worksheetShapes = .Shapes

            shpName = projectboardShapes.calcStatusShapeName(hproj.name, heuteColumn)
            ' existiert das schon ? 
            Try
                statusShape = worksheetShapes.Item(shpName)
                statusShape.Delete()
            Catch ex As Exception
                'statusShape = Nothing
            End Try

            'If statusShape Is Nothing Then

            hproj.calculateStatusCoord(timeAtStatus, top, left, width, height)
            statusShape = .Shapes.AddShape(Type:=Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                                            Left:=CSng(left), Top:=CSng(top), Width:=CSng(width), Height:=CSng(height))

            With statusShape
                .Name = shpName
                .Title = "Status"
                .AlternativeText = CInt(PTshty.status).ToString
            End With

            Call defineStatusAppearance(hproj, number, statusShape)

            ' jetzt der Liste der ProjectboardShapes hinzufügen
            projectboardShapes.add(statusShape)

            'shapesCollection.Add(resultShape.Name)

            'End If


        End With

        ' jetzt müssen die Charts ggf wieder nach vorne gebracht werden 
        Call bringChartsToFront(statusShape)


    End Sub


    ''' <summary>
    ''' zeichnet für alle selektierten Projekte die Phasen, die in Namelist angegeben sind;  
    ''' wenn namelist leer ist, werden alle Phasen des Projektes angezeigt
    ''' </summary>
    ''' <param name="nameList">enthält die Namen plus die Breadcrumbs der Phasen, die gezeichnet werden sollen; alle, wenn leer</param>
    ''' <param name="numberIt">gibt an, ob di ePhasen nummeriert werden sollen</param>
    ''' <param name="deleteOtherShapes">gibt an, ob die anderen Phasen-Shapes gelöscht werden sollen</param>
    ''' <remarks></remarks>
    Public Sub awinZeichnePhasen(ByVal nameList As Collection, ByVal numberIt As Boolean, ByVal deleteOtherShapes As Boolean)

        'Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim hproj As New clsProjekt
        Dim vglName As String = " "
        Dim pName As String
        Dim ok As Boolean = True
        Dim msNumber As Integer = 1

        Dim awinSelection As Excel.ShapeRange

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = True

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try


        If Not awinSelection Is Nothing Then

            Dim anzSelect As Integer = awinSelection.Count

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                ok = True
                With singleShp
                    If isProjectType(kindOfShape(singleShp)) Then


                        Try
                            hproj = ShowProjekte.getProject(singleShp.Name, True)
                        Catch ex As Exception
                            ok = False
                        End Try

                        If ok Then

                            If deleteOtherShapes Then
                                Call awinDeleteProjectChildShapes(singleShp, 3)
                            End If

                            Try
                                pName = hproj.name
                                Call zeichnePhasenInProjekt(hproj, nameList, False, msNumber)

                            Catch ex As Exception

                            End Try


                        End If

                    End If
                End With
            Next

            If msNumber = 1 Then
                If nameList.Count > 1 Then
                    Call MsgBox("Auswahl enthält diese Phasen nicht")
                Else
                    Call MsgBox("Auswahl enthält diese Phase nicht:  " & nameList.Item(1))
                End If
            End If

        Else
            ' tue es für alle Projekte in Showprojekte 


            Dim todoListe As New SortedList(Of Long, clsProjekt)
            Dim key As Long

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                todoListe.Add(key, kvp.Value)

            Next


            For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

                If deleteOtherShapes Then
                    singleShp = ShowProjekte.getShape(kvp.Value.name)
                    Call awinDeleteProjectChildShapes(singleShp, 3)
                End If

                Call zeichnePhasenInProjekt(kvp.Value, nameList, False, msNumber, showRangeLeft, showRangeRight)

            Next


            If msNumber = 1 Then
                If nameList.Count > 1 Then
                    Call MsgBox("im gewählten Zeitraum gibt es diese Phasen nicht")
                Else
                    Call MsgBox("im gewählten Zeitraum gibt es diese Phase nicht: " & nameList.Item(1))
                End If
            End If


        End If


        'ur: 17.7.2015: für PlanElemente visualisieren für Einzelprojekt-Info sollte nach zeichnen der Phasen nicht deselektiert werden
        '' ''Call awinDeSelect()


        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU



    End Sub


    ''' <summary>
    ''' zeichnet für interaktiven wie Report Modus die Milestones 
    ''' 0: grau, 1: grün, 2: gelb, 3:rot, 4: alle
    ''' </summary>
    ''' <param name="farbTyp">welcher Typus soll gezeichnet werden </param>
    ''' <remarks></remarks>
    Public Sub awinZeichneMilestones(ByVal nameList As Collection, ByVal farbTyp As Integer, ByVal numberIt As Boolean, ByVal deleteOtherShapes As Boolean)

        'Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As Excel.Shape
        Dim hproj As New clsProjekt
        Dim vglName As String = " "
        Dim pName As String
        Dim ok As Boolean = True
        Dim msNumber As Integer = 1

        Dim awinSelection As Excel.ShapeRange

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                ok = True
                With singleShp

                    If isProjectType(kindOfShape(singleShp)) Then

                        Try
                            hproj = ShowProjekte.getProject(singleShp.Name, True)
                        Catch ex As Exception
                            ok = False
                        End Try

                        If ok Then

                            If deleteOtherShapes Then
                                Call awinDeleteProjectChildShapes(singleShp, 1)
                            End If

                            Try
                                pName = hproj.name
                                Call zeichneMilestonesInProjekt(hproj, nameList, farbTyp, 0, 0, False, msNumber, False)
                            Catch ex As Exception
                                Dim a As Integer = 0
                            End Try


                        End If

                    End If
                End With
            Next

            If msNumber = 1 Then
                If nameList.Count > 1 Then
                    Call MsgBox("Auswahl enthält  diese Meilensteine nicht")
                ElseIf nameList.Count = 1 Then
                    Call MsgBox("Auswahl enthält keinen Meilenstein " & nameList.Item(1))

                End If
            End If

        Else


            If ShowProjekte.Count > 0 Then

                ' tue es für alle Projekte in Showprojekte 


                Dim todoListe As New SortedList(Of Long, clsProjekt)
                Dim key As Long

                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
                    todoListe.Add(key, kvp.Value)

                Next


                For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

                    If deleteOtherShapes Then
                        singleShp = ShowProjekte.getShape(kvp.Value.name)
                        Call awinDeleteProjectChildShapes(singleShp, 1)
                    End If

                    Call zeichneMilestonesInProjekt(kvp.Value, nameList, farbTyp, showRangeLeft, showRangeRight, numberIt, msNumber, False)

                Next


                If msNumber = 1 Then
                    If nameList.Count > 1 Then
                        Call MsgBox("im gewählten Zeitraum gibt es diese Meilensteine nicht")
                    ElseIf nameList.Count = 1 Then
                        Call MsgBox("im gewählten Zeitraum gibt es keinen Meilenstein " & nameList.Item(1))
                    End If
                End If

            Else
                Call MsgBox("Es sind keine Projekte geladen!")
            End If


        End If

        'ur: 17.7.2015: für PlanElemente visualisieren sollte nach zeichnen der Meilensteine nicht deselektiert werden
        '' ''Call awinDeSelect()

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU



    End Sub


    ''' <summary>
    ''' löscht das Projectboard und zeichnet dann die Constellation 
    ''' zeichnet das Projectboard, baut es von Scratch an auf ... 
    ''' in AlleProjekte sind die zu zeichnenden Projekte drin ...
    ''' </summary>
    Public Sub awinZeichnePlanTafel(ByVal myConstellation As clsConstellation)


        ' wenn das nicht so aufgesetzt, dann werden alle aktuellen Constelaltions-Items im Attribut auf false gesetzt ...
        Call clearProjectBoard(updateCurrentConstellation:=False)

        ' wenn Nothing übergeben wird, dann wird die currentConstellation gezeichnet

        Dim zeile As Integer = 2
        Dim hproj As clsProjekt = Nothing

        Dim listOfPnames As String() = myConstellation.sortListe.Values.ToArray

        For Each tmpPname As String In listOfPnames

            Dim cItem As clsConstellationItem = myConstellation.getShownItem(tmpPname)

            If Not IsNothing(cItem) Then
                ' diese Abfrage ist eigentlich redundant, aber sicher ist sicher 
                If cItem.show Then
                    Dim key As String = calcProjektKey(cItem.projectName, cItem.variantName)
                    hproj = AlleProjekte.getProject(key)

                    hproj.tfZeile = zeile

                    ShowProjekte.Add(hproj)

                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, zeile, tmpCollection, tmpCollection)

                    ' zeile soviel weiterschalten, wie Platz benötigt wird ...
                    zeile = zeile + 1


                End If

            End If


        Next


    End Sub


    ''' <summary>
    ''' setzt alle angezeigten Projekte, also ShowProjekte,  zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearProjectBoard(Optional ByVal updateCurrentConstellation As Boolean = True)

        Call awinClearPlanTafel()

        ShowProjekte.Clear(updateCurrentConstellation)
        projectboardShapes.clear()

        selectedProjekte.Clear(False)
        ImportProjekte.Clear(False)


    End Sub

    Public Sub awinZeichnePlanTafel(ByVal fromScratch As Boolean)

        Dim todoListe As New SortedList(Of Double, String)
        Dim key As Double
        Dim pname As String

        Dim lastZeileOld As Integer
        Dim hproj As clsProjekt
        Dim positionsKennzahl As Double

        Dim notOK As Boolean = True
        Dim tryExceptionCounts As Integer = 0


        If fromScratch Then
            Dim zeile As Integer
            Dim lastBU As String = ""

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                notOK = True

                With kvp.Value

                    'positionsKennzahl = calcKennziffer(kvp.Value)
                    If projectConstellations.Contains(currentConstellationPvName) Then
                        positionsKennzahl = projectConstellations.getConstellation(currentConstellationPvName).getBoardZeile(kvp.Key)
                    Else
                        positionsKennzahl = currentSessionConstellation.getBoardZeile(kvp.Key)
                    End If


                    Do While notOK
                        Try
                            If todoListe.ContainsKey(positionsKennzahl) Then
                                positionsKennzahl = positionsKennzahl + 0.00001
                            Else
                                todoListe.Add(positionsKennzahl, .name)
                                notOK = False
                            End If
                        Catch ex As Exception
                            positionsKennzahl = positionsKennzahl + 0.00001
                            tryExceptionCounts = tryExceptionCounts + 1
                        End Try
                    Loop


                End With

            Next

            zeile = 2
            Dim i As Integer

            For i = 1 To todoListe.Count

                pname = todoListe.ElementAt(i - 1).Value

                Try
                    hproj = ShowProjekte.getProject(pname)

                    hproj.tfZeile = zeile

                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, pname, zeile, tmpCollection, tmpCollection)

                    ' zeile soviel weiterschalten, wie Platz benötigt wird ...
                    zeile = zeile + hproj.calcNeededLines(tmpCollection, tmpCollection, hproj.extendedView Or awinSettings.drawphases, False)

                Catch ex As Exception
                    tryExceptionCounts = tryExceptionCounts + 1
                End Try

            Next


        Else

            Dim zeile As Integer, lastzeile As Integer, curzeile As Integer, max As Integer
            ' so wurde es bisher gemacht ... bis zum 17.1.15
            ' aufbauen der todoListe, so daß nachher die Projekte von oben nach unten gezeichnet werden können 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                Try
                    With kvp.Value
                        key = 10000 * .tfZeile + kvp.Value.Start
                        Do While todoListe.ContainsKey(key)
                            key = key + 0.000001
                        Loop
                        todoListe.Add(key, .name)
                    End With
                Catch ex As Exception

                    tryExceptionCounts = tryExceptionCounts + 1
                    'Call MsgBox("Fehler in awinZeichnePlanTafel")

                End Try


            Next

            zeile = 2
            lastzeile = 0


            'If ProjectBoardDefinitions.My.Settings.drawPhases = True Then
            ' dann sollen die Projekte im extended mode gezeichnet werden 
            ' jetzt erst mal die Konstellation "last" speichern
            ' 3.11.14 Auskommentiert: Zeichnen sollte nichts zu tun haben mit dem Verwalten von Konstellationen 
            ' Call storeSessionConstellation(ShowProjekte, "Last")

            ' jetzt die todoListe abarbeiten
            Dim i As Integer
            For i = 1 To todoListe.Count
                pname = todoListe.ElementAt(i - 1).Value

                Try
                    hproj = ShowProjekte.getProject(pname)

                    If i = 1 Then
                        curzeile = hproj.tfZeile
                        lastZeileOld = hproj.tfZeile
                        lastzeile = curzeile
                        max = curzeile
                    Else
                        If lastZeileOld = hproj.tfZeile Then
                            curzeile = lastzeile
                        Else
                            lastzeile = max
                            lastZeileOld = hproj.tfZeile
                        End If

                    End If

                    ' Änderung 9.10.14, damit die Spaces in einer 
                    'If hproj.tfZeile >= curZeile + 1 Then
                    '    curZeile = curZeile + 1
                    'End If
                    ' Ende Änderung
                    hproj.tfZeile = curzeile
                    lastzeile = curzeile
                    'Call ZeichneProjektinPlanTafel2(pname, curZeile)
                    ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                    ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, pname, curzeile, tmpCollection, tmpCollection)
                    curzeile = lastzeile + hproj.calcNeededLines(tmpCollection, tmpCollection, hproj.extendedView Or awinSettings.drawphases, False)


                    If curzeile > max Then
                        max = curzeile
                    End If
                Catch ex As Exception
                    tryExceptionCounts = tryExceptionCounts + 1
                End Try



            Next
        End If

        'If tryExceptionCounts > 0 Then
        '    Call MsgBox("Anzahl: " & tryExceptionCounts)
        'End If

        'Call MsgBox("Ende: " & Date.Now.TimeOfDay.ToString)

    End Sub

    ''' <summary>
    ''' löscht die zeicherische Darstellung des Projektes auf der Plantafel 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearProjektinPlantafel(ByVal pname As String)
        Dim eeWasTrue As Boolean = False
        Dim suWasTrue As Boolean = False
        'Dim XPos As Integer, YPos As Integer
        'Dim laenge As Integer
        'Dim tmpshapes As Excel.Shapes = appInstance.ActiveSheet.shapes
        Dim tmpshapes As Excel.Shapes = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes
        Dim shpelement As Excel.Shape

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        'Dim formerSU As Boolean = appInstance.ScreenUpdating
        'appInstance.ScreenUpdating = False

        ' Lösche das Shape Element
        Try
            shpelement = tmpshapes.Item(pname)
            With shpelement
                projectboardShapes.remove(shpelement)
            End With
        Catch ex As Exception

        End Try

        Try

            If ShowProjekte.contains(pname) Then
                Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
                Dim shpuid As String = hproj.shpUID
                hproj.shpUID = ""
                ShowProjekte.shpListe.Remove(shpuid)
            End If

        Catch ex As Exception

        End Try



        appInstance.EnableEvents = formerEE
        'appInstance.ScreenUpdating = formerSU


    End Sub
End Module
