Imports System.Math
Imports xlNS = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports Microsoft.Office.Core
'Imports System.Int32

Public Class clsEventsPrcCharts


    Public WithEvents PrcChartEvents As xlNS.Chart
    Public WithEvents PrcChartRightClick As CommandBarButton


    Private Sub PrcChartEvents_Activate() Handles PrcChartEvents.Activate

        Dim chtobj As xlNS.ChartObject

        Try
            chtobj = CType(Me.PrcChartEvents.Parent, Microsoft.Office.Interop.Excel.ChartObject)
            If selectedCharts.Contains(chtobj.Name) Then
                ' nichst tun 
            Else
                ' aufnehmen 
                selectedCharts.Add(chtobj.Name, chtobj.Name)
            End If
        Catch ex As Exception

        End Try
        If selectedProjekte.Count > 0 Then
            selectedProjekte.Clear(False)
            Call awinNeuZeichnenDiagramme(8)
        End If


    End Sub


    Private Sub PrcChartEvents_BeforeDoubleClick(ByVal ElementID As Integer, ByVal Arg1 As Integer, ByVal Arg2 As Integer, ByRef Cancel As Boolean) Handles PrcChartEvents.BeforeDoubleClick


        Dim chtobj As xlNS.ChartObject
        Dim IDKennung As String
        Dim foundDiagram As clsDiagramm = Nothing


        Cancel = True

        Try
            chtobj = CType(Me.PrcChartEvents.Parent, Microsoft.Office.Interop.Excel.ChartObject)

            IDKennung = chtobj.Name
            If DiagramList.contains(IDKennung) Then
                foundDiagram = DiagramList.getDiagramm(IDKennung)
                With foundDiagram
                    .top = chtobj.Top
                    .left = chtobj.Left
                    .width = chtobj.Width
                    .height = chtobj.Height
                End With
            End If
            
        Catch ex As Exception
            'Call MsgBox("konnte das Chart nicht in der Diagramm-Liste finden ...")
        End Try


    End Sub

    Private Sub PrcChartEvents_BeforeRightClick(ByRef Cancel As Boolean) Handles PrcChartEvents.BeforeRightClick

        ' tk Änderung 19.1.15 ; Right Click de-aktiviert 
        'Cancel = True
        'appInstance.CommandBars("awinRightClickinPRCChart").ShowPopup()

    End Sub


    Private Sub PrcChartRightClick_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles PrcChartRightClick.Click


        Dim diagrammTyp As String = " "
        Dim myCollection As New Collection
        'Dim i As Integer
        Dim found As Boolean
        Dim rcIdentifier As Integer
        Dim chtobj As xlNS.ChartObject
        Dim isCC As Boolean
        Dim bestaetigeOptimierung As New frmconfirmOptimierung
        Dim returnValue As DialogResult
        'Dim chtTitle As String
        Dim name As String = ""
        Dim IDkennung As String
        Dim foundDiagramm As clsDiagramm


        CancelDefault = True
        rcIdentifier = 0

        Try
            chtobj = CType(appInstance.ActiveChart.Parent, Microsoft.Office.Interop.Excel.ChartObject)
            IDkennung = chtobj.Name
            foundDiagramm = DiagramList.getDiagramm(IDkennung)
            If Not IsNothing(foundDiagramm) Then
                found = True
                With foundDiagramm
                    diagrammTyp = .diagrammTyp
                    myCollection = .gsCollection
                    isCC = .isCockpitChart
                End With
            End If
            


            If IsNothing(myCollection) Then
                myCollection = New Collection
                myCollection.Add("Ergebniskennzahl")
            Else
                If myCollection.Count < 1 Then
                    name = ""
                ElseIf myCollection.Count = 1 Then
                    name = CStr(myCollection.Item(1))
                ElseIf myCollection.Count > 1 Then
                    name = "Collection"
                End If
            End If
            


        Catch ex As Exception

            Call MsgBox("Diagramm wurde nicht erkannt ... PRCChartRightClick .... Abbruch")
            Exit Sub
        End Try

        'Try
        '    chtTitle = chtobj.Chart.ChartTitle.Text

        'Catch ex As Exception
        '    chtTitle = " "
        'End Try

        'i = 1
        'found = False

        'While i <= DiagramList.Count And Not found

        '    If chtTitle Like (DiagramList.getDiagramm(i).DiagrammTitel & "*") Then

        '        found = True



        '        'If diagrammTyp = DiagrammTypen(1) Then ' Rolle
        '        '    rcIdentifier = RoleDefinitions.getRoledef(myCollection.Item(1)).UID
        '        '    'selectedRoleNeeds = rcIdentifier
        '        '    'selectedCostNeeds = 0
        '        'ElseIf diagrammTyp = DiagrammTypen(2) Then ' Kostenart
        '        '    rcIdentifier = RoleDefinitions.Count + CostDefinitions.getCostdef(myCollection.Item(1)).UID
        '        '    'selectedCostNeeds = rcIdentifier
        '        '    'selectedRoleNeeds = 0
        '        'ElseIf diagrammTyp = DiagrammTypen(4) Then ' Ergebnis
        '        '    rcIdentifier = RoleDefinitions.Count + CostDefinitions.Count + 1
        '        '    'selectedCostNeeds = rcIdentifier
        '        '    'selectedRoleNeeds = 0
        '        'End If

        '    Else
        '        i = i + 1
        '    End If
        'End While

        'If Not found Then
        '    Call MsgBox("Diagramm wurde nicht erkannt ... PRCChartRightClick .... Abbruch")
        '    Exit Sub
        'End If


        Select Case Ctrl.Tag

            Case "Löschen"
                Call awinDeleteChart(chtobj)
                chtobj = Nothing

            Case "Bedarf anzeigen"
                ' wenn der gleiche bereits gezeigt wird: wieder ausschalten 

                If diagrammTyp = DiagrammTypen(0) Then

                    Call MsgBox("für diesen Diagramm-Typ nutzen Sie bitte die Funktion " & vbLf & _
                                 "Visualisieren - Phasen")

                ElseIf name <> "" Then
                    Dim screenUpdateFormerState As Boolean = appInstance.ScreenUpdating
                    appInstance.ScreenUpdating = False

                    With roentgenBlick
                        If .isOn And .name = name And .type = diagrammTyp Then
                            .isOn = False
                            .name = ""
                            .myCollection = Nothing
                            .type = ""
                            Call awinNoshowProjectNeeds()
                        Else
                            If .isOn Then
                                Call awinNoshowProjectNeeds()
                            End If
                            .isOn = True
                            .name = name
                            .myCollection = myCollection
                            .type = diagrammTyp
                            Call awinShowProjectNeeds1(myCollection, diagrammTyp)
                        End If
                    End With


                    appInstance.ScreenUpdating = screenUpdateFormerState
                Else
                    Call MsgBox("für dieses Diagramm ist der Röntgenblick nicht verfügbar")
                End If

            Case "Varianten optimieren"

                enableOnUpdate = False
                'Call awinCalcOptimizationVarianten(diagrammTyp, myCollection)
                enableOnUpdate = False


            Case "Optimieren"
                ' hier werden die Werte der Optimierung vermerkt: welche Projekte müssen verschoben werden , um welchen Offset 

                Dim OptimierungsErgebnis As New SortedList(Of String, clsOptimizationObject)
                'Dim shpElement As Microsoft.Office.Interop.Excel.Shape

                enableOnUpdate = False

                Call awinCalcOptimizationFreiheitsgrade(diagrammTyp, myCollection, OptimierungsErgebnis)

                If OptimierungsErgebnis.Count > 0 Then


                    returnValue = bestaetigeOptimierung.ShowDialog
                    If returnValue = DialogResult.OK Then

                        Call ClearPlanTafelfromOptArrows()

                        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                            Try

                                With kvp.Value

                                    If .StartOffset <> 0 And .Status = ProjektStatus(0) Then
                                        .startDate = .startDate.AddMonths(.StartOffset)

                                        If .StartOffset < 0 Then
                                            .earliestStart = .earliestStart - .StartOffset
                                            .latestStart = .latestStart - .StartOffset
                                        Else
                                            .latestStart = .latestStart - .StartOffset
                                            .earliestStart = .earliestStart - .StartOffset
                                        End If
                                        .StartOffset = 0

                                        If Not .isConsistent Then
                                            Call .syncXWertePhases()
                                        End If

                                        
                                        Dim phaseList As Collection = projectboardShapes.getPhaseList(.name)
                                        Dim milestoneList As Collection = projectboardShapes.getMilestoneList(.name)

                                        Call clearProjektinPlantafel(.name)

                                        ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                                        ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                                        Dim tmpCollection As New Collection
                                        Call ZeichneProjektinPlanTafel(tmpCollection, .name, .tfZeile, phaseList, milestoneList)
                                    End If

                                End With
                            Catch ex As Exception
                                Call MsgBox("Projekt: " & kvp.Key & " : Startzeitpunkt liegt in der Vergangenheit ")
                            End Try

                        Next

                        Call visualisiereErgebnis()
                        OptimierungsErgebnis.Clear()
                    Else

                        Call ClearPlanTafelfromOptArrows()

                        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                            With kvp.Value
                                .StartOffset = 0
                            End With
                        Next

                        Call visualisiereErgebnis()
                        OptimierungsErgebnis.Clear()
                    End If

                Else
                    MsgBox("es waren keine Verbesserungen zu erzielen")
                End If

                enableOnUpdate = True

        End Select



    End Sub



    Private Sub PrcChartEvents_MouseUp(Button As Integer, Shift As Integer, x As Integer, y As Integer) Handles PrcChartEvents.MouseUp
        
    End Sub



    Private Sub PrcChartEvents_Resize() Handles PrcChartEvents.Resize


        Dim chtobj As xlNS.ChartObject, chtobj1 As xlNS.ChartObject
        Dim IDKennung As String
        Dim foundDiagram As clsDiagramm
        Dim kFontsize As Double
        Dim achsenFontsize As Double
        Dim axisTitleFontsize As Double
        Try
            chtobj = CType(Me.PrcChartEvents.Parent, Microsoft.Office.Interop.Excel.ChartObject)
            Try
                chtobj1 = CType(appInstance.ActiveChart.Parent, Microsoft.Office.Interop.Excel.ChartObject)
            Catch ex As Exception

            End Try

            IDKennung = chtobj.Name
            foundDiagram = DiagramList.getDiagramm(IDKennung)

            kFontsize = (chtobj.Width / foundDiagram.width)
            'kHeight = (chtobj.Height / foundDiagram.height)

            With chtobj.Chart

                ' Schriftgröße der Überschrift anpassen
                If .HasTitle Then
                    .ChartTitle.Format.TextFrame2.TextRange.Font.Size = CType(.ChartTitle.Format.TextFrame2.TextRange.Font.Size * kFontsize, Single)
                End If

                ' Schriftgröße der Legende anpassen
                If .HasLegend Then
                    With .Legend
                        .Format.TextFrame2.TextRange.Font.Size = CType(.Format.TextFrame2.TextRange.Font.Size * kFontsize, Single)
                    End With
                End If

                ' Schriftgröße der x-Achse anpassen
                Try
                    With CType(.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary), Excel.Axis)
                        achsenFontsize = CType(.Ticklabels.Font.Size * kFontsize, Double)
                        .TickLabels.Font.Size = .TickLabels.Font.Size * kFontsize
                        If .HasTitle Then
                            With .AxisTitle
                                axisTitleFontsize = CType(.Characters.Font.Size * kFontsize, Double)
                                .Characters.Font.Size = .Characters.Font.Size * kFontsize
                            End With
                        End If
                    End With
                Catch ex As Exception

                End Try


                ' Schriftgröße der y-Achse anpassen
                Try
                    With CType(.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary), Excel.Axis)
                        '.TickLabels.Font.Size = .Ticklabels.Font.Size * kFontsize
                        .TickLabels.Font.Size = achsenFontsize
                        If .HasTitle Then
                            With .AxisTitle
                                .Characters.Font.Size = axisTitleFontsize
                            End With
                        End If
                     
                    End With
                Catch ex As Exception

                End Try

                ' Schriftgröße der eingezeichenten Daten bestimmen
                If .SeriesCollection.Count > 0 Then

                    For j = 1 To .SeriesCollection.Count

                        With .SeriesCollection(j)
                            If .HasDataLabels = True Then
                                .ApplyDataLabels()
                                For i = 1 To .Points.count
                                    With .Points(i)
                                        .DataLabel.Font.Size = .DataLabel.Font.Size * kFontsize
                                    End With
                                Next i
                            End If

                        End With

                    Next j

                End If

            End With


            With foundDiagram
                .top = chtobj.Top
                .left = chtobj.Left
                .width = chtobj.Width
                .height = chtobj.Height
            End With
        Catch ex As Exception
            'Call MsgBox("konnte das Chart nicht in der Diagramm-Liste finden ...")
        End Try

        enableOnUpdate = True

    End Sub


    Private Sub PrcChartEvents_Select(ElementID As Integer, Arg1 As Integer, Arg2 As Integer) Handles PrcChartEvents.Select


        Dim chtobjname As String
        Dim diagOBJ As clsDiagramm
        Dim msNumber As Integer = 1
        Dim chtobj As xlNS.ChartObject

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False

        Try
            chtobjname = CType(Me.PrcChartEvents.Parent, Microsoft.Office.Interop.Excel.ChartObject).Name

            chtobj = CType(Me.PrcChartEvents.Parent, Microsoft.Office.Interop.Excel.ChartObject)
            Dim IDKennung As String
            IDKennung = chtobj.Name


            diagOBJ = DiagramList.getDiagramm(chtobjname)




            If (ElementID = xlNS.XlChartItem.xlSeries) And Arg2 > 0 Then
                ' hier wird der farbige Balken gezeichnet 
                Dim selMonth As Integer = showRangeLeft + Arg2 - 1

                Call awinShowSelectedMonth(selMonth)


            ElseIf (ElementID = xlNS.XlChartItem.xlSeries) And Arg2 = -1 Then

                ' ggf Röntgenblick einschalten 
                ' jetzt sind alle Balken selektiert 
                ' im Falle Rolle / Kostenarten wird jetzt Röntgenblick eingeschaltet 

                Try

                    ' zeichne Phasen
                    If diagOBJ.diagrammTyp = DiagrammTypen(0) And diagOBJ.gsCollection.Count > 0 Then

                        Call awinDeleteProjectChildShapes(3)

                        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                            Call zeichnePhasenInProjekt(kvp.Value, diagOBJ.gsCollection, False, msNumber, showRangeLeft, showRangeRight)

                        Next


                    ElseIf (diagOBJ.diagrammTyp = DiagrammTypen(1) Or diagOBJ.diagrammTyp = DiagrammTypen(2)) And _
                        diagOBJ.gsCollection.Count > 0 Then

                        ' zeichne Rollen oder Kostenarten


                        Dim name As String = ""

                        If diagOBJ.gsCollection.Count < 1 Then
                            name = ""
                        ElseIf diagOBJ.gsCollection.Count = 1 Then
                            name = CStr(diagOBJ.gsCollection.Item(1))
                        ElseIf diagOBJ.gsCollection.Count > 1 Then
                            name = "Collection"
                        End If


                        With roentgenBlick
                            If .isOn Then
                                Call awinNoshowProjectNeeds()
                            End If
                            .isOn = True
                            .name = name
                            .myCollection = diagOBJ.gsCollection
                            .type = diagOBJ.diagrammTyp
                            Call awinShowProjectNeeds1(diagOBJ.gsCollection, diagOBJ.diagrammTyp)
                            'End If
                        End With




                    ElseIf diagOBJ.diagrammTyp = DiagrammTypen(5) And diagOBJ.gsCollection.Count > 0 Then

                        Call awinDeleteProjectChildShapes(1)

                        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                            ' hier wird die zeichneMilestones aufgerufen mit den Element-Namen und nicht den Element-IDs
                            ' d.h es ist wichtig, daß die Zeichen-Routine so schlau ist, im Falle des Aufrufes mit den Namen alle 
                            ' Namen durch ihre auftretenden IDs zu ersetzen.  
                            Call zeichneMilestonesInProjekt(kvp.Value, diagOBJ.gsCollection, 4, showRangeLeft, showRangeRight, False, msNumber, False)

                        Next

                    End If
                Catch ex As Exception

                End Try


            End If


        Catch ex As Exception

        End Try


        appInstance.ScreenUpdating = formerSU
        


    End Sub

  
End Class
