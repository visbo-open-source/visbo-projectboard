Imports System.Math
Imports xlNS = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports System.Windows.Forms

'Imports System.Int32

Public Class clsEventsPrcCharts


    Public WithEvents PrcChartEvents As xlNS.Chart
    Public WithEvents PrcChartRightClick As CommandBarButton


    Private Sub PrcChartEvents_Activate() Handles PrcChartEvents.Activate


        If selectedProjekte.Count > 0 Then
            selectedProjekte.Clear()
            Call awinNeuZeichnenDiagramme(8)
        End If


    End Sub


    Private Sub PrcChartEvents_BeforeDoubleClick(ByVal ElementID As Integer, ByVal Arg1 As Integer, ByVal Arg2 As Integer, ByRef Cancel As Boolean) Handles PrcChartEvents.BeforeDoubleClick


        Dim chtobj As xlNS.ChartObject
        Dim IDKennung As String
        Dim foundDiagram As clsDiagramm


        Cancel = True

        Try
            chtobj = Me.PrcChartEvents.Parent

            IDKennung = chtobj.Name
            foundDiagram = DiagramList.getDiagramm(IDKennung)
            With foundDiagram
                .top = chtobj.Top
                .left = chtobj.Left
                .width = chtobj.Width
                .height = chtobj.Height
            End With
        Catch ex As Exception
            'Call MsgBox("konnte das Chart nicht in der Diagramm-Liste finden ...")
        End Try



        'Dim i As Integer, p As Integer
        'Dim von As Integer, bis As Integer
        'Dim left As Double, top As Double, height As Double, width As Double
        'Dim found As Boolean
        'Dim diagrammTyp As String
        'Dim myCollection As New Collection
        'Dim chtobj As ChartObject
        ''Dim chtTitle As String
        'Dim repObj As Object = Nothing
        'Dim IDkennung As String

        'Cancel = True
        'diagrammTyp = " "
        ''
        '' die Werte des Charts bestimmen, aus dem heraus der Event aufgerufen wurde ...
        'Try
        '    chtobj = Me.PrcChartEvents.Parent
        '    IDkennung = chtobj.Name
        'Catch ex As NullReferenceException
        '    Call MsgBox("in PRC ChartEvents, BeforeDoubleClick: kein Chart-Objekt ...")
        '    Exit Sub
        'End Try


        '' ist es überhaupt ein Cockpit chart ?
        'If Not istCockpitDiagramm(chtobj) Then
        '    Exit Sub
        'End If

        'von = showRangeLeft
        'bis = showRangeRight

        'If istSummenDiagramm(chtobj, p) Then

        '    height = 2 * miniHeight
        '    top = WertfuerTop() + awinSettings.ChartHoehe1
        '    left = linkerRandCpPfChart + 5 * miniWidth
        '    width = 300
        '    Call awinCreatePersCostStructureDiagramm(top, left, width, height, False)


        'Else
        '    '
        '    ' Bestimmen der Breite und Position des Diagrammes
        '    '
        '    ' start_top = WertfuerTop + HoehePrcChart

        '    height = awinSettings.ChartHoehe1
        '    top = WertfuerTop()
        '    If von > 1 Then
        '        left = ((von - 1) / 3 - 1) * 3 * boxWidth + 32.8 + von * screen_correct
        '    Else
        '        left = 0
        '    End If

        '    width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct


        '    i = 1
        '    found = False

        '    Dim foundDiagramm As clsDiagramm

        '    'Try
        '    '    chtTitle = chtobj.Chart.ChartTitle.Text
        '    'Catch ex As Exception
        '    '    chtTitle = " "
        '    'End Try

        '    'While i <= DiagramList.Count And Not found
        '    '    If chtTitle Like (DiagramList.getDiagramm(i).DiagrammTitel & "*") Then
        '    '        diagrammTyp = DiagramList.getDiagramm(i).diagrammTyp
        '    '        myCollection = DiagramList.getDiagramm(i).gsCollection
        '    '        found = True
        '    '    Else
        '    '        i = i + 1
        '    '    End If
        '    'End While


        '    Try
        '        foundDiagramm = DiagramList.getDiagramm(IDkennung)
        '        diagrammTyp = foundDiagramm.diagrammTyp
        '        myCollection = foundDiagramm.gsCollection
        '        found = True
        '    Catch ex As Exception

        '    End Try

        '    If found Then
        '        Call awinCreateprcCollectionDiagram(myCollection, repObj, top, left, width, height, False, diagrammTyp)
        '    End If

        '    'myCollection.Clear()
        'End If

        ''chtobj = Nothing

    End Sub

    Private Sub PrcChartEvents_BeforeRightClick(ByRef Cancel As Boolean) Handles PrcChartEvents.BeforeRightClick

        Cancel = True
        appInstance.CommandBars("awinRightClickinPRCChart").ShowPopup()

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
            chtobj = appInstance.ActiveChart.Parent
            IDkennung = chtobj.Name
            foundDiagramm = DiagramList.getDiagramm(IDkennung)
            found = True
            With foundDiagramm
                diagrammTyp = .diagrammTyp
                myCollection = .gsCollection
                isCC = .isCockpitChart
            End With

            If myCollection.Count < 1 Then
                name = ""
            ElseIf myCollection.Count = 1 Then
                name = myCollection.Item(1)
            ElseIf myCollection.Count > 1 Then
                name = "Collection"
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

                If name <> "" Then
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


            Case "Optimieren"
                ' hier werden die Werte der Optimierung vermerkt: welche Projekte müssen verschoben werden , um welchen Offset 

                Dim OptimierungsErgebnis As New SortedList(Of String, clsOptimizationObject)

                enableOnUpdate = False

                Call awinCalculateOptimization1(diagrammTyp, myCollection, OptimierungsErgebnis)

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


                                        ' jetzt wird das Shape in der Plantafel gelöscht 
                                        Call clearProjektinPlantafel(.name)
                                        Call ZeichneProjektinPlanTafel(.name, .tfZeile)
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


        Try
            chtobj = Me.PrcChartEvents.Parent
            Try
                chtobj1 = appInstance.ActiveChart.Parent
            Catch ex As Exception

            End Try

            IDKennung = chtobj.Name
            foundDiagram = DiagramList.getDiagramm(IDKennung)
            With foundDiagram
                .top = chtobj.Top
                .left = chtobj.Left
                .width = chtobj.Width
                .height = chtobj.Height
            End With
        Catch ex As Exception
            'Call MsgBox("konnte das Chart nicht in der Diagramm-Liste finden ...")
        End Try


        'Dim chtobj As ChartObject
        ''Dim spanDiff As Integer



        'Try
        '    chtobj = Me.PrcChartEvents.Parent
        '    With chtobj
        '        Call MsgBox("top: " & .Top & vbLf & _
        '                     "left: " & .Left & vbLf & _
        '                     "width: " & .Width & vbLf & _
        '                     "height: " & .Height)
        '    End With

        'Catch ex As Exception
        '    Exit Sub
        'End Try

        'With chtobj
        '    width = .Width
        '    left = .Left
        'End With


        'If width <> previousWidth Then
        '    If left = previousLeft Then
        '        ' es wurde der rechte Rand verändert 
        '        spanDiff = (width - previousWidth) / boxWidth
        '        If spanDiff <> 0 Then
        '            Call awinChangeTimeSpan(showRangeLeft, showRangeRight + spanDiff)
        '        End If

        '    Else
        '        ' es wurde der linke Rand verändert 
        '        spanDiff = (width - previousWidth) / boxWidth
        '        If spanDiff <> 0 Then
        '            Call awinChangeTimeSpan(showRangeLeft + spanDiff, showRangeRight)
        '        End If

        '    End If

        'End If

        'With chtobj
        '    previousLeft = .Left
        '    previousTop = .Top
        '    previousWidth = .Width
        '    previousHeight = .Height
        'End With



    End Sub


    Private Sub PrcChartEvents_Select(ElementID As Integer, Arg1 As Integer, Arg2 As Integer) Handles PrcChartEvents.Select


        ' in ARG2 steht, das wievielte Element selektiert wurde ...

        If (ElementID = xlNS.XlChartItem.xlSeries) And Arg1 = 1 And Arg2 > 0 Then
            'Dim i As Integer
            Dim chtobjname As String
            Dim diagOBJ As clsDiagramm
            Dim msNumber As Integer
            Dim selMonth As Integer = showRangeLeft + Arg2 - 1
            Dim chtobj As xlNS.ChartObject

            'Dim formerUpdate As Boolean = appInstance.ScreenUpdating
            'appInstance.ScreenUpdating = False

            'Jetzt muss bestimmt werden , um welches Chart es sich handelt 


            Call awinDeleteMilestoneShapes(3)

            Try
                chtobjname = Me.PrcChartEvents.Parent.Name

                chtobj = Me.PrcChartEvents.Parent
                Dim IDKennung As String
                IDKennung = chtobj.Name


                diagOBJ = DiagramList.getDiagramm(chtobjname)




                '
                ' nur bei Phasen wird aktuell etwas gemacht 
                '
                If diagOBJ.diagrammTyp = DiagrammTypen(0) And diagOBJ.gsCollection.Count > 0 Then


                    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                        Call zeichnePhasenInProjekt(kvp.Value, diagOBJ.gsCollection, selMonth, selMonth, False, msNumber)

                    Next

                End If

                

            Catch ex As Exception

            End Try



        End If


    End Sub
End Class
