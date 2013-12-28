Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Module awinGUI


    Sub awinDeletePortfolioDiagram()
        Dim diagramTitle As String = "strategischer Fit, Risiko & Marge"
        Dim anzdiagrams As Integer
        Dim found As Boolean
        Dim i As Integer
        Dim chttitle As String = " "


        With appInstance.Worksheets(arrWsNames(3))
            anzdiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzdiagrams And Not found
                Try
                    chttitle = .ChartObjects(i).Chart.ChartTitle.text
                Catch ex As Exception
                    chttitle = " "
                End Try

                If chttitle Like ("*" & diagramTitle & "*") Then
                    found = True
                Else
                    i = i + 1
                End If
            End While

            If found Then
                .ChartObjects(i).delete()
            End If

            Try
                DiagramList.Remove(diagramTitle)
            Catch ex As Exception

            End Try

        End With

    End Sub
    '
    ' Prozedur für das Anzeigen der Diagramme
    ' awinCreatePortfolioDiagrams(ShowProjekte, TypeCollection, top, left, width, height)
    '
    Sub awinCreatePortfolioDiagrams(ByRef Projekte As clsProjekte, ByRef TypeCollection As Collection, top As Double, left As Double, width As Double, height As Double)

        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim pname As String
        Dim hproj As New clsProjekt
        Dim anzBubbles As Integer
        Dim riskValues() As Double, strategicValues() As Double, bubbleValues() As Double, tempArray() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim diagramTitle As String
        Dim pfDiagram As clsDiagramm
        Dim pfChart As clsEventsPfCharts
        Dim ptype As String
        Dim chtTitle As String
        Dim chtobjName As String = windowNames(3)






        diagramTitle = "strategischer Fit, Risiko & Marge"



        ' hier werden die Werte bestimmt ...

        ReDim riskValues(ShowProjekte.Count - 1)
        ReDim strategicValues(ShowProjekte.Count - 1)
        ReDim bubbleValues(ShowProjekte.Count - 1)
        ReDim nameValues(ShowProjekte.Count - 1)
        ReDim colorValues(ShowProjekte.Count - 1)
        ReDim PfChartBubbleNames(ShowProjekte.Count - 1)

        anzBubbles = 0



        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            hproj = kvp.Value
            pname = kvp.Key
            ptype = hproj.VorlagenName
            If istinStringCollection(ptype, TypeCollection) Then
                riskValues(anzBubbles) = hproj.Risiko
                strategicValues(anzBubbles) = hproj.StrategicFit
                bubbleValues(anzBubbles) = hproj.ProjectMarge
                nameValues(anzBubbles) = hproj.name
                colorValues(anzBubbles) = hproj.farbe
                PfChartBubbleNames(anzBubbles) = hproj.name
                anzBubbles = anzBubbles + 1
            End If
        Next kvp

        'hproj = Nothing

        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle Like ("*" & diagramTitle & "*") Then
                    found = True
                    Exit Sub
                Else
                    i = i + 1
                End If
            End While


            ReDim tempArray(anzBubbles - 1)

            appInstance.EnableEvents = False

            With appInstance.Charts.Add

                .SeriesCollection.NewSeries()
                .SeriesCollection(1).name = diagramTitle
                .SeriesCollection(1).ChartType = Excel.XlChartType.xlBubble3DEffect

                For i = 1 To anzBubbles
                    tempArray(i - 1) = strategicValues(i - 1)
                Next i
                .SeriesCollection(1).XValues = tempArray ' strategic

                For i = 1 To anzBubbles
                    tempArray(i - 1) = riskValues(i - 1)
                Next i
                .SeriesCollection(1).Values = tempArray

                For i = 1 To anzBubbles
                    tempArray(i - 1) = bubbleValues(i - 1)
                Next i
                .SeriesCollection(1).BubbleSizes = tempArray

                Dim series1 As Excel.Series = _
                        CType(.SeriesCollection(1),  _
                                Excel.Series)
                Dim point1 As Excel.Point = _
                            CType(series1.Points(1), Excel.Point)

                'Dim testName As String
                For i = 1 To anzBubbles
                    With .SeriesCollection(1).Points(i)
                        .HasDataLabel = False
                        '.DataLabel.text = PfChartBubbleNames(i - 1)
                        'testName = .DataLabel.text
                        .Interior.color = colorValues(i - 1)
                    End With
                Next i

                'With series1
                '    .ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowNone)
                'End With

                .ChartGroups(1).BubbleScale = 20
                .HasAxis(Excel.XlAxisType.xlCategory) = True
                .HasAxis(Excel.XlAxisType.xlValue) = True
                .Axes(Excel.XlAxisType.xlCategory).HasMajorGridlines = False
                .Axes(Excel.XlAxisType.xlValue).HasMajorGridlines = False


                With .Axes(Excel.XlAxisType.xlCategory)
                    .HasTitle = True
                    .MinimumScale = 0
                    .MaximumScale = 11
                    With .AxisTitle
                        .Characters.text = "strategischer Fit"
                        .Characters.Font.Size = 18
                        .Characters.Font.Bold = False
                    End With
                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .Size = 12
                    End With

                End With
                With .Axes(Excel.XlAxisType.xlValue)
                    .HasTitle = True
                    .MinimumScale = 0
                    .MaximumScale = 11
                    ' .ReversePlotOrder = True
                    With .AxisTitle
                        .Characters.text = "Risiko"
                        .Characters.Font.Size = 18
                        .Characters.Font.Bold = False
                    End With

                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .Size = 12
                    End With
                End With
                .HasLegend = False
                .HasTitle = True
                .ChartTitle.text = diagramTitle
                .ChartTitle.Characters.Font.Size = 18
                .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
            End With

            appInstance.EnableEvents = True
            appInstance.ShowChartTipNames = False
            appInstance.ShowChartTipValues = False

            With .ChartObjects(anzDiagrams + 1)
                .top = top
                .left = left
                .width = width
                .height = height
                .name = chtobjName
            End With

            With appInstance.ActiveSheet
                Try
                    With appInstance.ActiveSheet
                        .Shapes(chtobjName).line.visible = False
                    End With
                Catch ex As Exception

                End Try
            End With

            pfDiagram = New clsDiagramm

            'pfChart = New clsAwinEvent
            pfChart = New clsEventsPfCharts
            pfChart.PfChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            'pfDiagram.setpfDiagramEvent = pfChart
            pfDiagram.setDiagramEvent = pfChart

            With pfDiagram
                .DiagrammTitel = diagramTitle
                .diagrammTyp = DiagrammTypen(3) ' Portfolio
                .gsCollection = TypeCollection
                .isCockpitChart = False
            End With

            DiagramList.Add(pfDiagram)
            'pfDiagram = Nothing


        End With



    End Sub

    Sub awinCreateComplexRiskVolumeDiagramm(ByRef ProjektListe As Collection, ByRef repChart As Object, isProjektCharakteristik As Boolean, _
                                     showNegativeValues As Boolean, showLabels As Boolean, chartBorderVisible As Boolean, _
                                     top As Double, left As Double, width As Double, height As Double)

        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim pname As String = ""
        Dim hproj As New clsProjekt
        Dim anzBubbles As Integer
        Dim riskValues() As Double, xAchsenValues() As Double, bubbleValues() As Double, tempArray() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String
        Dim pfDiagram As clsDiagramm
        Dim pfChart As clsEventsPfCharts
        'Dim ptype As String

        Dim chtobjName As String = windowNames(3)
        Dim smallfontsize As Double, titlefontsize As Double

        Dim singleProject As Boolean



        If ProjektListe.Count > 1 Then
            singleProject = False
        Else
            singleProject = True
        End If


        If width > 450 Then
            titlefontsize = 20
            smallfontsize = 10
        ElseIf width > 250 Then
            titlefontsize = 14
            smallfontsize = 8
        Else
            titlefontsize = 12
            smallfontsize = 8
        End If






        ' hier werden die Werte bestimmt ...
        Try
            ReDim riskValues(ProjektListe.Count - 1)
            ReDim xAchsenValues(ProjektListe.Count - 1)
            ReDim bubbleValues(ProjektListe.Count - 1)
            ReDim nameValues(ProjektListe.Count - 1)
            ReDim colorValues(ProjektListe.Count - 1)
            ReDim PfChartBubbleNames(ProjektListe.Count - 1)
            ReDim positionValues(ProjektListe.Count - 1)
        Catch ex As Exception

            Throw New ArgumentException("Fehler in CreatePortfolioDiagramm " & ex.Message)

        End Try


        anzBubbles = 0


        For i = 1 To ProjektListe.Count
            pname = ProjektListe.Item(i)
            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    riskValues(anzBubbles) = .Risiko
                    xAchsenValues(anzBubbles) = .complexity
                    bubbleValues(anzBubbles) = .volume
                    nameValues(anzBubbles) = .name
                    colorValues(anzBubbles) = .farbe
                    PfChartBubbleNames(anzBubbles) = hproj.name & _
                            " (" & Format(bubbleValues(anzBubbles) / 1000, "##0.#") & " T)"

                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next

        If isProjektCharakteristik Then
            diagramTitle = portfolioDiagrammtitel(PTpfdk.ComplexRisiko)
        Else
            diagramTitle = portfolioDiagrammtitel(PTpfdk.ComplexRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)
        End If


        If singleProject Then
            chtobjName = pname & portfolioDiagrammtitel(PTpfdk.ComplexRisiko)
        Else
            chtobjName = portfolioDiagrammtitel(PTpfdk.ComplexRisiko)
        End If



        ' bestimmen der besten Position für die Werte ...
        Dim labelPosition(4) As String
        labelPosition(0) = "oben"
        labelPosition(1) = "rechts"
        labelPosition(2) = "unten"
        labelPosition(3) = "links"
        labelPosition(4) = "mittig"

        For i = 0 To anzBubbles - 1

            positionValues(i) = pfchartIstFrei(i, xAchsenValues, riskValues)

        Next



        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found


                If chtobjName = .chartObjects(i).name Then
                    found = True
                    repChart = .ChartObjects(i)
                    Exit Sub
                Else
                    i = i + 1
                End If
            End While


            ReDim tempArray(anzBubbles - 1)


            With appInstance.Charts.Add

                .SeriesCollection.NewSeries()
                .SeriesCollection(1).name = diagramTitle
                .SeriesCollection(1).ChartType = Excel.XlChartType.xlBubble3DEffect

                For i = 1 To anzBubbles
                    tempArray(i - 1) = xAchsenValues(i - 1)
                Next i
                .SeriesCollection(1).XValues = tempArray ' strategic

                For i = 1 To anzBubbles
                    tempArray(i - 1) = riskValues(i - 1)
                Next i
                .SeriesCollection(1).Values = tempArray

                For i = 1 To anzBubbles
                    If bubbleValues(i - 1) < 0.01 And bubbleValues(i - 1) > -0.01 Then
                        tempArray(i - 1) = 0.01
                    Else
                        tempArray(i - 1) = bubbleValues(i - 1)
                    End If
                Next i
                .SeriesCollection(1).BubbleSizes = tempArray

                Dim series1 As Excel.Series = _
                        CType(.SeriesCollection(1),  _
                                Excel.Series)
                Dim point1 As Excel.Point = _
                            CType(series1.Points(1), Excel.Point)

                'Dim testName As String
                For i = 1 To anzBubbles

                    With .SeriesCollection(1).Points(i)

                        If showLabels Then
                            Try
                                .HasDataLabel = True
                                With .DataLabel
                                    .text = PfChartBubbleNames(i - 1)
                                    If singleProject Then
                                        .font.size = awinSettings.CPfontsizeItems + 4
                                    Else
                                        .font.size = awinSettings.CPfontsizeItems
                                    End If

                                    Select Case positionValues(i - 1)
                                        Case labelPosition(0)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                        Case labelPosition(1)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionRight
                                        Case labelPosition(2)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionBelow
                                        Case labelPosition(3)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionLeft
                                        Case Else
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionCenter
                                    End Select
                                End With
                            Catch ex As Exception

                            End Try
                        Else
                            .HasDataLabel = False
                        End If

                        .Interior.color = colorValues(i - 1)
                    End With
                Next i



                '.ChartGroups(1).BubbleScale = sollte in Abhängigkeit der width gemacht werden 
                With .ChartGroups(1)
                    If singleProject Then
                        .BubbleScale = 20
                    Else
                        .BubbleScale = 20
                    End If

                    .SizeRepresents = Microsoft.Office.Interop.Excel.XlSizeRepresents.xlSizeIsArea
                    If showNegativeValues Then
                        .shownegativeBubbles = True
                    Else
                        .shownegativeBubbles = False
                    End If
                End With


                .HasAxis(Excel.XlAxisType.xlCategory) = True
                .HasAxis(Excel.XlAxisType.xlValue) = True
                .Axes(Excel.XlAxisType.xlCategory).HasMajorGridlines = False
                .Axes(Excel.XlAxisType.xlValue).HasMajorGridlines = False


                With .Axes(Excel.XlAxisType.xlCategory)
                    .HasTitle = True
                    .MinimumScale = 0
                    .MaximumScale = 1.1
                    With .AxisTitle
                        .Characters.text = "Komplexität"
                        .Characters.Font.Size = titlefontsize
                        .Characters.Font.Bold = False
                    End With
                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .Bold = True
                        .Size = awinSettings.fontsizeItems

                    End With

                End With
                With .Axes(Excel.XlAxisType.xlValue)
                    .HasTitle = True
                    .MinimumScale = 0
                    .MaximumScale = 11
                    ' .ReversePlotOrder = True
                    With .AxisTitle
                        .Characters.text = "Risiko"
                        .Characters.Font.Size = titlefontsize
                        .Characters.Font.Bold = False
                    End With

                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .bold = True
                        .Size = awinSettings.fontsizeItems
                    End With
                End With
                .HasLegend = False
                .HasTitle = True
                .ChartTitle.text = diagramTitle
                .ChartTitle.Characters.Font.Size = awinSettings.fontsizeTitle
                .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
            End With


            appInstance.ShowChartTipNames = False
            appInstance.ShowChartTipValues = False

            With .ChartObjects(anzDiagrams + 1)
                .top = top
                .left = left
                .width = width
                .height = height
                .name = chtobjName
            End With



            With appInstance.ActiveSheet
                Try
                    With appInstance.ActiveSheet
                        .Shapes(chtobjName).line.visible = chartBorderVisible
                    End With
                Catch ex As Exception

                End Try
            End With

            pfDiagram = New clsDiagramm

            'pfChart = New clsAwinEvent
            pfChart = New clsEventsPfCharts
            pfChart.PfChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            'pfDiagram.setpfDiagramEvent = pfChart
            pfDiagram.setDiagramEvent = pfChart

            With pfDiagram
                .DiagrammTitel = chtobjName
                .diagrammTyp = DiagrammTypen(3) ' Portfolio
                .gsCollection = ProjektListe
                .isCockpitChart = False
            End With

            DiagramList.Add(pfDiagram)
            'pfDiagram = Nothing

            repChart = .ChartObjects(anzDiagrams + 1)

        End With


    End Sub


    Sub awinCreateZeitRiskVolumeDiagramm(ByRef ProjektListe As Collection, ByRef repChart As Object, isProjektCharakteristik As Boolean, _
                                     showNegativeValues As Boolean, showLabels As Boolean, chartBorderVisible As Boolean, _
                                     top As Double, left As Double, width As Double, height As Double)

        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim pname As String = ""
        Dim hproj As New clsProjekt
        Dim anzBubbles As Integer
        Dim riskValues() As Double, xAchsenValues() As Double, bubbleValues() As Double, tempArray() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String
        Dim pfDiagram As clsDiagramm
        Dim pfChart As clsEventsPfCharts
        'Dim ptype As String

        Dim chtobjName As String = windowNames(3)
        Dim smallfontsize As Double, titlefontsize As Double

        Dim singleProject As Boolean



        If ProjektListe.Count > 1 Then
            singleProject = False
        Else
            singleProject = True
        End If


        If width > 450 Then
            titlefontsize = 20
            smallfontsize = 10
        ElseIf width > 250 Then
            titlefontsize = 14
            smallfontsize = 8
        Else
            titlefontsize = 12
            smallfontsize = 8
        End If






        ' hier werden die Werte bestimmt ...
        Try
            ReDim riskValues(ProjektListe.Count - 1)
            ReDim xAchsenValues(ProjektListe.Count - 1)
            ReDim bubbleValues(ProjektListe.Count - 1)
            ReDim nameValues(ProjektListe.Count - 1)
            ReDim colorValues(ProjektListe.Count - 1)
            ReDim PfChartBubbleNames(ProjektListe.Count - 1)
            ReDim positionValues(ProjektListe.Count - 1)
        Catch ex As Exception

            Throw New ArgumentException("Fehler in CreatePortfolioDiagramm " & ex.Message)

        End Try


        anzBubbles = 0

        Dim tmpstr(10) As String

        For i = 1 To ProjektListe.Count
            pname = ProjektListe.Item(i)
            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    riskValues(anzBubbles) = .Risiko
                    xAchsenValues(anzBubbles) = .dauerInDays / 365 * 12
                    bubbleValues(anzBubbles) = System.Math.Round(.volume / 10000) * 10

                    tmpstr = .name.Split(New Char() {" "}, 10)
                    nameValues(anzBubbles) = tmpstr(0) & " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)"
                    colorValues(anzBubbles) = .farbe
                    PfChartBubbleNames(anzBubbles) = .name & _
                            " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)"

                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next

        If isProjektCharakteristik Then
            diagramTitle = portfolioDiagrammtitel(PTpfdk.ZeitRisiko)
        Else
            diagramTitle = portfolioDiagrammtitel(PTpfdk.ZeitRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)
        End If


        If singleProject Then
            chtobjName = pname & portfolioDiagrammtitel(PTpfdk.ZeitRisiko)
        Else
            chtobjName = portfolioDiagrammtitel(PTpfdk.ZeitRisiko)
        End If



        ' bestimmen der besten Position für die Werte ...
        Dim labelPosition(4) As String
        labelPosition(0) = "oben"
        labelPosition(1) = "rechts"
        labelPosition(2) = "unten"
        labelPosition(3) = "links"
        labelPosition(4) = "mittig"

        For i = 0 To anzBubbles - 1

            positionValues(i) = pfchartIstFrei(i, xAchsenValues, riskValues)

        Next



        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found


                If chtobjName = .chartObjects(i).name Then
                    found = True
                    repChart = .ChartObjects(i)
                    Exit Sub
                Else
                    i = i + 1
                End If
            End While


            ReDim tempArray(anzBubbles - 1)


            With appInstance.Charts.Add

                .SeriesCollection.NewSeries()
                .SeriesCollection(1).name = diagramTitle
                .SeriesCollection(1).ChartType = Excel.XlChartType.xlBubble3DEffect

                For i = 1 To anzBubbles
                    tempArray(i - 1) = xAchsenValues(i - 1)
                Next i
                .SeriesCollection(1).XValues = tempArray ' strategic

                For i = 1 To anzBubbles
                    tempArray(i - 1) = riskValues(i - 1)
                Next i
                .SeriesCollection(1).Values = tempArray

                For i = 1 To anzBubbles
                    If bubbleValues(i - 1) < 0.01 And bubbleValues(i - 1) > -0.01 Then
                        tempArray(i - 1) = 0.01
                    Else
                        tempArray(i - 1) = bubbleValues(i - 1)
                    End If
                Next i
                .SeriesCollection(1).BubbleSizes = tempArray

                Dim series1 As Excel.Series = _
                        CType(.SeriesCollection(1),  _
                                Excel.Series)
                Dim point1 As Excel.Point = _
                            CType(series1.Points(1), Excel.Point)

                'Dim testName As String
                For i = 1 To anzBubbles

                    With .SeriesCollection(1).Points(i)

                        If showLabels Then
                            Try
                                .HasDataLabel = True
                                With .DataLabel
                                    '.text = PfChartBubbleNames(i - 1)
                                    .text = nameValues(i - 1)
                                    If singleProject Then
                                        .font.size = awinSettings.CPfontsizeItems + 4
                                    Else
                                        .font.size = awinSettings.CPfontsizeItems
                                    End If

                                    Select Case positionValues(i - 1)
                                        Case labelPosition(0)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                        Case labelPosition(1)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionRight
                                        Case labelPosition(2)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionBelow
                                        Case labelPosition(3)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionLeft
                                        Case Else
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionCenter
                                    End Select
                                End With
                            Catch ex As Exception

                            End Try
                        Else
                            .HasDataLabel = False
                        End If

                        .Interior.color = colorValues(i - 1)
                    End With
                Next i



                '.ChartGroups(1).BubbleScale = sollte in Abhängigkeit der width gemacht werden 
                With .ChartGroups(1)
                    If singleProject Then
                        .BubbleScale = 20
                    Else
                        .BubbleScale = 20
                    End If

                    .SizeRepresents = Microsoft.Office.Interop.Excel.XlSizeRepresents.xlSizeIsArea
                    If showNegativeValues Then
                        .shownegativeBubbles = True
                    Else
                        .shownegativeBubbles = False
                    End If
                End With


                .HasAxis(Excel.XlAxisType.xlCategory) = True
                .HasAxis(Excel.XlAxisType.xlValue) = True
                .Axes(Excel.XlAxisType.xlCategory).HasMajorGridlines = False
                .Axes(Excel.XlAxisType.xlValue).HasMajorGridlines = False


                With .Axes(Excel.XlAxisType.xlCategory)
                    .HasTitle = True
                    .MinimumScale = 15.0
                    .MaximumScale = System.Math.Round(xAchsenValues.Max / 10 + 0.5) * 10
                    With .AxisTitle
                        .Characters.text = "Projekt-Dauer"
                        .Characters.Font.Size = titlefontsize
                        .Characters.Font.Bold = False
                    End With
                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .Bold = True
                        .Size = awinSettings.fontsizeItems

                    End With

                End With
                With .Axes(Excel.XlAxisType.xlValue)
                    .HasTitle = True
                    .MinimumScale = 0
                    .MaximumScale = 11
                    ' .ReversePlotOrder = True
                    With .AxisTitle
                        .Characters.text = "Risiko"
                        .Characters.Font.Size = titlefontsize
                        .Characters.Font.Bold = False
                    End With

                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .bold = True
                        .Size = awinSettings.fontsizeItems
                    End With
                End With
                .HasLegend = False
                .HasTitle = True
                .ChartTitle.text = diagramTitle
                .ChartTitle.Characters.Font.Size = awinSettings.fontsizeTitle
                .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
            End With


            appInstance.ShowChartTipNames = False
            appInstance.ShowChartTipValues = False

            With .ChartObjects(anzDiagrams + 1)
                .top = top
                .left = left
                .width = width
                .height = height
                .name = chtobjName
            End With



            With appInstance.ActiveSheet
                Try
                    With appInstance.ActiveSheet
                        .Shapes(chtobjName).line.visible = chartBorderVisible
                    End With
                Catch ex As Exception

                End Try
            End With

            pfDiagram = New clsDiagramm

            'pfChart = New clsAwinEvent
            pfChart = New clsEventsPfCharts
            pfChart.PfChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            'pfDiagram.setpfDiagramEvent = pfChart
            pfDiagram.setDiagramEvent = pfChart

            With pfDiagram
                .DiagrammTitel = chtobjName
                .diagrammTyp = DiagrammTypen(3) ' Portfolio
                .gsCollection = ProjektListe
                .isCockpitChart = False
            End With

            DiagramList.Add(pfDiagram)
            'pfDiagram = Nothing

            repChart = .ChartObjects(anzDiagrams + 1)

        End With


    End Sub
   
    '
    ' Prozedur für das Anzeigen des Bubble Charts
    ' awinCreatePortfolioDiagrams(ShowProjekte, TypeCollection, top, left, width, height)
    '
    Sub awinCreateStratRisikMargeDiagramm(ByRef ProjektListe As Collection, ByRef repChart As Object, isProjektCharakteristik As Boolean, _
                                     showNegativeValues As Boolean, showLabels As Boolean, chartBorderVisible As Boolean, _
                                     top As Double, left As Double, width As Double, height As Double)

        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim pname As String
        Dim hproj As New clsProjekt
        Dim anzBubbles As Integer
        Dim riskValues() As Double, strategicValues() As Double, bubbleValues() As Double, tempArray() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String
        Dim pfDiagram As clsDiagramm
        Dim pfChart As clsEventsPfCharts
        'Dim ptype As String
        Dim chtTitle As String
        Dim chtobjName As String = windowNames(3)
        Dim smallfontsize As Double, titlefontsize As Double
        Dim kennung As String
        Dim singleProject As Boolean



        If ProjektListe.Count > 1 Then
            singleProject = False
        Else
            singleProject = True
        End If


        If width > 450 Then
            titlefontsize = 20
            smallfontsize = 10
        ElseIf width > 250 Then
            titlefontsize = 14
            smallfontsize = 8
        Else
            titlefontsize = 12
            smallfontsize = 8
        End If

        If isProjektCharakteristik Then

            diagramTitle = "Charakteristik " & summentitel2
            kennung = "Strategie"
        Else
            diagramTitle = summentitel2 & vbLf & textZeitraum(showRangeLeft, showRangeRight)
        End If




        ' hier werden die Werte bestimmt ...
        Try
            ReDim riskValues(ProjektListe.Count - 1)
            ReDim strategicValues(ProjektListe.Count - 1)
            ReDim bubbleValues(ProjektListe.Count - 1)
            ReDim nameValues(ProjektListe.Count - 1)
            ReDim colorValues(ProjektListe.Count - 1)
            ReDim PfChartBubbleNames(ProjektListe.Count - 1)
            ReDim positionValues(ProjektListe.Count - 1)
        Catch ex As Exception

            Throw New ArgumentException("Fehler in CreatePortfolioDiagramm " & ex.Message)

        End Try


        anzBubbles = 0


        For i = 1 To ProjektListe.Count
            pname = ProjektListe.Item(i)
            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    riskValues(anzBubbles) = .Risiko
                    strategicValues(anzBubbles) = .StrategicFit
                    bubbleValues(anzBubbles) = .ProjectMarge
                    nameValues(anzBubbles) = .name
                    colorValues(anzBubbles) = .farbe
                    If singleProject Then
                        PfChartBubbleNames(anzBubbles) = Format(bubbleValues(anzBubbles), "##0.#%")
                    Else
                        PfChartBubbleNames(anzBubbles) = hproj.name & _
                            " (" & Format(bubbleValues(anzBubbles), "##0.#%") & ")"
                    End If

                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next

        ' jetzt werden die negativen Werte alle auf den größten vorkommenden Wert gesetzt .. und mit roter Farbe markiert ..
        Dim maxWert As Double = bubbleValues.Max

        For i = 0 To anzBubbles - 1
            If bubbleValues(i) < 0 Then
                colorValues(i) = awinSettings.AmpelRot
                bubbleValues(i) = maxWert
            End If
        Next


        ' bestimmen der besten Position für die Werte ...
        Dim labelPosition(4) As String
        labelPosition(0) = "oben"
        labelPosition(1) = "rechts"
        labelPosition(2) = "unten"
        labelPosition(3) = "links"
        labelPosition(4) = "mittig"

        For i = 0 To anzBubbles - 1

            positionValues(i) = pfchartIstFrei(i, strategicValues, riskValues)

        Next



        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                Try
                    chtTitle = .ChartObjects(i).Chart.ChartTitle.text
                Catch ex As Exception
                    chtTitle = " "
                End Try

                If chtTitle Like ("*" & diagramTitle & "*") Then
                    found = True
                    repChart = .ChartObjects(i)
                    Exit Sub
                Else
                    i = i + 1
                End If
            End While


            ReDim tempArray(anzBubbles - 1)


            With appInstance.Charts.Add

                .SeriesCollection.NewSeries()
                .SeriesCollection(1).name = diagramTitle
                .SeriesCollection(1).ChartType = Excel.XlChartType.xlBubble3DEffect

                For i = 1 To anzBubbles
                    tempArray(i - 1) = strategicValues(i - 1)
                Next i
                .SeriesCollection(1).XValues = tempArray ' strategic

                For i = 1 To anzBubbles
                    tempArray(i - 1) = riskValues(i - 1)
                Next i
                .SeriesCollection(1).Values = tempArray

                For i = 1 To anzBubbles
                    If bubbleValues(i - 1) < 0.01 And bubbleValues(i - 1) > -0.01 Then
                        tempArray(i - 1) = 0.01
                    Else
                        tempArray(i - 1) = bubbleValues(i - 1)
                    End If
                Next i
                .SeriesCollection(1).BubbleSizes = tempArray

                Dim series1 As Excel.Series = _
                        CType(.SeriesCollection(1),  _
                                Excel.Series)
                Dim point1 As Excel.Point = _
                            CType(series1.Points(1), Excel.Point)

                'Dim testName As String
                For i = 1 To anzBubbles

                    With .SeriesCollection(1).Points(i)

                        If showLabels Then
                            Try
                                .HasDataLabel = True
                                With .DataLabel
                                    .text = PfChartBubbleNames(i - 1)
                                    If singleProject Then
                                        .font.size = awinSettings.CPfontsizeItems + 4
                                    Else
                                        .font.size = awinSettings.CPfontsizeItems
                                    End If

                                    Select Case positionValues(i - 1)
                                        Case labelPosition(0)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                        Case labelPosition(1)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionRight
                                        Case labelPosition(2)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionBelow
                                        Case labelPosition(3)
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionLeft
                                        Case Else
                                            .Position = Excel.XlDataLabelPosition.xlLabelPositionCenter
                                    End Select
                                End With
                            Catch ex As Exception

                            End Try
                        Else
                            .HasDataLabel = False
                        End If

                        .Interior.color = colorValues(i - 1)
                    End With
                Next i



                '.ChartGroups(1).BubbleScale = sollte in Abhängigkeit der width gemacht werden 
                With .ChartGroups(1)
                    If singleProject Then
                        .BubbleScale = 20
                    Else
                        .BubbleScale = 20
                    End If

                    .SizeRepresents = Microsoft.Office.Interop.Excel.XlSizeRepresents.xlSizeIsArea
                    If showNegativeValues Then
                        .shownegativeBubbles = True
                    Else
                        .shownegativeBubbles = False
                    End If
                End With


                .HasAxis(Excel.XlAxisType.xlCategory) = True
                .HasAxis(Excel.XlAxisType.xlValue) = True
                .Axes(Excel.XlAxisType.xlCategory).HasMajorGridlines = False
                .Axes(Excel.XlAxisType.xlValue).HasMajorGridlines = False


                With .Axes(Excel.XlAxisType.xlCategory)
                    .HasTitle = True
                    .MinimumScale = 0
                    .MaximumScale = 11
                    With .AxisTitle
                        .Characters.text = "strategischer Fit"
                        .Characters.Font.Size = titlefontsize
                        .Characters.Font.Bold = False
                    End With
                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .Bold = True
                        .Size = awinSettings.fontsizeItems

                    End With

                End With
                With .Axes(Excel.XlAxisType.xlValue)
                    .HasTitle = True
                    .MinimumScale = 0
                    .MaximumScale = 11
                    ' .ReversePlotOrder = True
                    With .AxisTitle
                        .Characters.text = "Risiko"
                        .Characters.Font.Size = titlefontsize
                        .Characters.Font.Bold = False
                    End With

                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .bold = True
                        .Size = awinSettings.fontsizeItems
                    End With
                End With
                .HasLegend = False
                .HasTitle = True
                .ChartTitle.text = diagramTitle
                .ChartTitle.Characters.Font.Size = awinSettings.fontsizeTitle
                .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
            End With


            appInstance.ShowChartTipNames = False
            appInstance.ShowChartTipValues = False

            With .ChartObjects(anzDiagrams + 1)
                .top = top
                .left = left
                .width = width
                .height = height
                .name = chtobjName
            End With



            With appInstance.ActiveSheet
                Try
                    With appInstance.ActiveSheet
                        .Shapes(chtobjName).line.visible = chartBorderVisible
                    End With
                Catch ex As Exception

                End Try
            End With

            pfDiagram = New clsDiagramm

            'pfChart = New clsAwinEvent
            pfChart = New clsEventsPfCharts
            pfChart.PfChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            'pfDiagram.setpfDiagramEvent = pfChart
            pfDiagram.setDiagramEvent = pfChart

            With pfDiagram
                .DiagrammTitel = diagramTitle
                .diagrammTyp = DiagrammTypen(3) ' Portfolio
                .gsCollection = ProjektListe
                .isCockpitChart = False
            End With

            DiagramList.Add(pfDiagram)
            'pfDiagram = Nothing

            repChart = .ChartObjects(anzDiagrams + 1)

        End With


    End Sub


    '
    ' Prozedur für den Update des Portfolio Diagramms
    '
    '
    Sub awinUpdatePortfolioDiagrams(ByVal chtobj As ChartObject)

        Dim i As Integer
        Dim pname As String
        Dim hproj As New clsProjekt
        Dim anzBubbles As Integer
        Dim riskValues() As Double, strategicValues() As Double, bubbleValues() As Double, tempArray() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String
        Dim showLabels As Boolean
        Dim showNegativeValues As Boolean = True
        Dim projektListe As Collection
        'Dim pfDiagram As clsDiagramm
        'Dim pfChart As clsEventsPfCharts
        'Dim TypeCollection As Collection


        diagramTitle = summentitel2 & vbLf & textZeitraum(showRangeLeft, showRangeRight)


        ' hier werden die Werte bestimmt ...

        ReDim riskValues(ShowProjekte.Count - 1)
        ReDim strategicValues(ShowProjekte.Count - 1)
        ReDim bubbleValues(ShowProjekte.Count - 1)
        ReDim nameValues(ShowProjekte.Count - 1)
        ReDim colorValues(ShowProjekte.Count - 1)
        ReDim PfChartBubbleNames(ShowProjekte.Count - 1)
        ReDim positionValues(ShowProjekte.Count - 1)

        anzBubbles = 0

        Dim selectionType As Integer = -1 ' keine Einschränkung
        projektListe = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        For i = 1 To projektListe.Count
            pname = projektListe.Item(i)
            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    riskValues(anzBubbles) = .Risiko
                    strategicValues(anzBubbles) = .StrategicFit
                    bubbleValues(anzBubbles) = .ProjectMarge
                    nameValues(anzBubbles) = .name
                    colorValues(anzBubbles) = .farbe
                    PfChartBubbleNames(anzBubbles) = hproj.name & " (" & Format(bubbleValues(anzBubbles), "##0.#%") & ")"
                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next

        ' jetzt werden die negativen Werte alle auf den größten vorkommenden Wert gesetzt .. und mit roter Farbe markiert ..
        Dim maxWert As Double = bubbleValues.Max
        For i = 0 To anzBubbles - 1
            If bubbleValues(i) < 0 Then
                colorValues(i) = awinSettings.AmpelRot
                bubbleValues(i) = maxWert
            End If
        Next



        ' bestimmen der besten Position für die Werte ...
        Dim labelPosition(3) As String
        labelPosition(0) = "oben"
        labelPosition(1) = "rechts"
        labelPosition(2) = "unten"
        labelPosition(3) = "links"

        For i = 0 To anzBubbles - 1

            positionValues(i) = pfchartIstFrei(i, strategicValues, riskValues)

        Next


        'hproj = Nothing

        'With appInstance.Worksheets(arrWsNames(3))


        ReDim tempArray(anzBubbles - 1)

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False



        With chtobj.Chart
            'With .SeriesCollection(1)
            '    showLabels = .HasDataLabels
            'End With

            showLabels = True

            ' remove old series
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete()
            Loop

            .SeriesCollection.NewSeries()
            .SeriesCollection(1).name = diagramTitle
            .SeriesCollection(1).ChartType = Excel.XlChartType.xlBubble3DEffect

            For i = 1 To anzBubbles
                tempArray(i - 1) = strategicValues(i - 1)
            Next i
            .SeriesCollection(1).XValues = tempArray ' strategic

            For i = 1 To anzBubbles
                tempArray(i - 1) = riskValues(i - 1)
            Next i
            .SeriesCollection(1).Values = tempArray

            For i = 1 To anzBubbles
                If bubbleValues(i - 1) < 0.01 And bubbleValues(i - 1) > -0.01 Then
                    tempArray(i - 1) = 0.01
                Else
                    tempArray(i - 1) = bubbleValues(i - 1)
                End If
            Next i
            .SeriesCollection(1).BubbleSizes = tempArray

            Dim series1 As Excel.Series = _
                    CType(.SeriesCollection(1),  _
                            Excel.Series)
            Dim point1 As Excel.Point = _
                        CType(series1.Points(1), Excel.Point)

            'Dim testName As String
            For i = 1 To anzBubbles

                With .SeriesCollection(1).Points(i)

                    If showLabels Then
                        Try
                            .HasDataLabel = True
                            With .DataLabel
                                .text = PfChartBubbleNames(i - 1)
                                .font.size = awinSettings.CPfontsizeItems

                                Select Case positionValues(i - 1)
                                    Case labelPosition(0)
                                        .Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                    Case labelPosition(1)
                                        .Position = Excel.XlDataLabelPosition.xlLabelPositionRight
                                    Case labelPosition(2)
                                        .Position = Excel.XlDataLabelPosition.xlLabelPositionBelow
                                    Case labelPosition(3)
                                        .Position = Excel.XlDataLabelPosition.xlLabelPositionLeft
                                    Case Else
                                        .Position = Excel.XlDataLabelPosition.xlLabelPositionCenter
                                End Select
                            End With
                        Catch ex As Exception

                        End Try
                    Else
                        .HasDataLabel = False
                    End If

                    .Interior.color = colorValues(i - 1)
                End With
            Next i



            '.ChartGroups(1).BubbleScale = sollte in Abhängigkeit der width gemacht werden 
            With .ChartGroups(1)
                
                .BubbleScale = 20
                .SizeRepresents = Microsoft.Office.Interop.Excel.XlSizeRepresents.xlSizeIsArea

                If showNegativeValues Then
                    .shownegativeBubbles = True
                Else
                    .shownegativeBubbles = False
                End If
            End With

            .ChartTitle.Text = diagramTitle
        End With

        appInstance.EnableEvents = formerEE
        



    End Sub

    Function pfchartIstFrei(index As Integer, ByRef sValues() As Double, ByRef rValues() As Double) As String
        Dim sfit As Double = sValues(index)
        Dim risk As Double = rValues(index)
        Dim istFrei As Boolean = False
        Dim anzahl As Integer = UBound(sValues)
        Dim k As Integer

        ' Routine bestimmt wo am meisten Platz frei ist 
        Dim korrfaktor As Double = 4.0
        Dim xOben As Double = 10 - risk, xUnten As Double = risk, xRight As Double = 10 - sfit, xLeft As Double = sfit
        Dim richtung As String = ""
        Dim mxAbstand As Double = -1

        For k = 0 To anzahl

            If k <> index Then
                If Math.Abs(sfit - sValues(k)) < 1.5 Then
                    If rValues(k) - risk > 0 Then
                        If (rValues(k) - risk) < xOben Then
                            xOben = (rValues(k) - risk)

                        End If
                    Else
                        If (risk - rValues(k)) > 0 And (risk - rValues(k)) < xUnten Then
                            xUnten = (risk - rValues(k))
                        End If
                    End If
                End If

                If Math.Abs(risk - rValues(k)) < 0.33 Then
                    If sValues(k) - sfit > 0 Then
                        If sValues(k) - sfit < xRight Then
                            xRight = sValues(k) - sfit

                        End If
                    Else
                        If sfit - sValues(k) < xLeft Then
                            xLeft = sfit - sValues(k)

                        End If
                    End If
                End If
            End If

        Next

        If Math.Max(xLeft, xRight) > 2 Then
            ' entweder links oder rechts platzieren 
            If xLeft > xRight And Not ((xRight = 10 - sfit) And xRight >= 0.4) Then
                richtung = "links"
            Else
                richtung = "rechts"
            End If
        Else
            If xOben > xUnten Or (xOben = 10 - risk) Then
                richtung = "oben"
            Else
                richtung = "unten"
            End If
        End If


        If richtung = "" Then
            pfchartIstFrei = "mittig"
        Else
            pfchartIstFrei = richtung
        End If


    End Function

    Function WertfuerTop() As Double


        Dim lastrow As Integer, lastcolumn As Integer



        lastrow = 2
        lastcolumn = 1
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            With kvp.Value

                If .tfZeile > lastrow Then
                    lastrow = .tfZeile
                End If

                If .tfSpalte + .Dauer - 1 > lastcolumn Then
                    lastcolumn = .tfSpalte + .Dauer - 1
                End If

            End With


        Next

        'With appInstance.Worksheets(arrWsNames(3))
        '    suchfeld = .Range("Projekteingabebereich")
        '    lastrow = suchfeld.Rows.Count
        '    lastcolumn = suchfeld.Columns.Count
        '    max = 0

        '    For spalte = 1 To lastcolumn
        '        current = .Cells(lastrow, spalte).End(XlDirection.xlUp).row
        '        If current > max Then
        '            max = current
        '        End If
        '    Next spalte

        'End With


        ''zeile = appInstance.ActiveSheet.UsedRange.Rows.Count
        'zeile = max
        WertfuerTop = lastrow * boxHeight + 60   ' starte oben


    End Function
    Function WertfuerTop(ByVal diagrammTyp As Integer) As Double

        WertfuerTop = 1000 + diagrammTyp * 100

    End Function

    'Function WertfuerletzteZeile() As Integer
    '    Dim spalte As Integer
    '    Dim max As Integer
    '    Dim current As Integer
    '    Dim lastrow As Integer, lastcolumn As Integer
    '    Dim suchfeld As Range


    '    With appInstance.Worksheets(arrWsNames(3))
    '        suchfeld = .Range("Projekteingabebereich")
    '        lastrow = suchfeld.Rows.Count
    '        lastcolumn = suchfeld.Columns.Count
    '        max = 0

    '        For spalte = 1 To lastcolumn
    '            current = .Cells(lastrow, spalte).End(XlDirection.xlUp).row
    '            If current > max Then
    '                max = current
    '            End If
    '        Next spalte

    '    End With


    '    WertfuerletzteZeile = max


    'End Function

    '
    '
    '
    '    Sub awinMoveChartUp()
    '        Dim chtobj As ChartObject



    '        On Error GoTo End_of_sub

    '        With appInstance.ActiveChart
    '            chtobj = .Parent
    '        End With

    '        With chtobj
    '            If .top < HoehePrcChart Then
    '                .top = 0
    '            Else
    '                .top = .top - HoehePrcChart
    '            End If
    '        End With

    'End_of_sub:


    '    End Sub

    '
    '
    '
    '    Sub awinMoveChartDown()
    '        Dim chtobj As ChartObject


    '        On Error GoTo End_of_sub

    '        With appInstance.ActiveChart
    '            chtobj = .Parent
    '        End With

    '        With chtobj
    '            .top = .top + HoehePrcChart
    '        End With

    'End_of_sub:


    '    End Sub

    '
    '
    '
    Sub ProjektEdit_DelKey()
        Dim anz_zeilen As Integer, anz_spalten As Integer
        Dim zeile As Integer, spalte As Integer
        Dim i As Integer
        Dim psel As Excel.Range


        appInstance.EnableEvents = False

        psel = appInstance.ActiveWindow.RangeSelection

        With appInstance.Worksheets(arrWsNames(11))
            .Unprotect()
            anz_zeilen = psel.Rows.Count
            anz_spalten = psel.Columns.Count
            zeile = psel.Row
            spalte = psel.Column
            If psel.Rows(1).Interior.color = iProjektFarbe Then
                ' ein ganzes Paket wurde selektiert
                psel.Clear()
                If .Rows(zeile).Interior.ColorIndex = Excel.Constants.xlNone Then
                    ' dann können die kompletten Zeilen gelöscht werden ...
                    For i = 1 To anz_zeilen + 1
                        .Rows(zeile).EntireRow.Delete()
                    Next i
                End If
            ElseIf spalte = 1 And psel(1, 1).Value <> "" Then
                ' die Zeile mit der Rolle soll gelöscht werden ...
                .Rows(zeile).EntireRow.Delete()

            End If
            .Protect()
        End With

        appInstance.EnableEvents = True

    End Sub

    '
    '
    '
    Sub ProjektEdit_InsKey()

        'anz_zeilen = psel.Rows.Count
        'anz_spalten = psel.Columns.Count
        'Application.EnableEvents = False


        'Application.EnableEvents = True

    End Sub




End Module
