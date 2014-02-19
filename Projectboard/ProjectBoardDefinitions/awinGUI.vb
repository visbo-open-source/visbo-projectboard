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


 
   
    ' Portfolio - Diagramme erstellen gemäß dem angegebenen charttype
    ' derzeit möglich: PTpfdk.FitRisiko; PTpfdk.ZeitRisiko; PTpfdk.ComplexRisiko
    ' bubbleColor kann derzeit PTpfdk.ProjektFarbe oder PTpfdk.AmpelFarbe sein.

    Sub awinCreatePortfolioDiagramms(ByRef ProjektListe As Collection, ByRef repChart As Object, isProjektCharakteristik As Boolean, _
                                     charttype As Integer, bubbleColor As Integer, showNegativeValues As Boolean, showLabels As Boolean, chartBorderVisible As Boolean, _
                                     top As Double, left As Double, width As Double, height As Double)

        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim pname As String = ""
        Dim hproj As New clsProjekt
        Dim anzBubbles As Integer
        Dim riskValues() As Double
        Dim xAchsenValues() As Double
        Dim bubbleValues() As Double, tempArray() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String = ""
        Dim pfDiagram As clsDiagramm
        Dim pfChart As clsEventsPfCharts
        Dim chtTitle As String
        Dim hilfsstring As String = ""
        Dim chtobjName As String = windowNames(3)
        Dim smallfontsize As Double, titlefontsize As Double
        Dim singleProject As Boolean
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents

        appInstance.ScreenUpdating = False

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

        Select Case charttype
            Case PTpfdk.FitRisiko

                If isProjektCharakteristik Then
                    diagramTitle = "Charakteristik " & summentitel2
                Else
                    diagramTitle = summentitel2 & vbLf & textZeitraum(showRangeLeft, showRangeRight)
                End If

            Case PTpfdk.ZeitRisiko

                If isProjektCharakteristik Then
                    diagramTitle = portfolioDiagrammtitel(PTpfdk.ZeitRisiko)
                    diagramTitle = "Charakteristik " & diagramTitle
                Else
                    diagramTitle = portfolioDiagrammtitel(PTpfdk.ZeitRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)
                End If

            Case PTpfdk.ComplexRisiko

                If isProjektCharakteristik Then
                    diagramTitle = portfolioDiagrammtitel(PTpfdk.ComplexRisiko)
                    diagramTitle = "Charakteristik " & diagramTitle
                Else
                    diagramTitle = portfolioDiagrammtitel(PTpfdk.ComplexRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)
                End If

            Case PTpfdk.FitRisikoVol
                If isProjektCharakteristik Then
                    diagramTitle = portfolioDiagrammtitel(PTpfdk.FitRisikoVol)
                    diagramTitle = "Charakteristik " & diagramTitle
                Else
                    diagramTitle = portfolioDiagrammtitel(PTpfdk.FitRisikoVol) & vbLf & textZeitraum(showRangeLeft, showRangeRight)
                End If

        End Select

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

        Dim tmpstr(10) As String                ' nur für Zeit/Risiko Chart erforderlich

        For i = 1 To ProjektListe.Count
            pname = ProjektListe.Item(i)
            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    riskValues(anzBubbles) = .Risiko

                    If bubbleColor = PTpfdk.ProjektFarbe Then

                        ' Projekttyp wird farblich gekennzeichent
                        colorValues(anzBubbles) = .farbe

                    Else ' bubbleColor ist AmpelFarbe

                        ' ProjektStatus wird farblich gekennzeichnet
                        Select Case hproj.ampelStatus
                            Case 0
                                '"Ampel nicht bewertet"
                                colorValues(anzBubbles) = awinSettings.AmpelNichtBewertet
                            Case 1
                                '"Ampel Grün"
                                colorValues(anzBubbles) = awinSettings.AmpelGruen
                            Case 2
                                '"Ampel Gelb"
                                colorValues(anzBubbles) = awinSettings.AmpelGelb
                            Case 3
                                '"Ampel Rot"
                                colorValues(anzBubbles) = awinSettings.AmpelRot
                        End Select
                    End If

                    Select Case charttype
                        Case PTpfdk.FitRisiko

                            xAchsenValues(anzBubbles) = .StrategicFit                                'Stragegie
                            bubbleValues(anzBubbles) = .ProjectMarge                                ' Marge
                            nameValues(anzBubbles) = .name
                            If singleProject Then
                                PfChartBubbleNames(anzBubbles) = Format(bubbleValues(anzBubbles), "##0.#%")
                            Else
                                PfChartBubbleNames(anzBubbles) = .name & _
                                    " (" & Format(bubbleValues(anzBubbles), "##0.#%") & ")"
                            End If
                        Case PTpfdk.FitRisikoVol

                            xAchsenValues(anzBubbles) = .StrategicFit                                'Stragegie
                            bubbleValues(anzBubbles) = .volume
                            nameValues(anzBubbles) = .name

                            PfChartBubbleNames(anzBubbles) = .name & _
                                " (" & Format(bubbleValues(anzBubbles) / 1000, "##0.#") & " T)"

                        Case PTpfdk.ZeitRisiko

                            xAchsenValues(anzBubbles) = .dauerInDays / 365 * 12                    'Zeit
                            bubbleValues(anzBubbles) = System.Math.Round(.volume / 10000) * 10
                            nameValues(anzBubbles) = .name & " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)"
                            PfChartBubbleNames(anzBubbles) = .name & _
                                    " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)"

                        Case PTpfdk.ComplexRisiko

                            xAchsenValues(anzBubbles) = .complexity                                'Complex
                            bubbleValues(anzBubbles) = .volume                                     'Volumen
                            nameValues(anzBubbles) = .name
                            PfChartBubbleNames(anzBubbles) = .name & _
                             " (" & Format(bubbleValues(anzBubbles) / 1000, "##0.#") & " T)"


                    End Select
                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next

        If singleProject Then
            chtobjName = getKennung("pf", charttype, ProjektListe)
        Else
            chtobjName = getKennung("pf", charttype, ProjektListe)
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
                Select Case charttype
                    Case PTpfdk.FitRisiko

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

                    Case Else
                        ' für Zeit/Risiko und 
                        ' für Complex/Risiko und 
                        ' für Strategie/FitRisikoVolume
                        hilfsstring = .chartObjects(i).name
                        If chtobjName = .chartObjects(i).name Then
                            found = True
                            repChart = .ChartObjects(i)
                            Exit Sub
                        Else
                            i = i + 1
                        End If
                End Select
            End While


            ReDim tempArray(anzBubbles - 1)


            With appInstance.Charts.Add

                .SeriesCollection.NewSeries()
                .SeriesCollection(1).name = diagramTitle
                .SeriesCollection(1).ChartType = Excel.XlChartType.xlBubble3DEffect

                For i = 1 To anzBubbles
                    tempArray(i - 1) = xAchsenValues(i - 1)
                Next i
                .SeriesCollection(1).XValues = tempArray

                For i = 1 To anzBubbles
                    tempArray(i - 1) = riskValues(i - 1)
                Next i
                .SeriesCollection(1).Values = tempArray

                For i = 1 To anzBubbles
                    If bubbleValues(i - 1) < 0.01 And bubbleValues(i - 1) > -0.01 Then
                        tempArray(i - 1) = 0.01
                    ElseIf bubbleValues(i - 1) < 0 Then
                        ' negative Werte werden Positiv dargestellt mit roten Beschriftung siehe unten
                        tempArray(i - 1) = System.Math.Abs(bubbleValues(i - 1))
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


                For i = 1 To anzBubbles

                    With CType(.SeriesCollection(1).Points(i), Excel.Point)

                        If showLabels Then
                            Try
                                .HasDataLabel = True

                                With .DataLabel
                                    .Text = PfChartBubbleNames(i - 1)
                                    '.Text = nameValues(i - 1)
                                    If singleProject Then
                                        .Font.Size = awinSettings.CPfontsizeItems + 4
                                    Else
                                        .Font.Size = awinSettings.CPfontsizeItems
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

                        .Interior.Color = colorValues(i - 1)

                        ' bei negativen Werten erfolgt die Beschriftung in roter Farbe  ..
                        If bubbleValues(i - 1) < 0 Then
                            .DataLabel.Font.Color = awinSettings.AmpelRot
                        End If

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
                Select Case charttype
                    Case PTpfdk.FitRisiko

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

                    Case PTpfdk.FitRisikoVol

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

                    Case PTpfdk.ZeitRisiko

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

                    Case PTpfdk.ComplexRisiko

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

                End Select


                ' für , Strategie/RisikoMarge,Strategie/RisikoVolume, Zeit/Risiko und Complex/Risiko gültig

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

                ' Events disablen, wegen Report erstellen
                appInstance.EnableEvents = False
                .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                appInstance.EnableEvents = formerEE
                ' Events sind wieder zurückgesetzt
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

            pfChart = New clsEventsPfCharts
            pfChart.PfChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            pfDiagram.setDiagramEvent = pfChart

            With pfDiagram

                .kennung = getKennung("pf", charttype, ProjektListe)
                .DiagrammTitel = diagramTitle
                .diagrammTyp = DiagrammTypen(3)                     ' Portfolio
                .gsCollection = ProjektListe
                .isCockpitChart = False

            End With

            DiagramList.Add(pfDiagram)
            repChart = .ChartObjects(anzDiagrams + 1)

        End With

        appInstance.ScreenUpdating = formerSU

    End Sub  ' Ende Prozedur awinCreatePortfolioChartDiagramm



    '
    ' Prozedur für den Update des Portfolio Diagramms
    '
    Sub awinUpdatePortfolioDiagrams(ByVal chtobj As ChartObject, bubbleColor As Integer)

        Dim i As Integer
        Dim pname As String
        Dim hproj As New clsProjekt
        Dim anzBubbles As Integer
        Dim riskValues() As Double, bubbleValues() As Double, tempArray() As Double
        Dim xAchsenValues() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String
        Dim showLabels As Boolean
        Dim showNegativeValues As Boolean = True
        Dim projektListe As Collection
        Dim charttype As Integer
        Dim chartkennung As String
        Dim tmpstr(3) As String
        'Dim foundDiagramm As Boolean
        'Dim pfDiagram As clsDiagramm
        'Dim pfChart As clsEventsPfCharts
        'Dim TypeCollection As Collection
        'Dim charttype As Integer

        ' hier wird in der Objektkennung nachgesehen, von welchem Typ dieses Portfolio-Diagramm ist
        ' PTpfdk.FitRisiko oder PTpfdk.ZeitRisiko oder PTpfdk.ComplexRisiko

        chartkennung = chtobj.Name
        tmpstr = chartkennung.Trim.Split(New Char() {"#"}, 3)
        charttype = tmpstr(1)

        'foundDiagramm = DiagramList.getDiagramm(chtobj.Name)
        ' event. für eine Erweiterung benötigt


        ' hier werden die Werte bestimmt ...
        Try
            ReDim riskValues(ShowProjekte.Count - 1)
            ReDim xAchsenValues(ShowProjekte.Count - 1)
            ReDim bubbleValues(ShowProjekte.Count - 1)
            ReDim nameValues(ShowProjekte.Count - 1)
            ReDim colorValues(ShowProjekte.Count - 1)
            ReDim PfChartBubbleNames(ShowProjekte.Count - 1)
            ReDim positionValues(ShowProjekte.Count - 1)
        Catch ex As Exception
            Throw New ArgumentException("Fehler in UpdatePortfolioDiagramm " & ex.Message)
        End Try




        anzBubbles = 0

        Dim selectionType As Integer = -1 ' keine Einschränkung
        projektListe = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        For i = 1 To projektListe.Count
            pname = projektListe.Item(i)
            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj
                    riskValues(anzBubbles) = .Risiko
                    If bubbleColor = PTpfdk.ProjektFarbe Then

                        ' Projekttyp wird farblich gekennzeichent
                        colorValues(anzBubbles) = .farbe

                    Else ' bubbleColor ist AmpelFarbe

                        ' ProjektStatus wird farblich gekennzeichnet
                        Select Case .ampelStatus
                            Case 0
                                '"Ampel nicht bewertet"
                                colorValues(anzBubbles) = awinSettings.AmpelNichtBewertet
                            Case 1
                                '"Ampel Grün"
                                colorValues(anzBubbles) = awinSettings.AmpelGruen
                            Case 2
                                '"Ampel Gelb"
                                colorValues(anzBubbles) = awinSettings.AmpelGelb
                            Case 3
                                '"Ampel Rot"
                                colorValues(anzBubbles) = awinSettings.AmpelRot
                        End Select
                    End If

                    Select Case charttype
                        Case PTpfdk.FitRisiko

                            xAchsenValues(anzBubbles) = .StrategicFit                                'Stragegie
                            bubbleValues(anzBubbles) = .ProjectMarge
                            nameValues(anzBubbles) = .name
                            PfChartBubbleNames(anzBubbles) = hproj.name & _
                                    " (" & Format(bubbleValues(anzBubbles), "##0.#%") & ")" 'Strategie/Rsiko

                        Case PTpfdk.FitRisikoVol

                            xAchsenValues(anzBubbles) = .StrategicFit                                'Stragegie
                            bubbleValues(anzBubbles) = .volume                               '   Volumen
                            nameValues(anzBubbles) = .name
                            PfChartBubbleNames(anzBubbles) = hproj.name & _
                                    " (" & Format(bubbleValues(anzBubbles) / 1000, "##0.#") & " T)"


                        Case PTpfdk.ZeitRisiko

                            xAchsenValues(anzBubbles) = .dauerInDays / 365 * 12                    'Zeit
                            bubbleValues(anzBubbles) = System.Math.Round(.volume / 10000) * 10
                            'tmpstr = .name.Split(New Char() {" "}, 10)                             'Zeit/Risiko
                            'nameValues(anzBubbles) = tmpstr(0) & " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)" 
                            nameValues(anzBubbles) = .name & " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)"
                            PfChartBubbleNames(anzBubbles) = .name & _
                                    " (" & Format(bubbleValues(anzBubbles), "##0.#") & " T)"

                        Case PTpfdk.ComplexRisiko

                            xAchsenValues(anzBubbles) = .complexity                                'Complex
                            bubbleValues(anzBubbles) = .volume                                     'Bubblegröße gemäß Volumen
                            nameValues(anzBubbles) = .name
                            PfChartBubbleNames(anzBubbles) = hproj.name & _
                             " (" & Format(bubbleValues(anzBubbles) / 1000, "##0.#") & " T)"


                    End Select
                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next

        Select Case charttype
            Case PTpfdk.FitRisiko

                diagramTitle = summentitel2 & vbLf & textZeitraum(showRangeLeft, showRangeRight)

            Case PTpfdk.FitRisikoVol

                diagramTitle = portfolioDiagrammtitel(PTpfdk.FitRisikoVol) & vbLf & textZeitraum(showRangeLeft, showRangeRight)

            Case PTpfdk.ZeitRisiko

                diagramTitle = portfolioDiagrammtitel(PTpfdk.ZeitRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)

            Case PTpfdk.ComplexRisiko

                diagramTitle = portfolioDiagrammtitel(PTpfdk.ComplexRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)
            Case Else
                diagramTitle = "Chart-Typ existiert nicht"
        End Select



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


        ReDim tempArray(anzBubbles - 1)

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False



        With chtobj.Chart
          
            showLabels = True

            ' remove old series
            Do Until .SeriesCollection.Count = 0
                .SeriesCollection(1).Delete()
            Loop

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

                With CType(.SeriesCollection(1).Points(i), Excel.Point)

                    If showLabels Then
                        Try
                            .HasDataLabel = True
                            With .DataLabel
                                .Text = PfChartBubbleNames(i - 1)
                                .Font.Size = awinSettings.CPfontsizeItems

                                ' bei negativen Werten erfolgt die Beschriftung in roter Farbe  ..
                                If bubbleValues(i - 1) < 0 Then
                                    .Font.Color = awinSettings.AmpelRot
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

                    .Interior.Color = colorValues(i - 1)
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

                If .tfspalte + .Dauer - 1 > lastcolumn Then
                    lastcolumn = .tfspalte + .Dauer - 1
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
