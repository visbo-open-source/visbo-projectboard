Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Module awinGUI


   
 
   

    ''' <summary>
    ''' Portfolio - Diagramme erstellen gemäß dem angegebenen charttype
    ''' derzeit möglich: PTpfdk.FitRisiko; PTpfdk.ZeitRisiko; PTpfdk.ComplexRisiko; PTpfdk.FitRisikoVol
    ''' bubbleColor kann derzeit PTpfdk.ProjektFarbe oder PTpfdk.AmpelFarbe sein.
    ''' wird auch aufgerufen aus Visualisieren-Projekt. wenn also projektliste nur ein Element enthält , dann bekommt 
    ''' das Chart eine kennung pr#type#auswahl ansonsten pf#type#auswahl
    ''' </summary>
    ''' <param name="ProjektListe"></param>
    ''' <param name="repChart"></param>
    ''' <param name="isProjektCharakteristik"></param>
    ''' <param name="charttype"></param>
    ''' <param name="bubbleColor"></param>
    ''' <param name="showNegativeValues"></param>
    ''' <param name="showLabels"></param>
    ''' <param name="chartBorderVisible"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Sub awinCreatePortfolioDiagrams(ByRef ProjektListe As Collection, ByRef repChart As Excel.ChartObject, isProjektCharakteristik As Boolean, _
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
        Dim ampelValues() As Long
        Dim positionValues() As String
        Dim diagramTitle As String = ""
        Dim pfDiagram As clsDiagramm
        Dim pfChart As clsEventsPfCharts
        'Dim chtTitle As String
        Dim hilfsstring As String = ""
        Dim chtobjName As String = windowNames(3)
        Dim smallfontsize As Double, titlefontsize As Double
        Dim singleProject As Boolean
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim kennung As String
        Dim tmpCollection As New Collection
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer

        ' Check zu Beginn: gibt es überhaupt etwas zu tun ? 
        ' wenn nein, sofortiger Exit 
        If ProjektListe.Count = 0 Then
            Exit Sub
        End If



        'Dim allOK As Boolean = False
        ' wenn der Charttype nicht bekannt ist : sofortiger Exit 
        If charttype = PTpfdk.FitRisiko Or _
            charttype = PTpfdk.FitRisikoDependency Or _
            charttype = PTpfdk.ZeitRisiko Or _
            charttype = PTpfdk.ComplexRisiko Or _
            charttype = PTpfdk.FitRisikoVol Or _
            charttype = PTpfdk.Dependencies Then

            'allOK = True
        Else
            Exit Sub
        End If

        Dim currentSheetName As String

        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentSheetName = arrWsNames(3)
        Else
            currentSheetName = arrWsNames(5)
        End If

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


        If isProjektCharakteristik And ProjektListe.Count = 1 Then
            pname = CStr(ProjektListe.Item(1))
            hproj = ShowProjekte.getProject(pname)
            tmpCollection.Add(hproj.getShapeText & "#0")
            ' ur: 21.07.2015: Versuch zur Korrektur:
            'kennung = calcChartKennung("pr", PTprdk.StrategieRisiko, tmpCollection)
            kennung = calcChartKennung("pr", charttype, tmpCollection)
        Else
            kennung = calcChartKennung("pf", charttype, ProjektListe)
        End If

        ' Änderung tk
        ' das folgende hängt ja nur ab von Charttype , deswegen ist das immer identisch 
        ' ausserdem ist jetzt sicher, dass charttype ein zulässiger Wert ist
        ' andernfalls wäre das Programm schon beendet
        If isProjektCharakteristik Then
            If ProjektListe.Count = 1 Then
                titelTeile(0) = portfolioDiagrammtitel(charttype) & " " & hproj.getShapeText & vbLf
                titelTeile(1) = " (" & hproj.timeStamp.ToString & ") "

            Else
                titelTeile(0) = portfolioDiagrammtitel(charttype) & vbLf
                titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
            End If

        Else
            titelTeile(0) = portfolioDiagrammtitel(charttype) & vbLf
            titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        End If




        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeilLaengen(1) = titelTeile(1).Length

        diagramTitle = titelTeile(0) & titelTeile(1)


        ' hier werden die Werte bestimmt ...
        Try
            ReDim riskValues(ProjektListe.Count - 1)
            ReDim xAchsenValues(ProjektListe.Count - 1)
            ReDim bubbleValues(ProjektListe.Count - 1)
            ReDim nameValues(ProjektListe.Count - 1)
            ReDim colorValues(ProjektListe.Count - 1)
            ReDim ampelValues(ProjektListe.Count - 1)
            ReDim PfChartBubbleNames(ProjektListe.Count - 1)
            ReDim positionValues(ProjektListe.Count - 1)
        Catch ex As Exception

            appInstance.ScreenUpdating = True
            'Throw New ArgumentException("Fehler in CreatePortfolioDiagramm " & ex.Message)
            Throw New ArgumentException(repMessages.getmsg(70) & ex.Message)

        End Try


        anzBubbles = 0
        ' neuer Typ: 8.3.14 Abhängigkeiten
        Dim tmpstr(10) As String                ' nur für Zeit/Risiko Chart erforderlich
        Dim activeDepIndex As Integer           ' Kennzahl: wieviel Projekte sind abhängig, wie stark strahlt das Projekt 
        Dim passiveDepIndex As Integer          ' Kennzahl: von wievielen Projekten abhängig
        Dim activeNumber As Integer             ' Kennzahl: auf wieviele Projekte strahlt es aus ?
        Dim passiveNumber As Integer            ' Kennzahl: von wievielen Projekten abhängig 
        Dim activeMax As Integer = 0
        Dim passiveMax As Integer = 0

        For i = 1 To ProjektListe.Count
            pname = CStr(ProjektListe.Item(i))
            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj

                    ' neuer Typ: 8.3.14 Abhängigkeiten
                    If charttype = PTpfdk.Dependencies Then
                        ' wird um eins erhöht damit es nicht auf der Nulllinie liegt 
                        activeDepIndex = allDependencies.activeIndex(pname, PTdpndncyType.inhalt) + 1
                        If activeMax < activeDepIndex Then
                            activeMax = activeDepIndex
                        End If
                        activeNumber = allDependencies.activeNumber(pname, PTdpndncyType.inhalt)
                        ' wird um eins erhöht damit es nicht auf der Nulllinie liegt 
                        passiveDepIndex = allDependencies.passiveIndex(pname, PTdpndncyType.inhalt) + 1
                        If passiveMax < passiveDepIndex Then
                            passiveMax = passiveDepIndex
                        End If
                        passiveNumber = allDependencies.passiveNumber(pname, PTdpndncyType.inhalt)
                        riskValues(anzBubbles) = activeDepIndex
                    Else
                        riskValues(anzBubbles) = .Risiko
                    End If


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

                    ' Änderung tk: in ampelValues werden jetzt die Ampelfarben gespeichert 
                    Select Case hproj.ampelStatus
                        Case 0
                            '"Ampel nicht bewertet"
                            ampelValues(anzBubbles) = awinSettings.AmpelNichtBewertet
                        Case 1
                            '"Ampel Grün"
                            ampelValues(anzBubbles) = awinSettings.AmpelGruen
                        Case 2
                            '"Ampel Gelb"
                            ampelValues(anzBubbles) = awinSettings.AmpelGelb
                        Case 3
                            '"Ampel Rot"
                            ampelValues(anzBubbles) = awinSettings.AmpelRot
                    End Select

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

                        Case PTpfdk.FitRisikoDependency
                            xAchsenValues(anzBubbles) = .StrategicFit                                'Stragegie
                            ' wird immer um 1 erhöht, damit der kleinste Wert 1 ist 
                            bubbleValues(anzBubbles) = allDependencies.activeNumber(pname, PTdpndncyType.inhalt) + 1
                            nameValues(anzBubbles) = .name
                            If singleProject Then
                                PfChartBubbleNames(anzBubbles) = " "
                            Else
                                'PfChartBubbleNames(anzBubbles) = .name & _
                                '    " (" & Format(bubbleValues(anzBubbles) - 1, "##0") & " Abh.)"
                                PfChartBubbleNames(anzBubbles) = .name & _
                                    " (" & Format(bubbleValues(anzBubbles) - 1, "##0") & repMessages.getmsg(71)
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


                        Case PTpfdk.Dependencies
                            ' neuer Typ: 8.3.14 Abhängigkeiten

                            xAchsenValues(anzBubbles) = passiveDepIndex                            'Abhängigkeiten
                            bubbleValues(anzBubbles) = .StrategicFit
                            nameValues(anzBubbles) = .name

                            PfChartBubbleNames(anzBubbles) = .name & _
                                " (" & passiveNumber.ToString & ", " & activeNumber.ToString & ")"


                    End Select
                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next


        chtobjName = kennung



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



        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentSheetName), Excel.Worksheet)

            Dim wasProtected As Boolean = .ProtectContents

            If .ProtectContents And visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
                .Unprotect(Password:="x")
                awinSettings.meEnableSorting = True
            End If

            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found
                If chtobjName = .ChartObjects(i).name Then
                    found = True
                    repChart = CType(.ChartObjects(i), Excel.ChartObject)
                    appInstance.ScreenUpdating = True
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

                        If awinSettings.mppShowAmpel Then
                            .Interior.Color = ampelValues(i - 1)
                        Else
                            .Interior.Color = colorValues(i - 1)
                        End If


                        ' alt: 
                        ' Änderung 30.12.15 
                        'If awinSettings.mppShowAmpel Then

                        '    With .Format.Glow
                        '        .Color.RGB = CInt(ampelValues(i - 1))
                        '        .Transparency = 0
                        '        .Radius = 3
                        '    End With

                        'End If
                        ' Ende Änderung 30.12.15

                        ' bei negativen Werten erfolgt die Beschriftung in roter Farbe  ..
                        If bubbleValues(i - 1) < 0 Then
                            Try
                                .DataLabel.Font.Color = awinSettings.AmpelRot
                            Catch ex As Exception

                            End Try
                        ElseIf bubbleValues(i - 1) > 0 Then
                            Try
                                .DataLabel.Font.Color = awinSettings.AmpelGruen
                            Catch ex As Exception

                            End Try
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

                With CType(.Axes(Excel.XlAxisType.xlCategory), Excel.Axis)
                    .HasMajorGridlines = False
                    If charttype = PTpfdk.Dependencies Then
                        .MajorTickMark = XlTickMark.xlTickMarkNone
                        .TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNone
                    End If

                End With

                With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                    .HasMajorGridlines = False
                    If charttype = PTpfdk.Dependencies Then
                        .MajorTickMark = XlTickMark.xlTickMarkNone
                        .TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNone
                    End If
                End With


                Select Case charttype
                    Case PTpfdk.FitRisiko

                        With .Axes(Excel.XlAxisType.xlCategory)
                            .HasTitle = True
                            .MinimumScale = 0
                            .MaximumScale = 11
                            With .AxisTitle
                                '.Characters.text = "strategischer Fit"
                                .Characters.text = repMessages.getmsg(72)
                                .Characters.Font.Size = titlefontsize
                                .Characters.Font.Bold = False
                            End With
                            With .TickLabels.Font
                                .FontStyle = "Normal"
                                .Bold = True
                                .Size = awinSettings.fontsizeItems

                            End With

                        End With

                    Case PTpfdk.FitRisikoDependency
                        With .Axes(Excel.XlAxisType.xlCategory)
                            .HasTitle = True
                            .MinimumScale = 0
                            .MaximumScale = 11
                            With .AxisTitle
                                '.Characters.text = "strategischer Fit"
                                .Characters.text = repMessages.getmsg(72)
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
                                '.Characters.text = "strategischer Fit"
                                .Characters.text = repMessages.getmsg(72)
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
                            .MaximumScale = System.Math.Round(xAchsenValues.Max)
                            With .AxisTitle
                                '.Characters.text = "Projekt-Dauer"
                                .Characters.text = repMessages.getmsg(73)
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
                                '.Characters.text = "Komplexität"
                                .Characters.text = repMessages.getmsg(74)
                                .Characters.Font.Size = titlefontsize
                                .Characters.Font.Bold = False
                            End With
                            With .TickLabels.Font
                                .FontStyle = "Normal"
                                .Bold = True
                                .Size = awinSettings.fontsizeItems

                            End With

                        End With

                    Case PTpfdk.Dependencies
                        With .Axes(Excel.XlAxisType.xlCategory)
                            .HasTitle = True
                            .MinimumScale = 0
                            .MaximumScale = passiveMax + 1

                            With .AxisTitle
                                '.Characters.text = "Abhängigkeit"
                                .Characters.text = repMessages.getmsg(75)
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

                Select Case charttype
                    Case PTpfdk.Dependencies
                        With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                            .HasTitle = True
                            .MinimumScale = 0
                            .MaximumScale = activeMax + 1

                            ' .ReversePlotOrder = True
                            With .AxisTitle
                                '.Characters.Text = "Ausstrahlung"
                                .Characters.Text = repMessages.getmsg(76)
                                .Characters.Font.Size = titlefontsize
                                .Characters.Font.Bold = False
                            End With

                            With .TickLabels.Font
                                .FontStyle = "Normal"
                                .Bold = True
                                .Size = awinSettings.fontsizeItems
                            End With
                        End With

                    Case Else
                        With .Axes(Excel.XlAxisType.xlValue)
                            .HasTitle = True
                            .MinimumScale = 0
                            .MaximumScale = 11
                            ' .ReversePlotOrder = True
                            With .AxisTitle
                                '.Characters.text = "Risiko"
                                .Characters.text = repMessages.getmsg(77)
                                .Characters.Font.Size = titlefontsize
                                .Characters.Font.Bold = False
                            End With

                            With .TickLabels.Font
                                .FontStyle = "Normal"
                                .bold = True
                                .Size = awinSettings.fontsizeItems
                            End With
                        End With
                End Select


                .HasLegend = False
                .HasTitle = True

                .ChartTitle.Text = diagramTitle
                .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend


                ' Events disablen, wegen Report erstellen
                appInstance.EnableEvents = False

                Dim achieved As Boolean = False
                Dim anzahlVersuche As Integer = 0
                Dim errmsg As String = ""
                Do While Not achieved And anzahlVersuche < 10
                    Try
                        'Call Sleep(100)
                        .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=currentSheetName)
                        achieved = True
                    Catch ex As Exception
                        errmsg = ex.Message
                        'Call Sleep(100)
                        anzahlVersuche = anzahlVersuche + 1
                    End Try
                Loop

                If Not achieved Then
                    Throw New ArgumentException("Chart-Fehler:" & errmsg)
                End If

                appInstance.EnableEvents = formerEE
                ' Events sind wieder zurückgesetzt
            End With


            'appInstance.ShowChartTipNames = False
            'appInstance.ShowChartTipValues = False

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


            If isProjektCharakteristik And ProjektListe.Count = 1 Then

                ' ur: 12.03.2015: testweise geändert, Diagramme mit nur einem selektierten Projekt gleich behandeln
                pfDiagram = New clsDiagramm
                ' Anfang Event Handling für Chart 
                pfChart = New clsEventsPfCharts
                pfChart.PfChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
                pfDiagram.setDiagramEvent = pfChart
                ' Ende Event Handling für Chart 

                With pfDiagram
                    .kennung = kennung
                    '.kennung = calcChartKennung("pr", PTprdk.StrategieRisiko, tmpCollection)
                    .DiagrammTitel = diagramTitle
                    .diagrammTyp = DiagrammTypen(3)                     ' Portfolio
                    .gsCollection = ProjektListe
                    .isCockpitChart = False
                    ' ur:09.03.2015: wegen Chart-Resize geändert
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                End With

                DiagramList.Add(pfDiagram)
            Else


                pfDiagram = New clsDiagramm
                ' Anfang Event Handling für Chart 
                pfChart = New clsEventsPfCharts
                pfChart.PfChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
                pfDiagram.setDiagramEvent = pfChart
                ' Ende Event Handling für Chart 

                With pfDiagram
                    .kennung = kennung
                    '.kennung = calcChartKennung("pf", charttype, ProjektListe)
                    .DiagrammTitel = diagramTitle
                    .diagrammTyp = DiagrammTypen(3)                     ' Portfolio
                    .gsCollection = ProjektListe
                    .isCockpitChart = False
                    ' ur:09.03.2015: wegen Chart-Resize geändert
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                End With

                DiagramList.Add(pfDiagram)
            End If

            ' wenn es geschützt war .. 
            ''If wasProtected And visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
            ''    .Protect(Password:="x", UserInterfaceOnly:=True, _
            ''                 AllowFormattingCells:=True, _
            ''                 AllowInsertingColumns:=False,
            ''                 AllowInsertingRows:=True, _
            ''                 AllowDeletingColumns:=False, _
            ''                 AllowDeletingRows:=True, _
            ''                 AllowSorting:=True, _
            ''                 AllowFiltering:=True)
            ''    .EnableSelection = XlEnableSelection.xlUnlockedCells
            ''    .EnableAutoFilter = True
            ''End If

            repChart = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

        End With

        appInstance.ScreenUpdating = formerSU

    End Sub  ' Ende Prozedur awinCreatePortfolioChartDiagramm

    
    ''' <summary>
    ''' aktualisiert die Portfolio Charts 
    ''' Strategie, Risiko, Marge oder oder andere, ähnliche 
    ''' </summary>
    ''' <param name="chtobj"></param>
    ''' <param name="bubbleColor">gibt an, ob di eAMpelfarbe des Projekts gezeigt werden soll</param>
    ''' <remarks></remarks>
    '''
    Sub awinUpdatePortfolioDiagrams(ByVal chtobj As ChartObject, bubbleColor As Integer)

        Dim i As Integer
        Dim pname As String
        Dim hproj As New clsProjekt
        Dim anzBubbles As Integer
        Dim riskValues() As Double, bubbleValues() As Double, tempArray() As Double
        Dim xAchsenValues() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim ampelValues() As Long
        Dim positionValues() As String
        Dim diagramTitle As String
        Dim showLabels As Boolean
        Dim showNegativeValues As Boolean = False
        Dim projektListe As New Collection
        Dim charttype As Integer
        Dim tmpstr(5) As String
        Dim isSingleProject As Boolean = False
        Dim datalabelSize As Integer = 0
        Dim datalabelColor As Double = 0
        Dim datalabelFont As Excel.Font
        Dim datalabel As Excel.DataLabel = Nothing

        'Dim pfDiagram As clsDiagramm
        'Dim pfChart As clsEventsPfCharts
        'Dim TypeCollection As Collection
        'Dim charttype As Integer

        ' hier wird in der Objektkennung nachgesehen, von welchem Typ dieses Portfolio-Diagramm ist
        ' PTpfdk.FitRisiko oder PTpfdk.ZeitRisiko oder PTpfdk.ComplexRisiko

        tmpstr = chtobj.Name.Trim.Split(New Char() {CChar("#")}, 4)
        If tmpstr(0) = "pr" Then
            isSingleProject = True
            Dim komplett As String = tmpstr(2)
            tmpstr = komplett.Split(New Char() {CChar("("), CChar(")")}, 4)
            If tmpstr.Length > 1 Then
                projektListe.Add(tmpstr(0))
            Else
                projektListe.Add(komplett)
            End If

        Else
            isSingleProject = False
            Dim selectionType As Integer = -1 ' keine Einschränkung
            projektListe = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)
        End If
        charttype = CInt(tmpstr(1))

        'Dim allOK As Boolean = False
        ' wenn der Charttype nicht bekannt ist : sofortiger Exit 
        If charttype = PTpfdk.FitRisiko Or _
            charttype = PTpfdk.FitRisikoDependency Or _
            charttype = PTpfdk.ZeitRisiko Or _
            charttype = PTpfdk.ComplexRisiko Or _
            charttype = PTpfdk.FitRisikoVol Or _
            charttype = PTpfdk.Dependencies Then

            'allOK = True
        Else
            Exit Sub
        End If


        'foundDiagramm = DiagramList.getDiagramm(chtobj.Name)
        ' event. für eine Erweiterung benötigt


        ' hier werden die Werte bestimmt ...
        Try
            ReDim riskValues(ShowProjekte.Count - 1)
            ReDim xAchsenValues(ShowProjekte.Count - 1)
            ReDim bubbleValues(ShowProjekte.Count - 1)
            ReDim nameValues(ShowProjekte.Count - 1)
            ReDim colorValues(ShowProjekte.Count - 1)
            ReDim ampelValues(ShowProjekte.Count - 1)
            ReDim PfChartBubbleNames(ShowProjekte.Count - 1)
            ReDim positionValues(ShowProjekte.Count - 1)
        Catch ex As Exception
            'Throw New ArgumentException("Fehler in UpdatePortfolioDiagramm " & ex.Message)
            Throw New ArgumentException(repMessages.getmsg(78) & ex.Message)
        End Try


        ' neuer Typ: 8.3.14 Abhängigkeiten
        Dim activeDepIndex As Integer           ' Kennzahl: wieviel Projekte sind abhängig, wie stark strahlt das Projekt 
        Dim passiveDepIndex As Integer          ' Kennzahl: von wievielen Projekten abhängig
        Dim activeNumber As Integer             ' Kennzahl: auf wieviele Projekte strahlt es aus ?
        Dim passiveNumber As Integer            ' Kennzahl: von wievielen Projekten abhängig 

        anzBubbles = 0

        ' Änderung 8.3 : hier muss die Unterscheidung gemacht werden, welche Projekte im Zeitraum denn überhaupt Abhängigkeiten haben  

        If charttype = PTpfdk.Dependencies Then
            Dim deleteList As New Collection
            For i = 1 To projektListe.Count
                pname = CStr(projektListe.Item(i))
                Try
                    hproj = ShowProjekte.getProject(pname)
                    activeNumber = allDependencies.activeNumber(pname, PTdpndncyType.inhalt)
                    passiveNumber = allDependencies.passiveNumber(pname, PTdpndncyType.inhalt)
                    If activeNumber = 0 And passiveNumber = 0 Then
                        deleteList.Add(pname)
                    End If
                Catch ex As Exception

                End Try
            Next

            ' jetzt müssen die Projekte rausgenommen werden, die keine Abhängigkeiten haben 
            For i = 1 To deleteList.Count
                pname = CStr(deleteList.Item(i))
                Try
                    projektListe.Remove(pname)
                Catch ex As Exception

                End Try
            Next
        End If


        For i = 1 To projektListe.Count
            pname = CStr(projektListe.Item(i))
            Try
                hproj = ShowProjekte.getProject(pname)
                With hproj

                    ' neuer Typ: 8.3.14 Abhängigkeiten
                    If charttype = PTpfdk.Dependencies Then
                        ' wird um eins erhöht , damit es nicht auf der Nullinie liegt 
                        activeDepIndex = allDependencies.activeIndex(pname, PTdpndncyType.inhalt) + 1
                        activeNumber = allDependencies.activeNumber(pname, PTdpndncyType.inhalt)
                        ' wird um eins erhöht , damit es nicht auf der Nullinie liegt 
                        passiveDepIndex = allDependencies.passiveIndex(pname, PTdpndncyType.inhalt) + 1
                        passiveNumber = allDependencies.passiveNumber(pname, PTdpndncyType.inhalt)
                        riskValues(anzBubbles) = activeDepIndex
                    Else
                        riskValues(anzBubbles) = .Risiko
                    End If

                    ' Änderung tk 2.6.15 es wird immer die Projektfarbe gezeigt, Ampelfarbe nur bei Anforderung
                    Select Case .ampelStatus
                        Case 0
                            '"Ampel nicht bewertet"
                            ampelValues(anzBubbles) = awinSettings.AmpelNichtBewertet
                        Case 1
                            '"Ampel Grün"
                            ampelValues(anzBubbles) = awinSettings.AmpelGruen
                        Case 2
                            '"Ampel Gelb"
                            ampelValues(anzBubbles) = awinSettings.AmpelGelb
                        Case 3
                            '"Ampel Rot"
                            ampelValues(anzBubbles) = awinSettings.AmpelRot
                    End Select


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

                        Case PTpfdk.FitRisikoDependency
                            xAchsenValues(anzBubbles) = .StrategicFit                                'Stragegie
                            ' wird immer um 1 erhöht, damit der kleinste Wert 1 ist 
                            bubbleValues(anzBubbles) = allDependencies.activeNumber(pname, PTdpndncyType.inhalt) + 1
                            nameValues(anzBubbles) = .name

                            'PfChartBubbleNames(anzBubbles) = .name & _
                            '        " (" & Format(bubbleValues(anzBubbles) - 1, "##0") & " Abh.)"
                            PfChartBubbleNames(anzBubbles) = .name & _
                                   " (" & Format(bubbleValues(anzBubbles) - 1, "##0") & repMessages.getmsg(71)

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


                        Case PTpfdk.Dependencies
                            ' neuer Typ: 8.3.14 Abhängigkeiten

                            xAchsenValues(anzBubbles) = passiveDepIndex                            'Abhängigkeiten
                            bubbleValues(anzBubbles) = .StrategicFit
                            nameValues(anzBubbles) = .name

                            PfChartBubbleNames(anzBubbles) = .name & _
                                " (" & passiveNumber.ToString & ", " & activeNumber.ToString & ")"

                    End Select
                End With
                anzBubbles = anzBubbles + 1
            Catch ex As Exception

            End Try
        Next


        ' Änderung tk 7.1.16
        ' das hängt ja nur von charttype ab ... 
        diagramTitle = portfolioDiagrammtitel(PTpfdk.FitRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)

        ' ab 7.1.16 auskommentiert 
        'Select Case charttype
        '    Case PTpfdk.FitRisiko

        '        diagramTitle = portfolioDiagrammtitel(PTpfdk.FitRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)

        '    Case PTpfdk.FitRisikoVol

        '        diagramTitle = portfolioDiagrammtitel(PTpfdk.FitRisikoVol) & vbLf & textZeitraum(showRangeLeft, showRangeRight)

        '    Case PTpfdk.ZeitRisiko

        '        diagramTitle = portfolioDiagrammtitel(PTpfdk.ZeitRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)

        '    Case PTpfdk.ComplexRisiko

        '        diagramTitle = portfolioDiagrammtitel(PTpfdk.ComplexRisiko) & vbLf & textZeitraum(showRangeLeft, showRangeRight)

        '    Case PTpfdk.Dependencies
        '        ' neuer Typ: 8.3.14 Abhängigkeiten

        '        diagramTitle = portfolioDiagrammtitel(PTpfdk.Dependencies) & vbLf & textZeitraum(showRangeLeft, showRangeRight)

        '    Case Else
        '        diagramTitle = "Chart-Typ existiert nicht"
        'End Select


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

            If projektListe.Count >= 0 Then

                ' Einstellungen der vorhandenen SeriesCollection merken

                Dim pts As Excel.Points = CType(.SeriesCollection(1).Points, Excel.Points)
                Dim dlFontSize As Double
                Dim dlFontBackground As Double
                Dim dlFontBold As Boolean
                Dim dlFontColorIndex As Integer
                Dim dlFontColor As Integer
                Dim dlFontFontStyle As String = ""
                Dim dlFontItalic As Boolean
                Dim dlFontStrikethrough As Boolean
                Dim dlFontSuperscript As Boolean
                Dim dlFontSubscript As Boolean
                Dim dlFontUnderline As Double

                For i = 1 To pts.Count

                    With CType(.SeriesCollection(1).Points(i), Excel.Point)

                        Try
                            If .HasDataLabel = True Then
                                datalabel = .DataLabel
                                With .DataLabel
                                    datalabelFont = .Font
                                    dlFontSize = CDbl(.Font.Size)
                                    dlFontBackground = CDbl(.Font.Background)
                                    dlFontBold = CBool(.Font.Bold)
                                    dlFontColorIndex = CInt(.Font.ColorIndex)
                                    dlFontFontStyle = CStr(.Font.FontStyle)
                                    dlFontItalic = CBool(.Font.Italic)
                                    dlFontStrikethrough = CBool(.Font.Strikethrough)
                                    dlFontSubscript = CBool(.Font.Subscript)
                                    dlFontSuperscript = CBool(.Font.Superscript)
                                    dlFontUnderline = CDbl(.Font.Underline)
                                    dlFontSize = CDbl(.Font.Size)
                                    If .Font.Color <> awinSettings.AmpelRot Then
                                        dlFontColor = CInt(.Font.Color)
                                    End If

                                End With
                            End If

                        Catch ex As Exception

                        End Try
                    End With

                Next i

                ' remove old series
                Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete()
                Loop

                ' nur dann neue Series-Collection aufbauen, wenn auch tatsächlich was in der Projektliste ist ..


                .SeriesCollection.NewSeries()
                .SeriesCollection(1).name = diagramTitle
                .SeriesCollection(1).ChartType = Excel.XlChartType.xlBubble3DEffect


                ' ur: 06.04.2017: nur wenn Werte für die SeriesCollection vorhanden sind
                ' d.h. ProjekteListe ist nicht leer
                If projektListe.Count > 0 Then

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

                    'Dim testName As String

                    Dim bubblePoint As Excel.Point
                    For i = 1 To anzBubbles

                        bubblePoint = CType(.SeriesCollection(1).Points(i), Excel.Point)

                        With CType(.SeriesCollection(1).Points(i), Excel.Point)

                            If showLabels Then
                                Try
                                    .HasDataLabel = True

                                    With .DataLabel
                                        .Text = PfChartBubbleNames(i - 1)
                                        .Font.Size = dlFontSize
                                        .Font.Background = dlFontBackground
                                        .Font.Bold = dlFontBold
                                        .Font.Color = dlFontColor
                                        .Font.ColorIndex = dlFontColorIndex
                                        .Font.FontStyle = dlFontFontStyle
                                        .Font.Italic = dlFontItalic
                                        .Font.Size = dlFontSize
                                        .Font.Strikethrough = dlFontStrikethrough
                                        .Font.Subscript = dlFontSubscript
                                        .Font.Superscript = dlFontSuperscript
                                        .Font.Underline = dlFontUnderline

                                        'ur: 17.7.2014: fontsize kommt vom existierenden Chart
                                        '.Font.Size = awinSettings.CPfontsizeItems

                                        ' bei negativen Werten erfolgt die Beschriftung in roter Farbe  ..

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

                                ' bei negativen Werten erfolgt die Beschriftung in roter Farbe  ..
                                If bubbleValues(i - 1) < 0 Then
                                    Try
                                        .DataLabel.Font.Color = awinSettings.AmpelRot
                                    Catch ex As Exception

                                    End Try
                                ElseIf bubbleValues(i - 1) > 0 Then
                                    Try
                                        .DataLabel.Font.Color = awinSettings.AmpelGruen
                                    Catch ex As Exception

                                    End Try
                                Else
                                    Try
                                        .DataLabel.Font.Color = System.Drawing.Color.Black
                                    Catch ex As Exception

                                    End Try
                                End If

                            Else
                                .HasDataLabel = False
                            End If

                            ' Änderung 30.12.15 
                            If awinSettings.mppShowAmpel Then
                                .Interior.Color = ampelValues(i - 1)
                            Else
                                .Interior.Color = colorValues(i - 1)
                            End If


                            ' Änderung wenn ampeln gezeigt werden sollen ...

                            'If awinSettings.mppShowAmpel Then

                            '    With .Format.Glow
                            '        .Color.RGB = CInt(ampelValues(i - 1))
                            '        .Transparency = 0
                            '        .Radius = 3
                            '    End With

                            'End If
                            ' Ende Äderung 30.12.15

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
                Else
                    ' Projektliste ist leer
                End If

            End If



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

                If .tfspalte + .anzahlRasterElemente - 1 > lastcolumn Then
                    lastcolumn = .tfspalte + .anzahlRasterElemente - 1
                End If

            End With


        Next

        WertfuerTop = lastrow * boxHeight + 60   ' starte oben


    End Function
    Function WertfuerTop(ByVal diagrammTyp As Integer) As Double

        WertfuerTop = 1000 + diagrammTyp * 100

    End Function




End Module
