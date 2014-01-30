Imports System.Math
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Public Module awinDiagrams

    '
    ' zeigt im Planungshorizont die Time Zone an - oder blendet sie aus, abhängg vom Wert showzone
    '
    Sub awinShowtimezone(ByVal von As Integer, ByVal bis As Integer, ByVal showzone As Boolean)
        Dim laenge As Integer



        laenge = bis - von

        If von > 0 And laenge > 0 Then

            With appInstance.Worksheets(arrWsNames(3))

                If showzone Then
                    '
                    ' die erste Zeile im Bereich einfärben
                    '
                    .Range(.Cells(1, von), .Cells(1, von).Offset(0, laenge)).Interior.color = showtimezone_color
                   
                Else
                    '
                    ' die erste Zeile im Bereich einfärben
                    '
                    .Range(.Cells(1, von), .Cells(1, von).Offset(0, laenge)).Interior.color = noshowtimezone_color
                    
                End If

            End With


        End If



    End Sub

    ''' <summary>
    ''' löscht Window und Cockpit Window vom Typ "prcTyp"
    ''' </summary>
    ''' <param name="prctyp"></param>
    ''' <remarks></remarks>
    Sub awinDeleteCockpitWindow(ByVal prctyp As Integer)
        Dim Test As Excel.Window

        Try
            Test = appInstance.ActiveWorkbook.Windows(windowNames(prctyp))
        Catch ex As Exception
            Exit Sub
        End Try

        Test.Close()
        Call awinLoescheCockpitCharts(prctyp)

    End Sub
    


    ''' <summary>
    ''' erzeugt ein Phasen-/Rollen-/Kostenart - Diagramm
    ''' bekommt Parameter für die darzustellenden Rollen mit, die Position, ob es ein Cockpit Chart ist und um welchen Diagramm-Typ es sich handelt
    ''' Diagramm-Typen:
    ''' 0 - Phase
    ''' 1 - Rolle
    ''' 2 - Kostenart
    ''' 3 - Portfolio
    ''' 4 - Summe  
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <param name="prcTyp"></param>
    ''' <remarks>ist nicht mehr notwendig - alles was für Anzeige der selektierten getan werden muss 
    ''' steht in der updatePRC CollectionDiagrams </remarks>
    Sub CSCreateprcCollectionDiagram(ByRef myCollection As Collection, ByRef repObj As Object, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                       ByVal isCockpitChart As Boolean, ByVal prcTyp As String, ByVal calledfromReporting As Boolean)

        Dim von As Integer, bis As Integer

        Dim anzDiagrams As Integer, i As Integer, m As Integer, r As Integer
        Dim found As Boolean
        'Dim korr_abstand As Double
        Dim minwert As Double, maxwert As Double
        Dim nr_pts As Integer
        Dim diagramTitle As String
        Dim objektFarbe As Object
        Dim Xdatenreihe() As String
        Dim datenreihe() As Double, edatenreihe() As Double, seriesSumDatenreihe() As Double
        Dim seldatenreihe() As Double, tmpdatenreihe() As Double
        Dim kdatenreihe() As Double ' nimmt die Kapa-Werte für das Diagramm auf
        Dim prcName As String
        Dim startdate As Date
        Dim diff As Object
        Dim mindone As Boolean, maxdone As Boolean
        Dim VarValues() As Double
        Dim prcDiagram As clsDiagramm
        'Dim prcChart As clsAwinEvent
        Dim prcChart As clsEventsPrcCharts
        Dim isPersCost As Boolean
        Dim isWeightedValues As Boolean = False
        Dim lastSC As Integer
        Dim titleZeitraum As String, titleSumme As String, einheit As String
        'Dim chtTitle As String
        Dim chtobjName As String
        Dim selectionFarbe As Long = awinSettings.AmpelNichtBewertet

        ' Debugging variable 
        Dim HDiagramList As clsDiagramme
        HDiagramList = DiagramList

        ' Farbe Null auf Standard zuweisen; wird dann später überschrieben; dient hier nur als definierter Start-Wert
        objektFarbe = 0




        von = showRangeLeft
        bis = showRangeRight
        einheit = " "


        ReDim Xdatenreihe(bis - von)
        ReDim datenreihe(bis - von)
        ReDim edatenreihe(bis - von)
        ReDim kdatenreihe(bis - von)
        ReDim seldatenreihe(bis - von)
        ReDim tmpdatenreihe(bis - von)
        ReDim seriesSumDatenreihe(bis - von)
        ReDim VarValues(bis - von)



        If myCollection.Count = 0 Then
            MsgBox("keine Phase / Rolle / Kostenart / Ergebnisart ausgewählt ...")
            Exit Sub
        End If

        'StartofCalendar = "1.1.2012"
        diff = -1
        startdate = StartofCalendar.AddMonths(diff)

        For m = von To bis
            Xdatenreihe(m - von) = startdate.AddMonths(m).ToString("MMM yy")
        Next m


        titleZeitraum = Xdatenreihe(0) & " - " & Xdatenreihe(bis - von)


        If prcTyp = DiagrammTypen(0) Then


            chtobjName = getKennung("pf", PTpfdk.Phasen, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Phasen)
            Else
                diagramTitle = myCollection.Item(1)
            End If

        ElseIf prcTyp = DiagrammTypen(1) Then

            chtobjName = getKennung("pf", PTpfdk.Rollen, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Rollen)
            Else
                diagramTitle = myCollection.Item(1)
            End If

        ElseIf prcTyp = DiagrammTypen(2) Then
            chtobjName = getKennung("pf", PTpfdk.Kosten, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Kosten)
            Else
                diagramTitle = myCollection.Item(1)
            End If


        ElseIf prcTyp = DiagrammTypen(4) Then
            chtobjName = "Ergebnis-Übersicht"
            diagramTitle = "Ergebnis-Übersicht"
        Else
            chtobjName = "Übersicht"
            diagramTitle = "Übersicht"
        End If


        ' jetzt prüfen, ob es bereits gespeicherte Werte für top, left, ... gibt ;
        ' Wenn ja : übernehmen

        If von > 1 Then
            left = ((von - 1) / 3 - 1) * 3 * boxWidth + 32.8 + von * screen_correct
        Else
            left = 0
        End If

        width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct



        If Not calledfromReporting Then
            Dim foundDiagramm As clsDiagramm

            Try
                foundDiagramm = DiagramList.getDiagramm(chtobjName)
                With foundDiagramm
                    top = .top
                    'left = .left
                    'width = .width
                    'height = .height
                End With
            Catch ex As Exception

            End Try

        End If



        If prcTyp = DiagrammTypen(1) Then
            kdatenreihe = ShowProjekte.getRoleKapasInMonth(myCollection)
        End If

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        With appInstance.Worksheets(arrWsNames(3))

            anzDiagrams = .ChartObjects.Count

            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found

                If .ChartObjects(i).name = chtobjName Then
                    found = True
                    repObj = .ChartObjects(i)
                Else
                    i = i + 1
                End If

            End While

            If Not found Then


                With CType(appInstance.Charts.Add, Excel.Chart)

                    If Not isCockpitChart Then
                        With .Axes(Excel.XlAxisType.xlValue)
                            .MinorUnit = 1
                        End With

                    End If

                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    For r = 1 To myCollection.Count

                        prcName = myCollection.Item(r)


                        If prcTyp = DiagrammTypen(0) Then
                            einheit = " "
                            objektFarbe = PhaseDefinitions.getPhaseDef(prcName).farbe
                            datenreihe = ShowProjekte.getCountPhasesInMonth(prcName)

                            tmpdatenreihe = selectedProjekte.getCountPhasesInMonth(prcName)
                            For ix = 0 To von - bis
                                datenreihe(ix) = datenreihe(ix) - tmpdatenreihe(ix)
                                seldatenreihe(ix) = seldatenreihe(ix) + tmpdatenreihe(ix)
                            Next

                        ElseIf prcTyp = DiagrammTypen(1) Then
                            einheit = " " & awinSettings.kapaEinheit
                            objektFarbe = RoleDefinitions.getRoledef(prcName).farbe
                            datenreihe = ShowProjekte.getRoleValuesInMonth(prcName)


                        ElseIf prcTyp = DiagrammTypen(2) Then
                            einheit = " T€"
                            If prcName = CostDefinitions.getCostdef(CostDefinitions.Count).name Then

                                ' es handelt sich um die Personalkosten, deshalb muss unterschieden werden zwischen internen und externen Kosten
                                isPersCost = True
                                objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                                datenreihe = ShowProjekte.getCostiValuesInMonth
                                edatenreihe = ShowProjekte.getCosteValuesInMonth
                                For i = 0 To bis - von
                                    seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + edatenreihe(i)
                                Next i

                            Else

                                ' es handelt sich nicht um die Personalkosten
                                isPersCost = False
                                objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                                datenreihe = ShowProjekte.getCostValuesInMonth(prcName)
                            End If
                        ElseIf prcTyp = DiagrammTypen(4) Then

                            ' es handelt sich um die Ergebnisse Earned Value bzw. Earned Value - gewichtet 
                            einheit = " T€"

                            objektFarbe = ergebnisfarbe1
                            datenreihe = ShowProjekte.getEarnedValuesInMonth()
                            ' jetzt müssen die - theoretischen Earned Values um die externen Kosten bereinigt werden, die abfallen, weil aufgrund 
                            ' bestimmter überlasteter Rollen externe , teurere Kräfte reingeholt werden müssen 

                            edatenreihe = ShowProjekte.getadditionalECostinMonth
                            For i = 0 To bis - von
                                datenreihe(i) = datenreihe(i) - edatenreihe(i)
                            Next

                            ' jetzt werdem die RiskValues bestimmt 
                            If prcName = ergebnisChartName(1) Then
                                isWeightedValues = True
                                edatenreihe = ShowProjekte.getWeightedRiskValuesInMonth
                                For i = 0 To bis - von
                                    If datenreihe(i) - edatenreihe(i) >= 0 Then
                                        datenreihe(i) = datenreihe(i) - edatenreihe(i)
                                    Else
                                        edatenreihe(i) = (edatenreihe(i) - datenreihe(i)) * -1
                                    End If

                                Next
                            Else
                                isWeightedValues = False
                            End If


                        End If

                        For i = 0 To bis - von
                            seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + datenreihe(i)
                        Next i


                        If isPersCost Then
                            With .SeriesCollection.NewSeries

                                .name = prcName & " intern "
                                .Interior.color = objektFarbe
                                .Values = datenreihe
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False
                            End With
                            With .SeriesCollection.NewSeries

                                .name = "externe Dienstleister "
                                .Interior.color = farbeExterne
                                .Values = edatenreihe
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False
                            End With
                        Else
                            With .SeriesCollection.NewSeries

                                .name = prcName
                                .Interior.color = objektFarbe
                                .Values = datenreihe
                                .XValues = Xdatenreihe
                                If myCollection.Count = 1 Then
                                    If isWeightedValues Then
                                        .ChartType = Excel.XlChartType.xlColumnStacked
                                    Else
                                        .ChartType = Excel.XlChartType.xlColumnClustered
                                    End If
                                Else
                                    .ChartType = Excel.XlChartType.xlColumnStacked
                                End If
                                .HasDataLabels = False
                            End With
                        End If

                    Next r

                    ' wenn es sich um die weighted Variante handelt
                    If isWeightedValues Then
                        With .SeriesCollection.NewSeries
                            .HasDataLabels = False
                            .name = "Risiko Abschlag"
                            .Interior.color = ergebnisfarbe2
                            .Values = edatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                        End With
                    End If

                    ' wenn der Wert größer ist als Null, dann Anzeigen ... 
                    If seldatenreihe.Sum > 0 Then
                        With .SeriesCollection.NewSeries
                            .HasDataLabels = False
                            .name = "Selected Projects"
                            .Interior.color = selectionFarbe
                            .Values = seldatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                        End With

                    End If

                    ' wenn es sich um ein Cockpit Chart handelt, dann wird der jeweilige Min, Max-Wert angezeigt

                    lastSC = .SeriesCollection.Count

                    If isCockpitChart Then
                        ' jetzt muss eine Dummy Series Collection eingeführt werde, damit das Datalabel über dem Balken angezeigt wird
                        If lastSC > 1 Then


                            maxwert = seriesSumDatenreihe.Max

                            For i = 0 To bis - von
                                VarValues(i) = 0.5 * maxwert
                            Next i

                            With .SeriesCollection.NewSeries
                                .name = "Dummy"
                                .Interior.color = RGB(255, 255, 255)
                                .Values = VarValues
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False
                            End With
                            lastSC = .SeriesCollection.Count

                        End If
                        With .SeriesCollection(lastSC)
                            .HasDataLabels = False
                            VarValues = seriesSumDatenreihe
                            nr_pts = .Points.Count

                            minwert = VarValues.Min
                            maxwert = VarValues.Max
                            mindone = False
                            maxdone = False
                            i = 1
                            While i <= nr_pts And (mindone = False Or maxdone = False)

                                If VarValues(i - 1) = minwert And Not mindone Then
                                    mindone = True
                                    With .Points(i)
                                        .HasDataLabel = True
                                        .DataLabel.text = Format(minwert, "##,##0")
                                        .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                                        Try
                                            .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                        Catch ex As Exception

                                        End Try


                                    End With
                                ElseIf VarValues(i - 1) = maxwert And Not maxdone Then
                                    maxdone = True
                                    With .Points(i)
                                        .HasDataLabel = True
                                        .DataLabel.text = Format(maxwert, "##,##0")
                                        .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                                        Try
                                            .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                        Catch ex As Exception

                                        End Try


                                    End With

                                End If
                                i = i + 1
                            End While
                        End With

                        ' es ist ein Cockpit-Diagramm, deswegen müssen folgende Einstellungen gelten:

                        .HasLegend = False
                        .HasAxis(Excel.XlAxisType.xlCategory) = False
                        .HasAxis(Excel.XlAxisType.xlValue) = False
                        .Axes(Excel.XlAxisType.xlCategory).HasMajorGridlines = False
                        With .Axes(Excel.XlAxisType.xlValue)
                            .HasMajorGridlines = False
                        End With

                    ElseIf myCollection.Count > 1 Then

                    End If


                    If prcTyp = DiagrammTypen(1) Then
                        With .SeriesCollection.NewSeries
                            .HasDataLabels = False
                            .name = "Gesamt-Kapazität"
                            .Border.color = rollenKapaFarbe
                            .Values = kdatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlLine
                            nr_pts = .Points.Count

                            With .Points(nr_pts)

                                .HasDataLabel = False

                            End With

                        End With
                    End If
                    .HasTitle = True
                    If prcTyp = DiagrammTypen(0) Or awinSettings.kapaEinheit = "ST" Then
                        titleSumme = ""
                    Else
                        titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & einheit & ")"
                    End If
                    .ChartTitle.Text = diagramTitle & titleSumme

                    If isCockpitChart Then

                        .ChartTitle.Font.Size = awinSettings.CPfontsizeTitle
                        .HasLegend = False

                    ElseIf lastSC > 1 Then

                        .HasLegend = True

                        .Legend.Position = Excel.Constants.xlTop

                        .Legend.Font.Size = awinSettings.fontsizeLegend
                    Else
                        .HasLegend = False
                    End If


                    .Name = prcTyp
                    .Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With


                With .ChartObjects(anzDiagrams + 1)

                    .Top = top
                    .Left = left
                    .Width = width
                    .Height = height
                    .Name = chtobjName
                    Dim tststr As String = .Name

                    '
                    ' diese Korrektur ist notwendig, um auszugleichen wenn die Achsenbeschriftungen größer als zweistellig werden
                    '
                    .Chart.Axes(Excel.XlAxisType.xlValue).minimumScale = 0

                    Dim korrAbstand As Double
                    korrAbstand = .Chart.Axes(Excel.XlAxisType.xlCategory).left - 22
                    If korrAbstand > 1 Then
                        .Left = left - korrAbstand
                        .Width = width + korrAbstand
                    End If
                End With


                ' wenn es ein Cockpit Chart ist: dann werden die Borderlines ausgeschaltet ...
                If isCockpitChart Then
                    Try
                        With appInstance.ActiveSheet
                            .Shapes(chtobjName).line.visible = False
                        End With
                    Catch ex As Exception

                    End Try
                Else
                    'Call awinScrollintoView()
                End If


                repObj = .ChartObjects(anzDiagrams + 1)

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then
                    prcDiagram = New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    prcChart = New clsEventsPrcCharts
                    prcChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart
                    prcDiagram.setDiagramEvent = prcChart
                    ' Ende Event Handling für Chart 


                    With prcDiagram
                        .DiagrammTitel = diagramTitle
                        .diagrammTyp = prcTyp
                        .gsCollection = myCollection
                        .isCockpitChart = isCockpitChart
                        .top = top
                        .left = left
                        .kennung = chtobjName
                        '.width = width
                        '.height = height

                    End With

                    ' eintragen in die sortierte Liste mit .kennung als dem Schlüssel 
                    ' wenn das Diagramm bereits existiert, muss es gelöscht werden, dann neu ergänzt ... 
                    Try
                        DiagramList.Add(prcDiagram)
                    Catch ex As Exception

                        Try
                            DiagramList.Remove(prcDiagram.kennung)
                            DiagramList.Add(prcDiagram)
                        Catch ex1 As Exception

                        End Try


                    End Try

                End If



            End If
        End With

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU





    End Sub


    ''' <summary>
    ''' erzeugt ein Phasen-/Rollen-/Kostenart - Diagramm
    ''' bekommt Parameter für die darzustellenden Rollen mit, die Position, ob es ein Cockpit Chart ist und um welchen Diagramm-Typ es sich handelt
    ''' Diagramm-Typen:
    ''' 0 - Phase
    ''' 1 - Rolle
    ''' 2 - Kostenart
    ''' 3 - Portfolio
    ''' 4 - Summe  
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <param name="prcTyp"></param>
    ''' <remarks></remarks>
    Sub awinCreateprcCollectionDiagram(ByRef myCollection As Collection, ByRef repObj As Object, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                       ByVal isCockpitChart As Boolean, ByVal prcTyp As String, ByVal calledfromReporting As Boolean)

        Dim von As Integer, bis As Integer

        Dim anzDiagrams As Integer, i As Integer, m As Integer, r As Integer
        Dim found As Boolean
        'Dim korr_abstand As Double
        Dim minwert As Double, maxwert As Double
        Dim nr_pts As Integer
        Dim diagramTitle As String
        Dim objektFarbe As Object
        Dim ampelfarbe(3) As Long
        Dim Xdatenreihe() As String
        Dim datenreihe() As Double, edatenreihe() As Double, seriesSumDatenreihe() As Double
        Dim kdatenreihe() As Double ' nimmt die Kapa-Werte für das Diagramm auf
        Dim msdatenreihe(,) As Double
        Dim prcName As String
        Dim startdate As Date
        Dim diff As Object
        Dim mindone As Boolean, maxdone As Boolean
        Dim VarValues() As Double
        Dim prcDiagram As clsDiagramm
        'Dim prcChart As clsAwinEvent
        Dim prcChart As clsEventsPrcCharts
        Dim isPersCost As Boolean
        Dim isWeightedValues As Boolean = False
        Dim lastSC As Integer
        Dim titleZeitraum As String, titleSumme As String, einheit As String
        'Dim chtTitle As String
        Dim chtobjName As String

        ' Debugging variable 
        Dim HDiagramList As clsDiagramme
        HDiagramList = DiagramList

        ' Farbe Null auf Standard zuweisen; wird dann später überschrieben; dient hier nur als definierter Start-Wert
        objektFarbe = 0

        With awinSettings
            ampelfarbe(0) = .AmpelNichtBewertet
            ampelfarbe(1) = .AmpelGruen
            ampelfarbe(2) = .AmpelGelb
            ampelfarbe(3) = .AmpelRot
        End With

        von = showRangeLeft
        bis = showRangeRight
        einheit = " "


        ReDim Xdatenreihe(bis - von)
        ReDim datenreihe(bis - von)
        ReDim edatenreihe(bis - von)
        ReDim kdatenreihe(bis - von)
        ReDim seriesSumDatenreihe(bis - von)
        ReDim VarValues(bis - von)
        ReDim msdatenreihe(3, bis - von)



        If myCollection.Count = 0 Then
            MsgBox("keine Phase / Rolle / Kostenart / Ergebnisart / Meilenstein ausgewählt ...")
            Exit Sub
        End If

        diff = -1
        startdate = StartofCalendar.AddMonths(diff)

        For m = von To bis
            Xdatenreihe(m - von) = startdate.AddMonths(m).ToString("MMM yy")
        Next m


        titleZeitraum = Xdatenreihe(0) & " - " & Xdatenreihe(bis - von)


        If prcTyp = DiagrammTypen(0) Then


            chtobjName = getKennung("pf", PTpfdk.Phasen, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Phasen)
            Else
                diagramTitle = myCollection.Item(1)
            End If

        ElseIf prcTyp = DiagrammTypen(1) Then

            chtobjName = getKennung("pf", PTpfdk.Rollen, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Rollen)
            Else
                diagramTitle = myCollection.Item(1)
            End If

        ElseIf prcTyp = DiagrammTypen(2) Then
            chtobjName = getKennung("pf", PTpfdk.Kosten, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Kosten)
            Else
                diagramTitle = myCollection.Item(1)
            End If


        ElseIf prcTyp = DiagrammTypen(4) Then
            chtobjName = "Ergebnis-Übersicht"
            diagramTitle = "Ergebnis-Übersicht"

        ElseIf prcTyp = DiagrammTypen(5) Then
            chtobjName = getKennung("pf", PTpfdk.Meilenstein, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Meilenstein)
            Else
                diagramTitle = myCollection.Item(1)
            End If

        Else
            chtobjName = "Übersicht"
            diagramTitle = "Übersicht"
        End If


        ' jetzt prüfen, ob es bereits gespeicherte Werte für top, left, ... gibt ;
        ' Wenn ja : übernehmen

        If von > 1 Then
            left = ((von - 1) / 3 - 1) * 3 * boxWidth + 32.8 + von * screen_correct
        Else
            left = 0
        End If

        width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct



        If Not calledfromReporting Then
            Dim foundDiagramm As clsDiagramm

            Try
                foundDiagramm = DiagramList.getDiagramm(chtobjName)
                With foundDiagramm
                    top = .top
                    'left = .left
                    'width = .width
                    'height = .height
                End With
            Catch ex As Exception

            End Try

        End If



        If prcTyp = DiagrammTypen(1) Then
            kdatenreihe = ShowProjekte.getRoleKapasInMonth(myCollection)
        End If

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        With appInstance.Worksheets(arrWsNames(3))

            anzDiagrams = .ChartObjects.Count

            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found

                If .ChartObjects(i).name = chtobjName Then
                    found = True
                    repObj = .ChartObjects(i)
                Else
                    i = i + 1
                End If

            End While

            If Not found Then


                With CType(appInstance.Charts.Add, Excel.Chart)

                    If Not isCockpitChart Then
                        With .Axes(Excel.XlAxisType.xlValue)
                            .MinorUnit = 1
                        End With

                    End If

                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    For r = 1 To myCollection.Count

                        prcName = myCollection.Item(r)


                        If prcTyp = DiagrammTypen(0) Then
                            einheit = " "
                            objektFarbe = PhaseDefinitions.getPhaseDef(prcName).farbe
                            datenreihe = ShowProjekte.getCountPhasesInMonth(prcName)

                        ElseIf prcTyp = DiagrammTypen(1) Then
                            einheit = " " & awinSettings.kapaEinheit
                            objektFarbe = RoleDefinitions.getRoledef(prcName).farbe
                            datenreihe = ShowProjekte.getRoleValuesInMonth(prcName)


                        ElseIf prcTyp = DiagrammTypen(2) Then
                            einheit = " T€"
                            If prcName = CostDefinitions.getCostdef(CostDefinitions.Count).name Then

                                ' es handelt sich um die Personalkosten, deshalb muss unterschieden werden zwischen internen und externen Kosten
                                isPersCost = True
                                objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                                datenreihe = ShowProjekte.getCostiValuesInMonth
                                edatenreihe = ShowProjekte.getCosteValuesInMonth
                                For i = 0 To bis - von
                                    seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + edatenreihe(i)
                                Next i

                            Else

                                ' es handelt sich nicht um die Personalkosten
                                isPersCost = False
                                objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                                datenreihe = ShowProjekte.getCostValuesInMonth(prcName)
                            End If
                        ElseIf prcTyp = DiagrammTypen(4) Then

                            ' es handelt sich um die Ergebnisse Earned Value bzw. Earned Value - gewichtet 
                            einheit = " T€"

                            objektFarbe = ergebnisfarbe1
                            datenreihe = ShowProjekte.getEarnedValuesInMonth()
                            ' jetzt müssen die - theoretischen Earned Values um die externen Kosten bereinigt werden, die abfallen, weil aufgrund 
                            ' bestimmter überlasteter Rollen externe , teurere Kräfte reingeholt werden müssen 

                            edatenreihe = ShowProjekte.getadditionalECostinMonth
                            For i = 0 To bis - von
                                datenreihe(i) = datenreihe(i) - edatenreihe(i)
                            Next

                            ' jetzt werdem die RiskValues bestimmt 
                            If prcName = ergebnisChartName(1) Then
                                isWeightedValues = True
                                edatenreihe = ShowProjekte.getWeightedRiskValuesInMonth
                                For i = 0 To bis - von
                                    If datenreihe(i) - edatenreihe(i) >= 0 Then
                                        datenreihe(i) = datenreihe(i) - edatenreihe(i)
                                    Else
                                        edatenreihe(i) = (edatenreihe(i) - datenreihe(i)) * -1
                                    End If

                                Next
                            Else
                                isWeightedValues = False
                            End If

                        ElseIf prcTyp = DiagrammTypen(5) Then

                            einheit = " "
                            msdatenreihe = ShowProjekte.getCountMilestonesInMonth(prcName)

                        End If

                        For i = 0 To bis - von
                            seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + datenreihe(i)
                        Next i


                        If isPersCost Then
                            With .SeriesCollection.NewSeries

                                .name = prcName & " intern "
                                .Interior.color = objektFarbe
                                .Values = datenreihe
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False
                            End With
                            With .SeriesCollection.NewSeries

                                .name = "externe Dienstleister "
                                .Interior.color = farbeExterne
                                .Values = edatenreihe
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False
                            End With
                        Else

                            If prcTyp = DiagrammTypen(5) Then

                                For c = 0 To 3

                                    For i = 0 To bis - von
                                        datenreihe(i) = msdatenreihe(c, i)
                                    Next

                                    With .SeriesCollection.NewSeries
                                        If c = 0 Then
                                            .name = prcName & ", ohne Ampel"
                                        ElseIf c = 1 Then
                                            .name = prcName & ", grüne Ampel"
                                        ElseIf c = 2 Then
                                            .name = prcName & ", gelbe Ampel"
                                        Else
                                            .name = prcName & ", rote Ampel"
                                        End If
                                        .Interior.color = ampelfarbe(c)
                                        .Values = datenreihe
                                        .XValues = Xdatenreihe
                                        .ChartType = Excel.XlChartType.xlColumnStacked
                                        .HasDataLabels = False
                                    End With


                                Next

                            Else

                                With .SeriesCollection.NewSeries
                                    .name = prcName
                                    .Interior.color = objektFarbe
                                    .Values = datenreihe
                                    .XValues = Xdatenreihe
                                    If myCollection.Count = 1 Then
                                        If isWeightedValues Then
                                            .ChartType = Excel.XlChartType.xlColumnStacked
                                        Else
                                            .ChartType = Excel.XlChartType.xlColumnClustered
                                        End If
                                    Else
                                        .ChartType = Excel.XlChartType.xlColumnStacked
                                    End If
                                    .HasDataLabels = False
                                End With

                            End If


                        End If

                    Next r

                    ' wenn es sich um die weighted Variante handelt
                    If isWeightedValues Then
                        With .SeriesCollection.NewSeries
                            .HasDataLabels = False
                            .name = "Risiko Abschlag"
                            .Interior.color = ergebnisfarbe2
                            .Values = edatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                        End With
                    End If

                    ' wenn es sich um ein Cockpit Chart handelt, dann wird der jeweilige Min, Max-Wert angezeigt

                    lastSC = .SeriesCollection.Count

                    If isCockpitChart Then
                        ' jetzt muss eine Dummy Series Collection eingeführt werde, damit das Datalabel über dem Balken angezeigt wird
                        If lastSC > 1 Then


                            maxwert = seriesSumDatenreihe.Max

                            For i = 0 To bis - von
                                VarValues(i) = 0.5 * maxwert
                            Next i

                            With .SeriesCollection.NewSeries
                                .name = "Dummy"
                                .Interior.color = RGB(255, 255, 255)
                                .Values = VarValues
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False
                            End With
                            lastSC = .SeriesCollection.Count

                        End If
                        With .SeriesCollection(lastSC)
                            .HasDataLabels = False
                            VarValues = seriesSumDatenreihe
                            nr_pts = .Points.Count

                            minwert = VarValues.Min
                            maxwert = VarValues.Max
                            mindone = False
                            maxdone = False
                            i = 1
                            While i <= nr_pts And (mindone = False Or maxdone = False)

                                If VarValues(i - 1) = minwert And Not mindone Then
                                    mindone = True
                                    With .Points(i)
                                        .HasDataLabel = True
                                        .DataLabel.text = Format(minwert, "##,##0")
                                        .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                                        Try
                                            .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                        Catch ex As Exception

                                        End Try


                                    End With
                                ElseIf VarValues(i - 1) = maxwert And Not maxdone Then
                                    maxdone = True
                                    With .Points(i)
                                        .HasDataLabel = True
                                        .DataLabel.text = Format(maxwert, "##,##0")
                                        .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                                        Try
                                            .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                        Catch ex As Exception

                                        End Try


                                    End With

                                End If
                                i = i + 1
                            End While
                        End With

                        ' es ist ein Cockpit-Diagramm, deswegen müssen folgende Einstellungen gelten:

                        .HasLegend = False
                        .HasAxis(Excel.XlAxisType.xlCategory) = False
                        .HasAxis(Excel.XlAxisType.xlValue) = False
                        .Axes(Excel.XlAxisType.xlCategory).HasMajorGridlines = False
                        With .Axes(Excel.XlAxisType.xlValue)
                            .HasMajorGridlines = False
                        End With

                    ElseIf myCollection.Count > 1 Then

                    End If


                    If prcTyp = DiagrammTypen(1) Then
                        With .SeriesCollection.NewSeries
                            .HasDataLabels = False
                            .name = "Gesamt-Kapazität"
                            .Border.color = rollenKapaFarbe
                            .Values = kdatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlLine
                            nr_pts = .Points.Count

                            With .Points(nr_pts)

                                .HasDataLabel = False

                            End With

                        End With
                    End If
                    .HasTitle = True

                    If prcTyp = DiagrammTypen(0) Or _
                        prcTyp = DiagrammTypen(5) Or _
                        awinSettings.kapaEinheit = "ST" Then
                        titleSumme = ""
                    Else
                        titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & einheit & ")"
                    End If

                    .ChartTitle.Text = diagramTitle & titleSumme

                    If isCockpitChart Then

                        .ChartTitle.Font.Size = awinSettings.CPfontsizeTitle
                        .HasLegend = False

                    ElseIf lastSC > 1 Then

                        .HasLegend = True

                        .Legend.Position = Excel.Constants.xlTop

                        .Legend.Font.Size = awinSettings.fontsizeLegend
                    Else
                        .HasLegend = False
                    End If


                    .Name = prcTyp
                    .Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With


                With .ChartObjects(anzDiagrams + 1)

                    .Top = top
                    .Left = left
                    .Width = width
                    .Height = height
                    .Name = chtobjName
                    Dim tststr As String = .Name

                    '
                    ' diese Korrektur ist notwendig, um auszugleichen wenn die Achsenbeschriftungen größer als zweistellig werden
                    '
                    .Chart.Axes(Excel.XlAxisType.xlValue).minimumScale = 0

                    Dim korrAbstand As Double
                    korrAbstand = .Chart.Axes(Excel.XlAxisType.xlCategory).left - 22
                    If korrAbstand > 1 Then
                        .Left = left - korrAbstand
                        .Width = width + korrAbstand
                    End If
                End With


                ' wenn es ein Cockpit Chart ist: dann werden die Borderlines ausgeschaltet ...
                If isCockpitChart Then
                    Try
                        With appInstance.ActiveSheet
                            .Shapes(chtobjName).line.visible = False
                        End With
                    Catch ex As Exception

                    End Try
                Else
                    'Call awinScrollintoView()
                End If


                repObj = .ChartObjects(anzDiagrams + 1)

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then
                    prcDiagram = New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    prcChart = New clsEventsPrcCharts
                    prcChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart
                    prcDiagram.setDiagramEvent = prcChart
                    ' Ende Event Handling für Chart 


                    With prcDiagram
                        .DiagrammTitel = diagramTitle
                        .diagrammTyp = prcTyp
                        .gsCollection = myCollection
                        .isCockpitChart = isCockpitChart
                        .top = top
                        .left = left
                        .kennung = chtobjName
                        '.width = width
                        '.height = height

                    End With

                    ' eintragen in die sortierte Liste mit .kennung als dem Schlüssel 
                    ' wenn das Diagramm bereits existiert, muss es gelöscht werden, dann neu ergänzt ... 
                    Try
                        DiagramList.Add(prcDiagram)
                    Catch ex As Exception

                        Try
                            DiagramList.Remove(prcDiagram.kennung)
                            DiagramList.Add(prcDiagram)
                        Catch ex1 As Exception

                        End Try


                    End Try

                End If



            End If
        End With

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU





    End Sub

    '
    ' aktualisiert ein Rollen-Diagramm
    '
    Sub awinUpdateprcCollectionDiagram(ByRef chtobj As ChartObject)

        Dim von As Integer, bis As Integer
        Dim i As Integer, m As Integer, d As Integer, r As Integer
        Dim found As Boolean

        Dim minwert As Double, maxwert As Double
        Dim nr_pts As Integer
        Dim diagramTitle As String

        Dim objektFarbe As Object
        Dim ampelfarbe(3) As Long
        Dim Xdatenreihe() As String
        Dim datenreihe() As Double, edatenreihe() As Double, seriesSumDatenreihe() As Double
        Dim msdatenreihe(,) As Double
        ' nimmt die Daten der selektierten Werte auf 
        Dim seldatenreihe() As Double, tmpdatenreihe() As Double
        Dim kdatenreihe() As Double
        Dim prcName As String
        Dim startdate As Date
        Dim diff As Object
        Dim mindone As Boolean, maxdone As Boolean
        Dim width As Double
        'Dim left As Double
        Dim myCollection As Collection
        Dim isCockpitChart As Boolean
        Dim isWeightedValues As Boolean = False
        Dim VarValues() As Double
        Dim prcTyp As String
        Dim isPersCost As Boolean
        Dim lastSC As Integer
        Dim titleSumme As String, einheit As String
        Dim selectionFarbe As Long = awinSettings.AmpelNichtBewertet

        'Dim chtTitle As String
        Dim chtobjName As String
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating


        ' Debugging variable 
        'Dim HDiagramList As clsDiagramme
        'HDiagramList = DiagramList

        chtobjName = chtobj.Name


        von = showRangeLeft
        bis = showRangeRight
        width = chtobj.Width

        Dim currentScale As Double
        Try
            With chtobj.Chart.Axes(Excel.XlAxisType.xlValue)
                currentScale = .MaximumScale
            End With
        Catch ex As Exception

        End Try


        ' Default Zuweisung ; wird später überschrieben ; verhindert , daß sie verwendet wird, ohne einen Wert zu haben 
        objektFarbe = 0

        With awinSettings
            ampelfarbe(0) = .AmpelNichtBewertet
            ampelfarbe(1) = .AmpelGruen
            ampelfarbe(2) = .AmpelGelb
            ampelfarbe(3) = .AmpelRot
        End With



        If istCockpitDiagramm(chtobj) Then
            ' dann ist es ein Cockpit Chart ....
            isCockpitChart = True
        Else
            isCockpitChart = False

            width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct

        End If

        ReDim Xdatenreihe(bis - von)
        ReDim datenreihe(bis - von)
        ReDim edatenreihe(bis - von)
        ReDim kdatenreihe(bis - von)
        ReDim seldatenreihe(bis - von)
        ReDim tmpdatenreihe(bis - von)
        ReDim seriesSumDatenreihe(bis - von)
        ReDim VarValues(bis - von)
        ReDim msdatenreihe(3, bis - von)


        found = False
        myCollection = New Collection
        einheit = " "
        prcTyp = " "
        d = 1
        Dim foundDiagram As clsDiagramm

        Try
            foundDiagram = DiagramList.getDiagramm(chtobjName)
            myCollection = foundDiagram.gsCollection
            prcTyp = foundDiagram.diagrammTyp
            found = True
        Catch ex As Exception
            Exit Sub
        End Try



        If Not found Then
            Exit Sub
        End If


        If myCollection.Count = 0 Then
            MsgBox("keine Phase-/Rolle-/Kostenart / Ergebnisart ausgewählt ...")
            Exit Sub
        End If

        diff = -1
        startdate = StartofCalendar.AddMonths(diff)


        For m = von To bis
            Xdatenreihe(m - von) = startdate.AddMonths(m).ToString("MMM yy")
        Next m

        If myCollection.Count > 1 Then
            If prcTyp = DiagrammTypen(0) Then
                diagramTitle = "Phasen-Übersicht"
            ElseIf prcTyp = DiagrammTypen(1) Then
                diagramTitle = "Rollen-Übersicht"
            ElseIf prcTyp = DiagrammTypen(2) Then
                diagramTitle = "Kosten-Übersicht"
            ElseIf prcTyp = DiagrammTypen(4) Then
                diagramTitle = "Ergebnis-Übersicht"
            ElseIf prcTyp = DiagrammTypen(5) Then
                chtobjName = getKennung("pf", PTpfdk.Meilenstein, myCollection)

                If myCollection.Count > 1 Then
                    diagramTitle = portfolioDiagrammtitel(PTpfdk.Meilenstein)
                Else
                    diagramTitle = myCollection.Item(1)
                End If
            Else
                diagramTitle = "Übersicht"
            End If
        Else
            diagramTitle = myCollection.Item(1)
        End If

        If prcTyp = DiagrammTypen(1) Then
            kdatenreihe = ShowProjekte.getRoleKapasInMonth(myCollection)
        End If


        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False


        With appInstance.Worksheets(arrWsNames(3))


            With chtobj.Chart

                ' remove old series
                Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete()
                Loop

                For r = 1 To myCollection.Count

                    prcName = myCollection.Item(r)

                    If prcTyp = DiagrammTypen(0) Then
                        einheit = " "
                        objektFarbe = PhaseDefinitions.getPhaseDef(prcName).farbe
                        datenreihe = ShowProjekte.getCountPhasesInMonth(prcName)
                        ' Ergänzung wegen Anzeige der selektierten Objekte ... 
                        tmpdatenreihe = selectedProjekte.getCountPhasesInMonth(prcName)
                        For ix = 0 To bis - von
                            datenreihe(ix) = datenreihe(ix) - tmpdatenreihe(ix)
                            seldatenreihe(ix) = seldatenreihe(ix) + tmpdatenreihe(ix)
                        Next

                    ElseIf prcTyp = DiagrammTypen(1) Then
                        einheit = " " & awinSettings.kapaEinheit
                        objektFarbe = RoleDefinitions.getRoledef(prcName).farbe
                        datenreihe = ShowProjekte.getRoleValuesInMonth(prcName)


                    ElseIf prcTyp = DiagrammTypen(2) Then
                        einheit = " T€"
                        If prcName = CostDefinitions.getCostdef(CostDefinitions.Count).name Then
                            ' es handelt sich um die Personalkosten, deshalb muss unterschieden werden zwischen internen und externen Kosten
                            isPersCost = True
                            objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                            datenreihe = ShowProjekte.getCostiValuesInMonth
                            edatenreihe = ShowProjekte.getCosteValuesInMonth
                            For i = 0 To bis - von
                                seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + edatenreihe(i)
                            Next i

                        Else
                            ' es handelt sich nicht um die Personalkosten
                            isPersCost = False
                            objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                            datenreihe = ShowProjekte.getCostValuesInMonth(prcName)
                        End If

                    ElseIf prcTyp = DiagrammTypen(4) Then
                        ' es handelt sich um die Ergebnisse Earned Value bzw. Earned Value - gewichtet 
                        einheit = " T€"

                        objektFarbe = ergebnisfarbe1
                        datenreihe = ShowProjekte.getEarnedValuesInMonth()
                        ' jetzt müssen die - theoretischen Earned Values um die externen Kosten bereinigt werden, die abfallen, weil aufgrund 
                        ' bestimmter überlasteter Rollen externe , teurere Kräfte reingeholt werden müssen 

                        edatenreihe = ShowProjekte.getadditionalECostinMonth
                        For i = 0 To bis - von
                            datenreihe(i) = datenreihe(i) - edatenreihe(i)
                        Next

                        ' jetzt werdem die RiskValues bestimmt 
                        If prcName = ergebnisChartName(1) Then
                            isWeightedValues = True
                            edatenreihe = ShowProjekte.getWeightedRiskValuesInMonth
                            For i = 0 To bis - von
                                If datenreihe(i) - edatenreihe(i) >= 0 Then
                                    datenreihe(i) = datenreihe(i) - edatenreihe(i)
                                Else
                                    edatenreihe(i) = (edatenreihe(i) - datenreihe(i)) * -1
                                End If

                            Next
                        Else
                            isWeightedValues = False
                        End If

                    ElseIf prcTyp = DiagrammTypen(5) Then

                        einheit = " "
                        msdatenreihe = ShowProjekte.getCountMilestonesInMonth(prcName)

                    End If

                    For i = 0 To bis - von
                        seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + datenreihe(i)
                    Next i


                    If isPersCost Then
                        With .SeriesCollection.NewSeries

                            .name = prcName & " intern "
                            .Interior.color = objektFarbe
                            .Values = datenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                            .HasDataLabels = False
                        End With
                        With .SeriesCollection.NewSeries

                            .name = "externe Dienstleister "
                            .Interior.color = farbeExterne
                            .Values = edatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                            .HasDataLabels = False
                        End With
                    Else
                        If prcTyp = DiagrammTypen(5) Then

                            For c = 0 To 3

                                For i = 0 To bis - von
                                    datenreihe(i) = msdatenreihe(c, i)
                                Next

                                With .SeriesCollection.NewSeries
                                    If c = 0 Then
                                        .name = prcName & ", ohne Ampel"
                                    ElseIf c = 1 Then
                                        .name = prcName & ", grüne Ampel"
                                    ElseIf c = 2 Then
                                        .name = prcName & ", gelbe Ampel"
                                    Else
                                        .name = prcName & ", rote Ampel"
                                    End If
                                    .Interior.color = ampelfarbe(c)
                                    .Values = datenreihe
                                    .XValues = Xdatenreihe
                                    .ChartType = Excel.XlChartType.xlColumnStacked
                                    .HasDataLabels = False
                                End With


                            Next

                        Else

                            With .SeriesCollection.NewSeries
                                .name = prcName
                                .Interior.color = objektFarbe
                                .Values = datenreihe
                                .XValues = Xdatenreihe
                                If myCollection.Count = 1 Then
                                    If isWeightedValues Then
                                        .ChartType = Excel.XlChartType.xlColumnStacked
                                    Else
                                        .ChartType = Excel.XlChartType.xlColumnClustered
                                    End If
                                Else
                                    .ChartType = Excel.XlChartType.xlColumnStacked
                                End If
                                .HasDataLabels = False
                            End With

                        End If

                    End If

                Next r

                ' wenn es sich um die weighted Variante handelt
                If isWeightedValues Then
                    With .SeriesCollection.NewSeries
                        .HasDataLabels = False
                        .name = "Risiko Abschlag"
                        .Interior.color = ergebnisfarbe2
                        .Values = edatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With
                End If


                ' Ergänzung wegen Anzeige selektierter Objekte 
                ' wenn der Wert größer ist als Null, dann Anzeigen ... 
                If seldatenreihe.Sum > 0 Then
                    With .SeriesCollection.NewSeries
                        .HasDataLabels = False
                        .name = "Selected Projects"
                        .Interior.color = selectionFarbe
                        .Values = seldatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                    End With

                End If


                ' wenn es sich um ein Cockpit Chart handelt, dann wird der jeweilige Min, Max-Wert angezeigt

                lastSC = .SeriesCollection.Count

                If isCockpitChart Then
                    ' jetzt muss eine Dummy Series Collection eingeführt werde, damit das Datalabel über dem Balken angezeigt wird
                    If lastSC > 1 Then

                        maxwert = appInstance.WorksheetFunction.Max(seriesSumDatenreihe)

                        For i = 0 To bis - von
                            VarValues(i) = 0.5 * maxwert
                        Next i

                        With .SeriesCollection.NewSeries
                            .name = "Dummy"
                            .Interior.color = RGB(255, 255, 255)
                            .Values = VarValues
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                            .HasDataLabels = False
                        End With
                        lastSC = .SeriesCollection.Count

                    End If
                    With .SeriesCollection(lastSC)
                        .HasDataLabels = False
                        VarValues = seriesSumDatenreihe
                        nr_pts = .Points.Count
                        minwert = VarValues.Min
                        maxwert = VarValues.Max

                        mindone = False
                        maxdone = False
                        i = 1
                        While i <= nr_pts And (mindone = False Or maxdone = False)

                            If VarValues(i - 1) = minwert And Not mindone Then
                                mindone = True
                                With .Points(i)
                                    .HasDataLabel = True
                                    .DataLabel.text = Format(minwert, "##,##0")

                                    .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                                    Try

                                        .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionBestFit
                                    Catch ex As Exception
                                    End Try


                                End With
                            ElseIf VarValues(i - 1) = maxwert And Not maxdone Then
                                maxdone = True
                                With .Points(i)
                                    .HasDataLabel = True
                                    .DataLabel.text = Format(maxwert, "##,##0")

                                    .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                                    Try

                                        .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionBestFit
                                    Catch ex As Exception
                                    End Try

                                End With

                            End If
                            i = i + 1
                        End While
                    End With

                    ' es ist ein Mini-Diagramm, deswegen müssen folgende Einstellungen gelten:

                    .HasLegend = False
                    .HasAxis(Excel.XlAxisType.xlCategory) = False
                    .HasAxis(Excel.XlAxisType.xlValue) = False
                    .Axes(Excel.XlAxisType.xlCategory).HasMajorGridlines = False
                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasMajorGridlines = False
                    End With

                ElseIf myCollection.Count > 1 Then

                End If

                If prcTyp = DiagrammTypen(1) Then
                    With .SeriesCollection.NewSeries
                        .HasDataLabels = False
                        .name = "Gesamt-Kapazität"
                        .Border.color = rollenKapaFarbe
                        .Values = kdatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlLine
                        nr_pts = .Points.Count

                        With .Points(nr_pts)

                            .HasDataLabel = False

                        End With

                    End With
                End If

                .HasTitle = True

                If prcTyp = DiagrammTypen(0) Or _
                        prcTyp = DiagrammTypen(5) Or _
                        awinSettings.kapaEinheit = "ST" Then
                    titleSumme = ""
                Else
                    titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & einheit & ")"
                End If

                .ChartTitle.Text = diagramTitle & titleSumme

                If isCockpitChart Then
                    .HasLegend = False
                ElseIf lastSC > 1 And seldatenreihe.Sum = 0 Then
                    .HasLegend = True
                    .Legend.Position = Excel.Constants.xlTop
                    .Legend.Font.Size = awinSettings.fontsizeLegend
                Else

                    .HasLegend = False
                End If

            End With


        End With

        With chtobj
            If Not isCockpitChart Then
                .Width = width
            End If
        End With

        ' Skalierung nur ändern, wenn erforderlich, weil der maxwert höher ist als die bisherige Skalierung ... 
        Dim hmxWert As Double = seriesSumDatenreihe.Max
        If hmxWert > currentScale Then
            With chtobj.Chart.Axes(Excel.XlAxisType.xlValue)
                .MaximumScale = hmxWert + 1
            End With
        Else
            With chtobj.Chart.Axes(Excel.XlAxisType.xlValue)
                .MaximumScale = currentScale
            End With
        End If


        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU




    End Sub


    '
    '
    '
    Sub awinCreateAuslastungsDiagramm(ByRef repObj As Excel.ChartObject, _
                                      ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                      ByVal calledfromReporting As Boolean)

        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim Xdatenreihe() As String
        Dim datenreihe() As Double
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim chtTitle As String

        Dim von As Integer, bis As Integer
        Dim diagramTitle As String
        Dim htxt As String
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim kennung As String

        Dim chtobjName As String
        Dim myCollection As New Collection
        myCollection.Add("Auslastung")
        chtobjName = getKennung("pf", PTpfdk.Auslastung, myCollection)
        myCollection.Clear()

        If Not calledfromReporting Then

            Dim foundDiagramm As clsDiagramm

            ' wenn die Werte für dieses Diagramm bereits einmal gespeichert wurden ... -> übernehmen 
            Try
                foundDiagramm = DiagramList.getDiagramm(chtobjName)
                With foundDiagramm
                    top = .top
                    left = .left
                    width = .width
                    height = .height
                End With
            Catch ex As Exception


            End Try
        End If



        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False


        titelTeile(0) = summentitel9 & " (" & awinSettings.kapaEinheit & ")"


        titelTeilLaengen(0) = titelTeile(0).Length + 1
        titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)

        von = showRangeLeft
        bis = showRangeRight



        ReDim Xdatenreihe(2)
        ReDim datenreihe(2)

        Xdatenreihe(0) = "Auslastung"
        Xdatenreihe(1) = "Über-Auslastung"
        Xdatenreihe(2) = "Unter-Auslastung"


        datenreihe(0) = ShowProjekte.getAuslastungsValues(0).Sum
        datenreihe(1) = ShowProjekte.getAuslastungsValues(1).Sum
        datenreihe(2) = ShowProjekte.getAuslastungsValues(2).Sum


        With appInstance.Worksheets(arrWsNames(3))

            anzDiagrams = .ChartObjects.Count

            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found

                Try
                    'chtTitle = .ChartObjects(i).Chart.ChartTitle.text
                    chtTitle = .ChartObjects(i).Name
                Catch ex As Exception
                    chtTitle = " "
                End Try


                If chtTitle = chtobjName Then
                    found = True
                    repObj = .chartobjects(i)
                Else
                    i = i + 1
                End If

            End While

            If Not found Then

                With appInstance.Charts.Add
                    .HasTitle = True
                    .HasLegend = True
                    .Legend.Position = Excel.Constants.xlRight
                    .Legend.Font.Size = awinSettings.fontsizeLegend + 2



                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    With .SeriesCollection.NewSeries
                        .name = "Auslastung"
                        .Values = datenreihe
                        .XValues = Xdatenreihe
                        .HasDataLabels = False
                        .ChartType = Excel.XlChartType.xlPie
                        .Points(1).Interior.color = awinSettings.AmpelGruen
                        .Points(2).Interior.color = awinSettings.AmpelRot
                        .Points(3).Interior.color = awinSettings.AmpelGelb


                        For i = 1 To 3
                            htxt = Format(datenreihe(i - 1), "###,###0")
                            With .Points(i)
                                .HasDataLabel = True
                                .DataLabel.text = htxt

                                .DataLabel.Font.Size = awinSettings.fontsizeItems + 2


                                Try
                                    .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                                Catch ex As Exception

                                End Try


                            End With
                        Next i

                    End With


                    .ChartTitle.text = diagramTitle
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                    .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                        titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend
                    .Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With
                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = awinSettings.ChartHoehe2
                    .name = chtobjName

                End With


                ' myCollection wird jetzt über alle Rollen aufgebaut ..
                myCollection.Clear()

                For i = 1 To RoleDefinitions.Count
                    Dim roleName As String
                    roleName = RoleDefinitions.getRoledef(i).name
                    Try
                        myCollection.Add(roleName, roleName)
                    Catch ex As Exception

                    End Try

                Next

                repObj = .ChartObjects(anzDiagrams + 1)

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then

                    Dim prcDiagram As New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    Dim prcChart As New clsEventsPrcCharts
                    prcChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart
                    prcDiagram.setDiagramEvent = prcChart
                    ' Ende Event Handling für Chart 


                    With prcDiagram
                        .DiagrammTitel = diagramTitle
                        .diagrammTyp = DiagrammTypen(4)
                        .gsCollection = myCollection
                        .isCockpitChart = False
                        .top = top
                        .left = left
                        .width = width
                        .height = height
                        .kennung = chtobjName
                    End With

                    ' eintragen in die sortierte Liste mit .kennung als dem Schlüssel 
                    ' wenn das Diagramm bereits existiert, muss es gelöscht werden, dann neu ergänzt ... 
                    Try
                        DiagramList.Add(prcDiagram)
                    Catch ex As Exception

                        Try
                            DiagramList.Remove(prcDiagram.kennung)
                            DiagramList.Add(prcDiagram)
                        Catch ex1 As Exception

                        End Try


                    End Try

                End If



            End If
        End With

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU




    End Sub

    ''' <summary>
    ''' erstellt ein Pie-Chart, das die Verteilung der Bewertungen anzeigt (wieviel grün, gelb, rot, grau) 
    ''' </summary>
    ''' <param name="repObj">Verweis auf das Chart Objekt - wird für die Report Erstellung benötigt </param>
    ''' <param name="future">-1 nur Vergangenheit, 1: nur Zukunft</param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <remarks></remarks>
    Sub awinCreateZielErreichungsDiagramm(ByRef repObj As Excel.ChartObject, ByVal future As Integer, _
                                                ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                                ByVal isCockpitChart As Boolean, ByVal calledfromReporting As Boolean)

        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim Xdatenreihe() As String
        Dim datenreihe() As Integer
        Dim htxt As String

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim diagramTitle As String
        Dim von As Integer, bis As Integer
        Dim chtTitle As String
        Dim chtobjName As String
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim kennung As String
        Dim heuteColumn As Integer = getColumnOfDate(Date.Now)
        'Dim sumDiagram As clsDiagramm
        'Dim sumChart As clsEventsPrcCharts

        ReDim Xdatenreihe(3)
        ReDim datenreihe(3)


        If future = -1 Then

            Dim myCollection As New Collection
            myCollection.Add("ZieleV")
            chtobjName = getKennung("pf", PTpfdk.ZieleV, myCollection)
            If showRangeLeft <= heuteColumn Then
                titelTeile(0) = summentitel6
                titelTeile(1) = textZeitraum(showRangeLeft, heuteColumn)
                Xdatenreihe(0) = "keine Information"
                Xdatenreihe(1) = "erreicht"
                Xdatenreihe(2) = "mit Einschränkungen"
                Xdatenreihe(3) = "nicht erreicht"
            Else
                Throw New ArgumentException("der betrachtete Bereich liegt vollständig in der Zukunft ... es gibt keine erreichten Ziele")
            End If


        ElseIf future = 1 Then
            Dim myCollection As New Collection
            myCollection.Add("ZieleF")
            chtobjName = getKennung("pf", PTpfdk.ZieleF, myCollection)
            If heuteColumn + 1 <= showRangeRight Then
                titelTeile(0) = summentitel7
                titelTeile(1) = textZeitraum(getColumnOfDate(Date.Now) + 1, showRangeRight)
                Xdatenreihe(0) = "keine Information"
                Xdatenreihe(1) = "wird erreicht"
                Xdatenreihe(2) = "Unsicherheiten"
                Xdatenreihe(3) = "erhebliche Risiken"
            Else
                Throw New ArgumentException("der betrachtete Bereich liegt vollständig in der Vergangenheit ... es gibt keine Prognose Werte")
            End If

        Else
            'titelTeile(0) = summentitel8
            'titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
            'Xdatenreihe(0) = "keine Information"
            'Xdatenreihe(1) = "wurde/wird erreicht"
            'Xdatenreihe(2) = "mit Einschränkungen/Unsicherheiten"
            'Xdatenreihe(3) = "nicht erreicht/erhebliche Risiken"
            Throw New ArgumentException("keine Angabe in Zielerreichungsdiagramm, ob Vergangenheit oder Zukunft betrachtet werden soll ")
        End If


        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeilLaengen(1) = titelTeile(1).Length



        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)


        von = showRangeLeft
        bis = showRangeRight

        ' jetzt prüfen, ob es bereits gespeicherte Werte für top, left, ... gibt ;
        ' Wenn ja : übernehmen


        If Not calledfromReporting Then
            Dim foundDiagramm As clsDiagramm

            Try
                foundDiagramm = DiagramList.getDiagramm(chtobjName)
                With foundDiagramm
                    top = .top
                    left = .left
                    width = .width
                    height = .height
                End With
            Catch ex As Exception

            End Try
        End If


        datenreihe(0) = ShowProjekte.getColorsInMonth(0, future).Sum
        datenreihe(1) = ShowProjekte.getColorsInMonth(1, future).Sum
        datenreihe(2) = ShowProjekte.getColorsInMonth(2, future).Sum
        datenreihe(3) = ShowProjekte.getColorsInMonth(3, future).Sum

        If datenreihe.Sum = 0 Then

            If future < 0 Then
                Call MsgBox("es gibt im betrachteten Zeitraum keine Ergebnisse aus der Vergangenheit ...")
            ElseIf future > 0 Then
                Call MsgBox("es gibt im betrachteten Zeitraum keine geplanten, zukünftigen Ergebnisse ...")
            Else
                Call MsgBox("es gibt im betrachteten Zeitraum keine vergangenen oder zukünftigen Ergebnisse ...")
            End If

        Else

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False

            With appInstance.Worksheets(arrWsNames(3))

                anzDiagrams = .ChartObjects.Count

                '
                ' um welches Diagramm handelt es sich ...
                '
                i = 1
                found = False
                While i <= anzDiagrams And Not found

                    Try
                        chtTitle = .ChartObjects(i).Name
                    Catch ex As Exception
                        chtTitle = " "
                    End Try


                    If chtobjName = .ChartObjects(i).Name Then
                        found = True
                        repObj = .Chartobjects(i)
                    Else
                        i = i + 1
                    End If

                End While

                If Not found Then

                    With appInstance.Charts.Add
                        .HasTitle = True
                        .ChartTitle.text = diagramTitle
                        If isCockpitChart Then
                            .HasLegend = False
                            .ChartTitle.font.size = awinSettings.CPfontsizeTitle
                        Else
                            .HasLegend = True
                            .Legend.Position = Excel.Constants.xlRight
                            .Legend.Font.Size = awinSettings.fontsizeLegend
                            .ChartTitle.font.size = awinSettings.fontsizeTitle
                        End If

                        Do Until .SeriesCollection.Count = 0
                            .SeriesCollection(1).Delete()
                        Loop


                        With .SeriesCollection.NewSeries
                            .name = "Status-Übersicht"
                            .Values = datenreihe
                            .XValues = Xdatenreihe
                            .HasDataLabels = False
                            .ChartType = Excel.XlChartType.xlPie

                            .Points(1).Interior.color = awinSettings.AmpelNichtBewertet
                            .Points(2).Interior.color = awinSettings.AmpelGruen
                            .Points(3).Interior.color = awinSettings.AmpelGelb
                            .Points(4).Interior.color = awinSettings.AmpelRot

                            For i = 1 To 4
                                htxt = Format(datenreihe(i - 1), "###,###0")
                                With .Points(i)
                                    .HasDataLabel = True
                                    .DataLabel.text = htxt
                                    If isCockpitChart Then
                                        '.DataLabel.Font.Size = 8
                                        .HasDataLabel = False
                                    Else
                                        .DataLabel.Font.Size = awinSettings.fontsizeItems
                                    End If

                                    Try
                                        .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                                    Catch ex As Exception

                                    End Try


                                End With
                            Next i

                        End With

                        .Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                    End With
                    With .ChartObjects(anzDiagrams + 1)
                        .top = top
                        .left = left
                        .width = width
                        .height = height
                        .name = chtobjName

                    End With

                    If isCockpitChart Then
                        Try
                            With appInstance.ActiveSheet
                                .Shapes(chtobjName).line.visible = False
                            End With
                        Catch ex As Exception

                        End Try
                    Else
                        'Call awinScrollintoView()
                    End If

                    repObj = .ChartObjects(anzDiagrams + 1)


                    ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                    ' aufgerufen wurde

                    If Not calledfromReporting Then

                        Dim prcDiagram As New clsDiagramm



                        ' Anfang Event Handling für Chart 
                        Dim prcChart As New clsEventsPrcCharts
                        prcChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart
                        prcDiagram.setDiagramEvent = prcChart
                        ' Ende Event Handling für Chart 


                        With prcDiagram
                            .DiagrammTitel = diagramTitle
                            .diagrammTyp = DiagrammTypen(4)
                            .gsCollection = Nothing
                            .isCockpitChart = isCockpitChart
                            .top = top
                            .left = left
                            .width = width
                            .height = height
                            .kennung = chtobjName
                        End With

                        ' eintragen in die sortierte Liste mit .kennung als dem Schlüssel 
                        ' wenn das Diagramm bereits existiert, muss es gelöscht werden, dann neu ergänzt ... 
                        Try
                            DiagramList.Add(prcDiagram)
                        Catch ex As Exception

                            Try
                                DiagramList.Remove(prcDiagram.kennung)
                                DiagramList.Add(prcDiagram)
                            Catch ex1 As Exception

                            End Try


                        End Try

                    End If


                    'sumDiagram = New clsDiagramm

                    'sumChart = New clsEventsPrcCharts
                    'sumChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart

                    'sumDiagram.setDiagramEvent = sumChart


                    'With sumDiagram
                    '    .DiagrammTitel = diagramTitle
                    '    .diagrammTyp = DiagrammTypen(4)
                    '    '.setCollection = myCollection
                    '    .isCockpitChart = isCockpitChart
                    'End With

                    'DiagramList.Add(sumDiagram)
                    'sumDiagram = Nothing


                End If
            End With

            appInstance.EnableEvents = formerEE
            appInstance.ScreenUpdating = formerSU

        End If


    End Sub

    Sub awinUpdateAuslastungsDiagramm(ByRef repObj As Excel.ChartObject)

        Dim anzDiagrams As Integer, i As Integer

        Dim Xdatenreihe() As String
        Dim datenreihe() As Double
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents


        Dim von As Integer, bis As Integer
        Dim diagramTitle As String
        Dim htxt As String
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim kennung As String



        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False


        titelTeile(0) = summentitel9 & " (" & awinSettings.kapaEinheit & ")"


        titelTeilLaengen(0) = titelTeile(0).Length + 1
        titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)



        von = showRangeLeft
        bis = showRangeRight


        ReDim Xdatenreihe(2)
        ReDim datenreihe(2)

        Xdatenreihe(0) = "Auslastung"
        Xdatenreihe(1) = "Über-Auslastung"
        Xdatenreihe(2) = "Unter-Auslastung"


        datenreihe(0) = ShowProjekte.getAuslastungsValues(0).Sum
        datenreihe(1) = ShowProjekte.getAuslastungsValues(1).Sum
        datenreihe(2) = ShowProjekte.getAuslastungsValues(2).Sum


        With appInstance.Worksheets(arrWsNames(3))




            With repObj.Chart


                Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete()
                Loop


                With .SeriesCollection.NewSeries
                    .name = "Auslastung"
                    .Values = datenreihe
                    .XValues = Xdatenreihe
                    .HasDataLabels = False
                    .ChartType = Excel.XlChartType.xlPie
                    .Points(1).Interior.color = awinSettings.AmpelGruen
                    .Points(2).Interior.color = awinSettings.AmpelRot
                    .Points(3).Interior.color = awinSettings.AmpelGelb


                    For i = 1 To 3
                        htxt = Format(datenreihe(i - 1), "###,###0")
                        With .Points(i)
                            .HasDataLabel = True
                            .DataLabel.text = htxt

                            .DataLabel.Font.Size = awinSettings.fontsizeItems + 2


                            Try
                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                            Catch ex As Exception

                            End Try


                        End With
                    Next i

                End With


                .ChartTitle.Text = diagramTitle
                .ChartTitle.Font.Size = awinSettings.fontsizeTitle
                .ChartTitle.Format.TextFrame2.TextRange.Characters(titelTeilLaengen(0) + 1, _
                    titelTeilLaengen(1)).Font.Size = awinSettings.fontsizeLegend

            End With

            ' myCollection wird jetzt über alle Rollen aufgebaut ..
            'Dim myCollection As New Collection

            'For i = 1 To RoleDefinitions.Count
            '    Dim roleName As String
            '    roleName = RoleDefinitions.getRoledef(i).name
            '    Try
            '        myCollection.Add(roleName, roleName)
            '    Catch ex As Exception

            '    End Try

            'Next

            repObj = .ChartObjects(anzDiagrams + 1)
            ' Änderung 31.7 dieses Diagramm muss nicht geupdated werden, ausserdem macht es keinen Sinn, den Roentgenblick anzuwenden 
            ' die Optimierung kann ebenso über die Summe der Rollen gemacht werden 

            'sumDiagram = New clsDiagramm

            'sumChart = New clsEventsPrcCharts
            'sumChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            'sumDiagram.setDiagramEvent = sumChart


            'With sumDiagram
            '    .DiagrammTitel = diagramTitle
            '    .diagrammTyp = DiagrammTypen(4)
            '    .gsCollection = myCollection
            '    .isCockpitChart = False
            'End With

            'DiagramList.Add(sumDiagram)




        End With

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU




    End Sub


    Sub awinUpdateColorDistributionDiagramm(ByRef repObj As Excel.ChartObject)

        Dim anzDiagrams As Integer, i As Integer
        Dim Xdatenreihe() As String
        Dim datenreihe() As Integer
        Dim htxt As String

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim diagramTitle As String
        Dim chtobjName As String = repObj.Name
        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim future As Integer
        Dim heuteColumn As Integer = getColumnOfDate(Date.Now)

        ReDim Xdatenreihe(3)
        ReDim datenreihe(3)



        If chtobjName = summentitel6 Then
            future = -1
            If showRangeLeft <= heuteColumn Then
                titelTeile(0) = summentitel6
                titelTeile(1) = textZeitraum(showRangeLeft, heuteColumn)
                Xdatenreihe(0) = "keine Information"
                Xdatenreihe(1) = "erreicht"
                Xdatenreihe(2) = "mit Einschränkungen"
                Xdatenreihe(3) = "nicht erreicht"
            Else
                Throw New ArgumentException("der betrachtete Bereich liegt vollständig in der Zukunft ... es gibt keine erreichten Ziele")
            End If


        ElseIf chtobjName = summentitel7 Then
            future = 1
            If heuteColumn + 1 <= showRangeRight Then
                titelTeile(0) = summentitel7
                titelTeile(1) = textZeitraum(getColumnOfDate(Date.Now) + 1, showRangeRight)
                Xdatenreihe(0) = "keine Information"
                Xdatenreihe(1) = "wird erreicht"
                Xdatenreihe(2) = "Unsicherheiten"
                Xdatenreihe(3) = "erhebliche Risiken"
            Else
                Throw New ArgumentException("der betrachtete Bereich liegt vollständig in der Vergangenheit ... es gibt keine Prognose Werte")
            End If

        Else   'If chtobjName = summentitel8 Then
            future = 0
            titelTeile(0) = summentitel8
            titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
            Xdatenreihe(0) = "keine Information"
            Xdatenreihe(1) = "wurde/wird erreicht"
            Xdatenreihe(2) = "mit Einschränkungen/Unsicherheiten"
            Xdatenreihe(3) = "nicht erreicht/erhebliche Risiken"
        End If



        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeilLaengen(1) = titelTeile(1).Length

        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)



        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False


        datenreihe(0) = ShowProjekte.getColorsInMonth(0, future).Sum
        datenreihe(1) = ShowProjekte.getColorsInMonth(1, future).Sum
        datenreihe(2) = ShowProjekte.getColorsInMonth(2, future).Sum
        datenreihe(3) = ShowProjekte.getColorsInMonth(3, future).Sum




        With appInstance.Worksheets(arrWsNames(3))

            anzDiagrams = .ChartObjects.Count



            With repObj.Chart
                .HasTitle = True
                .ChartTitle.Text = diagramTitle


                Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete()
                Loop


                With .SeriesCollection.NewSeries
                    .name = "Status-Übersicht"
                    .Values = datenreihe
                    .XValues = Xdatenreihe
                    .HasDataLabels = False
                    .ChartType = Excel.XlChartType.xlPie

                    .Points(1).Interior.color = awinSettings.AmpelNichtBewertet
                    .Points(2).Interior.color = awinSettings.AmpelGruen
                    .Points(3).Interior.color = awinSettings.AmpelGelb
                    .Points(4).Interior.color = awinSettings.AmpelRot

                    For i = 1 To 4
                        htxt = Format(datenreihe(i - 1), "###,###0")
                        With .Points(i)
                            .HasDataLabel = True
                            .DataLabel.text = htxt

                            .DataLabel.Font.Size = awinSettings.fontsizeItems

                            Try
                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                            Catch ex As Exception

                            End Try


                        End With
                    Next i

                End With

            End With

        End With

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU





    End Sub

    Sub awinCreatePersCostStructureDiagramm(ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, ByVal isCockpitChart As Boolean)


        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim Xdatenreihe() As String
        Dim datenreihe() As Double
        Dim htxt As String

        Dim updateScreenWasTrue As Boolean
        Dim diagramTitle As String
        Dim k0sum As Double, k1sum As Double, k2sum As Double, k3sum As Double
        Dim von As Integer, bis As Integer
        Dim chtTitle As String
        Dim chtobjName As String



        If appInstance.ScreenUpdating = True Then
            updateScreenWasTrue = True
            appInstance.ScreenUpdating = False
        Else
            updateScreenWasTrue = False
        End If


        width = 450
        diagramTitle = summentitel4 & " (T€)"
        chtobjName = diagramTitle
        von = showRangeLeft
        bis = showRangeRight

        appInstance.EnableEvents = False




        k0sum = System.Math.Round(ShowProjekte.getCostiValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        k2sum = System.Math.Round(ShowProjekte.getadditionalECostinMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        k1sum = System.Math.Round(ShowProjekte.getCosteValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10 - k2sum
        k3sum = System.Math.Round(ShowProjekte.getCostoValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10



        ReDim Xdatenreihe(3)
        ReDim datenreihe(3)
        Xdatenreihe(0) = "Auslastung"
        Xdatenreihe(1) = "Über-Auslastung (bewertet zu internen Kosten)"
        Xdatenreihe(2) = "Über-Auslastung (Mehrkosten)"
        Xdatenreihe(3) = "Unter-Auslastung"

        datenreihe(0) = k0sum
        datenreihe(1) = k1sum
        datenreihe(2) = k2sum
        datenreihe(3) = k3sum



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


                If ((chtTitle Like (diagramTitle & "*")) And _
                         (isCockpitChart = istCockpitDiagramm(.ChartObjects(i)))) Then
                    found = True
                Else
                    i = i + 1
                End If

            End While

            If Not found Then

                With appInstance.Charts.Add
                    .HasTitle = True
                    .ChartTitle.text = diagramTitle
                    If isCockpitChart Then
                        .HasLegend = False
                        .ChartTitle.font.size = awinSettings.CPfontsizeTitle
                    Else
                        .HasLegend = True
                        .Legend.Position = Excel.Constants.xlRight
                        .Legend.Font.Size = awinSettings.fontsizeLegend
                        .ChartTitle.font.size = awinSettings.fontsizeTitle
                    End If

                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    With .SeriesCollection.NewSeries
                        .name = "Projekt-Effizienz"
                        .Values = datenreihe
                        .XValues = Xdatenreihe
                        .HasDataLabels = False
                        .ChartType = Excel.XlChartType.xlPie
                        '.Points(1).Interior.color = CostDefinitions.getCostdef(CostDefinitions.Count).farbe
                        '.Points(2).Interior.color = iWertFarbe
                        '.Points(3).Interior.color = farbeExterne
                        '.Points(4).Interior.color = farbeInternOP

                        .Points(1).Interior.color = awinSettings.AmpelGruen
                        .Points(2).Interior.color = awinSettings.AmpelNichtBewertet
                        .Points(3).Interior.color = awinSettings.AmpelRot
                        .Points(4).Interior.color = awinSettings.AmpelGelb

                        For i = 1 To 4
                            htxt = Format(datenreihe(i - 1), "###,###0")
                            With .Points(i)
                                .HasDataLabel = True
                                .DataLabel.text = htxt
                                If isCockpitChart Then
                                    '.DataLabel.Font.Size = 8
                                    .HasDataLabel = False
                                Else
                                    .DataLabel.Font.Size = awinSettings.fontsizeItems
                                End If

                                Try
                                    .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                                Catch ex As Exception

                                End Try


                            End With
                        Next i

                    End With

                    .Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)

                End With
                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = awinSettings.ChartHoehe2
                    .name = chtobjName

                End With

                If isCockpitChart Then
                    Try
                        With appInstance.ActiveSheet
                            .Shapes(chtobjName).line.visible = False
                        End With
                    Catch ex As Exception

                    End Try
                Else
                    'Call awinScrollintoView()
                End If

                ' Änderung 31.7 : ohne Right Klick , ohne Optimierung 
                ' myCollection wird jetzt über alle Rollen aufgebaut ..
                'Dim myCollection As New Collection
                'Dim roleName As String

                'For i = 1 To RoleDefinitions.Count

                '    roleName = RoleDefinitions.getRoledef(i).name
                '    Try
                '        myCollection.Add(roleName, roleName)
                '    Catch ex As Exception

                '    End Try

                'Next


                'sumDiagram = New clsDiagramm

                'sumChart = New clsEventsPrcCharts
                'sumChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart

                'sumDiagram.setDiagramEvent = sumChart


                'With sumDiagram
                '    .DiagrammTitel = diagramTitle
                '    .diagrammTyp = DiagrammTypen(4)
                '    .gsCollection = myCollection
                '    .isCockpitChart = isCockpitChart
                'End With

                'DiagramList.Add(sumDiagram)
                'sumDiagram = Nothing


            End If
        End With

        appInstance.EnableEvents = True
        If updateScreenWasTrue Then
            appInstance.ScreenUpdating = True
        End If


    End Sub
    '
    '
    '
    Sub awinUpdatePersCostStructureDiagramm(ByRef chtobj As ChartObject)


        Dim Xdatenreihe() As String
        Dim datenreihe() As Double
        Dim k0sum, k1sum As Double, k2sum As Double, k3sum As Double
        Dim von As Integer, bis As Integer
        Dim htxt As String
        Dim i As Integer
        Dim isCockpitChart As Boolean


        von = showRangeLeft
        bis = showRangeRight

        If istCockpitDiagramm(chtobj) Then
            ' dann ist es ein Cockpit Chart ....
            isCockpitChart = True
        Else
            isCockpitChart = False
        End If


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False





        k0sum = ShowProjekte.getCostiValuesInMonth.Sum
        k2sum = ShowProjekte.getadditionalECostinMonth.Sum
        k1sum = ShowProjekte.getCosteValuesInMonth.Sum - k2sum
        k3sum = ShowProjekte.getCostoValuesInMonth.Sum



        ReDim Xdatenreihe(3)
        ReDim datenreihe(3)
        Xdatenreihe(0) = "Über interne Ressourcen geleistet"
        Xdatenreihe(1) = "Über externe Ressourcen geleistet (zu internen Kosten)"
        Xdatenreihe(2) = "Mehrkosten durch Überauslastung"
        Xdatenreihe(3) = "Mehrkosten durch Unterauslastung"

        datenreihe(0) = k0sum
        datenreihe(1) = k1sum
        datenreihe(2) = k2sum
        datenreihe(3) = k3sum






        With appInstance.Worksheets(arrWsNames(3))

            With chtobj

                ' hier müssen die seriescollection daten gelöscht und komplett neu aufgebaut werden ....
                '
                '
                '
                With .Chart
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop

                    With .SeriesCollection.NewSeries
                        .name = "Projekt-Effizienz"
                        .HasDataLabels = False
                        .Values = datenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlPie
                        .Points(1).Interior.color = awinSettings.AmpelGruen
                        .Points(2).Interior.color = awinSettings.AmpelNichtBewertet
                        .Points(3).Interior.color = awinSettings.AmpelRot
                        .Points(4).Interior.color = awinSettings.AmpelGelb

                        For i = 1 To 4
                            htxt = Format(datenreihe(i - 1), "###,###0")
                            With .Points(i)
                                .HasDataLabel = True
                                .DataLabel.text = htxt
                                If isCockpitChart Then
                                    '.DataLabel.Font.Size = 8
                                    .HasDataLabel = False
                                Else
                                    .DataLabel.Font.Size = awinSettings.fontsizeItems
                                End If

                                Try
                                    .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                                Catch ex As Exception

                                End Try


                            End With
                        Next i

                    End With
                End With
            End With


        End With

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU



    End Sub

    '
    '
    '
    Sub awinUpdateEffizienzDiagramm2(ByRef chtobj As ChartObject)


        Dim Xdatenreihe(0) As String
        'Dim datenreihe() As Double
        Dim von As Integer, bis As Integer
        Dim updateScreenWasTrue As Boolean
        'Dim htxt As String
        'Dim i As Integer
        Dim isCockpitChart As Boolean
        Dim earnedValue As Double, earnedValueweighted As Double, diff As Double
        Dim minscale As Double


        von = showRangeLeft
        bis = showRangeRight

        If istCockpitDiagramm(chtobj) Then
            ' dann ist es ein Cockpit Chart ....
            isCockpitChart = True
        Else
            isCockpitChart = False
        End If


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        If appInstance.ScreenUpdating = True Then
            updateScreenWasTrue = True
            appInstance.ScreenUpdating = False
        Else
            updateScreenWasTrue = False
        End If





        Xdatenreihe(0) = "Zeitraum" & vbLf & StartofCalendar.AddMonths(showRangeLeft - 1).ToString("MMM yy") & " - " & _
                    StartofCalendar.AddMonths(showRangeRight - 1).ToString("MMM yy")


        earnedValue = ShowProjekte.getEarnedValuesInMonth.Sum
        diff = ShowProjekte.getWeightedRiskValuesInMonth.Sum
        earnedValueweighted = earnedValue - diff

        If earnedValueweighted < 0 Then
            minscale = earnedValueweighted
        Else
            minscale = 0
        End If


        With appInstance.Worksheets(arrWsNames(3))

            With chtobj

                ' hier müssen die seriescollection daten gelöscht und komplett neu aufgebaut werden ....
                '
                '
                '
                Dim htxt As String
                With .Chart
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    'series
                    With .SeriesCollection.NewSeries
                        .name = ergebnisChartName(0)
                        .hasdatalabels = True
                        .Interior.color = ergebnisfarbe1
                        .Values = earnedValue
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnClustered

                        htxt = Format(earnedValue, "###,###0") & " T€"
                        With .Points(1)
                            .HasDataLabel = True
                            .DataLabel.text = htxt
                            If isCockpitChart Then
                                .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                            Else
                                .DataLabel.Font.Size = awinSettings.fontsizeItems
                            End If

                            Try
                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionBestFit
                            Catch ex As Exception

                            End Try


                        End With
                    End With


                    With .SeriesCollection.NewSeries
                        .name = ergebnisChartName(1)
                        .hasdatalabels = True
                        .Interior.color = ergebnisfarbe2
                        .Values = earnedValueweighted
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnClustered

                        htxt = Format(earnedValueweighted, "###,###0") & " T€"
                        With .Points(1)
                            .HasDataLabel = True
                            .DataLabel.text = htxt
                            If isCockpitChart Then
                                .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                            Else
                                .DataLabel.Font.Size = awinSettings.fontsizeItems
                            End If

                            Try
                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionBestFit
                            Catch ex As Exception

                            End Try


                        End With
                    End With

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        If minscale < 0 Then
                            .TickLabelPosition = Excel.Constants.xlLow
                        End If
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        If minscale < 0 Then
                            .MinimumScale = System.Math.Round(minscale - 1, mode:=MidpointRounding.ToEven)
                        Else
                            .MinimumScale = 0
                        End If

                    End With

                End With
            End With


        End With

        appInstance.EnableEvents = formerEE
        If updateScreenWasTrue Then
            appInstance.ScreenUpdating = True
        End If


    End Sub

    Sub awinUpdateErgebnisDiagramm(ByRef chtobj As ChartObject)


        Dim diagramTitle As String

        Dim minScale As Double
        Dim Xdatenreihe(3) As String
        Dim valueDatenreihe1(3) As Double
        Dim valueDatenreihe2(3) As Double
        Dim itemColor(3) As Object
        Dim itemValue(3) As Double
        Dim earnedValue As Double, additionalCostExt As Double, riskValue As Double, internwithoutProject As Double
        Dim ertragsWert As Double

        Dim mycollection As New Collection


        Xdatenreihe(0) = "Summe Projekt-Ergebnisse (Risiko-gewichtet)"
        'Xdatenreihe(1) = "Risiko-Abschlag"
        Xdatenreihe(1) = "Mehrkosten wegen Überauslastung"
        Xdatenreihe(2) = "Opportunitätskosten durch Unterauslastung"
        Xdatenreihe(3) = "Ergebnis-Kennzahl"


        Dim positiv As Boolean = True

        ' das ist der Earned Value 
        earnedValue = System.Math.Round(ShowProjekte.getEarnedValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        ' das ist der Risiko Abschlag  
        riskValue = System.Math.Round(ShowProjekte.getWeightedRiskValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10

        itemValue(0) = earnedValue - riskValue
        If itemValue(0) >= 0 Then
            itemColor(0) = ergebnisfarbe1
        Else
            itemColor(0) = farbeExterne
        End If

        Dim currentWert As Double = itemValue(0)


        ' das sind die Zusatzkosten, die durch Externe (wg Überauslastung) verursacht werden
        additionalCostExt = System.Math.Round(ShowProjekte.getadditionalECostinMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        itemValue(1) = additionalCostExt
        itemColor(1) = farbeExterne

        ' das sind die durch Unterauslastung verursachten Kosten , also Personal-Kosten von Leuten, die in keinem Projekt sind
        internwithoutProject = System.Math.Round(ShowProjekte.getCostoValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        itemValue(2) = internwithoutProject
        itemColor(2) = awinSettings.AmpelGelb

        ' das ist der Ertrag 
        ertragsWert = earnedValue - (riskValue + additionalCostExt + internwithoutProject)
        itemValue(3) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(3) = ergebnisfarbe2
        Else
            itemColor(3) = farbeExterne
        End If


        diagramTitle = summentitel1 & " " & textZeitraum(showRangeLeft, showRangeRight)

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False

        If ertragsWert < 0 Then
            minScale = System.Math.Round(ertragsWert / 10, mode:=MidpointRounding.ToEven) * 10
        Else
            minScale = 0
        End If

        'Dim htxt As String

        Dim valueCrossesNull As Boolean = False


        With appInstance.Worksheets(arrWsNames(3))

            With chtobj.Chart
                ' remove extra series
                Do Until .SeriesCollection.Count = 0
                    .SeriesCollection(1).Delete()
                Loop
                Dim crossindex As Integer = -1

                ' bestimmen des Anfangs  
                Dim iv = 0
                valueDatenreihe1(iv) = 0
                valueDatenreihe2(iv) = itemValue(iv)
                currentWert = itemValue(iv)
                Dim formerValue As Double = currentWert
                Dim negativeFromNull As Boolean = False

                ' alle nächsten Zwischen-Werte 
                For iv = 1 To 2
                    If formerValue <= 0 Then
                        negativeFromNull = True
                    Else
                        negativeFromNull = False
                    End If

                    currentWert = currentWert - itemValue(iv)
                    valueCrossesNull = (currentWert + itemValue(iv) > 0) And (currentWert < 0)

                    If currentWert >= 0 Then
                        valueDatenreihe1(iv) = currentWert
                        valueDatenreihe2(iv) = itemValue(iv)
                    ElseIf valueCrossesNull Then
                        valueDatenreihe1(iv) = currentWert
                        valueDatenreihe2(iv) = itemValue(iv) - currentWert * (-1) ' notwendig da currentWert ja negativ ist ..
                        crossindex = iv + 1
                    ElseIf negativeFromNull Then
                        valueDatenreihe1(iv) = formerValue
                        valueDatenreihe2(iv) = itemValue(iv) * (-1)
                    Else
                        valueDatenreihe1(iv) = currentWert
                        valueDatenreihe2(iv) = itemValue(iv) * (-1)
                    End If

                    formerValue = currentWert
                Next

                ' bestimmen des Ende 
                iv = 3
                valueDatenreihe1(iv) = 0
                valueDatenreihe2(iv) = itemValue(iv)



                'series
                With .SeriesCollection.NewSeries
                    .name = "Bottom"
                    .HasDataLabels = False
                    .Interior.colorindex = -4142
                    .Values = valueDatenreihe1
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlColumnStacked
                    If crossindex > 0 Then
                        ' es gab einen Übergang , dort muss Bottom auf die entsprechende Farbe gesetzt werden 
                        With .Points(crossindex)
                            .Interior.color = itemColor(crossindex - 1)
                        End With
                    End If

                End With

                With .SeriesCollection.NewSeries
                    .name = "Top"
                    .HasDataLabels = True
                    .Values = valueDatenreihe2
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlColumnStacked

                    For iv = 0 To 3

                        With .Points(iv + 1)
                            .HasDataLabel = True
                            .DataLabel.text = Format(itemValue(iv), "###,###0") & " T€"
                            .Interior.color = itemColor(iv)
                            .DataLabel.Font.Size = awinSettings.fontsizeLegend
                            Try
                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                            Catch ex As Exception

                            End Try
                        End With

                    Next

                End With
                .HasAxis(Excel.XlAxisType.xlCategory) = True
                .HasAxis(Excel.XlAxisType.xlValue) = False

                With .Axes(Excel.XlAxisType.xlCategory)
                    .HasTitle = False
                    If minScale < 0 Then
                        .TickLabelPosition = Excel.Constants.xlLow
                    End If
                    '.MinimumScale = 0

                End With

                'Dim hax As Excel.Axis
                'With hax
                '    .HasMajorGridlines
                '    .hasminor()
                'End With

                Try
                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        .HasMajorGridlines = False
                        .hasminorgridlines = False
                        If minScale < 0 Then
                            .MinimumScale = System.Math.Round((minScale - 1) / 10, mode:=MidpointRounding.ToEven) * 10
                        Else
                            .MinimumScale = 0
                        End If
                    End With
                Catch ex As Exception

                End Try


                .HasLegend = False
                'With .Legend
                '    .Position = XlConstants.xlTop
                '    .Font.Size = 8
                'End With
                .HasTitle = True
                .ChartTitle.Text = diagramTitle


            End With
        End With


        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    ''' <summary>
    ''' zeigt die Earned Values / Earned Values gewichtet an 
    ''' </summary>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <remarks></remarks>
    Sub awinCreateEffizienzDiagramm2(ByRef repObj As Object, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, ByVal isCockpitChart As Boolean)

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean
        'Dim plen As Integer
        Dim i As Integer
        Dim minScale As Double
        Dim Xdatenreihe(0) As String
        Dim earnedValue As Double, riskValue As Double
        Dim earnedValueWeighted As Double
        'Dim top As Double, left As Double, width As Double, height As Double
        Dim chtTitle As String
        'Dim pstart As Integer
        Dim mycollection As New Collection
        'Dim catName As String
        Dim sumDiagram As clsDiagramm
        Dim sumChart As clsEventsPrcCharts

        'Dim hproj As clsProjekt
        Dim ErgebnisListeR As New Collection
        Dim updateScreenwastrue As Boolean


        If appInstance.ScreenUpdating = True Then
            updateScreenwastrue = True
            appInstance.ScreenUpdating = False
        Else
            updateScreenwastrue = False
        End If




        Xdatenreihe(0) = "Zeitraum" & vbLf & StartofCalendar.AddMonths(showRangeLeft - 1).ToString("MMM yy") & " - " & _
                    StartofCalendar.AddMonths(showRangeRight - 1).ToString("MMM yy")


        earnedValue = ShowProjekte.getEarnedValuesInMonth.Sum - ShowProjekte.getadditionalECostinMonth.Sum
        riskValue = ShowProjekte.getWeightedRiskValuesInMonth.Sum


        earnedValueWeighted = earnedValue - riskValue





        diagramTitle = ergebnisChartName(0) & " / " & ergebnisChartName(1)




        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


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


                If ((chtTitle Like (diagramTitle & "*")) And _
                         (isCockpitChart = istCockpitDiagramm(.ChartObjects(i)))) Then
                    found = True
                    repObj = .ChartObjects(i)
                Else
                    i = i + 1
                End If

            End While



            If found Then

                MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                If earnedValueWeighted < 0 Then
                    minScale = earnedValueWeighted
                Else
                    minScale = 0
                End If

                Dim htxt As String

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    'series
                    With .SeriesCollection.NewSeries
                        .name = ergebnisChartName(0)
                        .HasDataLabels = True
                        .Interior.color = ergebnisfarbe1
                        .Values = earnedValue
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnClustered

                        htxt = Format(earnedValue, "###,###0") & " T€"
                        With .Points(1)
                            .HasDataLabel = True
                            .DataLabel.text = htxt
                            If isCockpitChart Then
                                .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                            Else
                                .DataLabel.Font.Size = awinSettings.fontsizeItems
                            End If

                            Try
                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionBestFit
                            Catch ex As Exception

                            End Try


                        End With
                    End With


                    With .SeriesCollection.NewSeries
                        .name = ergebnisChartName(1)
                        .HasDataLabels = True
                        .Interior.color = ergebnisfarbe2
                        .Values = earnedValueWeighted
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnClustered

                        htxt = Format(earnedValueWeighted, "###,###0") & " T€"
                        With .Points(1)
                            .HasDataLabel = True
                            .DataLabel.text = htxt
                            If isCockpitChart Then
                                .DataLabel.Font.Size = awinSettings.CPfontsizeItems
                            Else
                                .DataLabel.Font.Size = awinSettings.fontsizeItems
                            End If

                            Try
                                .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionBestFit
                            Catch ex As Exception

                            End Try


                        End With
                    End With

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = True

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        If minScale < 0 Then
                            .TickLabelPosition = Excel.Constants.xlLow
                        End If
                        '.MinimumScale = 0
                        'With .AxisTitle
                        '    .Characters.text = "Monate"
                        '    .Font.Size = 8
                        'End With
                    End With

                    With .Axes(Excel.XlAxisType.xlValue)
                        .HasTitle = False
                        If minScale < 0 Then
                            .MinimumScale = System.Math.Round(minScale - 1, mode:=MidpointRounding.ToEven)
                            '.AxisTitle.Position = Excel.XlChartElementPosition.xlChartElementPositionCustom
                            '.AxisTitle.Position = XlConstants.xlBottom
                        Else
                            .MinimumScale = 0
                        End If

                        'Dim hax As Excel.Axis
                        'With hax
                        '    .AxisTitle.Position = Excel.XlChartElementPosition.xlChartElementPositionCustom
                        'End With

                        'With .AxisTitle
                        '    If auswahl = 1 Then
                        '        .Characters.text = "Ressourcen"
                        '    Else
                        '        .Characters.text = "Personalkosten"
                        '    End If
                        '    .Font.Size = 8
                        'End With
                    End With

                    'If auswahl = 2 Then
                    '    .HasLegend = True
                    '    With .Legend
                    '        .Position = XlConstants.xlTop
                    '        .Font.Size = 8
                    '    End With
                    'Else
                    '    .HasLegend = False
                    'End If
                    .HasLegend = True
                    With .Legend
                        .Position = Excel.Constants.xlTop
                        .Font.Size = 8
                    End With
                    .HasTitle = True

                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.font.size = awinSettings.fontsizeTitle
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                ' jetzt kommt die Korrektur der Größe; herausfinden, wieviel Raum die Axis Beschriftung einnimmt ... 
                'With .ChartObjects(anzDiagrams + 1)
                '    .top = top
                '    .height = 2 * height

                '    Dim axleft As Double, axwidth As Double
                '    If .Chart.HasAxis(Excel.XlAxisType.xlValue) = True Then
                '        With .Chart.Axes(Excel.XlAxisType.xlValue)
                '            axleft = .left
                '            axwidth = .width
                '        End With
                '        If left - axwidth < 1 Then
                '            left = 1
                '            width = width + left + 9
                '        Else
                '            left = left - axwidth
                '            width = width + axwidth + 9
                '        End If

                '    End If

                '    .left = left
                '    .width = width


                'End With


                'With .ChartObjects(anzDiagrams + 1)
                '    .top = top
                '    .left = left
                '    .height = 2 * height
                '    .width = width
                'End With

                '
                ' wenn Auswahl = 2 : dann wird ein zweites Diagramm gezeichnet - gewichtete Earned Value plus Risiko Abschlag 
                '
                'If auswahl = 2 Then
                '    diagramTitle = "Earned Value + Risiko Abschlag" & gesamtSumme & " T€" & vbLf & pname
                '    'ReDim earnedValues(1)
                '    ReDim Xdatenreihe(1)

                '    Xdatenreihe(0) = "Earned Values - gewichtet"
                '    Xdatenreihe(1) = "Risiko Abschlag"

                '    'earnedValues(0) = hsum(0)
                '    'earnedValues(1) = hsum(1)


                '    With appInstance.Charts.Add
                '        ' remove extra series

                '        Do Until .SeriesCollection.Count = 0
                '            .SeriesCollection(1).Delete()
                '        Loop

                '        With .SeriesCollection.NewSeries
                '            .name = pname
                '            .Values = hsum
                '            .XValues = Xdatenreihe
                '            .ChartType = Excel.XlChartType.xlPie
                '            .HasDataLabels = True
                '            .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                '        End With

                '        With .SeriesCollection(1).Points(1)
                '            .Interior.color = ergebnisfarbe1
                '            .DataLabel.Font.Size = 10
                '        End With

                '        With .SeriesCollection(1).Points(2)
                '            .Interior.color = ergebnisfarbe2
                '            .DataLabel.Font.Size = 10
                '        End With

                '        .HasLegend = True
                '        With .Legend
                '            .Position = XlConstants.xlTop
                '            .Font.Size = 8
                '        End With
                '        .HasTitle = True
                '        .ChartTitle.text = diagramTitle
                '        .ChartTitle.font.size = 10
                '        .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                '    End With
                '    With .ChartObjects(anzDiagrams + 2)
                '        .top = top
                '        .left = left + width
                '        .height = 10 * boxHeight
                '        .width = 12 * boxWidth
                '    End With

                'Else

                'End If

            End If

            repObj = .ChartObjects(anzDiagrams + 1)

            sumDiagram = New clsDiagramm

            sumChart = New clsEventsPrcCharts
            sumChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            sumDiagram.setDiagramEvent = sumChart


            With sumDiagram
                .DiagrammTitel = diagramTitle
                .diagrammTyp = DiagrammTypen(4)
                '.setCollection = myCollection
                .isCockpitChart = isCockpitChart
            End With

            DiagramList.Add(sumDiagram)
            'sumDiagram = Nothing

        End With

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        If updateScreenwastrue Then
            appInstance.ScreenUpdating = True
        End If

    End Sub

    ''' <summary>
    ''' zeigt die Ergebnis Abschätzung an: Summe Projekt-Ergebnisse - (Mehrkosten durch Überauslastung + Deckungsbeitragspotenzial wg. Unterauslastung) 
    ''' </summary>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <remarks></remarks>
    Sub awinCreateErgebnisDiagramm(ByRef repObj As Object, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                   ByVal isCockpitChart As Boolean, ByVal calledfromReporting As Boolean)

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean
        'Dim plen As Integer
        Dim i As Integer
        Dim minScale As Double
        Dim Xdatenreihe(3) As String
        Dim valueDatenreihe1(3) As Double
        Dim valueDatenreihe2(3) As Double
        Dim itemColor(3) As Object
        Dim itemValue(3) As Double
        Dim earnedValue As Double, additionalCostExt As Double, riskValue As Double, internwithoutProject As Double
        Dim ertragsWert As Double

        Dim mycollection As New Collection
        Dim chtobjName As String

        'Dim hproj As clsProjekt
        Dim ErgebnisListeR As New Collection

        mycollection.Add("Ergebnis")
        chtobjName = getKennung("pf", PTpfdk.Auslastung, mycollection)
        mycollection.Clear()

        If Not calledfromReporting Then

            Dim foundDiagramm As clsDiagramm

            ' wenn die Werte für dieses Diagramm bereits einmal gespeichert wurden ... -> übernehmen 
            Try
                foundDiagramm = DiagramList.getDiagramm(chtobjName)
                With foundDiagramm
                    top = .top
                    left = .left
                    width = .width
                    height = .height
                End With
            Catch ex As Exception


            End Try
        End If


        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False



        Xdatenreihe(0) = "Summe Projekt-Ergebnisse (Risiko-gewichtet)"
        'Xdatenreihe(1) = "Risiko-Abschlag"
        Xdatenreihe(1) = "Mehrkosten wegen Überauslastung"
        Xdatenreihe(2) = "Opportunitätskosten durch Unterauslastung"
        Xdatenreihe(3) = "Ergebnis-Kennzahl"

        Dim positiv As Boolean = True

        ' das ist der Earned Value 
        earnedValue = System.Math.Round(ShowProjekte.getEarnedValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        ' das ist der Risiko Abschlag  
        riskValue = System.Math.Round(ShowProjekte.getWeightedRiskValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10

        itemValue(0) = earnedValue - riskValue
        If itemValue(0) >= 0 Then
            itemColor(0) = ergebnisfarbe1
        Else
            itemColor(0) = farbeExterne
        End If

        Dim currentWert As Double = itemValue(0)


        ' das sind die Zusatzkosten, die durch Externe (wg Überauslastung) verursacht werden
        additionalCostExt = System.Math.Round(ShowProjekte.getadditionalECostinMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        itemValue(1) = additionalCostExt
        itemColor(1) = farbeExterne

        ' das sind die durch Unterauslastung verursachten Kosten , also Personal-Kosten von Leuten, die in keinem Projekt sind
        internwithoutProject = System.Math.Round(ShowProjekte.getCostoValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        itemValue(2) = internwithoutProject
        itemColor(2) = awinSettings.AmpelGelb

        ' das ist der Ertrag 
        ertragsWert = earnedValue - (riskValue + additionalCostExt + internwithoutProject)
        itemValue(3) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(3) = ergebnisfarbe2
        Else
            itemColor(3) = farbeExterne
        End If

        diagramTitle = summentitel1 & " " & textZeitraum(showRangeLeft, showRangeRight)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count

            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found


                If .ChartObjects(i).Name = chtobjName Then
                    found = True
                Else
                    i = i + 1
                End If

            End While



            If found Then
                repObj = .ChartObjects(i)
                'MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                If ertragsWert < 0 Then
                    minScale = System.Math.Round(ertragsWert / 10, mode:=MidpointRounding.ToEven) * 10
                Else
                    minScale = 0
                End If

                'Dim htxt As String
                Dim valueCrossesNull As Boolean = False

                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop
                    Dim crossindex As Integer = -1

                    ' bestimmen des Anfangs  
                    Dim iv = 0
                    valueDatenreihe1(iv) = 0
                    valueDatenreihe2(iv) = itemValue(iv)
                    currentWert = itemValue(iv)
                    Dim formerValue As Double = currentWert
                    Dim negativeFromNull As Boolean = False

                    ' alle nächsten Zwischen-Werte 
                    For iv = 1 To 2
                        If formerValue <= 0 Then
                            negativeFromNull = True
                        Else
                            negativeFromNull = False
                        End If

                        currentWert = currentWert - itemValue(iv)
                        valueCrossesNull = (currentWert + itemValue(iv) > 0) And (currentWert < 0)

                        If currentWert >= 0 Then
                            valueDatenreihe1(iv) = currentWert
                            valueDatenreihe2(iv) = itemValue(iv)
                        ElseIf valueCrossesNull Then
                            valueDatenreihe1(iv) = currentWert
                            valueDatenreihe2(iv) = itemValue(iv) - currentWert * (-1) ' notwendig da currentWert ja negativ ist ..
                            crossindex = iv + 1
                        ElseIf negativeFromNull Then
                            valueDatenreihe1(iv) = formerValue
                            valueDatenreihe2(iv) = itemValue(iv) * (-1)
                        Else
                            valueDatenreihe1(iv) = currentWert
                            valueDatenreihe2(iv) = itemValue(iv) * (-1)
                        End If

                        formerValue = currentWert
                    Next

                    ' bestimmen des Ende 
                    iv = 3
                    valueDatenreihe1(iv) = 0
                    valueDatenreihe2(iv) = itemValue(iv)



                    'series
                    With .SeriesCollection.NewSeries
                        .name = "Bottom"
                        .HasDataLabels = False
                        .Interior.colorindex = -4142
                        .Values = valueDatenreihe1
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                        If crossindex > 0 Then
                            ' es gab einen Übergang , dort muss Bottom auf die entsprechende Farbe gesetzt werden 
                            With .Points(crossindex)
                                .Interior.color = itemColor(crossindex - 1)
                            End With
                        End If

                    End With

                    With .SeriesCollection.NewSeries
                        .name = "Top"
                        .HasDataLabels = True
                        .Values = valueDatenreihe2
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked

                        For iv = 0 To 3

                            With .Points(iv + 1)
                                .HasDataLabel = True
                                .DataLabel.text = Format(itemValue(iv), "###,###0") & " T€"
                                .Interior.color = itemColor(iv)
                                .DataLabel.Font.Size = awinSettings.fontsizeLegend
                                Try
                                    .DataLabel.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove
                                Catch ex As Exception

                                End Try
                            End With

                        Next

                    End With

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = False

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        If minScale < 0 Then
                            .TickLabelPosition = Excel.Constants.xlLow
                        End If
                        '.MinimumScale = 0

                    End With

                    'Dim hax As Excel.Axis
                    'With hax
                    '    .HasMajorGridlines
                    '    .hasminor()
                    'End With

                    Try
                        With .Axes(Excel.XlAxisType.xlValue)
                            .HasTitle = False
                            .HasMajorGridlines = False
                            .hasminorgridlines = False
                            If minScale < 0 Then
                                .MinimumScale = System.Math.Round((minScale - 1) / 10, mode:=MidpointRounding.ToEven) * 10
                            Else
                                .MinimumScale = 0
                            End If
                        End With
                    Catch ex As Exception

                    End Try


                    .HasLegend = False
                    'With .Legend
                    '    .Position = XlConstants.xlTop
                    '    .Font.Size = 8
                    'End With
                    .HasTitle = True

                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.font.size = awinSettings.fontsizeTitle
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .name = chtobjName
                End With

                repObj = .ChartObjects(anzDiagrams + 1)

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then

                    Dim prcDiagram As New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    Dim prcChart As New clsEventsPrcCharts
                    prcChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart
                    prcDiagram.setDiagramEvent = prcChart
                    ' Ende Event Handling für Chart 


                    With prcDiagram
                        .DiagrammTitel = diagramTitle
                        .diagrammTyp = DiagrammTypen(4)
                        .gsCollection = Nothing
                        .isCockpitChart = False
                        .top = top
                        .left = left
                        .width = width
                        .height = height
                        .kennung = chtobjName
                    End With

                    ' eintragen in die sortierte Liste mit .kennung als dem Schlüssel 
                    ' wenn das Diagramm bereits existiert, muss es gelöscht werden, dann neu ergänzt ... 
                    Try
                        DiagramList.Add(prcDiagram)
                    Catch ex As Exception

                        Try
                            DiagramList.Remove(prcDiagram.kennung)
                            DiagramList.Add(prcDiagram)
                        Catch ex1 As Exception

                        End Try


                    End Try

                End If

            End If


        End With

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    Sub awinCreateVerbesserungsPotentialDiagramm(ByRef repObj As Object, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, ByVal isCockpitChart As Boolean)

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean

        Dim i As Integer

        Dim Xdatenreihe(1) As String
        Dim itemColor(1) As Object
        Dim itemValue(1) As Double
        Dim additionalCostExt As Double, internwithoutProject As Double
        Dim chtTitle As String
        Dim mycollection As New Collection
        Dim sumDiagram As clsDiagramm
        Dim sumChart As clsEventsPrcCharts
        Dim ErgebnisListeR As New Collection
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False



        Xdatenreihe(0) = "Mehrkosten wegen Überauslastung"
        Xdatenreihe(1) = "Opportunitätskosten durch Unterauslastung"


        Dim positiv As Boolean = True



        ' das sind die Zusatzkosten, die durch Externe (wg Überauslastung) verursacht werden
        additionalCostExt = System.Math.Round(ShowProjekte.getadditionalECostinMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10

        itemValue(0) = additionalCostExt
        itemColor(0) = awinSettings.AmpelRot

        ' das sind die durch Unterauslastung verursachten Kosten , also Personal-Kosten von Leuten, die in keinem Projekt sind
        internwithoutProject = System.Math.Round(ShowProjekte.getCostoValuesInMonth.Sum / 10, mode:=MidpointRounding.ToEven) * 10
        itemValue(1) = internwithoutProject
        itemColor(1) = awinSettings.AmpelGelb


        diagramTitle = summentitel5 & " (T€) " & vbLf & StartofCalendar.AddMonths(showRangeLeft - 1).ToString("MMM yy") & " - " & _
                                                 StartofCalendar.AddMonths(showRangeRight - 1).ToString("MMM yy")


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


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


                If ((chtTitle Like (diagramTitle & "*")) And _
                         (isCockpitChart = istCockpitDiagramm(.ChartObjects(i)))) Then
                    found = True
                Else
                    i = i + 1
                End If

            End While



            If found Then
                repObj = .ChartObjects(i)
                'MsgBox(" Diagramm wird bereits angezeigt ...")
            Else



                With appInstance.Charts.Add
                    ' remove extra series
                    Do Until .SeriesCollection.Count = 0
                        .SeriesCollection(1).Delete()
                    Loop


                    'series
                    With .SeriesCollection.NewSeries
                        .name = "Potentiale"
                        .HasDataLabels = True
                        .Datalabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                        .Values = itemValue
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnClustered
                        .Points(1).interior.color = itemColor(0)
                        .Points(2).interior.color = itemColor(1)
                    End With

                    .HasAxis(Excel.XlAxisType.xlCategory) = True
                    .HasAxis(Excel.XlAxisType.xlValue) = False

                    With .Axes(Excel.XlAxisType.xlCategory)
                        .HasTitle = False
                        .TickLabelPosition = Excel.Constants.xlLow
                    End With


                    Try
                        With .Axes(Excel.XlAxisType.xlValue)
                            .HasTitle = False
                            .HasMajorGridlines = False
                            .hasminorgridlines = False
                            .MinimumScale = 0
                            .MaximumScale = Round((Max(itemValue(0), itemValue(1)) + 99.9) / 200, mode:=MidpointRounding.ToEven) * 200
                        End With
                    Catch ex As Exception

                    End Try


                    .HasLegend = False
                    .HasTitle = True

                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.font.size = awinSettings.fontsizeTitle
                    .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With

                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .name = diagramTitle

                End With

                repObj = .ChartObjects(anzDiagrams + 1)


            End If

            sumDiagram = New clsDiagramm

            sumChart = New clsEventsPrcCharts
            sumChart.PrcChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            sumDiagram.setDiagramEvent = sumChart


            With sumDiagram
                .DiagrammTitel = diagramTitle
                .diagrammTyp = DiagrammTypen(4)
                '.setCollection = myCollection
                .isCockpitChart = isCockpitChart
            End With

            DiagramList.Add(sumDiagram)
            'sumDiagram = Nothing

        End With

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    '
    ' zeigt für alle Projekte die Bedarfe für die jeweilige Rolle an
    '
    Sub awinShowProjectNeeds1(ByRef mycollection As Collection, type As String)
        Dim formerSU As Boolean = appInstance.ScreenUpdating

        appInstance.ScreenUpdating = False

        ' jetzt alle Shapes unsichtbar machen, die im Zeitraum liegen 

        ' dann die Werte in die Excel Zellen schreiben 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            Call awinShowNeedsofProject1(mycollection, type, kvp.Key)
        Next kvp


        ' jetzt wieder alle Shapes sichtbar machen 

        appInstance.ScreenUpdating = formerSU


    End Sub

    '
    ' zeigt für alle Projekte die Bedarfe für die jeweilige Rolle an
    '
    'Sub awinShowProjectNeeds(ByVal rolle As Integer)
    '    Dim updateScreenWasTrue As Boolean

    '    If appInstance.ScreenUpdating = True Then
    '        updateScreenWasTrue = True
    '        appInstance.ScreenUpdating = False
    '    Else
    '        updateScreenWasTrue = False
    '    End If

    '    appInstance.ScreenUpdating = False

    '    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
    '        Call awinShowNeedsofProject(rolle, kvp.Key)

    '    Next kvp

    '    If updateScreenWasTrue Then
    '        appInstance.ScreenUpdating = True
    '    End If

    'End Sub

    '
    ' löscht für alle Projekte die Bedarfe für die jeweilige Rolle an
    '
    Sub awinNoshowProjectNeeds()
        Dim updateScreenWasTrue As Boolean

        If appInstance.ScreenUpdating = True Then
            updateScreenWasTrue = True
            appInstance.ScreenUpdating = False
        Else
            updateScreenWasTrue = False
        End If

        Call diagramsVisible(False)

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            Call NoshowNeedsofProject(kvp.Key)
        Next kvp

        Call diagramsVisible(True)

        If updateScreenWasTrue Then
            appInstance.ScreenUpdating = True
        End If

    End Sub

    '
    ' zeigt für das gewählte Projekt die Bedarfe für die angegebene Rolle an
    '
    ''' <summary>
    ''' zeigt für das entsprechende Diagramm-Typ und jeweiligen prcname die entsprechenden Werte  
    ''' </summary>
    ''' <param name="mycollection">enthält ggf die zu betrachtende Menge an Werten</param>
    ''' <param name="type">wert aus DiagrammTypen 0..4 </param>
    ''' <param name="projektname">NAme des Projekts aus ShowProjekte</param>
    ''' <remarks></remarks>
    Sub awinShowNeedsofProject1(ByRef mycollection As Collection, ByVal type As String, ByVal projektname As String)

        Dim i As Integer, k As Integer, l As Integer, m As Integer

        Dim tempArray() As Double
        Dim pname As String = " "
        'Dim showKostenart As Boolean
        Dim hproj As New clsProjekt
        Dim sfarbe As Object
        Dim sgroesse As Double
        'Dim prcName As String
        'Dim itemName As String
        Dim persCost As String = CostDefinitions.getCostdef(CostDefinitions.Count).name
        Dim shpelement As Excel.Shape
        Dim tmpshapes As Excel.Shapes = appInstance.ActiveSheet.shapes


        Try
            hproj = ShowProjekte.getProject(projektname)
        Catch ex As Exception
            Call MsgBox("Projekt nicht gefunden (in ShowNeedsofProject): " & projektname)
            Exit Sub
        End Try



        Try
            shpelement = tmpshapes.Item(projektname)
            With shpelement
                .Fill.Transparency = 0.8
                '.Shadow.Transparency = 0.8
                .TextFrame2.TextRange.Text = ""
            End With

        Catch ex As Exception

        End Try

        If Not hproj Is Nothing Then
            With hproj
                sfarbe = RGB(0, 0, 0) '.Schriftfarbe
                sgroesse = .Schrift
                ' in L steht jetzt die Lä nge
                l = .Dauer
                i = .tfZeile + 1
                k = .tfspalte
            End With

            ReDim tempArray(l - 1)

            tempArray = hproj.getBedarfeInMonths(mycollection, type)

            Dim formerEE = appInstance.EnableEvents
            appInstance.EnableEvents = False

            ' hier muss jetzt tempArray gesetzt werden

            With appInstance.Worksheets(arrWsNames(3))

                'Call diagramsVisible("False")
                '
                ' jetzt in Planungs-Horizont eintragen
                '
                ' vorher noch den Projektnamen in das Kommentar Feld des ersten Feldes schreiben

                'If istInTimezone(k) Then
                '    ' Projektname in Kommentar wegschreiben
                '    If .Cells(i, k).Comment Is Nothing Then
                '        .Cells(i, k).AddComment(projektname)
                '    Else
                '        Try
                '            .Cells(i, k).ClearComments()
                '            .Cells(i, k).AddComment(projektname)
                '        Catch ex As Exception

                '        End Try

                '    End If
                'End If

                For m = 1 To l
                    If tempArray(m - 1) > 0 And istInTimezone(k + m - 1) Then
                        .Cells(i, k).Offset(0, m - 1).Value = tempArray(m - 1)
                    End If
                Next m

                Dim tmpgroesse As Integer
                If tempArray.Max > 999 Or tempArray.Min < -999 Then
                    tmpgroesse = sgroesse - 2
                ElseIf tempArray.Max > 9999 Or tempArray.Min < -9999 Then
                    tmpgroesse = sgroesse - 4
                Else
                    tmpgroesse = sgroesse
                End If
                .range(.Cells(i, k), .cells(i, k).Offset(0, l - 1)).font.color = sfarbe
                .range(.Cells(i, k), .cells(i, k).Offset(0, l - 1)).font.size = tmpgroesse
            End With

            appInstance.EnableEvents = formerEE

        End If



    End Sub

    '
    ' zeigt für das gewählte Projekt die Bedarfe für die angegebene Rolle an
    '
    'Sub awinShowNeedsofProject(ByVal rolle As Integer, ByVal projektname As String)

    '    Dim i As Integer, k As Integer, l As Integer, m As Integer, rk As Integer

    '    Dim tempArray() As Double
    '    Dim pname As String = " "
    '    Dim showKostenart As Boolean
    '    Dim hproj As New clsProjekt
    '    Dim sfarbe As Object
    '    Dim sgroesse As Double

    '    Dim tmpshapes As Excel.Shapes = appInstance.ActiveSheet.shapes
    '    Dim shpelement As Excel.Shape

    '    Try
    '        hproj = ShowProjekte.getProject(projektname)
    '    Catch ex As Exception
    '        Call MsgBox("Projekt nicht gefunden (in ShowNeedsofProject): " & projektname)
    '        Exit Sub
    '    End Try

    '    Try

    '        shpelement = tmpshapes.Item(projektname)
    '        With shpelement
    '            .TextFrame2.TextRange.Text = ""
    '            .Fill.Transparency = 0.8
    '        End With

    '    Catch ex As Exception

    '    End Try


    '    If Not hproj Is Nothing Then
    '        With hproj
    '            sfarbe = RGB(0, 0, 0) '.Schriftfarbe
    '            sgroesse = .Schrift
    '            ' in L steht jetzt die Länge
    '            l = .Dauer
    '            i = .tfZeile
    '            k = .tfSpalte
    '        End With

    '        If rolle > RoleDefinitions.Count Then
    '            showKostenart = True
    '            rk = rolle - RoleDefinitions.Count

    '        Else
    '            showKostenart = False
    '            rk = rolle
    '        End If


    '        ReDim tempArray(l - 1)

    '        '
    '        ' jetzt werden die Bedarfe für das angebene Projekt und die angebene Rolle gesucht
    '        '
    '        If showKostenart Then
    '            If rk < CostDefinitions.Count Then
    '                tempArray = hproj.getKostenBedarf(rk)
    '            Else
    '                tempArray = hproj.getAllPersonalKosten
    '            End If
    '        Else
    '            tempArray = hproj.getRessourcenBedarf(rk)
    '        End If

    '        appInstance.EnableEvents = False

    '        ' hier  muss jetzt tempArray gesetzt werden

    '        With appInstance.Worksheets(arrWsNames(3))

    '            'Call diagramsVisible("False")
    '            '
    '            ' jetzt in Planungs-Horizont eintragen
    '            '
    '            ' vorher noch den Projektnamen in das Kommentar Feld des ersten Feldes schreiben

    '            If istInTimezone(k) Then
    '                ' Projektname in Kommentar wegschreiben
    '                If .Cells(i, k).Comment Is Nothing Then
    '                    .Cells(i, k).AddComment(projektname)
    '                Else
    '                    .Cells(i, k).ClearComments()
    '                    .Cells(i, k).AddComment(projektname)
    '                End If
    '            End If

    '            For m = 1 To l
    '                If tempArray(m - 1) > 0 And istInTimezone(k + m - 1) Then
    '                    .Cells(i, k).Offset(0, m - 1).Value = tempArray(m - 1)
    '                End If
    '            Next m
    '            .range(.Cells(i, k), .cells(i, k).Offset(0, l - 1)).font.color = sfarbe
    '            .range(.Cells(i, k), .cells(i, k).Offset(0, l - 1)).font.size = sgroesse
    '        End With

    '        appInstance.EnableEvents = True

    '    End If


    'End Sub

    '
    ' löscht für das gewählte Projekt die Bedarfe für die angegebene Rolle
    '
    Sub NoshowNeedsofProject(ByVal projektname As String)
        Dim hproj As clsProjekt
        Dim sfarbe As Object
        Dim sgroesse As Double
        Dim i As Integer, k As Integer, l As Integer, m As Integer
        Dim shpelement As Excel.Shape
        Dim worksheetShapes As Excel.Shapes
        Dim pStatus As String


        Try

            worksheetShapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes

        Catch ex As Exception
            Throw New Exception("in NoshowNeedsofProject: keine Shapes Zuordnung möglich ")
        End Try


        Try
            hproj = ShowProjekte.getProject(projektname)
            pStatus = hproj.Status
        Catch ex As Exception
            Call MsgBox("Projekt nicht gefunden (in NoShowNeedsofProject): " & projektname)
            Exit Sub
        End Try


        Try
            'tmpshapes = CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Shapes
            shpelement = worksheetShapes.Item(projektname)
            With shpelement

                Try
                    If .GroupItems.Count > 1 Then
                        .GroupItems(1).TextFrame2.TextRange.Text = projektname
                        For i = 1 To .GroupItems.Count
                            If pStatus = ProjektStatus(0) Then
                                .GroupItems(i - 1).Fill.Transparency = 0.35
                            Else
                                .GroupItems(i - 1).Fill.Transparency = 0.0
                            End If
                        Next
                    Else
                        .TextFrame2.TextRange.Text = projektname
                        If pStatus = ProjektStatus(0) Then
                            .Fill.Transparency = 0.35
                        Else
                            .Fill.Transparency = 0.0
                        End If
                    End If

                Catch ex1 As Exception

                    .TextFrame2.TextRange.Text = projektname
                    If pStatus = ProjektStatus(0) Then
                        .Fill.Transparency = 0.35
                    Else
                        .Fill.Transparency = 0.0
                    End If

                End Try

                
                '.Shadow.Transparency = 0.0
            End With

        Catch ex As Exception

        End Try

        ' jetzt muss das Shape wieder auf "ohne Transparenz" gesetzt werden 

        If Not hproj Is Nothing Then
            With hproj
                sfarbe = RGB(0, 0, 0) '.Schriftfarbe
                sgroesse = .Schrift
                ' in L steht jetzt die Länge
                l = .Dauer
                i = .tfZeile + 1
                k = .tfspalte
            End With

            With appInstance.Worksheets(arrWsNames(3))

                appInstance.EnableEvents = False

                For m = 1 To l
                    If istInTimezone(k + m - 1) Then
                        .Cells(i, k).Offset(0, m - 1).Value = ""
                    End If
                Next m

                ' jetzt den Projektnamen in das erste Feld schreiben
                '.Cells(i, k).Value = projektname
                '.cells(i, k).font.size = sgroesse
                '.cells(i, k).font.color = RGB(0, 0, 0) ' sfarbe

                'With .Cells(i, k)

                '    Try
                '        .ClearComments()
                '    Catch ex As Exception

                '    End Try

                'End With

                appInstance.EnableEvents = True

            End With


        End If

    End Sub


    Function pnameInComment(ByVal c As Excel.Range, ByVal projektname As String) As Boolean
        Dim empty_str As String
        appInstance.EnableEvents = False

        empty_str = ""
        With c
            If .Comment Is Nothing Then
                pnameInComment = False
            ElseIf .Comment.Text = projektname Then
                pnameInComment = True
            Else
                pnameInComment = False
            End If
        End With

        appInstance.EnableEvents = True

    End Function
    '
    ' Funktion prüft , ob die Spalte angezeigt werden muss, also ob sie in der Time Zone enthalten ist
    '
    Function istInTimezone(ByVal spalte As Integer) As Boolean

        If spalte >= showRangeLeft And spalte <= showRangeRight Then
            istInTimezone = True
        Else
            istInTimezone = False
        End If

    End Function

    Function istBereichInTimezone(ByVal anfang As Integer, ByVal ende As Integer) As Boolean


        If ((ende) < showRangeLeft) Or (anfang > showRangeRight) Then
            istBereichInTimezone = False
        Else
            istBereichInTimezone = True
        End If


    End Function



    Sub diagramsVisible(ByVal show As Boolean)

        Dim anzDiagrams As Integer
        With appInstance.Worksheets(arrWsNames(3))

            anzDiagrams = .ChartObjects.Count

            For i = 1 To anzDiagrams
                .ChartObjects(i).Visible = show
            Next i
        End With

    End Sub



    '
    ' zeichnet alle dargestellten Diagramme neu
    '
    Sub awinNeuZeichnenDiagramme(ByVal typus As Integer)
        Dim anz_diagrams As Integer
        Dim chtobj As ChartObject
        Dim i As Integer, r As Integer, p As Integer, k As Integer, e As Integer


        ' typus:
        ' 1 - verschieben
        ' 2 - einfügen
        ' 3 - löschen
        ' 4 - betrachteten Zeitraum ändern
        ' 5 - Stammdaten ändern
        ' 6 - Ressourcen-Bedarfe, Kapas ändern
        ' 7 - Kosten-Bedarfe , Budgets ändern
        '




        With appInstance.Worksheets(arrWsNames(3))
            anz_diagrams = .ChartObjects.Count
            For i = 1 To anz_diagrams
                chtobj = .ChartObjects(i)
                Select Case typus
                    '
                    '
                    Case 1 ' Projekt wurde verschoben
                        If istRollenDiagramm(chtobj, r) Or istKostenartDiagramm(chtobj, k) Or _
                            istPhasenDiagramm(chtobj, p) Or istErgebnisDiagramm(chtobj, e) Then

                            Call awinUpdateprcCollectionDiagram(chtobj)


                        ElseIf istSummenDiagramm(chtobj, p) Then

                            If p = 1 Then
                                Call awinUpdateErgebnisDiagramm(chtobj)
                            ElseIf p = 2 Then
                                Call awinUpdatePortfolioDiagrams(chtobj)
                            ElseIf p = 4 Then
                                Call awinUpdatePersCostStructureDiagramm(chtobj)
                            ElseIf p = 5 Then
                                Call awinUpdateEffizienzDiagramm2(chtobj)
                            ElseIf p = 6 Or p = 7 Or p = 8 Then
                                Try
                                    Call awinUpdateColorDistributionDiagramm(chtobj)
                                Catch ex As Exception

                                End Try

                            ElseIf p = 9 Then
                                Try
                                    Call awinUpdateAuslastungsDiagramm(chtobj)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 10 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 1)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 11 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 2)
                                Catch ex As Exception

                                End Try

                            End If


                            'ElseIf istKostenartDiagramm(chtobj, k) Then ' muss ggf das mini-Chart aktualisiert werden
                            '    If chtobj.width <= miniWidth * 1.05 Then
                            '        Call awin_aktualisiereMiniChart(chtobj)
                            '    End If

                            'ElseIf istPhasenDiagramm(chtobj, p) Then

                            '    Set phaseCollection = New clsPhasen
                            '    phaseCollection.Add PhaseDefinitions.getPhaseDef(p)
                            '    Call awinShowPhaseCollectionDiagram(phaseCollection)
                            '    Set phaseCollection = Nothing


                        ElseIf istPortfolioDiagramm(chtobj, p) Then
                            ' nichts


                        Else ' ist Projekt-Charakteristik Diagramm

                        End If
                        '
                        '
                    Case 2 ' Projekt wurde eingefügt
                        '
                        If istRollenDiagramm(chtobj, r) Or istKostenartDiagramm(chtobj, k) Or _
                            istPhasenDiagramm(chtobj, p) Or istErgebnisDiagramm(chtobj, e) Then

                            Call awinUpdateprcCollectionDiagram(chtobj)


                        ElseIf istSummenDiagramm(chtobj, p) Then

                            If p = 1 Then
                                Call awinUpdateErgebnisDiagramm(chtobj)

                            ElseIf p = 2 Then
                                Call awinUpdatePortfolioDiagrams(chtobj)

                            ElseIf p = 4 Then
                                Call awinUpdatePersCostStructureDiagramm(chtobj)

                            ElseIf p = 5 Then
                                Call awinUpdateEffizienzDiagramm2(chtobj)

                            ElseIf p = 6 Or p = 7 Or p = 8 Then
                                Try
                                    Call awinUpdateColorDistributionDiagramm(chtobj)
                                Catch ex As Exception

                                End Try

                            ElseIf p = 9 Then
                                Try
                                    Call awinUpdateAuslastungsDiagramm(chtobj)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 10 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 1)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 11 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 2)
                                Catch ex As Exception

                                End Try
                            End If


                        ElseIf istPortfolioDiagramm(chtobj, p) Then

                            Call awinUpdatePortfolioDiagrams(chtobj)


                        Else ' ist Projekt-Charakteristik Diagramm
                        End If

                    Case 3 ' Projekt wurde gelöscht
                        If istRollenDiagramm(chtobj, r) Or istKostenartDiagramm(chtobj, k) Or _
                            istPhasenDiagramm(chtobj, p) Or istErgebnisDiagramm(chtobj, e) Then

                            Call awinUpdateprcCollectionDiagram(chtobj)


                        ElseIf istSummenDiagramm(chtobj, p) Then

                            If p = 1 Then
                                Call awinUpdateErgebnisDiagramm(chtobj)
                            ElseIf p = 2 Then
                                Call awinUpdatePortfolioDiagrams(chtobj)
                            ElseIf p = 4 Then
                                Call awinUpdatePersCostStructureDiagramm(chtobj)
                            ElseIf p = 5 Then
                                Call awinUpdateEffizienzDiagramm2(chtobj)
                            ElseIf p = 6 Or p = 7 Or p = 8 Then
                                Try
                                    Call awinUpdateColorDistributionDiagramm(chtobj)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 9 Then
                                Try
                                    Call awinUpdateAuslastungsDiagramm(chtobj)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 10 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 1)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 11 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 2)
                                Catch ex As Exception

                                End Try
                            End If


                        ElseIf istPortfolioDiagramm(chtobj, p) Then

                            Call awinUpdatePortfolioDiagrams(chtobj)


                        Else ' ist Projekt-Charakteristik Diagramm
                        End If

                    Case 4 ' betrachteter Zeitraum wurde geändert
                        If istRollenDiagramm(chtobj, r) Or istKostenartDiagramm(chtobj, k) Or _
                            istPhasenDiagramm(chtobj, p) Or istErgebnisDiagramm(chtobj, e) Then

                            Call awinUpdateprcCollectionDiagram(chtobj)

                        ElseIf istSummenDiagramm(chtobj, p) Then

                            If p = 1 Then
                                Call awinUpdateErgebnisDiagramm(chtobj)
                            ElseIf p = 2 Then
                                Call awinUpdatePortfolioDiagrams(chtobj)
                            ElseIf p = 4 Then
                                Call awinUpdatePersCostStructureDiagramm(chtobj)
                            ElseIf p = 5 Then
                                Call awinUpdateEffizienzDiagramm2(chtobj)
                            ElseIf p = 6 Or p = 7 Or p = 8 Then
                                Try
                                    Call awinUpdateColorDistributionDiagramm(chtobj)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 9 Then
                                Try
                                    Call awinUpdateAuslastungsDiagramm(chtobj)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 10 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 1)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 11 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 2)
                                Catch ex As Exception

                                End Try
                            End If

                        ElseIf istPortfolioDiagramm(chtobj, p) Then



                        Else ' ist Projekt-Charakteristik Diagramm
                        End If

                    Case 5 ' Stammdaten wurden geändert
                        If istRollenDiagramm(chtobj, r) Then


                        ElseIf istKostenartDiagramm(chtobj, k) Then


                        ElseIf istSummenDiagramm(chtobj, p) Then

                            If p = 1 Then
                                Call awinUpdateErgebnisDiagramm(chtobj)
                            ElseIf p = 2 Then
                                Call awinUpdatePortfolioDiagrams(chtobj)
                            ElseIf p = 4 Then
                                Call awinUpdatePersCostStructureDiagramm(chtobj)
                            ElseIf p = 5 Then
                                Call awinUpdateEffizienzDiagramm2(chtobj)
                            ElseIf p = 9 Then
                                Try
                                    Call awinUpdateAuslastungsDiagramm(chtobj)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 10 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 1)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 11 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 2)
                                Catch ex As Exception

                                End Try
                            End If



                        ElseIf istPortfolioDiagramm(chtobj, p) Then

                            Call awinUpdatePortfolioDiagrams(chtobj)

                        ElseIf istErgebnisDiagramm(chtobj, e) Then

                            Call awinUpdateprcCollectionDiagram(chtobj)

                        Else ' ist Projekt-Charakteristik Diagramm
                        End If

                    Case 6 ' Ressourcen Bedarf eines existierenden Projektes wurde geändert

                        If istRollenDiagramm(chtobj, r) Or istKostenartDiagramm(chtobj, k) Or _
                            istPhasenDiagramm(chtobj, p) Or istErgebnisDiagramm(chtobj, e) Then

                            Call awinUpdateprcCollectionDiagram(chtobj)

                        ElseIf istSummenDiagramm(chtobj, p) Then

                            If p = 1 Then
                                Call awinUpdateErgebnisDiagramm(chtobj)
                            ElseIf p = 2 Then
                                Call awinUpdatePortfolioDiagrams(chtobj)
                            ElseIf p = 4 Then
                                Call awinUpdatePersCostStructureDiagramm(chtobj)
                            ElseIf p = 5 Then
                                Call awinUpdateEffizienzDiagramm2(chtobj)
                            ElseIf p = 9 Then
                                Try
                                    Call awinUpdateAuslastungsDiagramm(chtobj)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 10 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 1)
                                Catch ex As Exception

                                End Try
                            ElseIf p = 11 Then
                                Try
                                    Call updateAuslastungsDetailPie(chtobj, 2)
                                Catch ex As Exception

                                End Try
                            End If



                        ElseIf istPortfolioDiagramm(chtobj, p) Then

                            Call awinUpdatePortfolioDiagrams(chtobj)


                        Else ' ist Projekt-Charakteristik Diagramm
                        End If

                    Case 7 ' Kosten Bedarf eines existierenden Projektes wurde geändert

                        If istRollenDiagramm(chtobj, r) Or istKostenartDiagramm(chtobj, k) Or _
                            istPhasenDiagramm(chtobj, p) Or istErgebnisDiagramm(chtobj, e) Then

                            Call awinUpdateprcCollectionDiagram(chtobj)

                        ElseIf istSummenDiagramm(chtobj, p) Then

                            If p = 1 Then
                                Call awinUpdateErgebnisDiagramm(chtobj)
                            ElseIf p = 2 Then
                                Call awinUpdatePortfolioDiagrams(chtobj)
                            ElseIf p = 4 Then
                                Call awinUpdatePersCostStructureDiagramm(chtobj)
                            ElseIf p = 5 Then
                                Call awinUpdateEffizienzDiagramm2(chtobj)
                            End If


                        ElseIf istPortfolioDiagramm(chtobj, p) Then

                            Call awinUpdatePortfolioDiagrams(chtobj)


                        Else ' ist Projekt-Charakteristik Diagramm oder Phasen Diagramm
                        End If

                    Case 8 ' Selection hat sich geändert 

                        If istRollenDiagramm(chtobj, r) Or istKostenartDiagramm(chtobj, k) Or _
                            istPhasenDiagramm(chtobj, p) Or istErgebnisDiagramm(chtobj, e) Then

                            Call awinUpdateprcCollectionDiagram(chtobj)

                        End If

                End Select

            Next i

        End With



    End Sub

    ''' <summary>
    ''' stellt das Fenster "Projekt Tafel" so ein, daß die gesamte Zeitleiste zu sehen ist und evtl das Diagramm
    ''' </summary>
    ''' <remarks></remarks>
    Sub awinScrollintoView()
        Dim ScrollColumn As Integer
        Dim zoom As Double
        Dim minWindowWidth As Double, minWindowHeight As Double

       
        Try
            appInstance.ActiveWorkbook.Windows(windowNames(5)).Activate()
        Catch ex As Exception
            Call MsgBox("Window " & windowNames(5) & " existiert nicht mehr !")
            Exit Sub
        End Try



        ScrollColumn = showRangeLeft - 12 ' war vorher 6
        If ScrollColumn <= 0 Then
            ScrollColumn = 1
        End If


        minWindowWidth = Max(boxWidth * (showRangeRight - showRangeLeft + 1 + 12), 60 * boxWidth)
        minWindowHeight = Max(WertfuerTop() + 30, 22 * boxHeight + 30)


        Dim shp As Excel.Shape
        For Each shp In appInstance.ActiveSheet.Shapes
            With shp
                If .BottomRightCell.Top > minWindowHeight And .BottomRightCell.Top < WertfuerTop() * boxHeight Then
                    minWindowHeight = .BottomRightCell.Top + 3 * boxHeight
                End If
                If .BottomRightCell.Left - (showRangeLeft - 6) * boxWidth > minWindowWidth Then
                    minWindowWidth = .BottomRightCell.Left + 3 * boxWidth - (showRangeLeft - 6) * boxWidth
                End If
            End With
        Next shp


        With appInstance.ActiveWindow
            If .UsableWidth / minWindowWidth < .UsableHeight / minWindowHeight Then
                ' Zoom an Breite orientieren ...
                Try
                    zoom = 100 * .UsableWidth / minWindowWidth
                    .Zoom = Min(zoom, 120)
                    If .Zoom < 60 Then
                        .Zoom = 60
                    End If
                Catch ex As Exception
                    If zoom < 20 Then
                        .Zoom = 20
                    ElseIf zoom > 400 Then
                        .Zoom = 400
                    Else
                        .Zoom = 100
                    End If
                End Try

            Else
                ' Zoom an Höhe orientieren 
                Try
                    zoom = 100 * .UsableHeight / minWindowHeight
                    .Zoom = Min(zoom, 120)
                    If .Zoom < 60 Then
                        .Zoom = 60
                    End If
                Catch ex As Exception
                    If zoom < 20 Then
                        .Zoom = 20
                    ElseIf zoom > 400 Then
                        .Zoom = 400
                    Else
                        .Zoom = 100
                    End If
                End Try

            End If
            If Abs(ScrollColumn - .ScrollColumn) > 2 Then
                .ScrollColumn = ScrollColumn
            End If
            .ScrollRow = 1
        End With


    End Sub


End Module
