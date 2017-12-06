Imports System.Math
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core

Public Module awinDiagrams

    '
    ' zeigt im Planungshorizont die Time Zone an - oder blendet sie aus, abhängig vom Wert showzone
    '
    Sub awinShowtimezone(ByVal von As Integer, ByVal bis As Integer, ByVal showzone As Boolean)
        Dim laenge As Integer

        laenge = bis - von

        If von > 0 And laenge > 0 Then

            With appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT))

                If showzone Then
                    '
                    ' die erste Zeile im Bereich einfärben
                    '
                    .Range(.Cells(1, von), .Cells(1, von).Offset(0, laenge)).Interior.color = showtimezone_color
                    If awinSettings.showTimeSpanInPT Then
                        .Range(.Cells(2, von), .Cells(5000, von).Offset(0, laenge)).Interior.color = awinSettings.timeSpanColor
                    End If

                Else
                    '
                    ' die erste Zeile im Bereich einfärben
                    '
                    .Range(.Cells(1, von), .Cells(1, von).Offset(0, laenge)).Interior.color = noshowtimezone_color
                    If awinSettings.showTimeSpanInPT Then
                        .Range(.Cells(2, von), .Cells(5000, von).Offset(0, laenge)).Interior.colorindex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone
                    End If

                End If

            End With

            visboZustaende.showTimeZoneBalken = False

        End If

    End Sub

    '
    ' zeigt im selektierten Zeitraum den Monat an, der gerade in einem Chart angeklickt wurde, so dass dass die 
    ' dort liegenden Elemente gezeigt werden 
    '
    Sub awinShowSelectedMonth(ByVal mon As Integer)
        Dim laenge As Integer
        Dim von As Integer = showRangeLeft
        Dim bis As Integer = showRangeRight

        Dim lastZeile As Integer = projectboardShapes.getMaxZeile

        If showRangeLeft = 0 Or showRangeRight = 0 Or showRangeLeft > showRangeRight Then
            Exit Sub
        End If

        laenge = showRangeRight - showRangeLeft

        If mon >= showRangeLeft And mon <= showRangeRight Then

            With appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT))

                '
                ' erst den Bereich einfärben  
                '
                .Range(.Cells(1, von), .Cells(1, von).Offset(0, laenge)).Interior.color = showtimezone_color
                If awinSettings.showTimeSpanInPT Then
                    .Range(.Cells(2, von), .Cells(5000, von).Offset(0, laenge)).Interior.color = awinSettings.timeSpanColor
                    .range(.cells(2, mon), .cells(lastZeile, mon)).interior.color = awinSettings.glowColor
                End If



            End With

            visboZustaende.showTimeZoneBalken = True

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
            Test = appInstance.Workbooks.Item(myProjektTafel).Windows(windowNames(prctyp))
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
    ''' 5 - Meilensteine 
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <param name="prcTyp"></param>
    ''' <remarks>myCollection am 23.5 per byval übergeben, damit im Falle der Rollen myCollection ausgeweitet werden kann ...</remarks>
    Sub awinCreateprcCollectionDiagram(ByVal myCollection As Collection, ByRef repObj As Excel.ChartObject, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                       ByVal isCockpitChart As Boolean, ByVal prcTyp As String, ByVal calledfromReporting As Boolean)

        Dim von As Integer, bis As Integer

        Dim anzDiagrams As Integer, i As Integer, m As Integer, r As Integer

        'Dim korr_abstand As Double
        Dim minwert As Double, maxwert As Double
        Dim nr_pts As Integer
        Dim diagramTitle As String = ""
        Dim objektFarbe As Object
        Dim ampelfarbe(3) As Long
        Dim Xdatenreihe() As String
        Dim datenreihe() As Double, edatenreihe() As Double, seriesSumDatenreihe() As Double
        Dim kdatenreihe() As Double ' nimmt die Kapa-Werte für das Diagramm auf
        Dim kdatenreihePlus() As Double ' nimmt die Kapa Werte inkl bereits beauftragter externer Ressourcen auf 
        Dim msdatenreihe(,) As Double
        Dim prcName As String = ""
        Dim startdate As Date
        Dim diff As Integer
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
        Dim breadcrumb As String = ""
        Dim newChtObj As Excel.ChartObject = Nothing


        Dim currentSheetName As String
        Dim found As Boolean = False

        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            If calledfromReporting Then
                currentSheetName = arrWsNames(ptTables.repCharts)
            Else
                currentSheetName = arrWsNames(ptTables.mptPfCharts)
            End If

        Else
            currentSheetName = arrWsNames(ptTables.meCharts)
        End If

       

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
        ReDim kdatenreihePlus(bis - von)
        ReDim seriesSumDatenreihe(bis - von)
        ReDim VarValues(bis - von)
        ReDim msdatenreihe(3, bis - von)



        If myCollection.Count = 0 Then
            'Call MsgBox("keine Phase / Rolle / Kostenart / Ergebnisart / Meilenstein ausgewählt ...")
            Call MsgBox(repMessages.getmsg(112))
            Exit Sub
        End If

        diff = -1
        startdate = StartofCalendar.AddMonths(diff)

        For m = von To bis
            Xdatenreihe(m - von) = startdate.AddMonths(m).ToString("MMM yy", repCult)
        Next m


        titleZeitraum = Xdatenreihe(0) & " - " & Xdatenreihe(bis - von)


        If prcTyp = DiagrammTypen(0) Then


            chtobjName = calcChartKennung("pf", PTpfdk.Phasen, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Phasen)
            Else
                diagramTitle = splitHryFullnameTo1(CStr(myCollection.Item(1)))
            End If

        ElseIf prcTyp = DiagrammTypen(1) Then

            chtobjName = calcChartKennung("pf", PTpfdk.Rollen, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Rollen)
            Else
                diagramTitle = CStr(myCollection.Item(1))
            End If

        ElseIf prcTyp = DiagrammTypen(2) Then
            chtobjName = calcChartKennung("pf", PTpfdk.Kosten, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Kosten)
            Else
                diagramTitle = CStr(myCollection.Item(1))
            End If


        ElseIf prcTyp = DiagrammTypen(4) Then
            'chtobjName = "Ergebnis-Übersicht"
            'diagramTitle = "Ergebnis-Übersicht"
            chtobjName = repMessages.getmsg(113)
            diagramTitle = repMessages.getmsg(113)


        ElseIf prcTyp = DiagrammTypen(5) Then
            chtobjName = calcChartKennung("pf", PTpfdk.Meilenstein, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Meilenstein)
            Else
                diagramTitle = splitHryFullnameTo1(CStr(myCollection.Item(1)))
            End If

        ElseIf prcTyp = DiagrammTypen(7) Then
            ' Phasen Kategorien
            chtobjName = calcChartKennung("pf", PTpfdk.PhaseCategories, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = "Phase-Categories"
            Else
                diagramTitle = "Category " & splitHryFullnameTo1(CStr(myCollection.Item(1)))
            End If

        ElseIf prcTyp = DiagrammTypen(8) Then
            ' Meilenstein-Kategorien
            chtobjName = calcChartKennung("pf", PTpfdk.MilestoneCategories, myCollection)

            If myCollection.Count > 1 Then
                diagramTitle = "Milestone-Categories"
            Else
                diagramTitle = "Category " & splitHryFullnameTo1(CStr(myCollection.Item(1)))
            End If

        Else
            chtobjName = repMessages.getmsg(114)
            diagramTitle = repMessages.getmsg(114)
        End If

        ' jetzt den Namen aus optischen Gründen ändern 
        If diagramTitle.Contains("#") Then
            diagramTitle = diagramTitle.Replace("#", "-")
        End If


        If prcTyp = DiagrammTypen(1) Then
            kdatenreihe = ShowProjekte.getRoleKapasInMonth(myCollection, False)
            kdatenreihePlus = ShowProjekte.getRoleKapasInMonth(myCollection, True)
        ElseIf prcTyp = DiagrammTypen(0) Then
            kdatenreihe = ShowProjekte.getPhaseSchwellWerteInMonth(myCollection)
        ElseIf prcTyp = DiagrammTypen(5) Then
            kdatenreihe = ShowProjekte.getMilestoneSchwellWerteInMonth(myCollection)
        End If


        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False



        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentSheetName), Excel.Worksheet)

            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count

            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            While i <= anzDiagrams And Not found

                If .ChartObjects(i).name = chtobjName Then
                    found = True
                    repObj = CType(.ChartObjects(i), Excel.ChartObject)
                Else
                    i = i + 1
                End If

            End While

            If Not found Then

                newChtObj = CType(.ChartObjects, Excel.ChartObjects).Add(left, top, width, height)

                With newChtObj.Chart


                    If Not isCockpitChart Then
                        With .Axes(Excel.XlAxisType.xlValue)
                            .MinorUnit = 1
                        End With

                    End If

                    ' remove old series
                    Try
                        Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                        Do While anz > 0
                            .SeriesCollection(1).Delete()
                            anz = anz - 1
                        Loop
                    Catch ex As Exception

                    End Try

                    ' wird benötigt, um zu entscheiden, ob es sich um eine SammelRolle handelt ... 
                    Dim sumRoleShowsPlaceHolderAndAssigned As Boolean
                    Dim pvName As String = ""
                    Dim type As Integer = -1
                    For r = 1 To myCollection.Count

                        pvName = ""
                        type = -1
                        sumRoleShowsPlaceHolderAndAssigned = False


                        If prcTyp = DiagrammTypen(0) Or prcTyp = DiagrammTypen(5) Then
                            ' Phasen oder Meilensteine ..
                            Call splitHryFullnameTo2(CStr(myCollection.Item(r)), prcName, breadcrumb, type, pvName)

                        ElseIf prcTyp = DiagrammTypen(7) Or prcTyp = DiagrammTypen(8) Then
                            Call splitHryFullnameTo2(CStr(myCollection.Item(r)), prcName, breadcrumb, type, pvName)
                            prcName = pvName ' der Name der Kategorie steht hier im pvName 

                        Else
                            prcName = CStr(myCollection.Item(r))
                        End If


                        If prcTyp = DiagrammTypen(0) Then
                            ' Phasen ...
                            einheit = " "
                            Dim tmpPhaseDef As clsPhasenDefinition = PhaseDefinitions.getPhaseDef(prcName)
                            If IsNothing(tmpPhaseDef) Then
                                If appearanceDefinitions.ContainsKey("Phasen Default") Then
                                    objektFarbe = appearanceDefinitions.Item("Phasen Default").form.Fill.ForeColor.RGB
                                Else
                                    objektFarbe = awinSettings.AmpelNichtBewertet
                                End If

                            Else
                                objektFarbe = tmpPhaseDef.farbe
                            End If

                            datenreihe = ShowProjekte.getCountPhasesInMonth(prcName, breadcrumb, type, pvName)

                        ElseIf prcTyp = DiagrammTypen(7) Then
                            ' Phasen-Kategorie 
                            einheit = " "

                            If appearanceDefinitions.ContainsKey(prcName) Then
                                objektFarbe = appearanceDefinitions.Item(prcName).form.Fill.ForeColor.RGB
                            Else
                                objektFarbe = awinSettings.AmpelNichtBewertet
                            End If

                            datenreihe = ShowProjekte.getCountPhaseCategoriesInMonth(prcName)

                        ElseIf prcTyp = DiagrammTypen(1) Then
                            ' Rollen 
                            einheit = " " & awinSettings.kapaEinheit
                            Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(prcName)
                            objektFarbe = tmpRole.farbe

                            If tmpRole.isCombinedRole Then
                                If awinSettings.showPlaceholderAndAssigned Then
                                    sumRoleShowsPlaceHolderAndAssigned = True
                                    datenreihe = ShowProjekte.getRoleValuesInMonth(roleID:=prcName, _
                                                                                   considerAllSubRoles:=True, _
                                                                                   type:=PTcbr.placeholders, _
                                                                                   excludedNames:=myCollection)
                                    edatenreihe = ShowProjekte.getRoleValuesInMonth(roleID:=prcName, _
                                                                                   considerAllSubRoles:=True, _
                                                                                   type:=PTcbr.realRoles, _
                                                                                   excludedNames:=myCollection)
                                Else
                                    datenreihe = ShowProjekte.getRoleValuesInMonth(roleID:=prcName, _
                                                                                   considerAllSubRoles:=True, _
                                                                                   type:=PTcbr.all, _
                                                                                   excludedNames:=myCollection)
                                End If

                            Else
                                datenreihe = ShowProjekte.getRoleValuesInMonth(prcName)
                            End If




                        ElseIf prcTyp = DiagrammTypen(2) Then
                            ' Kostenarten 
                            einheit = " T€"
                            If prcName = CostDefinitions.getCostdef(CostDefinitions.Count).name Then

                                ' es handelt sich um die Personalkosten, deshalb muss unterschieden werden zwischen internen und externen Kosten
                                isPersCost = True
                                objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                                'datenreihe = ShowProjekte.getCostiValuesInMonth
                                'edatenreihe = ShowProjekte.getCosteValuesInMonth
                                datenreihe = ShowProjekte.getCostGpValuesInMonth

                                ' Änderung tk: das wird doch hier nicht benötigt, ist eh Null, ausserdem wird das später nochmal gemacht 
                                'For i = 0 To bis - von
                                '    seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + edatenreihe(i)
                                'Next i

                            Else

                                ' es handelt sich nicht um die Personalkosten
                                isPersCost = False
                                objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                                datenreihe = ShowProjekte.getCostValuesInMonth(prcName)
                            End If
                        ElseIf prcTyp = DiagrammTypen(4) Then
                            ' Portfolio Charts wie Ergebnis 

                            ' es handelt sich um die Ergebnisse Earned Value bzw. Earned Value - gewichtet 
                            einheit = " T€"

                            objektFarbe = ergebnisfarbe1
                            datenreihe = ShowProjekte.getEarnedValuesInMonth()
                            ' jetzt müssen die - theoretischen Earned Values um die externen Kosten bereinigt werden, die abfallen, weil aufgrund 
                            ' bestimmter überlasteter Rollen externe , teurere Kräfte reingeholt werden müssen 

                            edatenreihe = ShowProjekte.getCosteValuesInMonth(True)
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
                            ' Meilensteine ... 

                            einheit = " "

                            Dim tmpMilestoneDef As clsMeilensteinDefinition = MilestoneDefinitions.getMilestoneDef(prcName)
                            If IsNothing(tmpMilestoneDef) Then
                                If appearanceDefinitions.ContainsKey("Meilenstein Default") Then
                                    objektFarbe = appearanceDefinitions.Item("Meilenstein Default").form.Fill.ForeColor.RGB
                                Else
                                    objektFarbe = awinSettings.AmpelNichtBewertet
                                End If

                            Else
                                objektFarbe = tmpMilestoneDef.farbe
                            End If

                            msdatenreihe = ShowProjekte.getCountMilestonesInMonth(prcName, breadcrumb, type, pvName)

                        ElseIf prcTyp = DiagrammTypen(8) Then
                            ' Meilenstein-Kategorie 
                            einheit = " "

                            If appearanceDefinitions.ContainsKey(prcName) Then
                                objektFarbe = appearanceDefinitions.Item(prcName).form.Fill.ForeColor.RGB
                            Else
                                objektFarbe = awinSettings.AmpelNichtBewertet
                            End If

                            datenreihe = ShowProjekte.getCountMilestoneCategoriesInMonth(prcName)
                        End If

                        If prcTyp = DiagrammTypen(1) And sumRoleShowsPlaceHolderAndAssigned Then
                            For i = 0 To bis - von
                                seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + datenreihe(i) + _
                                                            edatenreihe(i)
                            Next i
                        Else
                            For i = 0 To bis - von
                                seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + datenreihe(i)
                            Next i
                        End If


                        If isPersCost Then
                            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                                '.name = prcName & " intern "
                                .Name = prcName & repMessages.getmsg(115)
                                .Interior.Color = objektFarbe
                                .Values = datenreihe
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False
                            End With
                            If edatenreihe.Sum > 0 Then
                                With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                                    '.name = "Kosten durch Überlastung "
                                    .Name = repMessages.getmsg(152)
                                    .Interior.Color = farbeExterne
                                    .Values = edatenreihe
                                    .XValues = Xdatenreihe
                                    .ChartType = Excel.XlChartType.xlColumnStacked
                                    .HasDataLabels = False
                                End With
                            End If

                        Else
                            Dim legendName As String = ""
                            ' tk: repmsg muss nagepasst werden, wenn es nicht da ist 
                            If repMessages.getmsg(275) <> "" Then
                                legendName = prcName & " " & repMessages.getmsg(275)
                            Else
                                If awinSettings.englishLanguage Then
                                    legendName = prcName & " " & "Sum over all projects"
                                Else
                                    legendName = prcName & " " & "Summe über alle Projekte"
                                End If
                            End If


                            If prcTyp = DiagrammTypen(5) Then

                                ' Änderung 8.10.14 die Zahl der MEilensteine insgesamt anzeigen 
                                ' nicht aufgeschlüsselt nach welcher MEilenstein , welche Farbe

                                For i = 0 To bis - von
                                    datenreihe(i) = 0
                                    For c = 0 To 3
                                        datenreihe(i) = datenreihe(i) + msdatenreihe(c, i)
                                    Next
                                Next

                                With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                                    .Name = legendName
                                    .Interior.Color = objektFarbe
                                    .Values = datenreihe
                                    .XValues = Xdatenreihe
                                    .ChartType = Excel.XlChartType.xlColumnStacked
                                    .HasDataLabels = False
                                End With


                            Else

                                With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)


                                    If prcTyp = DiagrammTypen(1) And sumRoleShowsPlaceHolderAndAssigned Then
                                        ' repmsg!
                                        ' tk: repmsg muss nagepasst werden, wenn es nicht da ist 
                                        If repMessages.getmsg(276) <> "" Then
                                            .Name = legendName & ": " & repMessages.getmsg(276)
                                        Else
                                            If awinSettings.englishLanguage Then
                                                .Name = legendName & ": placeholder"
                                            Else
                                                .Name = legendName & ": Platzhalter"
                                            End If
                                        End If


                                    Else
                                        .Name = legendName
                                    End If

                                    .Interior.Color = objektFarbe
                                    .Values = datenreihe
                                    .XValues = Xdatenreihe
                                    If myCollection.Count = 1 Then
                                        If isWeightedValues Or sumRoleShowsPlaceHolderAndAssigned Then
                                            .ChartType = Excel.XlChartType.xlColumnStacked
                                        Else
                                            .ChartType = Excel.XlChartType.xlColumnClustered
                                        End If
                                    Else
                                        .ChartType = Excel.XlChartType.xlColumnStacked
                                    End If
                                    .HasDataLabels = False
                                End With

                                If prcTyp = DiagrammTypen(1) And sumRoleShowsPlaceHolderAndAssigned Then
                                    ' alle anderen zeigen 
                                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)

                                        ' tk: repmsg muss angepasst werden ... wenn es nicht da ist ... 
                                        If repMessages.getmsg(277) <> "" Then
                                            .Name = legendName & ": " & repMessages.getmsg(277)
                                        Else
                                            If awinSettings.englishLanguage Then
                                                .Name = legendName & ": assigned"
                                            Else
                                                .Name = legendName & ": zugeordnet"
                                            End If
                                        End If

                                        .Interior.Color = awinSettings.AmpelNichtBewertet
                                        .Values = edatenreihe
                                        .XValues = Xdatenreihe
                                        .ChartType = Excel.XlChartType.xlColumnStacked
                                        .HasDataLabels = False

                                    End With

                                End If

                            End If

                        End If

                    Next r

                    ' wenn es sich um die weighted Variante handelt
                    If isWeightedValues Then
                        With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                            .HasDataLabels = False
                            '.name = "Risiko Abschlag"
                            .Name = repMessages.getmsg(117)
                            .Interior.Color = ergebnisfarbe2
                            .Values = edatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                        End With
                    End If

                    ' wenn es sich um ein Cockpit Chart handelt, dann wird der jeweilige Min, Max-Wert angezeigt

                    lastSC = CType(.SeriesCollection, Excel.SeriesCollection).Count

                    If isCockpitChart Then
                        ' jetzt muss eine Dummy Series Collection eingeführt werde, damit das Datalabel über dem Balken angezeigt wird
                        If lastSC > 1 Then


                            maxwert = seriesSumDatenreihe.Max

                            For i = 0 To bis - von
                                VarValues(i) = 0.5 * maxwert
                            Next i

                            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                                .Name = "Dummy"
                                .Interior.Color = RGB(255, 255, 255)
                                .Values = VarValues
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False
                            End With
                            lastSC = CType(.SeriesCollection, Excel.SeriesCollection).Count

                        End If
                        With CType(.SeriesCollection(lastSC), Excel.Series)
                            .HasDataLabels = False
                            VarValues = seriesSumDatenreihe
                            nr_pts = CType(.Points, Excel.Points).Count

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

                    ' nur wenn auch Externe Ressourcen definiert / beauftragt sind, auch anzeigen
                    ' ansonsten werden nur die internen Kapazitäten angezeigt 
                    If prcTyp = DiagrammTypen(1) Then
                        If kdatenreihe.Sum < kdatenreihePlus.Sum Then
                            ' es gibt geplante externe Ressourcen ... 
                            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                                .HasDataLabels = False
                                '.name = "Kapazität incl. Externe"
                                .Name = repMessages.getmsg(118)

                                .Values = kdatenreihePlus
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlLine
                                With .Format.Line
                                    .DashStyle = MsoLineDashStyle.msoLineSysDot
                                    .ForeColor.RGB = XlRgbColor.rgbFuchsia
                                    .Weight = 2
                                End With
                                nr_pts = CType(.Points, Excel.Points).Count
                            End With
                        End If
                    End If

                    If prcTyp = DiagrammTypen(1) Or _
                        (prcTyp = DiagrammTypen(0) And kdatenreihe.Sum > 0) Or _
                        (prcTyp = DiagrammTypen(5) And kdatenreihe.Sum > 0) Then
                        With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                            .HasDataLabels = False

                            If prcTyp = DiagrammTypen(0) Or prcTyp = DiagrammTypen(5) Then
                                '.name = "Leistbarkeitsgrenze"
                                .Name = repMessages.getmsg(119)
                            Else
                                '.name = "Interne Kapazität"
                                .Name = repMessages.getmsg(260)
                            End If

                            '.Border.Color = rollenKapaFarbe
                            .Values = kdatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlLine
                            With .Format.Line
                                .DashStyle = MsoLineDashStyle.msoLineSolid
                                .ForeColor.RGB = XlRgbColor.rgbFireBrick
                                .Weight = 1.5
                            End With

                            nr_pts = CType(.Points, Excel.Points).Count

                            With .Points(nr_pts)

                                .HasDataLabel = False

                            End With

                        End With

                    End If
                    .HasTitle = True

                    If prcTyp = DiagrammTypen(0) Or _
                        prcTyp = DiagrammTypen(5) Or _
                        prcTyp = DiagrammTypen(7) Or _
                        prcTyp = DiagrammTypen(8) Then
                        titleSumme = ""

                    ElseIf prcTyp = DiagrammTypen(1) Then
                        einheit = awinSettings.kapaEinheit
                        titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & " / " & _
                                            Format(kdatenreihe.Sum, "##,##0") & " " & einheit & ")"

                    ElseIf prcTyp = DiagrammTypen(2) Then
                        einheit = "T€"
                        titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & " " & einheit & ")"
                    Else
                        titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & einheit & ")"
                    End If

                    .ChartTitle.Text = diagramTitle & titleSumme
                    ' lastSC muss  bestimmt werden 
                    lastSC = CType(.SeriesCollection, Excel.SeriesCollection).Count

                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle

                    If isCockpitChart Then

                        .ChartTitle.Font.Size = awinSettings.CPfontsizeTitle
                        .HasLegend = False

                    Else

                        'ElseIf lastSC > 1 Then

                        .HasLegend = True

                        .Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop
                        .Legend.Font.Size = awinSettings.fontsizeLegend
                        'Else
                        '    .HasLegend = False
                    End If

                End With



                With newChtObj
                    .Name = chtobjName
                    .Chart.Axes(Excel.XlAxisType.xlValue).minimumScale = 0
                End With


                ' wenn es ein Cockpit Chart ist: dann werden die Borderlines ausgeschaltet ...
                If isCockpitChart Then
                    Try
                        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentSheetName), Excel.Worksheet)
                            .Shapes.Item(chtobjName).Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                        End With
                    Catch ex As Exception

                    End Try
                Else
                    'Call awinScrollintoView()
                End If


                'repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                repObj = newChtObj

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then
                    prcDiagram = New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    prcChart = New clsEventsPrcCharts
                    'prcChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
                    prcChart.PrcChartEvents = newChtObj.Chart
                    prcDiagram.setDiagramEvent = prcChart
                    ' Ende Event Handling für Chart 


                    With prcDiagram
                        .DiagrammTitel = diagramTitle
                        .diagrammTyp = prcTyp
                        For ik As Integer = 1 To myCollection.Count
                            Dim tmpName As String = CStr(myCollection.Item(ik))
                            If Not .gsCollection.Contains(tmpName) Then
                                .gsCollection.Add(tmpName, tmpName)
                            End If
                        Next
                        ' das obige wurde gemacht, um myCollection nicht per Ref übergeben zu müssen ... 
                        '.gsCollection = myCollection
                        .isCockpitChart = isCockpitChart
                        .top = top
                        .left = left
                        .kennung = chtobjName
                        ' ur:09.03.2015: wegen Chart-Resize geändert
                        .width = width
                        .height = height

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
    ''' <summary>
    ''' aktualisiert ein Phasen-/Meilenstein-/Rollen-/Kosten-Diagramm
    ''' die optionalen Parameter sind relevant, wenn es um das Chart in Massen-Edit geht ... 
    ''' 
    ''' </summary>
    ''' <param name="chtobj"></param>
    ''' <remarks></remarks>
    Sub awinUpdateprcCollectionDiagram(ByVal chtobj As ChartObject, _
                                       ByVal roleCost As String, _
                                       ByVal isRole As Boolean)

        Dim von As Integer, bis As Integer
        Dim i As Integer, m As Integer, d As Integer, r As Integer
        Dim found As Boolean
        Dim hmxWert As Double = -10000.0 ' nimmt den Max-Wert der Datenreihe auf

        'Dim minwert As Double, maxwert As Double
        'Dim nr_pts As Integer
        Dim diagramTitle As String

        Dim objektFarbe As Object
        Dim ampelfarbe(3) As Long
        Dim Xdatenreihe() As String
        Dim datenreihe() As Double, edatenreihe() As Double, seriesSumDatenreihe() As Double
        Dim msdatenreihe(,) As Double
        ' nimmt die Daten der selektierten Werte auf 
        Dim seldatenreihe() As Double, tmpdatenreihe() As Double

        Dim kdatenreihe() As Double
        Dim kdatenreihePlus() As Double ' nimmt die Kapa Werte inkl bereits beauftragter externer Ressourcen auf 
        Dim prcName As String = ""

        Dim breadcrumb As String = ""
        Dim startdate As Date
        Dim diff As Integer
        'Dim mindone As Boolean, maxdone As Boolean
        'Dim width As Double
        'Dim left As Double
        Dim myCollection As Collection
        Dim isCockpitChart As Boolean
        Dim isWeightedValues As Boolean = False
        Dim VarValues() As Double
        Dim prcTyp As String
        Dim isPersCost As Boolean
        Dim lastSC As Integer
        Dim titleSumme As String, einheit As String
        Dim selectionFarbe As Long = awinSettings.AmpelRot

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
        'width = chtobj.Width

        Dim currentScale As Double
        Try
            With CType(chtobj.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
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

            'width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct

        End If

        ReDim Xdatenreihe(bis - von)
        ReDim datenreihe(bis - von)
        ReDim edatenreihe(bis - von)
        ReDim kdatenreihe(bis - von)
        ReDim kdatenreihePlus(bis - von)
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
        Dim foundDiagram As clsDiagramm = Nothing

        ' bestimmen, ob man sich auf der Projekt-Tafel befindet oder aber im MassEdit Ressourcen, Termine, Attribute
        Try
            If visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
                ' bestimmen des prcTyp
                If isRole Then
                    prcTyp = DiagrammTypen(1)
                Else
                    prcTyp = DiagrammTypen(2)
                End If
                If Not IsNothing(roleCost) Then
                    myCollection.Add(roleCost)
                End If
                found = True

            Else
                If DiagramList.contains(chtobjName) Then
                    foundDiagram = DiagramList.getDiagramm(chtobjName)

                    myCollection = foundDiagram.gsCollection
                    prcTyp = foundDiagram.diagrammTyp
                    found = True
                End If
                
            End If

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
            Xdatenreihe(m - von) = startdate.AddMonths(m).ToString("MMM yy", repCult)
        Next m

        If myCollection.Count > 1 Then
            If prcTyp = DiagrammTypen(0) Then
                'diagramTitle = "Phasen-Übersicht"
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Phasen)
            ElseIf prcTyp = DiagrammTypen(1) Then
                'diagramTitle = "Rollen-Übersicht"
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Rollen)
            ElseIf prcTyp = DiagrammTypen(2) Then
                'diagramTitle = "Kosten-Übersicht"
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Kosten)
            ElseIf prcTyp = DiagrammTypen(4) Then
                'diagramTitle = "Ergebnis-Übersicht"
                diagramTitle = repMessages.getmsg(113)
            ElseIf prcTyp = DiagrammTypen(5) Then
                chtobjName = calcChartKennung("pf", PTpfdk.Meilenstein, myCollection)
                diagramTitle = portfolioDiagrammtitel(PTpfdk.Meilenstein)
            
            Else
                diagramTitle = repMessages.getmsg(114)
            End If
        Else
            diagramTitle = splitHryFullnameTo1(CStr(myCollection.Item(1)))
        End If

        ' jetzt den Namen aus optischen Gründen ändern 
        If diagramTitle.Contains("#") Then
            diagramTitle = diagramTitle.Replace("#", "-")
        End If

        ' Änderung tk 26.10.15 
        ' damit Diagramm-Title manuell geändert werden kann und dann beim Update , bis auf die Summe 
        ' unverändert bleibt, wird der hier rausgelesen; das darf aber nicht im Massen-Edit sein ....
        If Not visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
            Dim tmpstr() As String = chtobj.Chart.ChartTitle.Text.Split(New Char() {CChar("("), CChar(")")}, 20)
            If tmpstr(0).Length > 0 Then
                diagramTitle = tmpstr(0).TrimEnd
            End If
        End If



        If prcTyp = DiagrammTypen(1) Then
            kdatenreihe = ShowProjekte.getRoleKapasInMonth(myCollection, False)
            kdatenreihePlus = ShowProjekte.getRoleKapasInMonth(myCollection, True)
        ElseIf prcTyp = DiagrammTypen(0) Then
            kdatenreihe = ShowProjekte.getPhaseSchwellWerteInMonth(myCollection)
        ElseIf prcTyp = DiagrammTypen(5) Then
            kdatenreihe = ShowProjekte.getMilestoneSchwellWerteInMonth(myCollection)
        End If

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False


        'With appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT))


        With chtobj.Chart

            ' remove old series
            Try
                Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                Do While anz > 0
                    .SeriesCollection(1).Delete()
                    anz = anz - 1
                Loop
            Catch ex As Exception

            End Try


            ' wird benötigt, um zu entscheiden, ob es sich um eine SammelRolle handelt ... 
            Dim sumRoleShowsPlaceHolderAndAssigned As Boolean

            For r = 1 To myCollection.Count

                Dim type As Integer = -1
                Dim pvname As String = ""
                sumRoleShowsPlaceHolderAndAssigned = False

                If prcTyp = DiagrammTypen(0) Or prcTyp = DiagrammTypen(5) Then
                    Call splitHryFullnameTo2(CStr(myCollection.Item(r)), prcName, breadcrumb, type, pvname)

                ElseIf prcTyp = DiagrammTypen(7) Or prcTyp = DiagrammTypen(8) Then
                    Call splitHryFullnameTo2(CStr(myCollection.Item(r)), prcName, breadcrumb, type, pvname)
                    prcName = pvname ' der Name der Kategorie steht hier im pvName 

                Else
                    prcName = CStr(myCollection.Item(r))
                End If


                If prcTyp = DiagrammTypen(0) Then
                    einheit = " "

                    Dim tmpPhaseDef As clsPhasenDefinition = PhaseDefinitions.getPhaseDef(prcName)

                    If IsNothing(tmpPhaseDef) Then
                        If appearanceDefinitions.ContainsKey("Phasen Default") Then
                            objektFarbe = appearanceDefinitions.Item("Phasen Default").form.Fill.ForeColor.RGB
                        Else
                            objektFarbe = awinSettings.AmpelNichtBewertet
                        End If

                    Else
                        objektFarbe = tmpPhaseDef.farbe
                    End If

                    datenreihe = ShowProjekte.getCountPhasesInMonth(prcName, breadcrumb, type, pvname)
                    hmxWert = Max(datenreihe.Max, hmxWert)

                    If awinSettings.showValuesOfSelected And myCollection.Count = 1 Then
                        ' Ergänzung wegen Anzeige der selektierten Objekte ... 
                        tmpdatenreihe = selectedProjekte.getCountPhasesInMonth(prcName, breadcrumb, type, pvname)
                        For ix = 0 To bis - von
                            datenreihe(ix) = datenreihe(ix) - tmpdatenreihe(ix)
                            seldatenreihe(ix) = seldatenreihe(ix) + tmpdatenreihe(ix)
                        Next
                    End If

                ElseIf prcTyp = DiagrammTypen(7) Then
                    ' Phasen-Kategorie 
                    einheit = " "

                    If appearanceDefinitions.ContainsKey(prcName) Then
                        objektFarbe = appearanceDefinitions.Item(prcName).form.Fill.ForeColor.RGB
                    Else
                        objektFarbe = awinSettings.AmpelNichtBewertet
                    End If

                    datenreihe = ShowProjekte.getCountPhaseCategoriesInMonth(prcName)

                ElseIf prcTyp = DiagrammTypen(1) Then
                    einheit = " " & awinSettings.kapaEinheit
                    Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(prcName)
                    objektFarbe = RoleDefinitions.getRoledef(prcName).farbe

                    If tmpRole.isCombinedRole Then

                        If awinSettings.showPlaceholderAndAssigned Then
                            sumRoleShowsPlaceHolderAndAssigned = True
                            datenreihe = ShowProjekte.getRoleValuesInMonth(roleID:=prcName, _
                                                                           considerAllSubRoles:=True, _
                                                                           type:=PTcbr.placeholders, _
                                                                           excludedNames:=myCollection)
                            edatenreihe = ShowProjekte.getRoleValuesInMonth(roleID:=prcName, _
                                                                           considerAllSubRoles:=True, _
                                                                           type:=PTcbr.realRoles, _
                                                                           excludedNames:=myCollection)
                        Else
                            datenreihe = ShowProjekte.getRoleValuesInMonth(roleID:=prcName, _
                                                                           considerAllSubRoles:=True, _
                                                                           type:=PTcbr.all, _
                                                                           excludedNames:=myCollection)
                        End If

                    Else
                        datenreihe = ShowProjekte.getRoleValuesInMonth(prcName)
                    End If


                    If (awinSettings.showValuesOfSelected) And myCollection.Count = 1 Then
                        ' Ergänzung wegen Anzeige der selektierten Objekte ... 
                        If tmpRole.isCombinedRole Then
                            tmpdatenreihe = selectedProjekte.getRoleValuesInMonth(roleID:=prcName, _
                                                                       considerAllSubRoles:=True, _
                                                                       type:=PTcbr.all, _
                                                                       excludedNames:=myCollection)
                        Else
                            tmpdatenreihe = selectedProjekte.getRoleValuesInMonth(prcName)
                        End If

                        For ix = 0 To bis - von
                            datenreihe(ix) = datenreihe(ix) - tmpdatenreihe(ix)

                            If tmpRole.isCombinedRole And sumRoleShowsPlaceHolderAndAssigned Then
                                ' in diesem Fall kann datenreihe(ix) auch negativ werden, muss also auch von edatenreihe abgezogen werden ...
                                If datenreihe(ix) < 0 Then
                                    ' datenreihe(ix) ist negativ, also heisst das abziehen 
                                    edatenreihe(ix) = edatenreihe(ix) + datenreihe(ix)
                                    datenreihe(ix) = 0
                                End If

                            End If

                            seldatenreihe(ix) = seldatenreihe(ix) + tmpdatenreihe(ix)
                        Next
                    End If

                ElseIf prcTyp = DiagrammTypen(2) Then
                    einheit = " T€"
                    If prcName = CostDefinitions.getCostdef(CostDefinitions.Count).name Then
                        ' es handelt sich um die Personalkosten, deshalb muss unterschieden werden zwischen internen und externen Kosten
                        isPersCost = True
                        objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                        'datenreihe = ShowProjekte.getCostiValuesInMonth
                        'edatenreihe = ShowProjekte.getCosteValuesInMonth
                        datenreihe = ShowProjekte.getCostGpValuesInMonth

                    Else
                        ' es handelt sich nicht um die Personalkosten
                        isPersCost = False
                        objektFarbe = CostDefinitions.getCostdef(prcName).farbe
                        datenreihe = ShowProjekte.getCostValuesInMonth(prcName)
                        hmxWert = datenreihe.Max

                        If (awinSettings.showValuesOfSelected) And myCollection.Count = 1 Then
                            ' Ergänzung wegen Anzeige der selektierten Objekte ... 
                            tmpdatenreihe = selectedProjekte.getCostValuesInMonth(prcName)
                            For ix = 0 To bis - von
                                datenreihe(ix) = datenreihe(ix) - tmpdatenreihe(ix)
                                seldatenreihe(ix) = seldatenreihe(ix) + tmpdatenreihe(ix)
                            Next
                        End If

                    End If

                ElseIf prcTyp = DiagrammTypen(4) Then
                    ' es handelt sich um die Ergebnisse Earned Value bzw. Earned Value - gewichtet 
                    einheit = " T€"

                    objektFarbe = ergebnisfarbe1
                    datenreihe = ShowProjekte.getEarnedValuesInMonth()
                    ' jetzt müssen die - theoretischen Earned Values um die externen Kosten bereinigt werden, die abfallen, weil aufgrund 
                    ' bestimmter überlasteter Rollen externe , teurere Kräfte reingeholt werden müssen 


                    edatenreihe = ShowProjekte.getCosteValuesInMonth(True)
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
                    Dim tmpMilestoneDef As clsMeilensteinDefinition = MilestoneDefinitions.getMilestoneDef(prcName)
                    If IsNothing(tmpMilestoneDef) Then
                        If appearanceDefinitions.ContainsKey("Meilenstein Default") Then
                            objektFarbe = appearanceDefinitions.Item("Meilenstein Default").form.Fill.ForeColor.RGB
                        Else
                            objektFarbe = awinSettings.AmpelNichtBewertet
                        End If

                    Else
                        objektFarbe = tmpMilestoneDef.farbe
                    End If
                    msdatenreihe = ShowProjekte.getCountMilestonesInMonth(prcName, breadcrumb, type, pvname)

                ElseIf prcTyp = DiagrammTypen(8) Then
                    ' Meilenstein-Kategorie 
                    einheit = " "

                    If appearanceDefinitions.ContainsKey(prcName) Then
                        objektFarbe = appearanceDefinitions.Item(prcName).form.Fill.ForeColor.RGB
                    Else
                        objektFarbe = awinSettings.AmpelNichtBewertet
                    End If

                    datenreihe = ShowProjekte.getCountMilestoneCategoriesInMonth(prcName)

                End If


                If isPersCost Then
                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)

                        '.name = prcName & " intern "
                        .Name = prcName & repMessages.getmsg(115)
                        .Interior.Color = objektFarbe
                        .Values = datenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked
                        .HasDataLabels = False
                    End With

                    If edatenreihe.Sum > 0 Then
                        With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                            '.name = "Kosten durch Überlastung "
                            .Name = repMessages.getmsg(152)
                            .Interior.Color = farbeExterne
                            .Values = edatenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                            .HasDataLabels = False
                        End With
                    End If

                Else
                    If prcTyp = DiagrammTypen(5) Then


                        ' Änderung 8.10.14 die Zahl der MEilensteine insgesamt anzeigen 
                        ' nicht aufgeschlüsselt nach welcher MEilenstein , welche Farbe

                        For i = 0 To bis - von
                            datenreihe(i) = 0
                            For c = 0 To 3
                                datenreihe(i) = datenreihe(i) + msdatenreihe(c, i)
                            Next
                        Next

                        With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                            If breadcrumb = "" Then
                                .Name = prcName
                            Else
                                .Name = breadcrumb & "-" & prcName
                            End If

                            '.Interior.color = ampelfarbe(0)
                            .Interior.Color = objektFarbe
                            .Values = datenreihe
                            .XValues = Xdatenreihe
                            .ChartType = Excel.XlChartType.xlColumnStacked
                            .HasDataLabels = False
                        End With


                    Else

                        ' Ergänzung wegen Anzeige selektierter Objekte 
                        ' wenn der Wert größer ist als Null, dann Anzeigen ... 
                        If myCollection.Count = 1 Then
                            If (awinSettings.showValuesOfSelected) And selectedProjekte.Count > 0 Then
                                With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                                    .HasDataLabels = False
                                    If selectedProjekte.Count = 1 Then
                                        .Name = selectedProjekte.getProject(1).name
                                    Else

                                        If awinSettings.englishLanguage Then
                                            .Name = "selected projects"
                                        Else
                                            .Name = "selektierte Projekte"
                                        End If
                                    End If
                                    .Interior.Color = selectionFarbe
                                    .Values = seldatenreihe
                                    .XValues = Xdatenreihe
                                    .ChartType = Excel.XlChartType.xlColumnStacked
                                End With

                            End If
                        End If
                        
                        Dim legendName As String = ""
                        If awinSettings.englishLanguage Then
                            legendName = "Sum over all projects"
                        Else
                            legendName = "Summe über alle Projekte"
                        End If

                        With CType(CType(chtobj.Chart.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)

                            If prcTyp = DiagrammTypen(1) And sumRoleShowsPlaceHolderAndAssigned Then
                                If awinSettings.englishLanguage Then
                                    .Name = legendName & ": placeholder"
                                Else
                                    .Name = legendName & ": Platzhalter"
                                End If

                            Else
                                If selectedProjekte.Count > 0 And myCollection.Count = 1 And awinSettings.showValuesOfSelected = True Then
                                    If awinSettings.englishLanguage Then
                                        .Name = "all other projects"
                                    Else
                                        .Name = "alle anderen Projekte"
                                    End If
                                Else
                                    If awinSettings.englishLanguage Then
                                        .Name = legendName
                                    Else
                                        .Name = legendName
                                    End If
                                End If
                            End If

                            .Interior.Color = objektFarbe
                            .Values = datenreihe
                            .XValues = Xdatenreihe
                            If myCollection.Count = 1 Then
                                If isWeightedValues Or sumRoleShowsPlaceHolderAndAssigned Or _
                                    (selectedProjekte.Count > 0 And awinSettings.showValuesOfSelected) Then
                                    .ChartType = Excel.XlChartType.xlColumnStacked
                                Else
                                    .ChartType = Excel.XlChartType.xlColumnClustered
                                End If
                            Else
                                .ChartType = Excel.XlChartType.xlColumnStacked
                            End If
                            .HasDataLabels = False
                        End With

                        If prcTyp = DiagrammTypen(1) And sumRoleShowsPlaceHolderAndAssigned Then
                            ' alle anderen zeigen 
                            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)

                                If awinSettings.englishLanguage Then
                                    .Name = legendName & ": assigned"
                                Else
                                    .Name = legendName & ": zugeordnet"
                                End If

                                .Interior.Color = awinSettings.AmpelNichtBewertet
                                .Values = edatenreihe
                                .XValues = Xdatenreihe
                                .ChartType = Excel.XlChartType.xlColumnStacked
                                .HasDataLabels = False

                            End With

                        End If

                    End If

                End If

                If prcTyp = DiagrammTypen(1) And sumRoleShowsPlaceHolderAndAssigned Then
                    For i = 0 To bis - von
                        seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + datenreihe(i) + _
                                                    edatenreihe(i)
                    Next i
                Else
                    For i = 0 To bis - von
                        seriesSumDatenreihe(i) = seriesSumDatenreihe(i) + datenreihe(i)
                    Next i
                End If

            Next r

            ' wenn es sich um die weighted Variante handelt
            If isWeightedValues Then
                With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                    .HasDataLabels = False
                    '.name = "Risiko Abschlag"
                    .Name = repMessages.getmsg(117)
                    .Interior.Color = ergebnisfarbe2
                    .Values = edatenreihe
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlColumnStacked
                End With
            End If


            ' wenn es sich um ein Cockpit Chart handelt, dann wird der jeweilige Min, Max-Wert angezeigt

            lastSC = CType(.SeriesCollection, Excel.SeriesCollection).Count


            ' nur wenn auch Externe Ressourcen definiert / beauftragt sind, auch anzeigen
            ' ansonsten werden nur die internen Kapazitäten angezeigt 
            ' hier werden die externen mitgezeichnet ....
            If prcTyp = DiagrammTypen(1) Then
                If kdatenreihe.Sum < kdatenreihePlus.Sum Then
                    'es gibt geplante externe Ressourcen ... 
                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        .HasDataLabels = False
                        '.name = "Kapazität incl. Externe"
                        .Name = repMessages.getmsg(118)

                        .Values = kdatenreihePlus
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlLine

                        'tk 28.3.17 soll bleiben wie es urspünglich war 
                        With .Format.Line
                            .DashStyle = MsoLineDashStyle.msoLineSysDot
                            .ForeColor.RGB = XlRgbColor.rgbFuchsia
                            .Weight = 2
                        End With

                        'nr_pts = CType(.Points, Excel.Points).Count
                    End With
                End If
            End If

            ' hier werde nur die internen gezeichnet ...
            If prcTyp = DiagrammTypen(1) Or _
                   (prcTyp = DiagrammTypen(0) And kdatenreihe.Sum > 0) Or _
                   (prcTyp = DiagrammTypen(5) And kdatenreihe.Sum > 0) Then
                With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                    .HasDataLabels = False

                    If prcTyp = DiagrammTypen(0) Or prcTyp = DiagrammTypen(5) Then
                        '.name = "Leistbarkeitsgrenze"
                        .Name = repMessages.getmsg(119)
                    Else
                        '.name = "Interne Kapazität"
                        .Name = repMessages.getmsg(260)
                    End If

                    '.Border.Color = rollenKapaFarbe
                    .Values = kdatenreihe
                    .XValues = Xdatenreihe
                    .ChartType = Excel.XlChartType.xlLine

                    ' tk: da es neu aufgebaut wird, muss es neu gezeichnet werden ..
                    With .Format.Line
                        .DashStyle = MsoLineDashStyle.msoLineSolid
                        .ForeColor.RGB = XlRgbColor.rgbFireBrick
                        .Weight = 1.5
                    End With

                    'nr_pts = CType(.Points, Excel.Points).Count

                    'With .Points(nr_pts)

                    '    .HasDataLabel = False

                    'End With

                End With

            End If

            .HasTitle = True


            If prcTyp = DiagrammTypen(0) Or _
                    prcTyp = DiagrammTypen(5) Or _
                    prcTyp = DiagrammTypen(7) Or _
                    prcTyp = DiagrammTypen(8) Then
                titleSumme = ""

            ElseIf prcTyp = DiagrammTypen(1) Then

                einheit = awinSettings.kapaEinheit
                If awinSettings.showValuesOfSelected And seldatenreihe.Sum > 0 Then
                    titleSumme = " (" & Format(seldatenreihe.Sum, "##,##0") & " / " & _
                                        Format(seriesSumDatenreihe.Sum, "##,##0") & " / " & _
                                        Format(kdatenreihe.Sum, "##,##0") & " " & einheit & ")"
                Else
                    titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & " / " & _
                                    Format(kdatenreihe.Sum, "##,##0") & " " & einheit & ")"
                End If


            ElseIf prcTyp = DiagrammTypen(2) Then

                einheit = "T€"
                If awinSettings.showValuesOfSelected And seldatenreihe.Sum > 0 Then
                    titleSumme = " (" & Format(seldatenreihe.Sum, "##,##0") & " / " & _
                                        Format(seriesSumDatenreihe.Sum, "##,##0") & einheit & ")"
                Else
                    titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & " " & einheit & ")"
                End If


            Else
                titleSumme = " (" & Format(seriesSumDatenreihe.Sum, "##,##0") & einheit & ")"
            End If


            .ChartTitle.Text = diagramTitle & titleSumme
            ' lastSC muss  bestimmt werden 
            lastSC = CType(.SeriesCollection, Excel.SeriesCollection).Count


            ' Änderung 18.3.15 tk: bei einem Update muss überhaupt nix geändert werden, was LEgende angeht ; 
            ' die ist entweder da und soll da bleiben oder sie ist nicht da und soll auch nicht kommen 
            'If isCockpitChart Then
            '    .HasLegend = False
            'ElseIf lastSC > 1 And seldatenreihe.Sum = 0 Then
            '    .HasLegend = True
            '    'ur: 11.03.2015: wenn ein Chart eine Legende hat, so soll sie bleiben wie zuletzt definiert, nicht jedesmal auf Ursprungszustand zurückgesetzt werden
            '    '.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop
            '    '.Legend.Font.Size = awinSettings.fontsizeLegend
            'Else
            '    .HasLegend = False
            'End If

        End With


        'End With ' with worksheet ...

        ' tk, darf nicht verändert werden, weil sonst ein defniertes Cockpit völlig aus dem Rahmen läuft  
        'With chtobj
        '    If Not isCockpitChart Then
        '        .Width = width
        '    End If
        'End With

        ' Skalierung nur ändern, wenn erforderlich, weil der maxwert höher ist als die bisherige Skalierung ... 
        If visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
            With CType(chtobj.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                .MaximumScaleIsAuto = True
                '.MaximumScale = hmxWert + 1
            End With
        Else
            hmxWert = Max(seriesSumDatenreihe.Max, hmxWert)
            If hmxWert > currentScale Then
                With chtobj.Chart.Axes(Excel.XlAxisType.xlValue)
                    .MaximumScale = hmxWert + 1
                End With
            Else
                With chtobj.Chart.Axes(Excel.XlAxisType.xlValue)
                    .MaximumScale = currentScale
                End With
            End If
        End If

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU




    End Sub


    ''' <summary>
    ''' aktualisiert das Budget Ergebnis Diagramm 
    ''' </summary>
    ''' <param name="chtObj">Verweis aus das zu aktualisierende Chart</param>
    ''' <remarks></remarks>
    Sub awinUpdateBudgetErgebnisDiagramm(ByVal chtObj As Excel.ChartObject)

        Dim diagramTitle As String
        Dim minScale As Double
        Dim maxscale As Double
        Dim Xdatenreihe(3) As String
        Dim valueDatenreihe1(3) As Double
        Dim valueDatenreihe2(3) As Double
        Dim itemColor(3) As Object
        Dim itemValue(3) As Double

        Dim budgetSum As Double, pCost As Double, oCost As Double
        Dim ertragsWert As Double
        Dim minColumn As Integer, maxColumn As Integer, heuteColumn As Integer, heuteIndex As Integer
        Dim future As Boolean = False

        heuteColumn = getColumnOfDate(Date.Today)
        heuteIndex = heuteColumn - showRangeLeft

        minColumn = showRangeLeft
        maxColumn = showRangeRight

        Dim mycollection As New Collection

        Dim ErgebnisListeR As New Collection



        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False


        Xdatenreihe(0) = repMessages.getmsg(49)
        Xdatenreihe(1) = repMessages.getmsg(51)
        Xdatenreihe(2) = repMessages.getmsg(52)
        Xdatenreihe(3) = repMessages.getmsg(53)



        Dim positiv As Boolean = True

        ' Ausrechnen amteiliges Budget, das i Zeitraum zur Verfügung steht und der im Zeitraum anfallenden Kosten  
        budgetSum = System.Math.Round(ShowProjekte.getBudgetValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        pCost = System.Math.Round(ShowProjekte.getCostGpValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        oCost = System.Math.Round(ShowProjekte.getOtherCostValuesInMonth.Sum, mode:=MidpointRounding.ToEven)

        ertragsWert = budgetSum - (pCost + oCost)

        If ertragsWert < 0 Then
            minScale = ertragsWert
        Else
            minScale = 0
        End If

        maxscale = budgetSum

        itemValue(0) = budgetSum
        itemColor(0) = ergebnisfarbe1


        Dim currentWert As Double = itemValue(0)



        ' das sind die Personalkosten
        itemValue(1) = pCost
        itemColor(1) = farbeExterne

        ' das sind die Other Cost 
        itemValue(2) = oCost
        itemColor(2) = farbeExterne

        ' das ist der Ertrag 
        itemValue(3) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(3) = ergebnisfarbe2
        Else
            itemColor(3) = farbeExterne
        End If

        'diagramTitle = portfolioDiagrammtitel(PTpfdk.Budget) & " " & textZeitraum(showRangeLeft, showRangeRight)
        If getColumnOfDate(Date.Now) > showRangeRight Then
            diagramTitle = "Portfolio " & textZeitraum(showRangeLeft, showRangeRight)
        Else
            diagramTitle = "Forecast Portfolio " & textZeitraum(showRangeLeft, showRangeRight)
        End If


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False



        If ertragsWert < 0 Then
            minScale = System.Math.Round(ertragsWert, mode:=MidpointRounding.ToEven)
        Else
            minScale = 0
        End If

        'Dim htxt As String
        Dim valueCrossesNull As Boolean = False

        With chtObj.Chart
            ' remove old series
            Try
                Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                Do While anz > 0
                    .SeriesCollection(1).Delete()
                    anz = anz - 1
                Loop
            Catch ex As Exception

            End Try
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
            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                .Name = "Bottom"
                .HasDataLabels = False
                .Interior.ColorIndex = -4142
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

            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                .Name = "Top"
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


            Try
                With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)

                    If minScale < .MinimumScale Then
                        .MinimumScale = minScale * 1.2
                    End If

                    If maxscale > .MaximumScale Then
                        .MaximumScale = maxscale * 1.2
                    End If
                End With
            Catch ex As Exception

            End Try


            .ChartTitle.Text = diagramTitle

        End With

        'End With


        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub


    ''' <summary>
    ''' aktualisiert das Budget Ergebnis Diagramm 
    ''' berücksichtigt Risiko Kosten - wenn die wieder mal aktiviert werden sollen, dann ... 
    ''' </summary>
    ''' <param name="chtObj">Verweis aus das zu aktualisierende Chart</param>
    ''' <remarks></remarks>
    Sub awinUpdateBudgetErgebnisDiagramm_deprecated(ByVal chtObj As Excel.ChartObject)

        Dim diagramTitle As String
        Dim minScale As Double
        Dim maxscale As Double
        Dim Xdatenreihe(4) As String
        Dim valueDatenreihe1(4) As Double
        Dim valueDatenreihe2(4) As Double
        Dim itemColor(4) As Object
        Dim itemValue(4) As Double

        Dim budgetSum As Double, pCost As Double, oCost As Double, riskValue As Double
        Dim ertragsWert As Double
        Dim minColumn As Integer, maxColumn As Integer, heuteColumn As Integer, heuteIndex As Integer
        Dim future As Boolean = False

        heuteColumn = getColumnOfDate(Date.Today)
        heuteIndex = heuteColumn - showRangeLeft

        minColumn = showRangeLeft
        maxColumn = showRangeRight

        Dim mycollection As New Collection

        Dim ErgebnisListeR As New Collection



        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False


        Xdatenreihe(0) = repMessages.getmsg(49)
        Xdatenreihe(1) = repMessages.getmsg(50)
        Xdatenreihe(2) = repMessages.getmsg(51)
        Xdatenreihe(3) = repMessages.getmsg(52)
        Xdatenreihe(4) = repMessages.getmsg(53)


        'Xdatenreihe(0) = "Budget Summe"
        'If heuteColumn >= minColumn + 1 And heuteColumn <= maxColumn Then
        '    Xdatenreihe(2) = "bisherige Kosten" & vbLf & textZeitraum(minColumn, heuteColumn - 1)
        '    Xdatenreihe(3) = "Prognose Kosten" & vbLf & textZeitraum(heuteColumn, maxColumn)
        'ElseIf heuteColumn > maxColumn Then
        '    future = False
        '    Xdatenreihe(2) = "bisherige Kosten" & vbLf & textZeitraum(minColumn, maxColumn)
        '    Xdatenreihe(3) = "Prognose Kosten" & vbLf & "existieren nicht"
        'ElseIf heuteColumn <= minColumn Then
        '    future = True
        '    Xdatenreihe(2) = "bisherige Kosten" & vbLf & "existieren nicht"
        '    Xdatenreihe(3) = "Prognose Kosten" & vbLf & textZeitraum(minColumn, maxColumn)
        'End If

        'Xdatenreihe(1) = "Risiko-Abschlag"
        'Xdatenreihe(4) = "Ergebnis"

        Dim positiv As Boolean = True

        ' Ausrechnen amteiliges Budget, das i Zeitraum zur Verfügung steht und der im Zeitraum anfallenden Kosten  
        budgetSum = System.Math.Round(ShowProjekte.getBudgetValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        pCost = System.Math.Round(ShowProjekte.getCostGpValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        oCost = System.Math.Round(ShowProjekte.getOtherCostValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        riskValue = System.Math.Round(ShowProjekte.getWeightedRiskValuesInMonth.Sum, mode:=MidpointRounding.ToEven)

        ertragsWert = budgetSum - (riskValue + pCost + oCost)

        If ertragsWert < 0 Then
            minScale = ertragsWert
        Else
            minScale = 0
        End If

        maxscale = budgetSum

        itemValue(0) = budgetSum
        itemColor(0) = ergebnisfarbe1


        Dim currentWert As Double = itemValue(0)


        ' das ist der Risiko-Abschlag 
        itemValue(1) = riskValue
        itemColor(1) = iProjektFarbe

        ' das sind die Personalkosten
        itemValue(2) = pCost
        itemColor(2) = farbeExterne

        ' das sind die Other Cost 
        itemValue(3) = oCost
        itemColor(3) = farbeExterne

        ' das ist der Ertrag 
        itemValue(4) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(4) = ergebnisfarbe2
        Else
            itemColor(4) = farbeExterne
        End If

        diagramTitle = portfolioDiagrammtitel(PTpfdk.Budget) & " " & textZeitraum(showRangeLeft, showRangeRight)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False





        If ertragsWert < 0 Then
            minScale = System.Math.Round(ertragsWert, mode:=MidpointRounding.ToEven)
        Else
            minScale = 0
        End If

        'Dim htxt As String
        Dim valueCrossesNull As Boolean = False

        With chtObj.Chart
            ' remove extra series
            ' remove old series
            Try
                Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                Do While anz > 0
                    .SeriesCollection(1).Delete()
                    anz = anz - 1
                Loop
            Catch ex As Exception

            End Try
            Dim crossindex As Integer = -1

            ' bestimmen des Anfangs  
            Dim iv = 0
            valueDatenreihe1(iv) = 0
            valueDatenreihe2(iv) = itemValue(iv)
            currentWert = itemValue(iv)
            Dim formerValue As Double = currentWert
            Dim negativeFromNull As Boolean = False

            ' alle nächsten Zwischen-Werte 
            For iv = 1 To 3
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
            iv = 4
            valueDatenreihe1(iv) = 0
            valueDatenreihe2(iv) = itemValue(iv)



            'series
            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                .Name = "Bottom"
                .HasDataLabels = False
                .Interior.ColorIndex = -4142
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

            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                .Name = "Top"
                .HasDataLabels = True
                .Values = valueDatenreihe2
                .XValues = Xdatenreihe
                .ChartType = Excel.XlChartType.xlColumnStacked

                For iv = 0 To 4

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

            ' Änderung tk: 15.9.16
            ' das muss ja eigentlich nicht angepasst werden, da es sich hier um Update handelt ... 
            ''.HasAxis(Excel.XlAxisType.xlCategory) = True
            ''.HasAxis(Excel.XlAxisType.xlValue) = False

            ''With .Axes(Excel.XlAxisType.xlCategory)
            ''    .HasTitle = False
            ''    If minScale < 0 Then
            ''        .TickLabelPosition = Excel.Constants.xlLow
            ''    End If
            ''    '.MinimumScale = 0

            ''End With


            Try
                With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)

                    If minScale < .MinimumScale Then
                        .MinimumScale = minScale * 1.2
                    End If

                    If maxscale > .MaximumScale Then
                        .MaximumScale = maxscale * 1.2
                    End If
                End With
            Catch ex As Exception

            End Try


            ''.HasLegend = False
            ''.HasTitle = True

            .ChartTitle.Text = diagramTitle
            '.ChartTitle.Font.Size = awinSettings.fontsizeTitle

            '
            ' tk : das gehört hier doch nicht hin , das ist doch cut&paste Fehler !? 
            ''Dim achieved As Boolean = False
            ''Dim anzahlVersuche As Integer = 0
            ''Dim errmsg As String = ""
            ''Do While Not achieved And anzahlVersuche < 10
            ''    Try
            ''        Call Sleep(100)
            ''        .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(ptTables.MPT)).name)
            ''        achieved = True
            ''    Catch ex As Exception
            ''        errmsg = ex.Message
            ''        Call Sleep(100)
            ''        anzahlVersuche = anzahlVersuche + 1
            ''    End Try
            ''Loop

            ''If Not achieved Then
            ''    Throw New ArgumentException("Chart-Fehler:" & errmsg)
            ''End If

        End With

        'End With


        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub


    ''' <summary>
    ''' zeigt für den betrachteten Zeitraum das Auslastungsdiagramm an
    ''' Rolle ist beauftragt, ist ohne Arbeit, ist überlastet 
    ''' </summary>
    ''' <param name="repObj"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="calledfromReporting"></param>
    ''' <remarks></remarks>
    Sub awinCreateAuslastungsDiagramm(ByRef repObj As Excel.ChartObject, _
                                          ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                          ByVal calledfromReporting As Boolean)

        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
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

        Dim chtobjName As String
        Dim myCollection As New Collection
        myCollection.Add("Auslastung")
        chtobjName = calcChartKennung("pf", PTpfdk.Auslastung, myCollection)
        myCollection.Clear()

        Dim currentSheetName As String

        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentSheetName = arrWsNames(ptTables.MPT)
        Else
            currentSheetName = arrWsNames(ptTables.meRC)
        End If


        If Not calledfromReporting Then

            Dim foundDiagramm As clsDiagramm = Nothing

            ' wenn die Werte für dieses Diagramm bereits einmal gespeichert wurden ... -> übernehmen 
            Try
                If DiagramList.contains(chtobjName) Then
                    foundDiagramm = DiagramList.getDiagramm(chtobjName)
                    With foundDiagramm
                        top = .top
                        left = .left
                        width = .width
                        height = .height
                    End With
                End If
                
            Catch ex As Exception


            End Try
        End If



        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False


        titelTeile(0) = portfolioDiagrammtitel(PTpfdk.Auslastung) & " (" & awinSettings.kapaEinheit & ")"


        titelTeilLaengen(0) = titelTeile(0).Length + 1
        titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)

        von = showRangeLeft
        bis = showRangeRight



        ReDim Xdatenreihe(2)
        ReDim datenreihe(2)

        'Xdatenreihe(0) = "Auslastung"
        'Xdatenreihe(1) = "Über-Auslastung"
        'Xdatenreihe(2) = "Unter-Auslastung"
        Xdatenreihe(0) = repMessages.getmsg(93)
        Xdatenreihe(1) = repMessages.getmsg(94)
        Xdatenreihe(2) = repMessages.getmsg(95)


        datenreihe(0) = ShowProjekte.getAuslastungsValues(0).Sum
        datenreihe(1) = ShowProjekte.getAuslastungsValues(1).Sum
        datenreihe(2) = ShowProjekte.getAuslastungsValues(2).Sum


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

                If .ChartObjects(i).Name = chtobjName Then
                    found = True
                    repObj = CType(.ChartObjects(i), Excel.ChartObject)
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


                    ' remove old series
                    Try
                        Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                        Do While anz > 0
                            .SeriesCollection(1).Delete()
                            anz = anz - 1
                        Loop
                    Catch ex As Exception

                    End Try


                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        .Name = "Auslastung"
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


                    Dim achieved As Boolean = False
                    Dim anzahlVersuche As Integer = 0
                    Dim errmsg As String = ""
                    Do While Not achieved And anzahlVersuche < 10
                        Try
                            'Call Sleep(100)
                            .Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=currentSheetName)
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

                repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then

                    Dim prcDiagram As New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    Dim prcChart As New clsEventsPrcCharts
                    prcChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
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

            ' nicht aktivieren, weil ein Chart dann im Mass-Edit Fenster nicht mehr selektierbar ist nicht mehr selektierbar ist ... 
            ' wenn es geschützt war .. 
            'If wasProtected And visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
            '    .Protect(Password:="x", UserInterfaceOnly:=True, _
            '                 AllowFormattingCells:=True, _
            '                 AllowInsertingColumns:=False,
            '                 AllowInsertingRows:=True, _
            '                 AllowDeletingColumns:=False, _
            '                 AllowDeletingRows:=True, _
            '                 AllowSorting:=True, _
            '                 AllowFiltering:=True)
            '    .EnableSelection = XlEnableSelection.xlUnlockedCells
            '    .EnableAutoFilter = True
            'End If

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

        Dim currentSheetName As String

        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentSheetName = arrWsNames(ptTables.MPT)
        Else
            currentSheetName = arrWsNames(ptTables.meRC)
        End If


        ReDim Xdatenreihe(3)
        ReDim datenreihe(3)


        If future = -1 Then

            Dim myCollection As New Collection
            myCollection.Add("ZieleV")
            chtobjName = calcChartKennung("pf", PTpfdk.ZieleV, myCollection)
            If showRangeLeft <= heuteColumn Then
                titelTeile(0) = summentitel6
                titelTeile(1) = textZeitraum(showRangeLeft, heuteColumn)
                'Xdatenreihe(0) = "keine Information"
                'Xdatenreihe(1) = "erreicht"
                'Xdatenreihe(2) = "mit Einschränkungen"
                'Xdatenreihe(3) = "nicht erreicht"
                Xdatenreihe(0) = repMessages.getmsg(97)
                Xdatenreihe(1) = repMessages.getmsg(98)
                Xdatenreihe(2) = repMessages.getmsg(99)
                Xdatenreihe(3) = repMessages.getmsg(100)
            Else
                'Throw New ArgumentException("der betrachtete Bereich liegt vollständig in der Zukunft ... es gibt keine erreichten Ziele")
                Throw New ArgumentException(repMessages.getmsg(101))
            End If


        ElseIf future = 1 Then
            Dim myCollection As New Collection
            myCollection.Add("ZieleF")
            chtobjName = calcChartKennung("pf", PTpfdk.ZieleF, myCollection)
            If heuteColumn + 1 <= showRangeRight Then
                titelTeile(0) = summentitel7
                titelTeile(1) = textZeitraum(getColumnOfDate(Date.Now) + 1, showRangeRight)
                'Xdatenreihe(0) = "keine Information"
                'Xdatenreihe(1) = "wird erreicht"
                'Xdatenreihe(2) = "Unsicherheiten"
                'Xdatenreihe(3) = "erhebliche Risiken"
                Xdatenreihe(0) = repMessages.getmsg(97)
                Xdatenreihe(1) = repMessages.getmsg(102)
                Xdatenreihe(2) = repMessages.getmsg(103)
                Xdatenreihe(3) = repMessages.getmsg(104)
            Else
                'Throw New ArgumentException("der betrachtete Bereich liegt vollständig in der Vergangenheit ... es gibt keine Prognose Werte")
                Throw New ArgumentException(repMessages.getmsg(105))
            End If

        Else
            'titelTeile(0) = summentitel8
            'titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
            'Xdatenreihe(0) = "keine Information"
            'Xdatenreihe(1) = "wurde/wird erreicht"
            'Xdatenreihe(2) = "mit Einschränkungen/Unsicherheiten"
            'Xdatenreihe(3) = "nicht erreicht/erhebliche Risiken"
            'Throw New ArgumentException("keine Angabe in Zielerreichungsdiagramm, ob Vergangenheit oder Zukunft betrachtet werden soll ")
            Throw New ArgumentException(repMessages.getmsg(106))
        End If


        'With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

        titelTeilLaengen(0) = titelTeile(0).Length
        titelTeilLaengen(1) = titelTeile(1).Length



        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)


        von = showRangeLeft
        bis = showRangeRight

        ' jetzt prüfen, ob es bereits gespeicherte Werte für top, left, ... gibt ;
        ' Wenn ja : übernehmen


        If Not calledfromReporting Then
            Dim foundDiagramm As clsDiagramm = Nothing

            Try
                If DiagramList.contains(chtobjName) Then
                    foundDiagramm = DiagramList.getDiagramm(chtobjName)
                    With foundDiagramm
                        top = .top
                        left = .left
                        width = .width
                        height = .height
                    End With
                End If
                
            Catch ex As Exception

            End Try
        End If


        datenreihe(0) = ShowProjekte.getColorsInMonth(0, future).Sum
        datenreihe(1) = ShowProjekte.getColorsInMonth(1, future).Sum
        datenreihe(2) = ShowProjekte.getColorsInMonth(2, future).Sum
        datenreihe(3) = ShowProjekte.getColorsInMonth(3, future).Sum

        If datenreihe.Sum = 0 Then

            If future < 0 Then
                'Throw New Exception("es gibt im betrachteten Zeitraum keine Ergebnisse aus der Vergangenheit ...")
                Throw New Exception(repMessages.getmsg(107))

            ElseIf future > 0 Then
                'Throw New Exception("es gibt im betrachteten Zeitraum keine geplanten, zukünftigen Ergebnisse ...")
                Throw New Exception(repMessages.getmsg(108))
            Else
                'Throw New Exception("es gibt im betrachteten Zeitraum keine vergangenen oder zukünftigen Ergebnisse ...")
                Throw New Exception(repMessages.getmsg(109))
            End If

        Else

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False

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

                    Try
                        chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Name
                    Catch ex As Exception
                        chtTitle = " "
                    End Try


                    If chtobjName = .ChartObjects(i).Name Then
                        found = True
                        repObj = CType(.ChartObjects(i), Excel.ChartObject)
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

                        ' remove old series
                        Try
                            Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                            Do While anz > 0
                                .SeriesCollection(1).Delete()
                                anz = anz - 1
                            Loop
                        Catch ex As Exception

                        End Try


                        With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                            .Name = "Status-Übersicht"
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

                        Dim achieved As Boolean = False
                        Dim anzahlVersuche As Integer = 0
                        Dim errmsg As String = ""
                        Do While Not achieved And anzahlVersuche < 10
                            Try
                                'Call Sleep(100)
                                .Location(Where:=Excel.XlChartLocation.xlLocationAsObject, Name:=currentSheetName)
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



                    End With
                    With CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                        .Top = top
                        .Left = left
                        .Width = width
                        .Height = height
                        .Name = chtobjName

                    End With

                    If isCockpitChart Then
                        Try
                            With CType(appInstance.ActiveSheet, Excel.Worksheet)
                                .Shapes.Item(chtobjName).Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                            End With
                        Catch ex As Exception

                        End Try
                    Else
                        'Call awinScrollintoView()
                    End If

                    repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)


                    ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                    ' aufgerufen wurde

                    If Not calledfromReporting Then

                        Dim prcDiagram As New clsDiagramm



                        ' Anfang Event Handling für Chart 
                        Dim prcChart As New clsEventsPrcCharts
                        prcChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
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

                End If

                ' Schutz nicht mehr aktivieren, weil charts dann nicht mehr selektierbar sind .  
                ' '' wenn es geschützt war .. 
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

            End With

            appInstance.EnableEvents = formerEE
            appInstance.ScreenUpdating = formerSU

        End If


    End Sub

    Sub awinUpdateAuslastungsDiagramm(ByVal repObj As Excel.ChartObject)

        Dim i As Integer

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


        titelTeile(0) = portfolioDiagrammtitel(PTpfdk.Auslastung) & " (" & awinSettings.kapaEinheit & ")"

        titelTeilLaengen(0) = titelTeile(0).Length + 1
        titelTeile(1) = textZeitraum(showRangeLeft, showRangeRight)
        titelTeilLaengen(1) = titelTeile(1).Length
        diagramTitle = titelTeile(0) & vbLf & titelTeile(1)
        kennung = titelTeile(0)



        von = showRangeLeft
        bis = showRangeRight


        ReDim Xdatenreihe(2)
        ReDim datenreihe(2)

        'Xdatenreihe(0) = "Auslastung"
        'Xdatenreihe(1) = "Über-Auslastung"
        'Xdatenreihe(2) = "Unter-Auslastung"
        Xdatenreihe(0) = repMessages.getmsg(93)
        Xdatenreihe(1) = repMessages.getmsg(94)
        Xdatenreihe(2) = repMessages.getmsg(95)


        datenreihe(0) = ShowProjekte.getAuslastungsValues(0).Sum
        datenreihe(1) = ShowProjekte.getAuslastungsValues(1).Sum
        datenreihe(2) = ShowProjekte.getAuslastungsValues(2).Sum


        'With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)




        With repObj.Chart


            ' remove old series
            Try
                Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                Do While anz > 0
                    .SeriesCollection(1).Delete()
                    anz = anz - 1
                Loop
            Catch ex As Exception

            End Try


            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                .Name = "Auslastung"
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

        'repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
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




        'End With

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU




    End Sub



    ''' <summary>
    ''' aktualisiert das Portfolio diagramm Ergebnis-Kennzahl mit Übersicht 
    ''' Projekt-Ergebnis, Kosten der Überauslastung, Unterauslastung, Ergbnis-Kennzahl  
    ''' </summary>
    ''' <param name="chtobj">Chart, das aktualisiert werden soll</param>
    ''' <remarks></remarks>
    Sub awinUpdateErgebnisDiagramm(ByVal chtobj As ChartObject)


        Dim diagramTitle As String

        Dim minScale As Double
        Dim maxscale As Double
        Dim Xdatenreihe(3) As String
        Dim valueDatenreihe1(3) As Double
        Dim valueDatenreihe2(3) As Double
        Dim itemColor(3) As Object
        Dim itemValue(3) As Double
        Dim earnedValue As Double, additionalCostExt As Double, internwithoutProject As Double
        Dim ertragsWert As Double
        Dim zeitraumBudget As Double, zeitraumCost As Double, zeitraumRisiko As Double

        Dim mycollection As New Collection


        Xdatenreihe(0) = "Summe Projekt-Ergebnisse (Risiko-gewichtet)"
        'Xdatenreihe(1) = "Risiko-Abschlag"
        Xdatenreihe(1) = "Mehrkosten wegen Überauslastung"
        Xdatenreihe(2) = "Opportunitätskosten durch Unterauslastung"
        Xdatenreihe(3) = "Ergebnis-Kennzahl"


        Dim positiv As Boolean = True


        ' neu 
        ' Ausrechnen amteiliges Budget, das im Zeitraum zur Verfügung steht und der im Zeitraum anfallenden Kosten  
        zeitraumBudget = System.Math.Round(ShowProjekte.getBudgetValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        zeitraumCost = System.Math.Round(ShowProjekte.getTotalCostValuesInMonth.Sum, mode:=MidpointRounding.ToEven)

        ' das ist der Risiko Abschlag  
        zeitraumRisiko = System.Math.Round(ShowProjekte.getWeightedRiskValuesInMonth.Sum, mode:=MidpointRounding.ToEven)


        ' das ist der Earned Value 
        earnedValue = zeitraumBudget - (zeitraumCost + zeitraumRisiko)

        itemValue(0) = earnedValue

        If earnedValue < 0 Then
            minScale = earnedValue * 1.2
        Else
            minScale = 0
        End If

        maxscale = zeitraumBudget * 1.2

        If itemValue(0) >= 0 Then
            itemColor(0) = ergebnisfarbe1
        Else
            itemColor(0) = farbeExterne
        End If

        Dim currentWert As Double = itemValue(0)


        ' das sind die Zusatzkosten, die durch Externe (wg Überauslastung) verursacht werden
        additionalCostExt = System.Math.Round(ShowProjekte.getCosteValuesInMonth(True).Sum, mode:=MidpointRounding.ToEven)
        itemValue(1) = additionalCostExt
        itemColor(1) = farbeExterne

        ' das sind die durch Unterauslastung verursachten Kosten , also Personal-Kosten von Leuten, die in keinem Projekt sind
        internwithoutProject = System.Math.Round(ShowProjekte.getCostoValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        itemValue(2) = internwithoutProject
        itemColor(2) = awinSettings.AmpelGelb

        ' das ist der Ertrag 
        ertragsWert = earnedValue - (additionalCostExt + internwithoutProject)
        itemValue(3) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(3) = ergebnisfarbe2
        Else
            itemColor(3) = farbeExterne
        End If


        diagramTitle = portfolioDiagrammtitel(PTpfdk.ErgebnisWasserfall) & " " & _
                                       textZeitraum(showRangeLeft, showRangeRight)

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False


        If ertragsWert < 0 Then
            minScale = System.Math.Round(ertragsWert, mode:=MidpointRounding.ToEven)
        Else
            minScale = 0
        End If

        'Dim htxt As String

        Dim valueCrossesNull As Boolean = False


        'With CType(appInstance.Workbooks.Item("Projectboard.xlsx").Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

        With chtobj.Chart
            ' remove old series
            Try
                Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                Do While anz > 0
                    .SeriesCollection(1).Delete()
                    anz = anz - 1
                Loop
            Catch ex As Exception

            End Try
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
            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                .Name = "Bottom"
                .HasDataLabels = False
                .Interior.ColorIndex = -4142
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

            With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                .Name = "Top"
                .HasDataLabels = True
                .Values = valueDatenreihe2
                .XValues = Xdatenreihe
                .ChartType = Excel.XlChartType.xlColumnStacked

                For iv = 0 To 3

                    With .Points(iv + 1)
                        .HasDataLabel = True
                        .DataLabel.text = Format(itemValue(iv), "###,###0") & " T€"
                        .Interior.color = itemColor(iv)
                        ' ur:17.7.2014 fontsize bei update nicht ändern für die Legend
                        '.DataLabel.Font.Size = awinSettings.fontsizeLegend
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
                        .MinimumScale = System.Math.Round((minScale - 1), mode:=MidpointRounding.ToEven)
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
        'End With


        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


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
    Sub awinCreateErgebnisDiagramm(ByRef repObj As Excel.ChartObject, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
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
        Dim earnedValue As Double, additionalCostExt As Double, internwithoutProject As Double
        Dim ertragsWert As Double
        Dim zeitraumBudget As Double, zeitraumCost As Double, zeitraumRisiko As Double


        Dim mycollection As New Collection
        Dim chtobjName As String

        Dim currentSheetName As String

        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentSheetName = arrWsNames(ptTables.MPT)
        Else
            currentSheetName = arrWsNames(ptTables.meRC)
        End If

        'Dim hproj As clsProjekt
        Dim ErgebnisListeR As New Collection

        mycollection.Add("Ergebniskennzahl")
        chtobjName = calcChartKennung("pf", PTpfdk.ErgebnisWasserfall, mycollection)
        mycollection.Clear()

        If Not calledfromReporting Then

            Dim foundDiagramm As clsDiagramm = Nothing

            ' wenn die Werte für dieses Diagramm bereits einmal gespeichert wurden ... -> übernehmen 
            Try
                If DiagramList.contains(chtobjName) Then
                    foundDiagramm = DiagramList.getDiagramm(chtobjName)
                    With foundDiagramm
                        top = .top
                        left = .left
                        width = .width
                        height = .height
                    End With
                End If
                
            Catch ex As Exception


            End Try
        End If


        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False



        'Xdatenreihe(0) = "Summe Projekt-Ergebnisse (Risiko-gewichtet)"
        ''Xdatenreihe(1) = "Risiko-Abschlag"
        'Xdatenreihe(1) = "Mehrkosten wegen Überauslastung"
        'Xdatenreihe(2) = "Opportunitätskosten durch Unterauslastung"
        'Xdatenreihe(3) = "Ergebnis-Kennzahl"

        Xdatenreihe(0) = repMessages.getmsg(151)
        'Xdatenreihe(1) = "Risiko-Abschlag"
        Xdatenreihe(1) = repMessages.getmsg(152)
        Xdatenreihe(2) = repMessages.getmsg(153)
        Xdatenreihe(3) = repMessages.getmsg(154)

        Dim positiv As Boolean = True


        ' Ausrechnen amteiliges Budget, das im Zeitraum zur Verfügung steht und der im Zeitraum anfallenden Kosten  
        zeitraumBudget = System.Math.Round(ShowProjekte.getBudgetValuesInMonth.Sum, mode:=MidpointRounding.ToEven)

        Dim pCost As Double = System.Math.Round(ShowProjekte.getCostGpValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        Dim oCost As Double = System.Math.Round(ShowProjekte.getOtherCostValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        'zeitraumCost = System.Math.Round(ShowProjekte.getTotalCostValuesInMonth.Sum, mode:=MidpointRounding.ToEven) 
        zeitraumCost = pCost + oCost


        ' das ist der Risiko Abschlag  
        zeitraumRisiko = System.Math.Round(ShowProjekte.getWeightedRiskValuesInMonth.Sum, mode:=MidpointRounding.ToEven)


        ' das ist der Earned Value 
        earnedValue = zeitraumBudget - (zeitraumCost + zeitraumRisiko)


        itemValue(0) = earnedValue

        If itemValue(0) >= 0 Then
            itemColor(0) = ergebnisfarbe1
        Else
            itemColor(0) = farbeExterne
        End If

        Dim currentWert As Double = itemValue(0)


        ' das sind die Zusatzkosten, die durch Überauslastung) verursacht werden
        additionalCostExt = System.Math.Round(ShowProjekte.getCosteValuesInMonth(True).Sum, mode:=MidpointRounding.ToEven)
        itemValue(1) = additionalCostExt
        itemColor(1) = farbeExterne

        ' das sind die durch Unterauslastung verursachten Kosten , also Personal-Kosten von Leuten, die in keinem Projekt sind
        internwithoutProject = System.Math.Round(ShowProjekte.getCostoValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        itemValue(2) = internwithoutProject
        itemColor(2) = awinSettings.AmpelGelb

        ' das ist der Ertrag 
        ertragsWert = earnedValue - (additionalCostExt + internwithoutProject)
        itemValue(3) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(3) = ergebnisfarbe2
        Else
            itemColor(3) = farbeExterne
        End If

        diagramTitle = portfolioDiagrammtitel(PTpfdk.ErgebnisWasserfall) & " " & _
                                        textZeitraum(showRangeLeft, showRangeRight)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


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


                If .ChartObjects(i).Name = chtobjName Then
                    found = True
                Else
                    i = i + 1
                End If

            End While



            If found Then
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                If ertragsWert < 0 Then
                    minScale = System.Math.Round(ertragsWert, mode:=MidpointRounding.ToEven)
                Else
                    minScale = 0
                End If

                'Dim htxt As String
                Dim valueCrossesNull As Boolean = False

                With appInstance.Charts.Add
                    ' remove extra series
                    ' remove old series
                    Try
                        Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                        Do While anz > 0
                            .SeriesCollection(1).Delete()
                            anz = anz - 1
                        Loop
                    Catch ex As Exception

                    End Try
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
                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Bottom"
                        .Name = repMessages.getmsg(149)
                        .HasDataLabels = False
                        .Interior.ColorIndex = -4142
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

                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Top"
                        .Name = repMessages.getmsg(150)
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
                                .MinimumScale = System.Math.Round((minScale - 1), mode:=MidpointRounding.ToEven)
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

                End With

                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .name = chtobjName
                End With

                repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then

                    Dim prcDiagram As New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    Dim prcChart As New clsEventsPrcCharts
                    prcChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
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

            ' Schutz nicht mehr aktivieren, weil Chart dann nicht mehr selektierbar ist ... 
            '' wenn es geschützt war .. 
            'If wasProtected And visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
            '    .Protect(Password:="x", UserInterfaceOnly:=True, _
            '                 AllowFormattingCells:=True, _
            '                 AllowInsertingColumns:=False,
            '                 AllowInsertingRows:=True, _
            '                 AllowDeletingColumns:=False, _
            '                 AllowDeletingRows:=True, _
            '                 AllowSorting:=True, _
            '                 AllowFiltering:=True)
            '    .EnableSelection = XlEnableSelection.xlUnlockedCells
            '    .EnableAutoFilter = True
            'End If

        End With

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    ''' <summary>
    ''' zeigt für das Portfolio an: Budget, Risiko, Personalkosten, Sonstige Kosten, Ergebnis 
    ''' wird momentan nicht benutzt, wenn mal wieder Risko Kosten berücksichtigt werden sollen , dann diese Routine reaktivieren ...  
    ''' </summary>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <remarks></remarks>
    Sub awinCreateBudgetErgebnisDiagramm_Deprecated(ByRef repObj As Excel.ChartObject, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                   ByVal isCockpitChart As Boolean, ByVal calledfromReporting As Boolean)

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean
        'Dim plen As Integer
        Dim i As Integer
        Dim minScale As Double
        Dim maxScale As Double
        Dim Xdatenreihe(4) As String
        Dim valueDatenreihe1(4) As Double
        Dim valueDatenreihe2(4) As Double
        Dim itemColor(4) As Object
        Dim itemValue(4) As Double

        Dim budgetSum As Double, pCost As Double, oCost As Double, riskValue As Double
        Dim ertragsWert As Double
        Dim minColumn As Integer, maxColumn As Integer, heuteColumn As Integer, heuteIndex As Integer
        Dim future As Boolean = False

        heuteColumn = getColumnOfDate(Date.Today)
        heuteIndex = heuteColumn - showRangeLeft

        minColumn = showRangeLeft
        maxColumn = showRangeRight

        Dim mycollection As New Collection
        Dim chtobjName As String

        'Dim hproj As clsProjekt
        Dim ErgebnisListeR As New Collection

        Dim currentSheetName As String

        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentSheetName = arrWsNames(ptTables.MPT)
        Else
            currentSheetName = arrWsNames(ptTables.meRC)
        End If

        mycollection.Add("Projektergebnisse")
        chtobjName = calcChartKennung("pf", PTpfdk.Budget, mycollection)
        mycollection.Clear()

        If Not calledfromReporting Then

            Dim foundDiagramm As clsDiagramm = Nothing

            ' wenn die Werte für dieses Diagramm bereits einmal gespeichert wurden ... -> übernehmen 
            Try
                If DiagramList.contains(chtobjName) Then
                    foundDiagramm = DiagramList.getDiagramm(chtobjName)
                    With foundDiagramm
                        top = .top
                        left = .left
                        width = .width
                        height = .height
                    End With
                End If

            Catch ex As Exception


            End Try
        End If


        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False

        Xdatenreihe(0) = repMessages.getmsg(49)
        Xdatenreihe(1) = repMessages.getmsg(50)
        Xdatenreihe(2) = repMessages.getmsg(51)
        Xdatenreihe(3) = repMessages.getmsg(52)
        Xdatenreihe(4) = repMessages.getmsg(53)

        Dim positiv As Boolean = True

        ' Ausrechnen amteiliges Budget, das i Zeitraum zur Verfügung steht und der im Zeitraum anfallenden Kosten  
        budgetSum = System.Math.Round(ShowProjekte.getBudgetValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        pCost = System.Math.Round(ShowProjekte.getCostGpValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        oCost = System.Math.Round(ShowProjekte.getOtherCostValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        riskValue = System.Math.Round(ShowProjekte.getWeightedRiskValuesInMonth.Sum, mode:=MidpointRounding.ToEven)

        ertragsWert = budgetSum - (riskValue + pCost + oCost)

        maxScale = budgetSum * 1.2
        If ertragsWert < 0 Then
            minScale = ertragsWert * 1.2
        Else
            minScale = 0
        End If


        itemValue(0) = budgetSum
        itemColor(0) = ergebnisfarbe1


        Dim currentWert As Double = itemValue(0)


        ' das ist der Risiko-Abschlag 
        itemValue(1) = riskValue
        itemColor(1) = iProjektFarbe

        ' das sind die Personalkosten
        itemValue(2) = pCost
        itemColor(2) = farbeExterne

        ' das sind die Other Cost 
        itemValue(3) = oCost
        itemColor(3) = farbeExterne

        ' das ist der Ertrag 
        itemValue(4) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(4) = ergebnisfarbe2
        Else
            itemColor(4) = farbeExterne
        End If

        diagramTitle = portfolioDiagrammtitel(PTpfdk.Budget) & " " & textZeitraum(showRangeLeft, showRangeRight)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


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


                If .ChartObjects(i).Name = chtobjName Then
                    found = True
                Else
                    i = i + 1
                End If

            End While



            If found Then
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                'If ertragsWert < 0 Then
                '    minScale = System.Math.Round(ertragsWert, mode:=MidpointRounding.ToEven)
                'Else
                '    minScale = 0
                'End If

                'Dim htxt As String
                Dim valueCrossesNull As Boolean = False

                With appInstance.Charts.Add
                    ' remove old series
                    Try
                        Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                        Do While anz > 0
                            .SeriesCollection(1).Delete()
                            anz = anz - 1
                        Loop
                    Catch ex As Exception

                    End Try
                    Dim crossindex As Integer = -1

                    ' bestimmen des Anfangs  
                    Dim iv = 0
                    valueDatenreihe1(iv) = 0
                    valueDatenreihe2(iv) = itemValue(iv)
                    currentWert = itemValue(iv)
                    Dim formerValue As Double = currentWert
                    Dim negativeFromNull As Boolean = False

                    ' alle nächsten Zwischen-Werte 
                    For iv = 1 To 3
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
                    iv = 4
                    valueDatenreihe1(iv) = 0
                    valueDatenreihe2(iv) = itemValue(iv)



                    'series
                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Bottom"
                        .Name = repMessages.getmsg(149)
                        .HasDataLabels = False
                        .Interior.ColorIndex = -4142
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

                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Top"
                        .Name = repMessages.getmsg(150)
                        .HasDataLabels = True
                        .Values = valueDatenreihe2
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked

                        For iv = 0 To 4

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
                        With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                            .HasTitle = False
                            .HasMajorGridlines = False
                            .HasMinorGridlines = False
                            .MinimumScale = minScale
                            .MaximumScale = maxScale
                            .MaximumScaleIsAuto = False
                            .MinimumScaleIsAuto = False

                            'If minScale < 0 Then
                            '    .MinimumScale = System.Math.Round((minScale - 1), mode:=MidpointRounding.ToEven)
                            'Else
                            '    .MinimumScale = 0
                            'End If
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

                End With

                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .name = chtobjName
                End With

                repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then

                    Dim prcDiagram As New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    Dim prcChart As New clsEventsPrcCharts
                    prcChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
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

            ' nicht mehr schützen, weil Charts dann nicht mehr selektierbar sind 
            '' wenn es geschützt war .. 
            'If wasProtected And visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
            '    .Protect(Password:="x", UserInterfaceOnly:=True, _
            '                 AllowFormattingCells:=True, _
            '                 AllowInsertingColumns:=False,
            '                 AllowInsertingRows:=True, _
            '                 AllowDeletingColumns:=False, _
            '                 AllowDeletingRows:=True, _
            '                 AllowSorting:=True, _
            '                 AllowFiltering:=True)
            '    .EnableSelection = XlEnableSelection.xlUnlockedCells
            '    .EnableAutoFilter = True
            'End If

        End With

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    ''' <summary>
    ''' zeigt für das Portfolio an: Budget, Personalkosten, Sonstige Kosten, Ergebnis 
    ''' zeigt das selbe an wie awinCreateErgebnisDiagramm, aber ohne Risiko BEitrag
    ''' </summary>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <remarks></remarks>
    Sub awinCreateBudgetErgebnisDiagramm(ByRef repObj As Excel.ChartObject, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                   ByVal isCockpitChart As Boolean, ByVal calledfromReporting As Boolean)

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean = False
        'Dim plen As Integer
        Dim i As Integer
        Dim minScale As Double
        Dim maxScale As Double
        Dim Xdatenreihe(3) As String
        Dim valueDatenreihe1(3) As Double
        Dim valueDatenreihe2(3) As Double
        Dim itemColor(3) As Object
        Dim itemValue(3) As Double

        Dim budgetSum As Double, pCost As Double, oCost As Double
        Dim ertragsWert As Double
        Dim minColumn As Integer, maxColumn As Integer, heuteColumn As Integer, heuteIndex As Integer
        Dim future As Boolean = False
        Dim newChtObj As Excel.ChartObject = Nothing

        heuteColumn = getColumnOfDate(Date.Today)
        heuteIndex = heuteColumn - showRangeLeft

        minColumn = showRangeLeft
        maxColumn = showRangeRight

        Dim mycollection As New Collection
        Dim chtobjName As String

        'Dim hproj As clsProjekt
        Dim ErgebnisListeR As New Collection

        Dim currentSheetName As String

        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            If calledfromReporting Then
                currentSheetName = arrWsNames(ptTables.repCharts)
            Else
                currentSheetName = arrWsNames(ptTables.mptPfCharts)
            End If

        Else
            currentSheetName = arrWsNames(ptTables.meCharts)
        End If

        mycollection.Add("Projektergebnisse")
        chtobjName = calcChartKennung("pf", PTpfdk.Budget, mycollection)
        mycollection.Clear()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False

        Xdatenreihe(0) = repMessages.getmsg(49)
        Xdatenreihe(1) = repMessages.getmsg(51)
        Xdatenreihe(2) = repMessages.getmsg(52)
        Xdatenreihe(3) = repMessages.getmsg(53)


        Dim positiv As Boolean = True

        ' Ausrechnen amteiliges Budget, das i Zeitraum zur Verfügung steht und der im Zeitraum anfallenden Kosten  
        budgetSum = System.Math.Round(ShowProjekte.getBudgetValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        pCost = System.Math.Round(ShowProjekte.getCostGpValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        oCost = System.Math.Round(ShowProjekte.getOtherCostValuesInMonth.Sum, mode:=MidpointRounding.ToEven)

        ertragsWert = budgetSum - (pCost + oCost)

        maxScale = budgetSum * 1.2
        If ertragsWert < 0 Then
            minScale = ertragsWert * 1.2
        Else
            minScale = 0
        End If


        itemValue(0) = budgetSum
        itemColor(0) = ergebnisfarbe1


        Dim currentWert As Double = itemValue(0)

        ' das sind die Personalkosten
        itemValue(1) = pCost
        itemColor(1) = farbeExterne

        ' das sind die Other Cost 
        itemValue(2) = oCost
        itemColor(2) = farbeExterne

        ' das ist der Ertrag 
        itemValue(3) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(3) = ergebnisfarbe2
        Else
            itemColor(3) = farbeExterne
        End If

        'diagramTitle = portfolioDiagrammtitel(PTpfdk.Budget) & " " & textZeitraum(showRangeLeft, showRangeRight)
        If getColumnOfDate(Date.Now) > showRangeRight Then
            diagramTitle = "Portfolio " & textZeitraum(showRangeLeft, showRangeRight)
        Else
            diagramTitle = "Forecast Portfolio " & textZeitraum(showRangeLeft, showRangeRight)
        End If


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


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
            While i <= anzDiagrams And Not found


                If .ChartObjects(i).Name = chtobjName Then
                    found = True
                Else
                    i = i + 1
                End If

            End While



            If found Then
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                Dim valueCrossesNull As Boolean = False

                newChtObj = CType(CType(CType(appInstance.Workbooks.Item(myProjektTafel),  _
                            Excel.Workbook).Worksheets.Item(currentSheetName),  _
                            Excel.Worksheet).ChartObjects, Excel.ChartObjects).Add(left, top, width, height)

                With newChtObj.Chart
                    ' remove old series
                    Try
                        Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                        Do While anz > 0
                            .SeriesCollection(1).Delete()
                            anz = anz - 1
                        Loop
                    Catch ex As Exception

                    End Try
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
                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Bottom"
                        .Name = repMessages.getmsg(149)
                        .HasDataLabels = False
                        .Interior.ColorIndex = -4142
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

                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Top"
                        .Name = repMessages.getmsg(150)
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
                        With CType(.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                            .HasTitle = False
                            .HasMajorGridlines = False
                            .HasMinorGridlines = False
                            .MinimumScale = minScale
                            .MaximumScale = maxScale
                            .MaximumScaleIsAuto = False
                            .MinimumScaleIsAuto = False

                            'If minScale < 0 Then
                            '    .MinimumScale = System.Math.Round((minScale - 1), mode:=MidpointRounding.ToEven)
                            'Else
                            '    .MinimumScale = 0
                            'End If
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
                    .ChartTitle.Font.Size = awinSettings.fontsizeTitle

                End With

                'With .ChartObjects(anzDiagrams + 1)
                With newChtObj
                    '.Top = top
                    '.Left = left
                    '.Width = width
                    '.Height = height
                    .Name = chtobjName
                End With

                'repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                repObj = newChtObj

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then

                    Dim prcDiagram As New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    Dim prcChart As New clsEventsPrcCharts
                    prcChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
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

        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    ''' <summary>
    ''' zeigt an, wieviel bisher vom Budget aufgebraucht wurde und was noch aussteht
    ''' mal noch drin lassen, ob das noch gebraucht wird ... 
    ''' </summary>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="isCockpitChart"></param>
    ''' <remarks></remarks>
    Sub awinCreateBudgetErgebnisDiagrammOld(ByRef repObj As Excel.ChartObject, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, _
                                   ByVal isCockpitChart As Boolean, ByVal calledfromReporting As Boolean)

        Dim diagramTitle As String
        Dim anzDiagrams As Integer
        Dim found As Boolean
        'Dim plen As Integer
        Dim i As Integer
        Dim minScale As Double
        Dim Xdatenreihe(4) As String
        Dim valueDatenreihe1(4) As Double
        Dim valueDatenreihe2(4) As Double
        Dim itemColor(4) As Object
        Dim itemValue(4) As Double

        Dim budgetSum As Double, costPast As Double, costFuture As Double, riskValue As Double
        Dim zeitraumCost As Double
        Dim costValues() As Double
        Dim ertragsWert As Double
        Dim minColumn As Integer, maxColumn As Integer, heuteColumn As Integer, heuteIndex As Integer
        Dim future As Boolean = False

        heuteColumn = getColumnOfDate(Date.Today)
        heuteIndex = heuteColumn - showRangeLeft

        minColumn = showRangeLeft
        maxColumn = showRangeRight

        Dim mycollection As New Collection
        Dim chtobjName As String

        'Dim hproj As clsProjekt
        Dim ErgebnisListeR As New Collection

        mycollection.Add("Projektergebnisse")
        chtobjName = calcChartKennung("pf", PTpfdk.Budget, mycollection)
        mycollection.Clear()

        If Not calledfromReporting Then

            Dim foundDiagramm As clsDiagramm = Nothing

            ' wenn die Werte für dieses Diagramm bereits einmal gespeichert wurden ... -> übernehmen 
            Try
                If DiagramList.contains(chtobjName) Then
                    foundDiagramm = DiagramList.getDiagramm(chtobjName)
                    With foundDiagramm
                        top = .top
                        left = .left
                        width = .width
                        height = .height
                    End With
                End If

            Catch ex As Exception


            End Try
        End If


        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False



        'Xdatenreihe(0) = "Budget Summe"
        'If heuteColumn >= minColumn + 1 And heuteColumn <= maxColumn Then
        '    Xdatenreihe(2) = "bisherige Kosten" & vbLf & textZeitraum(minColumn, heuteColumn - 1)
        '    Xdatenreihe(3) = "Prognose Kosten" & vbLf & textZeitraum(heuteColumn, maxColumn)
        'ElseIf heuteColumn > maxColumn Then
        '    future = False
        '    Xdatenreihe(2) = "bisherige Kosten" & vbLf & textZeitraum(minColumn, maxColumn)
        '    Xdatenreihe(3) = "Prognose Kosten" & vbLf & "existieren nicht"
        'ElseIf heuteColumn <= minColumn Then
        '    future = True
        '    Xdatenreihe(2) = "bisherige Kosten" & vbLf & "existieren nicht"
        '    Xdatenreihe(3) = "Prognose Kosten" & vbLf & textZeitraum(minColumn, maxColumn)
        'End If


        Xdatenreihe(0) = repMessages.getmsg(144)
        If heuteColumn >= minColumn + 1 And heuteColumn <= maxColumn Then
            Xdatenreihe(2) = repMessages.getmsg(145) & vbLf & textZeitraum(minColumn, heuteColumn - 1)
            Xdatenreihe(3) = repMessages.getmsg(146) & vbLf & textZeitraum(heuteColumn, maxColumn)
        ElseIf heuteColumn > maxColumn Then
            future = False
            Xdatenreihe(2) = repMessages.getmsg(145) & vbLf & textZeitraum(minColumn, maxColumn)
            Xdatenreihe(3) = repMessages.getmsg(146) & vbLf & repMessages.getmsg(147)
        ElseIf heuteColumn <= minColumn Then
            future = True
            Xdatenreihe(2) = repMessages.getmsg(145) & vbLf & repMessages.getmsg(147)
            Xdatenreihe(3) = repMessages.getmsg(146) & vbLf & textZeitraum(minColumn, maxColumn)
        End If

        'Xdatenreihe(1) = "Risiko-Abschlag"
        'Xdatenreihe(4) = "Ergebnis"
        Xdatenreihe(1) = repMessages.getmsg(50)
        Xdatenreihe(4) = repMessages.getmsg(148)

        Dim positiv As Boolean = True

        ' Ausrechnen amteiliges Budget, das i Zeitraum zur Verfügung steht und der im Zeitraum anfallenden Kosten  
        budgetSum = System.Math.Round(ShowProjekte.getBudgetValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        costValues = ShowProjekte.getTotalCostValuesInMonth
        zeitraumCost = System.Math.Round(costValues.Sum, mode:=MidpointRounding.ToEven)


        Dim zeitraumLaenge = costValues.Length - 1
        costPast = 0
        For i = 0 To Min(heuteIndex - 1, zeitraumLaenge)
            costPast = costPast + costValues(i)
        Next
        costPast = System.Math.Round(costPast, mode:=MidpointRounding.ToEven)


        costFuture = 0
        For i = Max(0, heuteIndex) To zeitraumLaenge
            costFuture = costFuture + costValues(i)
        Next
        costFuture = System.Math.Round(costFuture, mode:=MidpointRounding.ToEven)

        Dim korrektur As Double = zeitraumCost - (costPast + costFuture)
        If future Then
            costFuture = costFuture + korrektur
        Else
            costPast = costPast + korrektur
            If costPast < 0 Then
                costFuture = costFuture + costPast
                costPast = 0
            End If
        End If

        ' das ist der Risiko Abschlag  
        riskValue = System.Math.Round(ShowProjekte.getWeightedRiskValuesInMonth.Sum, mode:=MidpointRounding.ToEven)

        itemValue(0) = budgetSum
        itemColor(0) = ergebnisfarbe1


        Dim currentWert As Double = itemValue(0)


        ' das ist der Risiko-Abschlag 
        itemValue(1) = riskValue
        itemColor(1) = iProjektFarbe

        ' das sind die Kosten der Vergangenheit
        itemValue(2) = costPast
        itemColor(2) = farbeExterne

        ' das sind die Kosten der Zukunft
        itemValue(3) = costFuture
        itemColor(3) = farbeExterne

        ' das ist der Ertrag 
        ertragsWert = budgetSum - (costPast + costFuture + riskValue)
        itemValue(4) = ertragsWert
        If ertragsWert > 0 Then
            itemColor(4) = ergebnisfarbe2
        Else
            itemColor(4) = farbeExterne
        End If

        diagramTitle = portfolioDiagrammtitel(PTpfdk.Budget) & " " & textZeitraum(showRangeLeft, showRangeRight)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count

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
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'MsgBox(" Diagramm wird bereits angezeigt ...")
            Else

                If ertragsWert < 0 Then
                    minScale = System.Math.Round(ertragsWert, mode:=MidpointRounding.ToEven)
                Else
                    minScale = 0
                End If

                'Dim htxt As String
                Dim valueCrossesNull As Boolean = False

                With appInstance.Charts.Add
                    ' remove old series
                    Try
                        Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                        Do While anz > 0
                            .SeriesCollection(1).Delete()
                            anz = anz - 1
                        Loop
                    Catch ex As Exception

                    End Try
                    Dim crossindex As Integer = -1

                    ' bestimmen des Anfangs  
                    Dim iv = 0
                    valueDatenreihe1(iv) = 0
                    valueDatenreihe2(iv) = itemValue(iv)
                    currentWert = itemValue(iv)
                    Dim formerValue As Double = currentWert
                    Dim negativeFromNull As Boolean = False

                    ' alle nächsten Zwischen-Werte 
                    For iv = 1 To 3
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
                    iv = 4
                    valueDatenreihe1(iv) = 0
                    valueDatenreihe2(iv) = itemValue(iv)



                    'series
                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Bottom"
                        .Name = repMessages.getmsg(149)
                        .HasDataLabels = False
                        .Interior.ColorIndex = -4142
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

                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Top"
                        .Name = repMessages.getmsg(150)
                        .HasDataLabels = True
                        .Values = valueDatenreihe2
                        .XValues = Xdatenreihe
                        .ChartType = Excel.XlChartType.xlColumnStacked

                        For iv = 0 To 4

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
                                .MinimumScale = System.Math.Round((minScale - 1), mode:=MidpointRounding.ToEven)
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

                    Dim achieved As Boolean = False
                    Dim anzahlVersuche As Integer = 0
                    Dim errmsg As String = ""
                    Do While Not achieved And anzahlVersuche < 10
                        Try
                            'Call Sleep(100)
                            .Location(Where:=XlChartLocation.xlLocationAsObject, Name:=appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)).name)
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

                End With

                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .name = chtobjName
                End With

                repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

                ' jetzt muss die letzte Position des Diagramms gespeichert werden , wenn es nicht aus der Reporting Engine 
                ' aufgerufen wurde
                If Not calledfromReporting Then

                    Dim prcDiagram As New clsDiagramm

                    ' Anfang Event Handling für Chart 
                    Dim prcChart As New clsEventsPrcCharts
                    prcChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart
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



    Sub awinCreateVerbesserungsPotentialDiagramm(ByRef repObj As Excel.ChartObject, ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double, ByVal isCockpitChart As Boolean)

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

        Dim currentSheetName As String

        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentSheetName = arrWsNames(ptTables.MPT)
        Else
            currentSheetName = arrWsNames(ptTables.meRC)
        End If

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False



        'Xdatenreihe(0) = "Mehrkosten wegen Überauslastung"
        'Xdatenreihe(1) = "Opportunitätskosten durch Unterauslastung"
        Xdatenreihe(0) = repMessages.getmsg(141)
        Xdatenreihe(1) = repMessages.getmsg(142)


        Dim positiv As Boolean = True



        ' das sind die Zusatzkosten, die durch Externe (wg Überauslastung) verursacht werden
        additionalCostExt = System.Math.Round(ShowProjekte.getCosteValuesInMonth(True).Sum, mode:=MidpointRounding.ToEven)

        itemValue(0) = additionalCostExt
        itemColor(0) = awinSettings.AmpelRot

        ' das sind die durch Unterauslastung verursachten Kosten , also Personal-Kosten von Leuten, die in keinem Projekt sind
        internwithoutProject = System.Math.Round(ShowProjekte.getCostoValuesInMonth.Sum, mode:=MidpointRounding.ToEven)
        itemValue(1) = internwithoutProject
        itemColor(1) = awinSettings.AmpelGelb


        diagramTitle = summentitel5 & " (T€) " & vbLf & StartofCalendar.AddMonths(showRangeLeft - 1).ToString("MMM yy", repCult) & " - " & _
                                                 StartofCalendar.AddMonths(showRangeRight - 1).ToString("MMM yy", repCult)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False


        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentSheetName), Excel.Worksheet)

            Dim wasProtected As Boolean = .ProtectContents

            If .ProtectContents And visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
                .Unprotect(Password:="x") ' damit Chart selektierbar ist ...
                awinSettings.meEnableSorting = True ' damit es konsistent ist mit Menu Anzeige 
            End If

            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count

            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False
            While i <= anzDiagrams And Not found

                Try
                    chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                Catch ex As Exception
                    chtTitle = " "
                End Try


                If ((chtTitle Like (diagramTitle & "*")) And _
                         (isCockpitChart = istCockpitDiagramm(CType(.ChartObjects(i), Excel.ChartObject)))) Then
                    found = True
                Else
                    i = i + 1
                End If

            End While



            If found Then
                repObj = CType(.ChartObjects(i), Excel.ChartObject)
                'MsgBox(" Diagramm wird bereits angezeigt ...")
            Else



                With appInstance.Charts.Add
                    ' remove old series
                    Try
                        Dim anz As Integer = CInt(CType(.SeriesCollection, Excel.SeriesCollection).Count)
                        Do While anz > 0
                            .SeriesCollection(1).Delete()
                            anz = anz - 1
                        Loop
                    Catch ex As Exception

                    End Try


                    'series
                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).NewSeries, Excel.Series)
                        '.name = "Potentiale"
                        .Name = repMessages.getmsg(143)
                        .HasDataLabels = True
                        .DataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
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


                End With

                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .name = diagramTitle

                End With

                repObj = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)


            End If

            sumDiagram = New clsDiagramm

            sumChart = New clsEventsPrcCharts
            sumChart.PrcChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart

            sumDiagram.setDiagramEvent = sumChart


            With sumDiagram
                .DiagrammTitel = diagramTitle
                .diagrammTyp = DiagrammTypen(4)
                '.setCollection = myCollection
                .isCockpitChart = isCockpitChart
            End With

            DiagramList.Add(sumDiagram)
            'sumDiagram = Nothing

            ' '' wenn es geschützt war .. 
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

        End With

        'Call awinScrollintoView()
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    ''
    '' zeigt für alle/die selektierten Projekte die Bedarfe für die jeweilige Rolle an
    ''
    'Sub awinShowProjectNeeds1(ByRef mycollection As Collection, type As String)
    '    Dim formerSU As Boolean = appInstance.ScreenUpdating

    '    appInstance.ScreenUpdating = False

    '    ' jetzt überprüfen, ob Projekte selektiert sind
    '    If selectedProjekte.Count > 0 Then
    '        ' dann die Werte in die Excel Zellen der selektierten Projekte schreiben 
    '        For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste
    '            Call awinShowNeedsofProject1(mycollection, type, kvp.Key)
    '        Next kvp
    '    Else

    '        ' sonst die Werte aller geladenen Projekte in die Excel Zellen schreiben 
    '        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
    '            Call awinShowNeedsofProject1(mycollection, type, kvp.Key)
    '        Next kvp
    '    End If


    '    ' jetzt wieder alle Shapes sichtbar machen 

    '    appInstance.ScreenUpdating = formerSU


    'End Sub

    ' tk 21.8.17 wird nicht mehr angeboten 
    '
    ' löscht für alle Projekte die Bedarfe für die jeweilige Rolle an
    '
    'Sub awinNoshowProjectNeeds()

    '    Dim formerSU As Boolean = appInstance.ScreenUpdating
    '    appInstance.ScreenUpdating = False


    '    Call diagramsVisible(False)

    '    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
    '        Call NoshowNeedsofProject(kvp.Key)
    '    Next kvp

    '    Call diagramsVisible(True)


    '    appInstance.ScreenUpdating = formerSU


    'End Sub

    ''
    '' zeigt für das gewählte Projekt die Bedarfe für die angegebene Rolle an
    ''
    ' ''' <summary>
    ' ''' zeigt für das entsprechende Diagramm-Typ und jeweiligen prcname die entsprechenden Werte  
    ' ''' </summary>
    ' ''' <param name="mycollection">enthält ggf die zu betrachtende Menge an Werten</param>
    ' ''' <param name="type">wert aus DiagrammTypen 0..4 </param>
    ' ''' <param name="projektname">NAme des Projekts aus ShowProjekte</param>
    ' ''' <remarks></remarks>
    'Sub awinShowNeedsofProject1(ByRef mycollection As Collection, ByVal type As String, ByVal projektname As String)

    '    Dim i As Integer, k As Integer, l As Integer, m As Integer

    '    Dim tempArray() As Double
    '    Dim pname As String = " "
    '    'Dim showKostenart As Boolean
    '    Dim hproj As New clsProjekt
    '    Dim sfarbe As Object
    '    Dim sgroesse As Integer
    '    'Dim prcName As String
    '    'Dim itemName As String
    '    Dim persCost As String = CostDefinitions.getCostdef(CostDefinitions.Count).name
    '    Dim shpelement As Excel.Shape
    '    Dim tmpshapes As Excel.Shapes = CType(appInstance.ActiveSheet, Excel.Worksheet).Shapes


    '    Try
    '        hproj = ShowProjekte.getProject(projektname)
    '    Catch ex As Exception
    '        Call MsgBox("Projekt nicht gefunden (in ShowNeedsofProject): " & projektname)
    '        Exit Sub
    '    End Try


    '    Dim anzahlTage As Integer = hproj.dauerInDays



    '    Try
    '        shpelement = tmpshapes.Item(projektname)
    '        ' jetzt muss unterschieden werden, um welche Art es sich handelt 

    '        With shpelement

    '            Try
    '                If .GroupItems.Count > 1 Then

    '                    If CBool(.GroupItems.Item(1).TextFrame2.HasText) And Not awinSettings.drawProjectLine Then
    '                        .GroupItems.Item(1).TextFrame2.TextRange.Text = ""
    '                    End If

    '                    For i = 1 To .GroupItems.Count

    '                        If .GroupItems.Item(i).AlternativeText = "(Projektname)" Then
    '                            .GroupItems.Item(i).Line.Transparency = 0.8
    '                            .GroupItems.Item(i).Fill.Transparency = 1.0
    '                            .TextFrame2.TextRange.Text = ""
    '                        Else
    '                            If awinSettings.drawProjectLine And i = 1 Then

    '                                .GroupItems.Item(i).Line.Transparency = 0.8

    '                            Else

    '                                .GroupItems.Item(i).Fill.Transparency = 0.8

    '                            End If
    '                        End If



    '                    Next
    '                Else
    '                    .Fill.Transparency = 0.8
    '                    .TextFrame2.TextRange.Text = ""
    '                End If

    '            Catch ex1 As Exception

    '                .Fill.Transparency = 0.8
    '                .TextFrame2.TextRange.Text = ""

    '            End Try

    '        End With

    '    Catch ex As Exception

    '    End Try

    '    If Not hproj Is Nothing Then
    '        With hproj
    '            sfarbe = RGB(0, 0, 0) '.Schriftfarbe
    '            sgroesse = .Schrift
    '            ' in L steht jetzt die Lä nge
    '            l = .anzahlRasterElemente
    '            i = .tfZeile + 1
    '            k = .tfspalte
    '        End With

    '        ReDim tempArray(l - 1)

    '        tempArray = hproj.getBedarfeInMonths(mycollection, type)

    '        Dim formerEE = appInstance.EnableEvents
    '        appInstance.EnableEvents = False

    '        ' hier muss jetzt tempArray gesetzt werden

    '        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

    '            Dim atleastOne As Boolean = False
    '            For m = 1 To l
    '                If tempArray(m - 1) > 0 And istInTimezone(k + m - 1) Then

    '                    Try
    '                        .Cells(i, k).Offset(0, m - 1).Value = tempArray(m - 1)
    '                        atleastOne = True
    '                    Catch ex As Exception

    '                    End Try

    '                End If
    '            Next m

    '            Dim tmpgroesse As Integer
    '            If tempArray.Max > 999 Or tempArray.Min < -999 Then
    '                tmpgroesse = sgroesse - 2
    '            ElseIf tempArray.Max > 9999 Or tempArray.Min < -9999 Then
    '                tmpgroesse = sgroesse - 4
    '            Else
    '                tmpgroesse = sgroesse
    '            End If

    '            If atleastOne Then

    '                Try
    '                    .Range(.Cells(i, k), .Cells(i, k).Offset(0, l - 1)).Font.Color = sfarbe
    '                    .Range(.Cells(i, k), .Cells(i, k).Offset(0, l - 1)).Font.Size = tmpgroesse
    '                Catch ex As Exception

    '                End Try

    '            End If

    '        End With

    '        appInstance.EnableEvents = formerEE

    '    End If



    'End Sub

    ' tk, 21.8.17 Funktion wird nicht mehr angebeoten 
    '
    ' löscht für das gewählte Projekt die Bedarfe für die angegebene Rolle
    '
    ''Sub NoshowNeedsofProject(ByVal projektname As String)
    ''    Dim hproj As clsProjekt
    ''    Dim sfarbe As Object
    ''    Dim sgroesse As Double
    ''    Dim i As Integer, k As Integer, l As Integer, m As Integer
    ''    Dim shpelement As Excel.Shape
    ''    Dim worksheetShapes As Excel.Shapes
    ''    Dim pStatus As String


    ''    Try

    ''        worksheetShapes = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes

    ''    Catch ex As Exception
    ''        Throw New Exception("in NoshowNeedsofProject: keine Shapes Zuordnung möglich ")
    ''    End Try


    ''    Try
    ''        hproj = ShowProjekte.getProject(projektname)
    ''        pStatus = hproj.Status
    ''    Catch ex As Exception
    ''        Call MsgBox("Projekt nicht gefunden (in NoShowNeedsofProject): " & projektname)
    ''        Exit Sub
    ''    End Try


    ''    Try
    ''        'tmpshapes = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Shapes
    ''        shpelement = worksheetShapes.Item(projektname)
    ''        With shpelement

    ''            Try
    ''                If .GroupItems.Count > 1 Then

    ''                    If CBool(.GroupItems.Item(1).TextFrame2.HasText) And Not awinSettings.drawProjectLine Then
    ''                        .GroupItems.Item(1).TextFrame2.TextRange.Text = projektname
    ''                    End If

    ''                    For i = 1 To .GroupItems.Count

    ''                        If .GroupItems.Item(i).AlternativeText = "(Projektname)" Then
    ''                            .GroupItems.Item(i).Line.Transparency = 0.0
    ''                            .GroupItems.Item(i).Fill.Transparency = 0.0
    ''                            .TextFrame2.TextRange.Text = hproj.getShapeText

    ''                        ElseIf awinSettings.drawProjectLine And i = 1 Then

    ''                            .GroupItems.Item(i).Line.Transparency = 0.0

    ''                        Else
    ''                            If pStatus = ProjektStatus(0) Then
    ''                                .GroupItems.Item(i).Fill.Transparency = 0.35
    ''                            Else
    ''                                .GroupItems.Item(i).Fill.Transparency = 0.0
    ''                            End If
    ''                        End If

    ''                    Next
    ''                Else

    ''                    If pStatus = ProjektStatus(0) Then
    ''                        .Fill.Transparency = 0.35
    ''                    Else
    ''                        .Fill.Transparency = 0.0
    ''                    End If

    ''                    .TextFrame2.TextRange.Text = projektname
    ''                End If

    ''            Catch ex1 As Exception

    ''                If pStatus = ProjektStatus(0) Then
    ''                    .Fill.Transparency = 0.35
    ''                Else
    ''                    .Fill.Transparency = 0.0
    ''                End If

    ''                .TextFrame2.TextRange.Text = projektname
    ''            End Try


    ''            '.Shadow.Transparency = 0.0
    ''        End With

    ''    Catch ex As Exception

    ''    End Try

    ''    ' jetzt muss das Shape wieder auf "ohne Transparenz" gesetzt werden 

    ''    If Not hproj Is Nothing Then
    ''        With hproj
    ''            sfarbe = RGB(0, 0, 0) '.Schriftfarbe
    ''            sgroesse = .Schrift
    ''            ' in L steht jetzt die Länge
    ''            l = .anzahlRasterElemente
    ''            i = .tfZeile + 1
    ''            k = .tfspalte
    ''        End With

    ''        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

    ''            appInstance.EnableEvents = False

    ''            For m = 1 To l
    ''                If istInTimezone(k + m - 1) Then
    ''                    .Cells(i, k).Offset(0, m - 1).Value = ""
    ''                End If
    ''            Next m


    ''            appInstance.EnableEvents = True

    ''        End With


    ''    End If

    ''End Sub




    ''' <summary>
    ''' Funktion prüft , ob die Spalte angezeigt werden muss, also ob sie in der Time Zone enthalten ist
    ''' </summary>
    ''' <param name="spalte"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function istInTimezone(ByVal spalte As Integer) As Boolean

        If showRangeLeft <= 0 And showRangeRight <= 0 Then
            istInTimezone = True
        ElseIf spalte >= showRangeLeft And spalte <= showRangeRight Then
            istInTimezone = True
        Else
            istInTimezone = False
        End If

    End Function

    ''' <summary>
    ''' prüft, ob der angegebene Bereich sich mit dem gewählten Zeitraum überlappt 
    ''' </summary>
    ''' <param name="anfang"></param>
    ''' <param name="ende"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function istBereichInTimezone(ByVal anfang As Integer, ByVal ende As Integer) As Boolean


        If ((ende) < showRangeLeft) Or (anfang > showRangeRight) Then
            istBereichInTimezone = False
        Else
            istBereichInTimezone = True
        End If


    End Function



    Sub diagramsVisible(ByVal show As Boolean)

        Dim anzDiagrams As Integer
        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)

            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count

            For i = 1 To anzDiagrams
                CType(.ChartObjects(i), Excel.ChartObject).Visible = show
            Next i
        End With

    End Sub

    ''' <summary>
    ''' löscht alle Charts im angegebenen Sheet 
    ''' </summary>
    ''' <param name="sheetName"></param>
    ''' <remarks></remarks>
    Public Sub deleteChartsInSheet(ByVal sheetName As String)

        Dim anzDiagrams As Integer
        Dim i As Integer = 1
        Dim chtobj As Excel.ChartObject

        Try
            With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(sheetName), Excel.Worksheet)

                anzDiagrams = CInt(CType(.ChartObjects, Excel.ChartObjects).Count)

                While i <= anzDiagrams

                    Try
                        chtobj = CType(.ChartObjects(1), Excel.ChartObject)
                        Call awinDeleteChart(chtobj)
                        i = i + 1
                    Catch ex As Exception
                        i = anzDiagrams + 1
                    End Try


                End While

            End With
        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' zeichnet alle dargestellten Portfolio ("Pf") Diagramme neu
    ''' die optionale Parameter werden im Fall MassenEdit benötigt - es wird dann mitgegeben, welche Rolle/Kostenart/Milestone/Phase aktualisiert werden soll und
    ''' welches Projekt im Chart ausgewiesen werden soll 
    ''' 99 - nur das Strategie-Risiko Chart neu zeichnen; hier sollen die Markierungen weggenommen werden 
    ''' </summary>
    ''' <param name="typus"></param>
    ''' <remarks></remarks>
    Sub awinNeuZeichnenDiagramme(ByVal typus As Integer, _
                                 Optional ByVal roleCost As String = Nothing)
        Dim anz_diagrams As Integer
        Dim chtobj As ChartObject
        Dim formerActiveSheet As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating

        Dim formerShowValuesOfSelected As Boolean = awinSettings.showValuesOfSelected
        Dim i As Integer, p As Integer

        Dim isRole As Boolean = True

        Dim currentSheetName As String
        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentSheetName = arrWsNames(ptTables.mptPfCharts)
            roleCost = Nothing
            Try
                If visboWindowExists(PTwindows.mpt) Then
                    Dim tmpmsg As String = ""
                    projectboardWindows(PTwindows.mpt).Caption = bestimmeWindowCaption(PTwindows.mpt, tmpmsg)
                End If
            Catch ex As Exception

            End Try
        Else
            ' roleCost wird übergeben, wenn man sich im modus <> graficboard befindet 
            If Not IsNothing(roleCost) Then
                If RoleDefinitions.containsName(roleCost) Then
                    isRole = True
                ElseIf CostDefinitions.containsName(roleCost) Then
                    isRole = False
                Else
                    roleCost = Nothing
                End If
            End If
            ' currentSheetName = arrWsNames(ptTables.meRC)
            awinSettings.showValuesOfSelected = True
            currentSheetName = arrWsNames(ptTables.meCharts)
        End If

        ' nur etwas tun, wenn ShowProjekte.count > 0 ...
        'If ShowProjekte.Count > 0 Then

        ' temporärer Check: 
        'If CType(appInstance.ActiveSheet, Excel.Worksheet).Name <> currentSheetName Then
        '    Call MsgBox("Fehler: " & currentSheetName & " ist ungleich " & _
        '                CType(appInstance.ActiveSheet, Excel.Worksheet).Name)
        'End If


        ' typus:
        ' 1 - verschieben
        ' 2 - einfügen
        ' 3 - löschen
        ' 4 - betrachteten Zeitraum ändern
        ' 5 - Stammdaten ändern
        ' 6 - Ressourcen-Bedarfe, Kapas ändern
        ' 7 - Kosten-Bedarfe , Budgets ändern
        ' 8 - Selektion geändert
        ' 9 - Cockpit wurde geladen; (alle Diagramme neuzeichnen)

        ' Schutz Funktion : wenn showrangeleft = 0 und showrangeright = 0 , dann nichts tun
        If showRangeRight - showRangeLeft >= minColumns - 1 Then

            ' wenn das ActiveSheet ungleich dem currentSheetName ist, muss gewechselt werden ... 
            If CType(appInstance.ActiveSheet, Excel.Worksheet).Name <> currentSheetName Then
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False
                CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentSheetName), Excel.Worksheet).Activate()
            End If

            With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentSheetName), Excel.Worksheet)

                ' 24.5.17 das wird nicht mehr benötigt, weil die Charts jetzt in einem eigenen Sheet sind ... 
                'Dim wasProtected As Boolean = False
                'If currentSheetName = arrWsNames(ptTables.meRC) Then
                '    wasProtected = .ProtectContents

                '    If wasProtected And visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
                '        .Unprotect(Password:="x")
                '    End If
                'End If

                anz_diagrams = CType(.ChartObjects, Excel.ChartObjects).Count
                For i = 1 To anz_diagrams
                    chtobj = CType(.ChartObjects(i), Excel.ChartObject)

                    Select Case typus
                        '
                        '
                        Case 8 ' Selection hat sich geändert 

                            If istRollenDiagramm(chtobj) Or istKostenartDiagramm(chtobj) Or _
                                istPhasenDiagramm(chtobj) Or istMileStoneDiagramm(chtobj) Then

                                Call awinUpdateprcCollectionDiagram(chtobj:=chtobj, _
                                                                    roleCost:=roleCost, _
                                                                    isRole:=isRole)

                            End If
                        Case 99
                            ' nur die Strategie - / Risiko Diagramme sollen neu gezeichnet, d.h die Markierungen zurückgesetzt werden 
                            If istSummenDiagramm(chtobj, p) Then
                                If p = PTpfdk.Dependencies Or _
                                       p = PTpfdk.FitRisiko Or _
                                       p = PTpfdk.FitRisikoVol Or _
                                       p = PTpfdk.ZeitRisiko Or _
                                       p = PTpfdk.ComplexRisiko Then
                                    Call awinUpdateMarkerInPortfolioDiagrams(chtobj)
                                    'Call awinUpdatePortfolioDiagrams(chtobj, PTpfdk.AmpelFarbe)
                                End If
                            End If

                        Case Else
                            ' 1: Projekt wurde verschoben
                            ' 2: Projekt wurde eingefügt
                            ' 3: Projekt wurde gelöscht
                            ' 4: betrachteter Zeitraum wurde geändert
                            ' 5: Stammdaten wurden geändert
                            ' 6: Ressourcen Bedarf eines existierenden Projektes wurde geändert
                            ' 7: Kosten Bedarf eines existierenden Projektes wurde geändert
                            ' 9: Cockpit wurde geladen; (alle Diagramme neuzeichnen)

                            If (typus <> 5) And (istRollenDiagramm(chtobj) Or istKostenartDiagramm(chtobj) Or _
                                istPhasenDiagramm(chtobj) Or istMileStoneDiagramm(chtobj)) Then

                                Call awinUpdateprcCollectionDiagram(chtobj:=chtobj, roleCost:=roleCost, isRole:=isRole)


                            ElseIf istSummenDiagramm(chtobj, p) Then

                                If p = PTpfdk.ErgebnisWasserfall Then
                                    Call awinUpdateErgebnisDiagramm(chtobj)

                                ElseIf p = PTpfdk.Dependencies Or _
                                       p = PTpfdk.FitRisiko Or _
                                       p = PTpfdk.FitRisikoVol Or _
                                       p = PTpfdk.ZeitRisiko Or _
                                       p = PTpfdk.ComplexRisiko Then

                                    Call awinUpdatePortfolioDiagrams(chtobj, PTpfdk.AmpelFarbe)

                                ElseIf p = PTpfdk.Auslastung Then
                                    Try
                                        Call awinUpdateAuslastungsDiagramm(chtobj)
                                    Catch ex As Exception

                                    End Try

                                ElseIf p = PTpfdk.UeberAuslastung Then
                                    Try
                                        Call updateAuslastungsDetailPie(chtobj, 1)
                                    Catch ex As Exception

                                    End Try
                                ElseIf p = PTpfdk.Unterauslastung Then
                                    Try
                                        Call updateAuslastungsDetailPie(chtobj, 2)
                                    Catch ex As Exception

                                    End Try

                                    ' p = 19 
                                ElseIf p = PTpfdk.Budget Then
                                    Try
                                        Call awinUpdateBudgetErgebnisDiagramm(chtobj)
                                    Catch ex As Exception

                                    End Try
                                End If


                            End If



                    End Select

                Next i

                ' wenn das ActiveSheet ungleich dem currentSheetName war, muss jetzt zurück gewechselt werden ... 
                Dim xName As String = CType(appInstance.ActiveSheet, Excel.Worksheet).Name
                If CType(appInstance.ActiveSheet, Excel.Worksheet).Name <> formerActiveSheet.Name Then
                    CType(formerActiveSheet, Excel.Worksheet).Activate()
                    If appInstance.EnableEvents <> formerEE Then
                        appInstance.EnableEvents = formerEE
                    End If
                    If appInstance.ScreenUpdating <> formerSU Then
                        appInstance.ScreenUpdating = formerSU
                    End If
                End If

                ' tk 24.5.17 das wird nicht mehr benötigt, weil die Charts jetzt in einem Extra Sheet sind ... 
                '' '' wenn es geschützt war .. 
                'If wasProtected And visboZustaende.projectBoardMode = ptModus.massEditRessCost Then
                '    .Protect(Password:="x", UserInterfaceOnly:=True, _
                '                 AllowFormattingCells:=True, _
                '                 AllowInsertingColumns:=False,
                '                 AllowInsertingRows:=True, _
                '                 AllowDeletingColumns:=False, _
                '                 AllowDeletingRows:=True, _
                '                 AllowSorting:=True, _
                '                 AllowFiltering:=True)
                '    .EnableSelection = XlEnableSelection.xlUnlockedCells
                '    .EnableAutoFilter = True
                'End If


            End With
        End If

        ' jetzt muss geprüft werden, ob sich awinsettings.showValuesofselected geändert hatte ... dann wieder zurücksetzen 
        If formerShowValuesOfSelected <> awinSettings.showValuesOfSelected Then
            awinSettings.showValuesOfSelected = formerShowValuesOfSelected
        End If

    End Sub

    ''' <summary>
    ''' setzt in den Pf Diagrammen eine evtl gesetzte Fill Farbe zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub unmarkPfDiagrams()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        Try
            Dim currentSheetName As String = arrWsNames(ptTables.mptPfCharts)
            With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentSheetName), Excel.Worksheet)
                Dim anz_diagrams As Integer = CType(.ChartObjects, Excel.ChartObjects).Count
                For i As Integer = 1 To anz_diagrams

                    Try
                        Dim chtobj As Excel.ChartObject = CType(.ChartObjects(i), Excel.ChartObject)
                        Dim curShape As Excel.Shape = .Shapes.Item(chtobj.Name)
                        With curShape.Fill
                            .Visible = MsoTriState.msoFalse
                            .ForeColor.RGB = RGB(255, 255, 255)
                        End With
                    Catch ex As Exception

                    End Try

                Next i

            End With
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = formerEE

    End Sub

    ' ''' <summary>
    ' ''' stellt das Fenster "Projekt Tafel" so ein, daß die gesamte Zeitleiste zu sehen ist und evtl das Diagramm
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Sub awinScrollintoView()
    '    Dim ScrollColumn As Integer
    '    Dim zoom As Double
    '    Dim minWindowWidth As Double, minWindowHeight As Double

    '    Dim testWBName As String = appInstance.ActiveWorkbook.Name
    '    Dim testWSName As String = CType(appInstance.ActiveSheet, Excel.Worksheet).Name
    '    Dim testEnable As Boolean = appInstance.EnableEvents
    '    Try
    '        appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)).activate()
    '        'appInstance.ActiveWorkbook.Windows(windowNames(5)).Activate()
    '    Catch ex As Exception
    '        Call MsgBox("Window " & windowNames(5) & " existiert nicht mehr !")
    '        Exit Sub
    '    End Try



    '    ScrollColumn = showRangeLeft - 12 ' war vorher 6
    '    If ScrollColumn <= 0 Then
    '        ScrollColumn = 1
    '    End If


    '    minWindowWidth = Max(boxWidth * (showRangeRight - showRangeLeft + 1 + 12), 60 * boxWidth)
    '    minWindowHeight = Max(WertfuerTop() + 30, 22 * boxHeight + 30)


    '    Dim shp As Excel.Shape
    '    For Each shp In appInstance.ActiveSheet.Shapes
    '        With shp
    '            If .BottomRightCell.Top > minWindowHeight And .BottomRightCell.Top < WertfuerTop() * boxHeight Then
    '                minWindowHeight = .BottomRightCell.Top + 3 * boxHeight
    '            End If
    '            If .BottomRightCell.Left - (showRangeLeft - 6) * boxWidth > minWindowWidth Then
    '                minWindowWidth = .BottomRightCell.Left + 3 * boxWidth - (showRangeLeft - 6) * boxWidth
    '            End If
    '        End With
    '    Next shp


    '    With appInstance.ActiveWindow
    '        If .UsableWidth / minWindowWidth < .UsableHeight / minWindowHeight Then
    '            ' Zoom an Breite orientieren ...
    '            Try
    '                zoom = 100 * .UsableWidth / minWindowWidth
    '                .Zoom = Min(zoom, 120)
    '                If .Zoom < 60 Then
    '                    .Zoom = 60
    '                End If
    '            Catch ex As Exception
    '                If zoom < 20 Then
    '                    .Zoom = 20
    '                ElseIf zoom > 400 Then
    '                    .Zoom = 400
    '                Else
    '                    .Zoom = 100
    '                End If
    '            End Try

    '        Else
    '            ' Zoom an Höhe orientieren 
    '            Try
    '                zoom = 100 * .UsableHeight / minWindowHeight
    '                .Zoom = Min(zoom, 120)
    '                If .Zoom < 60 Then
    '                    .Zoom = 60
    '                End If
    '            Catch ex As Exception
    '                If zoom < 20 Then
    '                    .Zoom = 20
    '                ElseIf zoom > 400 Then
    '                    .Zoom = 400
    '                Else
    '                    .Zoom = 100
    '                End If
    '            End Try

    '        End If
    '        If Abs(ScrollColumn - .ScrollColumn) > 2 Then
    '            .ScrollColumn = ScrollColumn
    '        End If
    '        .ScrollRow = 1
    '    End With


    'End Sub


End Module
