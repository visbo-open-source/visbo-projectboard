
Imports ProjectBoardDefinitions
'Imports DBAccLayer
Imports ProjectboardReports
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
'Imports System.ComponentModel
'Imports System.Windows
Imports System.Windows.Forms
Imports System.Security.Principal
Imports System.Text.RegularExpressions

'Imports System
'Imports System.Runtime.Serialization
'Imports System.Xml
'Imports System.Xml.Serialization
'Imports System.IO
'Imports System.Drawing
'Imports System.Globalization

'Imports Microsoft.VisualBasic
'Imports System.Security.Principal




Public Module awinGeneralModules

    Private Enum ptInventurSpalten
        Name = 0
        Vorlage = 1
        Start = 2
        Ende = 3
        startElement = 4
        endElement = 5
        Dauer = 6
        Budget = 7
        Risiko = 8
        Strategie = 9
        Kapazitaet = 10
        Businessunit = 11
        Beschreibung = 12
        KostenExtern = 13
    End Enum

    Private Enum allianzSpalten
        Name = 0
        AmpelText = 1
        BusinessUnit = 2
        Description = 3
        Responsible = 4
        Budget = 5
        Projektnummer = 6
        Status = 7
        itemType = 8
        pvBudget = 9
    End Enum

    Private Enum ptModuleSpalten
        produktlinie = 0
        name = 1
        projektTyp = 2
        abhaengigVon = 3
        strategicFit = 4
        risiko = 5
        volume = 6
        budget = 7
    End Enum

    ''' <summary>
    ''' erstellt die CacheProjekte Liste , vorläufig erstmal aus der Session 
    ''' </summary>
    ''' <param name="todoListe">enthält die pvnames der Projekte</param>
    ''' <remarks></remarks>
    Public Sub buildCacheProjekte(ByVal todoListe As Collection,
                                  Optional ByVal namesArePvNames As Boolean = False)

        Dim err As New clsErrorCodeMsg

        Dim pName As String

        For i As Integer = 1 To todoListe.Count
            pName = CStr(todoListe.Item(i))
            Dim hproj As clsProjekt = Nothing

            Try
                If namesArePvNames Then
                    hproj = AlleProjekte.getProject(pName)
                Else
                    If ShowProjekte.contains(pName) Then
                        hproj = ShowProjekte.getProject(pName, True)
                    End If
                End If

                ' in CacheProjekte soll jetzt der aktuelle Stand der Session , nicht der Stand aus der Datenbank 
                If Not IsNothing(hproj) Then
                    Dim oldProj As clsProjekt = hproj.createVariant("$cache$", hproj.variantDescription)
                    oldProj.variantName = hproj.variantName
                    sessionCacheProjekte.upsert(oldProj)
                End If

                'tk 19.4.19 auskommentiert , weil 
                'If Not IsNothing(hproj) Then
                '    If Not noDB Then
                '        wenn es In der DB existiert, dann im Cache aufbauen 

                '        If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(hproj.name, hproj.variantName, Date.Now, err) Then
                '            für den Datenbank Cache aufbauen 
                '            Dim dbProj As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, Date.Now, err)
                '            dbCacheProjekte.upsert(dbProj)
                '        End If
                '    End If
                'End If

            Catch ex As Exception

            End Try
        Next

    End Sub

    ''' <summary>
    ''' prüft ob das Laden und Anzeigen eines Projektes aus der Datenbank in Showprojekte überhaupt zulässig ist ... 
    ''' ist dann nicht zulässig, wenn pName bereits in einem Summenprojekt aus Showprojekte referenziert oder ein pname bereits in Showprojekte ist  
    ''' Es reicht, wenn der pName identisch ist; 
    ''' in diesem Fall wird das Projekt in seiner gewünschten Variante vname in Alleprojekte geladen ...
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="vname"></param>
    ''' <returns></returns>
    Public Function loadIsAllowed(ByVal pname As String, ByVal vname As String) As Boolean
        Dim tmpResult As Boolean = True

        loadIsAllowed = tmpResult
    End Function
    ''' <summary>
    ''' baut eine todoliste der Projekte auf. Wenn in todoliste ein Summary Projekt enthalten ist, wird es durch seine Projekte, die im Show liegen ersetzt
    ''' jeder Eintrag der todoListe ist der vollständige key in der Form pName#vname
    ''' </summary>
    ''' <param name="todoListe"></param>
    ''' <returns></returns>
    Public Function substituteListeByPVNameIDs(ByVal todoListe As Collection,
                                                  Optional ByVal noNeedtoBeInShowProjekte As Boolean = False) As Collection

        Dim err As New clsErrorCodeMsg

        Dim tmpCollection As New Collection
        Dim key As String = ""

        For Each pName As String In todoListe

            If ShowProjekte.contains(pName) Then
                Dim hproj As clsProjekt = ShowProjekte.getProject(pName)

                If Not IsNothing(hproj) Then

                    If hproj.projectType = ptPRPFType.project Or myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                        key = calcProjektKey(hproj)
                        If Not tmpCollection.Contains(key) Then
                            tmpCollection.Add(key, key)
                        End If

                    Else
                        ' muss ersetzt werden 
                        Try
                            Dim currentPortfolio As clsConstellation = projectConstellations.getConstellation(pName)
                            Dim projektNamen As SortedList(Of String, String) = currentPortfolio.getProjectNames(considerShowAttribute:=True,
                                                                                                                 showAttribute:=True,
                                                                                                                 fullNameKeys:=True)

                            For Each kvp As KeyValuePair(Of String, String) In projektNamen
                                Dim pproj As clsProjekt = AlleProjekte.getProject(kvp.Key)

                                If IsNothing(pproj) Then
                                    ' aus Datenbank nachladen ... 
                                    If Not noDB Then
                                        Dim dbPName As String = getPnameFromKey(kvp.Key)
                                        Dim dbVName As String = getVariantnameFromKey(kvp.Key)

                                        ' wenn es in der DB existiert, dann im Cache aufbauen 
                                        If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(dbPName, dbVName, Date.Now, err) Then
                                            ' jetzt aus datenbank holen und in AlleProjekte eintragen 
                                            pproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(dbPName, dbVName, "", Date.Now, err)
                                            If Not IsNothing(pproj) Then
                                                If Not AlleProjekte.Containskey(kvp.Key) Then
                                                    AlleProjekte.Add(pproj, False)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If

                                If Not IsNothing(pproj) Then
                                    ' jetzt ist das Projekt in AlleProjekte nachgeladen ...
                                    If Not pproj.projectType = ptPRPFType.portfolio Then
                                        ' wenn kein Summary Projekt: direkt eintragen
                                        If Not tmpCollection.Contains(kvp.Key) Then
                                            tmpCollection.Add(kvp.Key, kvp.Key)
                                        End If
                                    Else
                                        ' es handelt sich um ein Summary Projekt, dann muss weiter ersetzt werden 
                                        Dim tmpTodoList As New Collection
                                        tmpTodoList.Add(kvp.Key, kvp.Key)
                                        Dim teilErgebnis As Collection = substituteListeByPVNameIDs(tmpTodoList, True)

                                        For Each pvname As String In teilErgebnis
                                            If Not tmpCollection.Contains(pvname) Then
                                                tmpCollection.Add(pvname, pvname)
                                            End If
                                        Next
                                    End If

                                Else
                                    Call MsgBox("Projekt nicht gefunden: " & kvp.Key)
                                End If

                            Next
                        Catch ex As Exception

                        End Try

                    End If
                End If
            End If

        Next

        substituteListeByPVNameIDs = tmpCollection
    End Function



    ''' <summary>
    ''' setzt die komplette Session zurück 
    ''' löscht alle Shapes, sofern noch welche vorhanden sind, löscht Showprojekte, alleprojekte, etc. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearCompleteSession()

        Dim err As New clsErrorCodeMsg

        Dim allShapes As Excel.Shapes
        appInstance.EnableEvents = False
        enableOnUpdate = False

        ' jetzt: Löschen der Session 

        Try

            allShapes = CType(appInstance.ActiveSheet, Excel.Worksheet).Shapes
            For Each element As Excel.Shape In allShapes
                element.Delete()
            Next

        Catch ex As Exception
            Call MsgBox("Fehler beim Löschen der Shapes ...")
        End Try

        ' Hier werden die Datenstrukturen alle zurückgesetzt ... 
        ' tk 24.10.19, die hier vorher stehenden Aufrufe wurden in emptyAllVISBOStructures  
        Call emptyAllVISBOStructures()


        ' spezifisch für Projectboard, also den Excel Add-In
        DiagramList.Clear()
        awinButtonEvents.Clear()
        projectboardShapes.clear()


        ' jetzt werden die temporären Schutz Mechanismen rausgenommen ...

        ' ur: 01.02.2019: es werden sonst zuviele Locks versucht zu löschen und Projekte geladen, die nie geladen waren
        'If CType(databaseAcc, DBAccLayer.Request).cancelWriteProtections(dbUsername, err) Then
        '    If awinSettings.visboDebug Then
        '        Call MsgBox("Ihre vorübergehenden Schreibsperren wurden aufgehoben")
        '    End If
        'End If


        ' tk, 10.11.16 allDependencies darf nicht gelöscht werden, weil das sonst nicht mehr vorhanden ist
        ' allDependencies wird aktull nur beim Start geladen - und das reicht ja auch ... 
        ' beim Laden eines Szenarios, beim Laden von Projekten wird das nicht mehr geladen ...
        ' auch die geladenen Konstellationen bleiben erhalten 
        ' alternativ könnte das Folgende aktiviert werden ..
        ''allDependencies.Clear()
        ''projectConstellations.Liste.Clear()
        ' '' hier werden jetzt wieder die in der Datenbank vorhandenen Abhängigkeiten und Szenarios geladen ...
        ''Call readInitConstellations()


        ' Löschen der Charts

        Try
            If visboZustaende.projectBoardMode = ptModus.graficboard Then
                Call deleteChartsInSheet(arrWsNames(ptTables.mptPfCharts))
                Call deleteChartsInSheet(arrWsNames(ptTables.mptPrCharts))
                Call deleteChartsInSheet(arrWsNames(ptTables.MPT))
                ' jetzt müssen alle Windows bis auf Window(0) = Multiprojekt-Tafel geschlossen werden 
                ' und mache ProjectboardWindows(mpt) great again ...
                Call closeAllWindowsExceptMPT()

            Else
                Call deleteChartsInSheet(arrWsNames(ptTables.meCharts))
            End If
        Catch ex As Exception
            Dim a As String = ex.Message
        End Try



        appInstance.EnableEvents = True
        enableOnUpdate = True
    End Sub

    ''' <summary>
    ''' macht den Teil des ClearSession, der so ggf auch in Powerpoint, Project etc gemacht werden kann, um 
    ''' alle Strukturen zurückzusetzen
    ''' </summary>
    Public Sub emptyAllVISBOStructures(Optional ByVal calledFromPPT As Boolean = False)

        ShowProjekte.Clear()
        AlleProjekte.Clear()
        writeProtections.Clear()
        selectedProjekte.Clear(False)
        ImportProjekte.Clear(False)

        ' die ProjectConstellations bleiben erhalten - aber sie sind einfach 
        ' projectConstellations.clearLoadedPortfolios()
        projectConstellations.Clear()


        ' es gibt ja nix mehr in der Session 
        currentConstellationPvName = ""

        ' Range zurücksetzen 
        If calledFromPPT Then
            showRangeLeft = 0
            showRangeRight = 0
        End If

        '
        ' jetzt den Datenbank Cache Löschen , aber nur wenn es nicht von Powerpoint Add-In aus aufgerufen wird
        If Not calledFromPPT Then
            Dim clearOK As Boolean = False
            Try
                clearOK = CType(databaseAcc, DBAccLayer.Request).clearCache()
            Catch ex As Exception
                Call MsgBox("Warning: no Cache clearing " & ex.Message)
            End Try
        End If
        '
        '


    End Sub

    ''' <summary>
    ''' setzt die Messages je nach Sprache 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub setLanguageMessages()
        'ergebnisChartName(0) = "Earned Value"
        'ergebnisChartName(1) = "Earned Value - gewichtet"
        'ergebnisChartName(2) = "Verbesserungs-Potential"
        'ergebnisChartName(3) = "Risiko-Abschlag"

        ergebnisChartName(0) = repMessages.getmsg(54)
        'Call MsgBox(ergebnisChartName(0))
        ergebnisChartName(1) = repMessages.getmsg(55)
        ergebnisChartName(2) = repMessages.getmsg(56)
        ergebnisChartName(3) = repMessages.getmsg(57)

        ' diese Variablen werden benötigt, um die Diagramme gemäß des gewählten Zeitraums richtig zu positionieren
        summentitel1 = repMessages.getmsg(249)
        summentitel2 = repMessages.getmsg(250)
        summentitel3 = repMessages.getmsg(251)
        summentitel4 = repMessages.getmsg(252)
        summentitel5 = repMessages.getmsg(253)
        summentitel6 = repMessages.getmsg(254)
        summentitel7 = repMessages.getmsg(255)
        summentitel8 = repMessages.getmsg(256)
        summentitel9 = repMessages.getmsg(257)
        summentitel10 = repMessages.getmsg(258)
        summentitel11 = repMessages.getmsg(259)



        ReDim portfolioDiagrammtitel(21)
        'portfolioDiagrammtitel(PTpfdk.Phasen) = "Phasen - Übersicht"
        'portfolioDiagrammtitel(PTpfdk.Rollen) = "Rollen - Übersicht"
        'portfolioDiagrammtitel(PTpfdk.Kosten) = "Kosten - Übersicht"
        'portfolioDiagrammtitel(PTpfdk.ErgebnisWasserfall) = summentitel1
        'portfolioDiagrammtitel(PTpfdk.FitRisiko) = summentitel2
        'portfolioDiagrammtitel(PTpfdk.Auslastung) = summentitel9
        'portfolioDiagrammtitel(PTpfdk.UeberAuslastung) = summentitel10
        'portfolioDiagrammtitel(PTpfdk.Unterauslastung) = summentitel11
        'portfolioDiagrammtitel(PTpfdk.ZieleV) = summentitel6
        'portfolioDiagrammtitel(PTpfdk.ZieleF) = summentitel7
        'portfolioDiagrammtitel(PTpfdk.ComplexRisiko) = "Komplexität, Risiko und Volumen"
        'portfolioDiagrammtitel(PTpfdk.ZeitRisiko) = "Zeit, Risiko und Volumen"
        'portfolioDiagrammtitel(PTpfdk.AmpelFarbe) = ""
        'portfolioDiagrammtitel(PTpfdk.ProjektFarbe) = ""
        'portfolioDiagrammtitel(PTpfdk.Meilenstein) = "Meilenstein - Übersicht"
        'portfolioDiagrammtitel(PTpfdk.FitRisikoVol) = "strategischer Fit, Risiko & Volumen"
        'portfolioDiagrammtitel(PTpfdk.Dependencies) = "Abhängigkeiten: Aktive bzw passive Beeinflussung"
        'portfolioDiagrammtitel(PTpfdk.betterWorseL) = "Abweichungen zum letztem Stand"
        'portfolioDiagrammtitel(PTpfdk.betterWorseB) = "Abweichungen zur Beauftragung"
        'portfolioDiagrammtitel(PTpfdk.Budget) = "Budget Übersicht"
        'portfolioDiagrammtitel(PTpfdk.FitRisikoDependency) = "strategischer Fit, Risiko & Ausstrahlung"

        portfolioDiagrammtitel(PTpfdk.Phasen) = repMessages.getmsg(58)
        portfolioDiagrammtitel(PTpfdk.Rollen) = repMessages.getmsg(59)
        portfolioDiagrammtitel(PTpfdk.Kosten) = repMessages.getmsg(60)
        portfolioDiagrammtitel(PTpfdk.ErgebnisWasserfall) = summentitel1
        portfolioDiagrammtitel(PTpfdk.FitRisiko) = summentitel2
        portfolioDiagrammtitel(PTpfdk.Auslastung) = summentitel9
        portfolioDiagrammtitel(PTpfdk.UeberAuslastung) = summentitel10
        portfolioDiagrammtitel(PTpfdk.Unterauslastung) = summentitel11
        portfolioDiagrammtitel(PTpfdk.ZieleV) = summentitel6
        portfolioDiagrammtitel(PTpfdk.ZieleF) = summentitel7
        portfolioDiagrammtitel(PTpfdk.ComplexRisiko) = repMessages.getmsg(61)
        portfolioDiagrammtitel(PTpfdk.ZeitRisiko) = repMessages.getmsg(62)
        portfolioDiagrammtitel(PTpfdk.AmpelFarbe) = ""
        portfolioDiagrammtitel(PTpfdk.ProjektFarbe) = ""
        portfolioDiagrammtitel(PTpfdk.Meilenstein) = repMessages.getmsg(63)
        portfolioDiagrammtitel(PTpfdk.FitRisikoVol) = repMessages.getmsg(64)
        portfolioDiagrammtitel(PTpfdk.Dependencies) = repMessages.getmsg(65)
        portfolioDiagrammtitel(PTpfdk.betterWorseL) = repMessages.getmsg(66)
        portfolioDiagrammtitel(PTpfdk.betterWorseB) = repMessages.getmsg(67)
        portfolioDiagrammtitel(PTpfdk.Budget) = repMessages.getmsg(68)
        portfolioDiagrammtitel(PTpfdk.FitRisikoDependency) = repMessages.getmsg(69)

    End Sub



    ''' <summary>
    ''' setzt Kalenderleiste und Spaltenbreite sowie -Höhe 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub prepareProjektTafel()


        ' bestimmen der Spaltenbreite und Spaltenhöhe ...
        Dim testCase As String = appInstance.ActiveWorkbook.Name
        If testCase <> myProjektTafel Then
            CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook).Activate()
        End If

        Dim wsName3 As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)),
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

        Dim tmpRange As Excel.Range
        Dim tempWSName As String = CType(appInstance.ActiveSheet, Excel.Worksheet).Name

        'Dim tmpStart As Date
        Try

            Call prepareCalendar(wsName3)

            With wsName3

                ' ur: 19.02.2019: wird nun mit "Call prepareCalendar(wsName3)" erledigt

                'Dim rng As Excel.Range
                ''Dim colDate As date
                'If awinSettings.zeitEinheit = "PM" Then
                '    ' die Kalender-Leiste schreiben 
                '    CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar
                '    CType(.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(1)
                '    rng = .Range(.Cells(1, 1), .Cells(1, 2))
                '    '' Deutsches Format:
                '    'rng.NumberFormat = "[$-407]mmm yy;@"
                '    ' Englische Format:
                '    rng.NumberFormat = "[$-409]mmm yy;@"

                '    Dim destinationRange As Excel.Range = .Range(.Cells(1, 1), .Cells(1, 720))
                '    With destinationRange
                '        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                '        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                '        '' Deutsches Format: 
                '        'rng.NumberFormat = "[$-407]mmm yy;@"
                '        ' Englische Format:
                '        .NumberFormat = "[$-409]mmm yy;@"
                '        .WrapText = False
                '        .Orientation = 90
                '        .AddIndent = False
                '        .IndentLevel = 0
                '        ' Änderung tk 14.11 - sonst können ja die Spaltenbreiten ud Höhen nicht explizit gesetzt werden 
                '        ' das ist vor allem auf der Zeichenfläche notwendig, weil sonst die Berechnung und Positionierung der Grafik Elemente nicht mehr stimmt 
                '        .ShrinkToFit = True
                '        .ReadingOrder = Excel.Constants.xlContext
                '        .MergeCells = False
                '        .Interior.Color = noshowtimezone_color
                '        .Font.Color = calendarFontColor
                '    End With

                '    rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)

                'ElseIf awinSettings.zeitEinheit = "PW" Then
                '    For i As Integer = 1 To 210
                '        CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddDays((i - 1) * 7)
                '    Next
                'ElseIf awinSettings.zeitEinheit = "PT" Then
                '    Dim workOnSat As Boolean = False
                '    Dim workOnSun As Boolean = False


                '    If Weekday(StartofCalendar, FirstDayOfWeek.Monday) > 3 Then
                '        tmpStart = StartofCalendar.AddDays(8 - Weekday(StartofCalendar, FirstDayOfWeek.Monday))
                '    Else
                '        tmpStart = StartofCalendar.AddDays(Weekday(StartofCalendar, FirstDayOfWeek.Monday) - 8)
                '    End If
                '    '
                '    ' jetzt ist tmpstart auf Montag ... 
                '    Dim tmpDay As Date
                '    Dim i As Integer = 1

                '    For w As Integer = 1 To 30
                '        For d As Integer = 0 To 4
                '            ' das sind Montag bis Freitag
                '            tmpDay = tmpStart.AddDays(d)
                '            If Not feierTage.Contains(tmpDay) Then
                '                CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                '                i = i + 1
                '            End If
                '        Next
                '        tmpDay = tmpStart.AddDays(5)
                '        If workOnSat Then
                '            CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                '            i = i + 1
                '        End If
                '        tmpDay = tmpStart.AddDays(6)
                '        If workOnSun Then
                '            CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                '            i = i + 1
                '        End If
                '        tmpStart = tmpStart.AddDays(7)
                '    Next


                'End If


                ' hier werden jetzt die Spaltenbreiten und Zeilenhöhen gesetzt 

                Dim maxRows As Integer = .Rows.Count
                Dim maxColumns As Integer = .Columns.Count

                tmpRange = CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range)
                CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
                CType(.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe2
                CType(.Columns, Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = awinSettings.spaltenbreite

                With CType(.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range)
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    .NumberFormat = "####0"
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    ' Änderung tk 14.11 - sonst können ja die Spaltenbreiten ud Höhen nicht explizit gesetzt werden 
                    ' das ist vor allem auf der Zeichenfläche notwendig, weil sonst die Berechnung und Positionierung der Grafik Elemente nicht mehr stimmt 
                    .ShrinkToFit = True
                    .ReadingOrder = Excel.Constants.xlContext
                    .MergeCells = False
                End With

                boxWidth = CDbl(CType(.Cells(3, 3), Global.Microsoft.Office.Interop.Excel.Range).Width)
                boxHeight = CDbl(CType(.Cells(3, 3), Global.Microsoft.Office.Interop.Excel.Range).Height)

                topOfMagicBoard = CDbl(CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range).Height) + 0.1 * boxHeight
                screen_correct = 0.1 * 19.3 / boxWidth


                Dim laenge As Integer
                laenge = showRangeRight - showRangeLeft

                If laenge > 0 And showRangeLeft > 0 Then

                    CType(.Range(.Cells(1, showRangeLeft), .Cells(1, showRangeLeft + laenge)), Excel.Range).Interior.Color = showtimezone_color
                    CType(.Range(.Cells(1, showRangeLeft), .Cells(1, showRangeLeft + laenge)), Excel.Range).Font.Color = calendarFontColor

                End If

            End With
        Catch ex As Exception

        End Try

    End Sub

    Public Sub prepareCalendar(ByVal wsname As Microsoft.Office.Interop.Excel.Worksheet)

        Dim tmpStart As Date
        Try
            With wsname
                Dim rng As Excel.Range
                'Dim colDate As date
                If awinSettings.zeitEinheit = "PM" Then
                    ' die Kalender-Leiste schreiben 
                    CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar
                    CType(.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(1)
                    rng = .Range(.Cells(1, 1), .Cells(1, 2))
                    '' Deutsches Format:
                    'rng.NumberFormat = "[$-407]mmm yy;@"
                    ' Englische Format:
                    rng.NumberFormat = "[$-409]mmm yy;@"

                    Dim destinationRange As Excel.Range = .Range(.Cells(1, 1), .Cells(1, 720))
                    With destinationRange
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        '' Deutsches Format: 
                        'rng.NumberFormat = "[$-407]mmm yy;@"
                        ' Englische Format:
                        .NumberFormat = "[$-409]mmm yy;@"
                        .WrapText = False
                        .Orientation = 90
                        .AddIndent = False
                        .IndentLevel = 0
                        ' Änderung tk 14.11 - sonst können ja die Spaltenbreiten ud Höhen nicht explizit gesetzt werden 
                        ' das ist vor allem auf der Zeichenfläche notwendig, weil sonst die Berechnung und Positionierung der Grafik Elemente nicht mehr stimmt 
                        .ShrinkToFit = True
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .Interior.Color = noshowtimezone_color
                        .Font.Color = calendarFontColor
                    End With

                    rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)

                ElseIf awinSettings.zeitEinheit = "PW" Then
                    For i As Integer = 1 To 210
                        CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddDays((i - 1) * 7)
                    Next
                ElseIf awinSettings.zeitEinheit = "PT" Then
                    Dim workOnSat As Boolean = False
                    Dim workOnSun As Boolean = False


                    If Weekday(StartofCalendar, FirstDayOfWeek.Monday) > 3 Then
                        tmpStart = StartofCalendar.AddDays(8 - Weekday(StartofCalendar, FirstDayOfWeek.Monday))
                    Else
                        tmpStart = StartofCalendar.AddDays(Weekday(StartofCalendar, FirstDayOfWeek.Monday) - 8)
                    End If
                    '
                    ' jetzt ist tmpstart auf Montag ... 
                    Dim tmpDay As Date
                    Dim i As Integer = 1

                    For w As Integer = 1 To 30
                        For d As Integer = 0 To 4
                            ' das sind Montag bis Freitag
                            tmpDay = tmpStart.AddDays(d)
                            If Not feierTage.Contains(tmpDay) Then
                                CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                                i = i + 1
                            End If
                        Next
                        tmpDay = tmpStart.AddDays(5)
                        If workOnSat Then
                            CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                            i = i + 1
                        End If
                        tmpDay = tmpStart.AddDays(6)
                        If workOnSun Then
                            CType(.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpDay.ToString("d")
                            i = i + 1
                        End If
                        tmpStart = tmpStart.AddDays(7)
                    Next


                End If
            End With
        Catch ex As Exception

        End Try


    End Sub



    ''' <summary>
    ''' liest die Konstellationen und Abhängigkeiten in der Datenbank 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub readInitConstellations()

        Dim err As New clsErrorCodeMsg

        ' Datenbank ist gestartet
        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

            ' alle Konstellationen laden 
            projectConstellations = CType(databaseAcc, DBAccLayer.Request).retrieveConstellationsFromDB(Date.Now, err)

            ' hier werden jetzt auch alle Abhängigkeiten geladen 
            allDependencies = CType(databaseAcc, DBAccLayer.Request).retrieveDependenciesFromDB()

            Dim axt As Integer = 9

        Else
            Throw New ArgumentException("Datenbank - Verbindung ist unterbrochen ...")
        End If

    End Sub
    '
    '
    '
    Public Sub awinChangeTimeSpan(ByVal von As Integer, ByVal bis As Integer,
                                  Optional ByVal noFurtherActions As Boolean = False)

        'Dim k As Integer

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim noTimeFrame As Boolean = False

        appInstance.EnableEvents = False



        If von < 1 Then
            von = 1
        End If

        If bis < von + minColumns - 1 Then
            noTimeFrame = True
        End If

        ' damit es nicht flackert, wenn zweimal hintereinander Zeitzone aufgehoben wird  
        If noTimeFrame Then
            If showRangeRight <> showRangeLeft Then
                appInstance.ScreenUpdating = False
            End If
        Else
            appInstance.ScreenUpdating = False
        End If


        If showRangeLeft <> von Or showRangeRight <> bis Or
            AlleProjekte.Count = 0 Then
            '
            ' wenn roentgenblick.ison , werden Bedarfe angezeigt - die müssen hier ausgeblendet werden - nachher mit den neuen Werten eingeblendet werden
            '



            '
            ' aktualisieren der Showtime zone, erst die alte ausblenden , dann die neue einblenden
            '
            Call awinShowtimezone(showRangeLeft, showRangeRight, False)

            If noTimeFrame Then

                showRangeLeft = 0
                showRangeRight = 0


            Else
                Call awinShowtimezone(von, bis, True)


                showRangeLeft = von
                showRangeRight = bis


                ' jetzt werden - falls nötig die Projekte nachgeladen ... 
                Try
                    If awinSettings.applyFilter Then
                        ' vorher hiess das loadprojectsonChange - jetzt ist es so: 
                        ' wenn applyFilter = true, dann soll nachgeladen werden unter Anwendung 
                        ' des Filters "Last"
                        Dim filter As New clsFilter
                        filter = filterDefinitions.retrieveFilter("Last")
                        Call awinProjekteImZeitraumLaden(awinSettings.databaseName, filter)

                        '' jetzt sind wieder alle Projekte des Zeitraums da - deswegen muss nicht ggf nachgeladen werden 
                        'DeletedProjekte.Clear()

                        '
                        '   wenn "selectedRoleNeeds" ungleich Null ist, werden Bedarfe angezeigt - die müssen hier wieder - mit den neuen Daten für show_range_lefet, .._right eingeblendet werden
                        '


                        '
                        ' wenn diagramme angezeigt sind - aktualisieren dieser Diagramme
                        '



                    End If

                    ' betrachteter Zeitraum wurde geändert - typus = 4
                    Call awinNeuZeichnenDiagramme(4)


                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try
            End If






        End If



        appInstance.EnableEvents = formerEE
        If appInstance.ScreenUpdating <> formerSU Then
            appInstance.ScreenUpdating = formerSU
        End If



    End Sub


    ''' <summary>
    ''' Es wird ein Projekt in  VISBO-Stuktur erzeugt.
    ''' msproj, ein Projekt von MS Project wird in ein Projekt hproj gemäß der Struktur mapStruktur gemappt.
    ''' die entsprechende Vorschriften sind in den Custom-Field visbo_mapping definiert.
    ''' Ergebnis: gemapptes Projekt
    ''' </summary>
    ''' <param name="msproj">MS Project - Projekt</param>
    ''' <param name="mapStruktur">Vorlage</param>
    ''' <param name="hproj">bereits von MS Project eingelesenes Projekt in VISBO-Struktur</param>
    ''' <param name="visbo_mapping">Definition des MSProject-CustomField</param>
    ''' <returns>gemapptes Projekt in VISBO-Struktur</returns>
    ''' <remarks></remarks>
    Public Function mappingProject(ByVal msproj As MSProject.Project,
                                   ByVal mapStruktur As String,
                                   ByVal hproj As clsProjekt,
                                   ByVal visbo_mapping As MSProject.PjCustomField,
                                   ByVal wbs_elemID_liste As SortedList(Of String, String)) As clsProjekt




        ' hproj ist das bereits aus msproj von MSProject erzeugte Original-Projekt

        ' mproj bezeichnet das gemappte Projekt, also das Ergebnis
        Dim mproj As clsProjekt = Nothing

        ' vproj bezeichnet die Vorlage für das Mapping hier das TMS-Projekt
        Dim vproj As clsProjekt = Nothing

        Dim mPhase As clsPhase = Nothing
        Dim mMilestone As clsMeilenstein = Nothing
        Dim vMappingText As String = ""

        Dim vmappingtext_inclBreadC As String = ""
        Dim aktPhaseBC As String = ""

        Dim msTask As MSProject.Task
        Dim anztasks As Integer = msproj.Tasks.Count


        ' vMapping = true, wenn Mapping-Spalte Inhalte hat
        Dim vMapping As Boolean = False

        ' Für Check-Message
        Dim outputCollection As New Collection
        Dim outputline As String = ""

        ' -------------------------------------------------------------------------
        ' Check, ob gemappt werden muss (visbo_mapping enthält Angaben zum Mapping)
        '
        ' Bestimmung von minDate und maxDate des gemappten Projektes
        ' -------------------------------------------------------------------------

        ' MINimum und MAXimum Datum für Start und Ende des TMS-Projektes zu finden

        Dim minDate As Date = hproj.endeDate
        Dim maxDate As Date = hproj.startDate

        For i = 1 To anztasks

            msTask = msproj.Tasks.Item(i)

            Try
                vMappingText = Trim(msTask.GetField(visbo_mapping))
            Catch ex As Exception
                vMappingText = ""
            End Try


            If vMappingText <> "" Then

                If minDate > msTask.Start Then
                    minDate = msTask.Start
                End If
                If maxDate < msTask.Finish Then
                    maxDate = msTask.Finish
                End If
                vMapping = vMapping Or True
            End If

        Next i

        ' von beiden Datum nur die Datumsvariante hernehmen
        minDate = minDate.Date
        maxDate = maxDate.Date

        ' ENDE min-max - Bestimmung
        ' ------------------------------


        If vMapping Then

            ' es wird kein existierendes Projekt als Vorlage verwendet 
            Dim myProject As clsProjekt = Nothing

            ' tk 7.6.21 
            Dim budgetVorgabe As Double = hproj.Erloes
            Try
                budgetVorgabe = hproj.getGesamtKostenBedarf.Sum
            Catch ex As Exception

            End Try
            vproj = erstelleProjektAusVorlage(myProject, "TMSHilfsproj", mapStruktur, minDate, maxDate, budgetVorgabe, 0,
                                     hproj.StrategicFit, hproj.Risiko, Nothing, hproj.description, hproj.businessUnit)

            If Not IsNothing(vproj) Then

                mproj = New clsProjekt(minDate, minDate, minDate)
                mproj.variantName = mapStruktur
                Try
                    With mproj
                        .name = hproj.name
                        .VorlagenName = vproj.VorlagenName
                        .startDate = vproj.startDate
                        .businessUnit = vproj.businessUnit
                        .Erloes = vproj.Erloes
                        .earliestStartDate = vproj.earliestStartDate
                        .latestStartDate = vproj.latestStartDate
                        .Status = vproj.Status
                        .description = vproj.description

                        .StrategicFit = vproj.StrategicFit
                        .Risiko = vproj.Risiko
                        'plen = .anzahlRasterElemente
                        'pcolor = .farbe
                    End With
                Catch ex As Exception

                End Try
                ' alle Phasen des TMS_Projektes vproj durchgehen, in das mproj eintragen und in MSProjekt 
                ' die zugehörigen Phasen und Meilensteine suchen und übernehmen aus hproj


                ' übernehmen der RootPhase aus vproj
                Dim cphase As New clsPhase(mproj)
                vproj.AllPhases.ElementAt(0).copyTo(cphase,
                                                    withoutNameID:=True,
                                                    withoutMS:=True,
                                                    withoutRolesCosts:=True)
                cphase.nameID = rootPhaseName
                mproj.AddPhase(cphase)

                For hi As Integer = 1 To vproj.AllPhases.Count - 1

                    Dim aktPhase As New clsPhase(mproj)
                    vproj.AllPhases.ElementAt(hi).copyTo(aktPhase,
                                                         withoutNameID:=True,
                                                         withoutMS:=True,
                                                         withoutRolesCosts:=True)

                    aktPhase.nameID = mproj.hierarchy.findUniqueElemKey(vproj.AllPhases.ElementAt(hi).name, False)
                    Dim parentID As String = vproj.hierarchy.getParentIDOfID(aktPhase.nameID)

                    ' aktuelle Phase des VorlagenProjekts in MappingProjekt übernehmen
                    mproj.AddPhase(aktPhase, aktPhase.name, parentID)


                    For i = 1 To anztasks

                        msTask = msproj.Tasks.Item(i)

                        Try
                            vMappingText = Trim(msTask.GetField(visbo_mapping))
                        Catch ex As Exception
                            vMappingText = ""
                        End Try



                        If vMappingText.Contains(".") Then
                            vmappingtext_inclBreadC = ".#" & vMappingText.Replace(".", "#")
                            aktPhaseBC = mproj.hierarchy.getBreadCrumb(aktPhase.nameID)
                        Else
                            vmappingtext_inclBreadC = ""
                            aktPhaseBC = ""
                        End If

                        If vMappingText = aktPhase.name _
                            Or vmappingtext_inclBreadC = (aktPhaseBC & "#" & aktPhase.name) Then

                            If Not CType(msTask.Milestone, Boolean) Or
                                (CType(msTask.Milestone, Boolean) And CType(msTask.Summary, Boolean)) Then

                                ' mstask ist Phase

                                mPhase = New clsPhase(mproj)

                                'ur: 27.02.2020: Korrektur des Mappings für BHTC und auch generell
                                Dim elemID As String = wbs_elemID_liste(msTask.WBS)
                                Dim hPhase As clsPhase = hproj.getPhaseByID(elemID)
                                'Dim hPhase As clsPhase = hproj.getPhase(msTask.Name)

                                hPhase.copyTo(mPhase,
                                              withoutNameID:=True,
                                              withoutMS:=True,
                                              withoutRolesCosts:=True)

                                mPhase.nameID = mPhase.parentProject.hierarchy.findUniqueElemKey(msTask.Name, False)
                                Try
                                    ' Berechnung Phasen-Start
                                    Dim mphaseStartOffset As Long
                                    Dim dauerIndays As Long
                                    mphaseStartOffset = DateDiff(DateInterval.Day, minDate, CDate(msTask.Start).Date)
                                    dauerIndays = calcDauerIndays(CDate(msTask.Start).Date, CDate(msTask.Finish).Date)
                                    mPhase.changeStartandDauer(mphaseStartOffset, dauerIndays)
                                    mPhase.offset = 0

                                    Dim mphasestart As Date = mPhase.getStartDate
                                    Dim mphaseende As Date = mPhase.getEndDate

                                    ' Verification Check
                                    If DateDiff(DateInterval.Day, CDate(msTask.Start).Date, mphasestart.Date) <> 0 Then
                                        outputline = "(Phase) : " & msTask.Name & "beginnt:(MSProject):" & CDate(msTask.Start).Date.ToShortDateString & " - " & "(VISBO):" & mphasestart.ToShortDateString
                                        outputCollection.Add(outputline)
                                    End If
                                    If DateDiff(DateInterval.Day, CDate(msTask.Finish).Date, mphaseende.Date) <> 0 Then
                                        outputline = "(Phase) : " & msTask.Name & "endet:(MSProject):" & CDate(msTask.Finish).Date.ToShortDateString & " - " & "(VISBO):" & mphaseende.ToShortDateString
                                        outputCollection.Add(outputline)
                                    End If



                                    ' eintragen Phase
                                    mproj.AddPhase(mPhase, msTask.Name, aktPhase.nameID)
                                Catch ex As Exception
                                    Call MsgBox(ex.Message)
                                End Try

                            Else
                                ' mstask ist Meilenstein

                                aktPhase = mproj.getPhaseByID(aktPhase.nameID)

                                mMilestone = New clsMeilenstein(aktPhase)

                                'ur: 27.02.2020: Korrektur Mapping für BHTC hier nachgezogen
                                Dim elemID As String = wbs_elemID_liste(msTask.WBS)
                                Dim hMilestone As clsMeilenstein = hproj.getMilestoneByID(elemID)
                                'Dim hMilestone As clsMeilenstein = hproj.getMilestone(msTask.Name)

                                Dim newMSNameID As String = aktPhase.parentProject.hierarchy.findUniqueElemKey(msTask.Name, True)
                                hMilestone.copyTo(mMilestone, newMSNameID)

                                Dim hMSDate As Date = hMilestone.getDate
                                mMilestone.setDate = hMSDate

                                Dim testDate As Date = mMilestone.getDate

                                ' Verification Check
                                If DateDiff(DateInterval.Day, hMilestone.getDate.Date, mMilestone.getDate.Date) <> 0 Then
                                    outputline = "Milestone : " & msTask.Name & " : (MSProject):" & hMilestone.getDate.ToShortDateString & " - " & "(VISBO):" & mMilestone.getDate.ToShortDateString
                                    outputCollection.Add(outputline)
                                End If


                                Try
                                    aktPhase.addMilestone(mMilestone, origName:=msTask.Name)

                                Catch ex As Exception
                                    Call MsgBox(ex.Message)
                                End Try

                            End If

                        End If

                    Next i


                Next hi

                mappingProject = mproj


                If outputCollection.Count > 0 Then
                    Call showOutPut(outputCollection, "Mapping " & mproj.name & " TMS-Variante", "folgende Ungereimtheiten In den Daten wurden festgestellt")
                    mappingProject = Nothing
                End If

            Else

                mappingProject = Nothing

            End If

        Else

            ' Kein Mapping definiert

            mappingProject = Nothing
        End If


    End Function


    ''' <summary>
    ''' Methode trägt alle Projekte aus ImportProjekte in AlleProjekte bzw. Showprojekte ein, sofern die Anzahl mit der myCollection übereinstimmt
    ''' die Projekte werden in der Reihenfolge auf das Board gezeichnet, wie sie in der ImportProjekte aufgeführt sind
    ''' wenn ein importiertes Projekt bereits in der Datenbank existiert und verändert ist, dann wird es markiert und gleichzeitig temporär geschützt 
    ''' wenn ein importiertes Projekt bereits in der Datenbank existiert, verändert wurde und von anderen geschützt ist, dann wird eine Variante angelegt 
    ''' </summary>
    ''' <param name="importDate"></param>
    ''' <param name="drawPlanTafel">sollen die PRojekte gezeichnet werden</param>
    ''' <param name="fileFrom3rdParty">stammt der Import von einer 3rd Party ab, müssen also evtl Ressourcen etc ergänzt werden</param>
    ''' <remarks></remarks>
    Public Sub importProjekteEintragen(ByVal importDate As Date, ByVal drawPlanTafel As Boolean,
                                       ByVal fileFrom3rdParty As Boolean,
                                       ByVal getSomeValuesFromOldProj As Boolean,
                                       Optional ByVal calledFromActualDataImport As Boolean = False)

        Dim err As New clsErrorCodeMsg

        Dim hproj As New clsProjekt, formerProj As New clsProjekt
        Dim fullName As String, vglName As String
        'Dim pname As String

        Dim anzAktualisierungen As Integer, anzNeuProjekte As Integer
        Dim tafelZeile As Integer = 2
        'Dim shpElement As Excel.Shape
        Dim phaseList As New Collection
        Dim milestoneList As New Collection
        Dim wasNotEmpty As Boolean

        Dim existsInSession As Boolean = False


        ' aus der Datenbank alle WriteProtections holen ...
        If Not noDB And AlleProjekte.Count > 0 Then
            writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)
        End If

        If AlleProjekte.Count > 0 Then
            wasNotEmpty = True
            tafelZeile = projectboardShapes.getMaxZeile
        Else
            wasNotEmpty = False
        End If


        Dim differentToPrevious As Boolean = False


        anzAktualisierungen = 0
        anzNeuProjekte = 0

        ' tk wenn es Namensänderungen gibt, dann sollen die hier angezeigt werden ... 

        Dim nameChangeCollection As New Collection

        ' jetzt werden alle importierten Projekte bearbeitet 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste

            ' wenn ein Projekt importiert wird, das durch andere geschützt ist , so wird eine neue Variante angelegt
            ' dann soll das ursprüngliche Projekt , sofern es in de rSession existiert, nicht aus der Session gelöscht werden 
            Dim newVariantGenerated As Boolean = False
            fullName = kvp.Key
            hproj = kvp.Value


            ' jetzt muss überprüft werden, ob dieses Projekt bereits in AlleProjekte / Showprojekte existiert 
            ' wenn ja, muss es um die entsprechenden Werte dieses Projektes (Status, etc)  ergänzt werden
            ' wenn nein, wird es im Show-Modus ergänzt 
            Dim searchPName As String = hproj.name
            Dim searchVName As String = hproj.variantName
            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And searchVName = "" Then
                ' das hier muss gemacht werden, weil man ja wissen will, inwieweit sich das Projekt im Vergleich zur Baseline / pfv verändert werden 
                searchVName = getDefaultVariantNameAccordingUserRole()
            End If

            vglName = calcProjektKey(searchPName, searchVName)
            Try
                If kvp.Value.kundenNummer = "" Then
                    formerProj = AlleProjekte.getProject(vglName)
                Else
                    formerProj = AlleProjekte.getProjectByKDNr(kvp.Value.kundenNummer)
                End If


                If IsNothing(formerProj) Then
                    ' jetzt muss geprüft werden, ob das Projekt bereits in der Datenbank existiert ... 
                    existsInSession = False
                    If Not noDB Then
                        formerProj = awinReadProjectFromDatabase(kvp.Value.kundenNummer, searchPName, searchVName, Date.Now)
                    End If
                Else
                    existsInSession = True
                End If

                ' ist es immer noch Nothing ? 
                If IsNothing(formerProj) Then
                    ' wenn es jetzt immer noch Nothing ist, dann existiert es weder in der Datenbank noch in der Session .... 

                    ' falls es sich um eine Variante handelt, muss jetzt geprüft werden, ob die Basis-Variante in Session oder DB existiert  
                    Dim baseProj As clsProjekt = Nothing

                    If searchVName <> "" Then
                        ' dann muss evtl aus dem Basis Projekt was geholt werden 
                        baseProj = AlleProjekte.getProject(calcProjektKey(hproj.name, ""))

                        If IsNothing(baseProj) Then
                            ' jetzt muss geprüft werden, ob das Projekt bereits in der Datenbank existiert ... 
                            If Not noDB Then
                                baseProj = awinReadProjectFromDatabase(kvp.Value.kundenNummer, hproj.name, "", Date.Now)
                            End If
                        End If
                    End If


                    Try
                        With hproj
                            ' 5.5.2014 ur: soll nicht wieder auf 0 gesetzt werden, sondern Einstellung beibehalten
                            '.earliestStart = 0
                            .earliestStartDate = .startDate
                            .latestStartDate = .startDate
                            .Id = vglName & "#" & importDate.ToString
                            ' 5.5.2014 ur: soll nicht wieder auf 0 gesetzt werden, sondern Einstellung beibehalten
                            '.latestStart = 0

                            ' Änderung tk 12.12.15: LeadPerson darf doch nicht auf leer gesetzt werden ...
                            '.leadPerson = " "
                            .shpUID = ""
                            .StartOffset = 0

                            ' ein importiertes Projekt soll normalerweise immer gleich  auf "beauftragt" gesetzt werden; 
                            ' das kann aber jetzt an der aufrufenden Stelle gesetzt werden 
                            ' Inventur: erst mal auf geplant, sonst beauftragt 
                            '.Status = pStatus
                            If Not IsNothing(baseProj) Then
                                .Status = baseProj.Status
                                If baseProj.name <> hproj.name Then

                                End If
                            End If

                            .tfZeile = tafelZeile
                            .timeStamp = importDate

                        End With

                        ' Workaround: 
                        Dim tmpValue As Integer = hproj.dauerInDays
                        ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                        'Call awinCreateBudgetWerte(hproj)
                        tafelZeile = tafelZeile + 1

                        anzNeuProjekte = anzNeuProjekte + 1
                    Catch ex1 As Exception
                        Throw New ArgumentException("Fehler bei Übernahme der Attribute des alten Projektes" & vbLf & ex1.Message)
                    End Try
                Else

                    ' jetzt sollen bestimmte Werte aus der früheren Projekt-Version übernommen werden 
                    ' das ist dann wichtig, wenn z.Bsp nur Rplan Excel Werte eingelesen werden, die enthalten ja nix ausser Termine ...
                    ' und in dem Fall können ja interaktiv bzw. über Export/Import Visbo Steckbrief Werte gesetzt worden sein 

                    ' prüfen: hat cproj invoices? und hproj keine ? 
                    Try
                        Dim array1 As Double() = formerProj.getInvoicesPenalties
                        Dim array2 As Double() = hproj.getInvoicesPenalties
                        If array1.Sum > 0 And array2.Sum = 0 Then
                            Call hproj.updateProjectwithInvoicesFrom(formerProj)
                        End If
                    Catch ex As Exception

                    End Try


                    Try
                        If getSomeValuesFromOldProj Then
                            Call awinAdjustValuesByExistingProj(hproj, formerProj, existsInSession, importDate, tafelZeile, fileFrom3rdParty)
                        End If

                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                    ' jetzt sicherstellen, dass der Vergleich nicht einfach aufgrund Unterschied im VarantName zu einem Unterschied, damit Markierung führt ... 
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And hproj.variantName <> formerProj.variantName Then
                        formerProj.variantName = hproj.variantName
                    End If

                    ' evtl ist jetzt der Name ein anderer , weil über P-Nr aus Datenbank geholt  
                    If Not IsNothing(formerProj) Then
                        ' Fall Telair oder auch andere: wenn PNr angeben ist und die PNames aus DB und Import nicht übereinstimmen, dann gewinnt DB 
                        If formerProj.name <> hproj.name Then
                            Dim txtMsg As String = "from " & hproj.name & " to " & formerProj.name
                            nameChangeCollection.Add(txtMsg)

                            hproj.name = formerProj.name
                        End If
                    End If

                    If Not hproj.isIdenticalTo(vProj:=formerProj) Then
                        ' das heisst, das Projekt hat sich verändert 
                        hproj.marker = True
                        'If hproj.Status = ProjektStatus(PTProjektStati.beauftragt) Then
                        '    hproj.Status = ProjektStatus(PTProjektStati.ChangeRequest)
                        'End If

                        ' wenn das Projekt bereits von anderen geschützt ist, soll es als Variante angelegt werden 
                        ' andernfalls soll es von mir geschützt werden ; allerdings soll es nur dann einen temporärewn Schutz bekommen, 
                        ' wenn es nicht schon von mir permanent geschützt ist 
                        If Not noDB Then
                            Dim wpItem As clsWriteProtectionItem

                            Dim isProtectedbyOthers As Boolean = Not tryToprotectProjectforMe(hproj.name, searchVName)

                            If isProtectedbyOthers Then

                                ' nicht erfolgreich, weil durch anderen geschützt ... 
                                ' oder aber noch gar nicht in Datenbank: aber das ist noch nicht berücksichtigt  
                                wpItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                                writeProtections.upsert(wpItem)

                                ' jetzt Variante anlegen 
                                Dim teilName As String = dbUsername
                                If dbUsername.Length > 4 Then
                                    teilName = dbUsername.Substring(0, 4)
                                End If
                                Dim newVname As String = "I" & teilName
                                hproj.variantName = newVname

                                ' jetzt das Flag setzen 
                                newVariantGenerated = True
                            End If

                        End If


                    Else
                        hproj.marker = False
                    End If

                    ' jetzt müssen hier noch die ID's stabiliert werden: gleicher BreadCrumb hat immer gleiche ID 
                    Dim baselineProj As clsProjekt = awinReadProjectFromDatabase(hproj.kundenNummer, hproj.name, ptVariantFixNames.pfv.ToString, Date.Now)
                    Dim lastProj As clsProjekt = awinReadProjectFromDatabase(hproj.kundenNummer, hproj.name, hproj.variantName, Date.Now)


                    Dim baseLineBreadCrumbIDList As New SortedList(Of String, String)
                    Dim lastProjBreadCrumbIDList As New SortedList(Of String, String)
                    Dim myBreadCrumbList As SortedList(Of String, String) = hproj.getBreadCrumbIDList

                    If Not IsNothing(baselineProj) Then
                        baseLineBreadCrumbIDList = baselineProj.getBreadCrumbIDList
                    End If

                    If Not IsNothing(lastProj) Then
                        lastProjBreadCrumbIDList = lastProj.getBreadCrumbIDList
                    End If

                    ' hproj wird in den NameIDs nur angepasst, wenn es tatsächlich auch Änderungen gibt ... 
                    If baseLineBreadCrumbIDList.Count + lastProjBreadCrumbIDList.Count > 0 Then
                        Dim tmpProj As clsProjekt = hproj.ensureStableIDs(baseLineBreadCrumbIDList, lastProjBreadCrumbIDList)

                        ' hier kommt die Prüfung ... 
                        If Not IsNothing(tmpProj) Then

                            Dim tmpBreadCrumbList As SortedList(Of String, String) = tmpProj.getBreadCrumbIDList
                            Dim outPutCollection As Collection = checkIDStability(tmpProj, baselineProj, lastProj)

                            ' wenn outputCollection keine Fehlermeldungen enthält , dann wird das tmpProj übernommen ..
                            If outPutCollection.Count > 0 Then
                                Call showOutPut(outPutCollection, "Fehler bei Enable Stable IDs", "")
                            Else

                                Dim wasMarked As Boolean = hproj.marker
                                hproj = tmpProj
                                hproj.marker = wasMarked


                            End If

                        End If

                    End If

                    If (Not calledFromActualDataImport) And (myCustomUserRole.customUserRole <> ptCustomUserRoles.PortfolioManager) Then

                        ' jetzt sicherstellen, dass das Projekt die Ist-Daten aus dem alten Projekt bekommt.  
                        Try
                            Dim oldVname As String = hproj.variantName
                            Dim wasMarked As Boolean = hproj.marker
                            Dim tmpProj As clsProjekt = hproj.createVariant("$MergeTemp", hproj.variantDescription)
                            Call tmpProj.mergeActualValues(formerProj)
                            ' wenn alles gut ging ...
                            hproj = tmpProj
                            hproj.variantName = oldVname
                            hproj.marker = Not hproj.isIdenticalTo(vProj:=formerProj)

                            If hproj.marker <> wasMarked Then
                                Dim txtMsg As String = hproj.name & ": old actual data was restored by former project description ... "
                                nameChangeCollection.Add(txtMsg)
                            End If

                        Catch ex As Exception
                            ' nichts tun ... 
                            Dim txtMsg As String = ex.Message & vbLf & "project " & hproj.name & " was imported without merging actual data"
                            nameChangeCollection.Add(txtMsg)
                        End Try

                    End If


                    anzAktualisierungen = anzAktualisierungen + 1

                    Try
                        If newVariantGenerated Then
                            ' das alte in AlleProjekte lassen 
                            ' das alte in ShowProjekte rausnehmen  
                            If ShowProjekte.contains(hproj.name) Then
                                ShowProjekte.Remove(hproj.name)
                            End If

                            ' das muss auch gemacht werden, wenn hproj.marker = true
                        ElseIf existsInSession Or hproj.marker = True Then
                            AlleProjekte.Remove(vglName)
                            If ShowProjekte.contains(hproj.name) Then
                                ShowProjekte.Remove(hproj.name, False)
                            End If
                        End If


                    Catch ex1 As Exception
                        Throw New ArgumentException("Fehler beim Update des Projektes " & ex1.Message)
                    End Try

                End If


            Catch ex As Exception

                Call MsgBox("Fehler in Import: Line Number 1438 " & vbLf & ex.Message)

            End Try

            ' in beiden Fällen - sowohl bei neu wie auch Aktualisierung muss jetzt das Projekt 
            ' sowohl auf der Plantafel eingetragen werden als auch in ShowProjekte und in alleProjekte eingetragen 

            ' bringe das neue Projekt in Showprojekte und in AlleProjekte
            If Not IsNothing(formerProj) Then
                ' Fall Telair oder auch andere: wenn PNr angeben ist und die PNames aus DB und Import nicht übereinstimmen, dann gewinnt DB 
                If formerProj.name <> hproj.name Then
                    hproj.name = formerProj.name
                End If
            End If



            Try
                vglName = calcProjektKey(hproj.name, hproj.variantName)
                If existsInSession Then
                    AlleProjekte.Add(hproj)
                    ShowProjekte.Add(hproj)
                Else
                    AlleProjekte.Add(hproj)
                    ShowProjekte.Add(hproj)
                    Try
                        Dim constItem As clsConstellationItem = currentSessionConstellation.getItem(vglName)
                        If Not IsNothing(constItem) Then
                            constItem.zeile = hproj.tfZeile
                        End If
                    Catch ex As Exception

                    End Try
                End If


            Catch ex As Exception
                'ur:16.1.2015: Dies ist kein Fehler sondern gewollt: 
                'Call MsgBox("Fehler bei Eintrag Showprojekte / Import " & hproj.name)
            End Try





        Next



        If ImportProjekte.Count < 1 Then
            If awinSettings.englishLanguage Then
                Call MsgBox(" no projects imported ...")
            Else
                Call MsgBox(" es wurden keine Projekte importiert ...")
            End If

        Else

            If awinSettings.englishLanguage Then

                Dim txtMsg As String = ImportProjekte.Count & " projects were read " & vbLf & vbLf &
                        anzNeuProjekte.ToString & " New projects" & vbLf &
                        anzAktualisierungen.ToString & " project updates"
                nameChangeCollection.Add(txtMsg)
                'Call MsgBox(ImportProjekte.Count & " projects were read " & vbLf & vbLf &
                '        anzNeuProjekte.ToString & " New projects" & vbLf &
                '        anzAktualisierungen.ToString & " project updates")
            Else

                Dim txtMsg As String = "es wurden " & ImportProjekte.Count & " Projekte bearbeitet!" & vbLf & vbLf &
                        anzNeuProjekte.ToString & " neue Projekte" & vbLf &
                        anzAktualisierungen.ToString & " Projekt-Aktualisierungen"

                nameChangeCollection.Add(txtMsg)

            End If

            If nameChangeCollection.Count > 0 Then
                Dim headerMsg As String = "project Names were changed To DB-Names"
                Call showOutPut(nameChangeCollection, headerMsg, "")
            End If


            If anzNeuProjekte > 0 Or anzAktualisierungen > 0 Then
                ' jetzt muss wieder entsprechend der 
                currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
            End If

            ' Änderung tk: jetzt wird das neu gezeichnet 
            ' wenn anzNeuProjekte > 0, dann hat sich die Konstellataion verändert 
            If currentConstellationPvName <> calcLastSessionScenarioName() Then
                currentConstellationPvName = calcLastSessionScenarioName()
            End If


            If drawPlanTafel Then
                If wasNotEmpty Then
                    Call awinClearPlanTafel()
                End If

                'Call awinZeichnePlanTafel(True)
                Call awinZeichnePlanTafel(True)
                Call awinNeuZeichnenDiagramme(2)
            End If

            'Call storeSessionConstellation("Last")

        End If



        ImportProjekte.Clear(False)

    End Sub

    ''' <summary>
    ''' prüft ob die IDs tatsächlich alle stabil sind
    ''' stabil heisst: jede Id aus der Baseline bzw. lastvpv , die im hproj existiert, hat den jeweils gleichen Breadcrumb
    ''' wenn Collection.count = 0 dann alles ok
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="baseLineProj">das Baseline Projekt</param>
    ''' <param name="lastProj">das last Projekt </param>
    ''' <returns></returns>
    Private Function checkIDStability(ByVal hproj As clsProjekt,
                                      ByVal baseLineProj As clsProjekt,
                                      ByVal lastProj As clsProjekt) As Collection

        ' die Prüfung, ob auch alle IDs jetzt die gleiche BreadCrumb haben 
        Dim result As New Collection
        Dim logMessage As String = ""

        Dim baseLineBreadCrumbIDList As New SortedList(Of String, String)
        Dim lastProjBreadCrumbIDList As New SortedList(Of String, String)

        If Not IsNothing(baseLineProj) Then
            baseLineBreadCrumbIDList = baseLineProj.getBreadCrumbIDList
        End If

        If Not IsNothing(lastProj) Then
            lastProjBreadCrumbIDList = lastProj.getBreadCrumbIDList
        End If

        For Each tstKvp As KeyValuePair(Of String, String) In baseLineBreadCrumbIDList
            If elemIDIstMeilenstein(tstKvp.Key) Then
                Dim tstMS As clsMeilenstein = hproj.getMilestoneByID(tstKvp.Key)
                If Not IsNothing(tstMS) Then
                    If Not hproj.getBcElemName(tstMS.nameID) = tstKvp.Value Then
                        logMessage = "baseline ungleich:   " & tstMS.nameID & "; " & tstKvp.Value
                        result.Add(logMessage)
                    End If
                End If
            Else
                Dim tstPh As clsPhase = hproj.getPhaseByID(tstKvp.Key)
                If Not IsNothing(tstPh) Then
                    If Not hproj.getBcElemName(tstPh.nameID) = tstKvp.Value Then
                        logMessage = "baseline ungleich " & tstPh.nameID & "; " & tstKvp.Value
                        result.Add(logMessage)
                    End If
                End If
            End If

        Next

        For Each tstKvp As KeyValuePair(Of String, String) In lastProjBreadCrumbIDList
            If elemIDIstMeilenstein(tstKvp.Key) Then
                Dim tstMS As clsMeilenstein = hproj.getMilestoneByID(tstKvp.Key)
                If Not IsNothing(tstMS) Then
                    If Not hproj.getBcElemName(tstMS.nameID) = tstKvp.Value Then
                        logMessage = "lastproj ungleich " & tstMS.nameID & "; " & tstKvp.Value
                        result.Add(logMessage)
                    End If
                End If
            Else
                Dim tstPh As clsPhase = hproj.getPhaseByID(tstKvp.Key)
                If Not IsNothing(tstPh) Then
                    If Not hproj.getBcElemName(tstPh.nameID) = tstKvp.Value Then
                        logMessage = "baseline ungleich " & tstPh.nameID & "; " & tstKvp.Value
                        result.Add(logMessage)
                    End If
                End If
            End If

        Next

        checkIDStability = result
    End Function
    ''' <summary>
    ''' is called when a 
    ''' </summary>
    ''' <param name="hproj"></param>
    Public Sub synchronizeWithValuesOFExisting(ByRef hproj As clsProjekt)

        Dim err As New clsErrorCodeMsg

        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() And Not noDB Then

            If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(hproj.name, hproj.variantName, hproj.timeStamp, err) Then
                ' prüfen, ob es Unterschied gibt 
                Dim standInDB As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, "", hproj.timeStamp, err)

                If Not IsNothing(standInDB) Then

                    Call awinAdjustValuesByExistingProj(hproj, standInDB, False, Date.Now, 2, True)

                End If


            End If

        End If


    End Sub

    ''' <summary>
    ''' übernimmt vom existierenden Projekt einige Werte wie Kundennummer, vpID, actualDataUntil: für VISBO steckbriefe oder Fremdsysteme
    ''' wenn vom Fremdsystem kommt: dann werden , wenn überhaupt keine Ressourcen (z.B MS Project)  da sind, die Ressourcen des vorherigen Standes genommen
    ''' ist vor allem dann relevant wenn nur ein RPLAN Excel mit gerademal Terminen eingelesen wird ....
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="formerProj"></param>
    ''' <param name="existsInSession"></param>
    ''' <param name="importDate"></param>
    ''' <param name="tafelZeile"></param>
    ''' <param name="fileFrom3rdParty">gibt an, ob es sich um einen VISBO Steckbrief handelt (false) oder einen fremden Steckbrief</param>
    ''' <remarks></remarks>
    Private Sub awinAdjustValuesByExistingProj(ByRef hproj As clsProjekt, ByVal formerProj As clsProjekt,
                                               ByVal existsInSession As Boolean, ByVal importDate As Date,
                                               ByRef tafelZeile As Integer,
                                               ByVal fileFrom3rdParty As Boolean)
        ' es existiert schon - deshalb müssen alle restlichen Werte aus dem cproj übernommen werden 
        Dim vglName As String = calcProjektKey(hproj)

        Try
            With hproj


                ' Änderung tk: das wird mit 28.12.16 nicht mehr benötigt ...  
                '.earliestStart = cproj.earliestStart
                '.earliestStartDate = cproj.earliestStartDate
                '.latestStart = cproj.latestStart
                '.latestStartDate = cproj.latestStartDate
                .earliestStartDate = .startDate
                .latestStartDate = .startDate

                .Id = vglName & "#" & importDate.ToString

                .StartOffset = 0
                .Status = formerProj.Status

                ' 
                ' jetzt muss in Abhäbgigeit von autoSetActualDate das actualData von cProj übernommen werden 
                ' das soll unabhängig vom autoSetActualData gemacht werden ... 
                hproj.actualDataUntil = formerProj.actualDataUntil

                ' übernehme die VPID 
                hproj.vpID = formerProj.vpID

                ' übernehme die Kunden-Nummer 
                hproj.kundenNummer = formerProj.kundenNummer


                If existsInSession Then
                    .shpUID = formerProj.shpUID
                    ' in diesem Fall heisst es ja genaus, dann ist es auch in der sortListe der Constellations bereits vorhanden ...
                    '.tfZeile = cproj.tfZeile
                Else
                    .shpUID = ""
                    .tfZeile = tafelZeile
                    tafelZeile = tafelZeile + 1
                End If


                .timeStamp = importDate
                .UID = formerProj.UID

                ' tk 19.2.20 nur wenn der Vorlagen-Name was anderes ist 
                If .VorlagenName = "" And formerProj.VorlagenName <> "" Then
                    .VorlagenName = formerProj.VorlagenName
                End If


                If .Erloes > 0 Then
                    ' Workaround: 
                    Dim tmpValue As Integer = hproj.dauerInDays
                    ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                    ' Call awinCreateBudgetWerte(hproj)

                End If

                ' macht er jetzt immer, wenn das cproj keine Ressourcenbedarfe enthält
                If hproj.getGesamtKostenBedarf.Sum = 0 And formerProj.getGesamtKostenBedarf.Sum > 0 Then
                    ' dann wurde in VISBO eine Ressourcen- und Kostenplanung gemacht , die jetzt übernommen werden muss
                    Try
                        Dim tmpProj As clsProjekt = hproj.updateProjectWithRessourcesFrom(formerProj)
                        If Not IsNothing(tmpProj) Then
                            hproj = tmpProj
                        End If
                    Catch ex As Exception
                        Call MsgBox("resources from former version could Not be copied ... ")
                    End Try

                End If


                If fileFrom3rdParty Then

                    '.farbe = cproj.farbe
                    .Schrift = formerProj.Schrift
                    .Schriftfarbe = formerProj.Schriftfarbe



                    ' jetzt müssen Verantwortlicher für Projekt, actualDataUntil, Budget, Risiko, Beschreibung, Ampel und Ampel-Text übernommen werden 
                    If hproj.leadPerson = "" And formerProj.leadPerson <> "" Then
                        hproj.leadPerson = formerProj.leadPerson
                    End If

                    If hproj.Erloes = 0 And formerProj.Erloes > 0 Then
                        hproj.Erloes = formerProj.Erloes
                    End If


                    If hproj.ampelStatus = 0 And hproj.ampelErlaeuterung = "" And formerProj.ampelStatus > 0 Then
                        hproj.ampelStatus = formerProj.ampelStatus
                        hproj.ampelErlaeuterung = formerProj.ampelErlaeuterung
                    End If

                    If hproj.description = "" And formerProj.description <> "" Then
                        hproj.description = formerProj.description
                    End If


                    hproj.Risiko = formerProj.Risiko
                    hproj.StrategicFit = formerProj.StrategicFit
                    hproj.projectType = formerProj.projectType

                End If

                ' jetzt muss noch geprüft werden, ob 


            End With

        Catch ex As Exception
            Throw New ArgumentException("Fehler bei Übernahme der Attribute des alten Projektes" & vbLf & ex.Message)
        End Try

    End Sub


    ''' <summary>
    ''' wenn das Projekt mit Namen pName und Varianten-Name vName und einem TimeStamp kleiner/gleich datum in der Datenbank existiert, 
    ''' wird das Projekt als Ergebnis zurückgegeben
    ''' Nothing sonst 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <param name="datum"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function awinReadProjectFromDatabase(ByVal pNr As String, ByVal pName As String, ByVal vName As String, ByVal datum As Date) As clsProjekt

        Dim err As New clsErrorCodeMsg

        Dim tmpResult As clsProjekt = Nothing
        Dim allNames As Collection = Nothing

        '
        ' prüfen, ob es in der Datenbank existiert ... wenn ja,  laden und anzeigen

        ' Stufe 1: gibt es die Projekt-Nummer bereits in der Datenbank? 
        If pNr <> "" Then
            Try
                ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                allNames = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(pNr, err)
                If allNames.Count > 1 Then
                    Dim errMsg As String = "Project-Number occurs more than once in DB" & pNr

                    Dim usedName As String = ""
                    For Each tmpName As String In allNames
                        If usedName = "" Then
                            usedName = tmpName
                        End If
                        errMsg = errMsg & vbLf & tmpName
                    Next
                    errMsg = errMsg & vbLf & vbLf & "used name " & usedName

                    Call MsgBox(errMsg)

                ElseIf allNames.Count = 1 Then
                    pName = allNames.Item(1)
                End If
            Catch ex As Exception

            End Try
        End If



        Try

            If Not noDB Then

                If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                    If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, datum, err) Then

                        ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                        tmpResult = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pName, vName, "", datum, err)

                    Else
                        ' nichts tun, tmpResult ist bereits Nothing 
                    End If
                Else
                    ' nichts tun, tmpResult ist bereits Nothing 
                End If
            End If


        Catch ex As Exception

        End Try

        awinReadProjectFromDatabase = tmpResult

    End Function

    ''' <summary>
    ''' liest den Wert eines Cusomized Flag. Das Ergebnis ist True oder False
    ''' </summary>
    ''' <param name="msTask"></param>
    ''' <param name="visboflag"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function readCustomflag(ByVal msTask As MSProject.Task, ByVal visboflag As MSProject.PjField) As Boolean

        Dim tskflag As Boolean = True
        Select Case visboflag
            Case MSProject.PjField.pjTaskFlag1
                tskflag = msTask.Flag1
            Case MSProject.PjField.pjTaskFlag2
                tskflag = msTask.Flag2
            Case MSProject.PjField.pjTaskFlag3
                tskflag = msTask.Flag3
            Case MSProject.PjField.pjTaskFlag4
                tskflag = msTask.Flag4
            Case MSProject.PjField.pjTaskFlag5
                tskflag = msTask.Flag5
            Case MSProject.PjField.pjTaskFlag6
                tskflag = msTask.Flag6
            Case MSProject.PjField.pjTaskFlag7
                tskflag = msTask.Flag7
            Case MSProject.PjField.pjTaskFlag8
                tskflag = msTask.Flag8
            Case MSProject.PjField.pjTaskFlag9
                tskflag = msTask.Flag9
            Case MSProject.PjField.pjTaskFlag10
                tskflag = msTask.Flag10
            Case MSProject.PjField.pjTaskFlag11
                tskflag = msTask.Flag11
            Case MSProject.PjField.pjTaskFlag12
                tskflag = msTask.Flag12
            Case MSProject.PjField.pjTaskFlag13
                tskflag = msTask.Flag13
            Case MSProject.PjField.pjTaskFlag14
                tskflag = msTask.Flag14
            Case MSProject.PjField.pjTaskFlag15
                tskflag = msTask.Flag15
            Case MSProject.PjField.pjTaskFlag16
                tskflag = msTask.Flag16
            Case MSProject.PjField.pjTaskFlag17
                tskflag = msTask.Flag17
            Case MSProject.PjField.pjTaskFlag18
                tskflag = msTask.Flag18
            Case MSProject.PjField.pjTaskFlag19
                tskflag = msTask.Flag19
            Case MSProject.PjField.pjTaskFlag20
                tskflag = msTask.Flag230

        End Select
        readCustomflag = tskflag
    End Function

    ''' <summary>
    ''' Prüft, ob eine Phase (elemID) aus dem Projekt hproj gelöscht werden kann, 
    ''' da weder sie selbst betrachtet werden soll, noch all ihre Kinder
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <param name="hproj"></param>
    ''' <param name="liste"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    '''
    Public Function isRemovable(ByVal elemID As String, ByVal hproj As clsProjekt, ByVal liste As SortedList(Of String, Boolean)) As Boolean

        Dim ind As Integer = 1
        Dim hrchynode As clsHierarchyNode = Nothing
        Dim result As Boolean

        result = True

        Try

            hrchynode = hproj.hierarchy.nodeItem(elemID)
            If hrchynode.childCount = 0 Then
                result = result And Not liste(elemID)
            End If
            If hrchynode.childCount > 0 And result Then

                While result And ind <= hrchynode.childCount

                    Dim nodeID As String = hrchynode.getChild(ind)
                    result = result And liste.ContainsKey(nodeID) And Not liste(nodeID)
                    result = result And isRemovable(nodeID, hproj, liste)
                    ind = ind + 1

                End While

            End If

        Catch ex As Exception
            Call MsgBox("Fehler bei der Prüfung, ob das Element elemID= " & elemID & " entfernt werden kann")
            Throw New ArgumentException("Fehler bei der Prüfung, ob das Element elemID entfernt werden kann")
        End Try

        isRemovable = result

    End Function
    ''' <summary>
    ''' berechnet in importAllianz3 aus einer sortierten Liste von Rollen und Array Namen den geldwerten Betrag  
    ''' </summary>
    ''' <param name="roleValues"></param>
    ''' <returns></returns>
    Public Function calcIstValueOf(ByVal roleValues As SortedList(Of String, Double())) As Double
        Dim tmpResult As Double = 0.0
        Dim hrole As clsRollenDefinition = Nothing

        For Each rvkvp As KeyValuePair(Of String, Double()) In roleValues
            Dim teamID As Integer = -1
            hrole = RoleDefinitions.getRoleDefByIDKennung(rvkvp.Key, teamID)
            If Not IsNothing(hrole) Then
                tmpResult = tmpResult + rvkvp.Value.Sum * hrole.tagessatzIntern
            End If
        Next

        calcIstValueOf = tmpResult
    End Function

    ''' <summary>
    ''' macht aus einem PName, dem evtl die Projektnummer hinten angehängt ist, nur den  PName 
    ''' </summary>
    ''' <param name="tmpPName"></param>
    ''' <param name="tmpPNr"></param>
    ''' <returns></returns>
    Friend Function getAllianzPNameFromPPN(ByVal tmpPName As String, ByVal tmpPNr As String) As String
        Dim tmpResult As String = ""

        Dim tmpStr() As String = tmpPName.Split(New Char() {CChar(" ")})
        If tmpStr.Length > 1 Then
            If tmpStr(tmpStr.Length - 1).Trim = tmpPNr Then
                For i As Integer = 0 To tmpStr.Length - 2
                    tmpResult = tmpResult & tmpStr(i)
                Next
            Else
                tmpResult = tmpPName
            End If
        Else
            tmpResult = tmpPName
        End If

        getAllianzPNameFromPPN = tmpResult
    End Function

    ''' <summary>
    ''' gibt den Allianz Rollen-Namen zurück, sofern 
    ''' </summary>
    ''' <param name="fullRName"></param>
    ''' <param name="isExtern"></param>
    ''' <returns></returns>
    Public Function getAllianzRoleNameFromValue(ByVal fullRName As String, ByVal isExtern As Boolean) As String
        Dim tmpResult As String = ""
        Dim found As Boolean = False
        Dim roleName As String = ""

        If isExtern Then
            If fullRName.StartsWith("*") Then
                fullRName = fullRName.Substring(1)
            End If
            'Dim tmpStr = fullRName.Split(New Char() {CChar("-"), CChar("("), CChar(")")})
            'If tmpStr.Length >= 3 Then
            '    roleName = tmpStr(1)
            'End If
            roleName = fullRName
        Else
            ' Prüfung 1: besteht es nur aus einem Wort ? 
            roleName = fullRName
        End If

        ' jetzt prüfen, ob es die Rolle gibt ... 
        If RoleDefinitions.containsName(roleName) Then
            tmpResult = roleName
        Else
            Dim a As Boolean = True
        End If


        getAllianzRoleNameFromValue = tmpResult
    End Function
    ''' <summary>
    ''' macht aus dem pName einen gültigen Projekt-Namen und
    ''' bestimmt ob es sich um ein bekanntes Projekt handelt: entweder bereits in projektliste geladen oder aber in der Datenbank vorhanden 
    ''' wenn nur in DB, wird es geladen und in projektliste abgelegt
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="projektKDNr">ist die optional anzugebende Projekt-Kunden-Nummer</param>
    ''' <param name="projektListe">ist AlleProjekte oder eine andere Instanz vom Typ clsProjekteAlle, in der das Projekt aufgenommen wird, wenn es existiert </param>
    ''' <param name="lookupTable">enthält die Mapping Informationen zu Projekten, also welcher Name ist welches Projekt</param>
    ''' <param name="createUnknownProject">gibt an, ob ein unbekanntes Projekt erzeugt und in projektliste aufgenommen werden soll </param>
    ''' <returns></returns>
    Public Function isKnownProject(ByRef pname As String, ByVal projektKDNr As String, ByRef projektListe As clsProjekteAlle,
                                    Optional ByVal lookupTable As SortedList(Of String, String) = Nothing,
                                    Optional ByVal createUnknownProject As Boolean = False) As Boolean

        Dim err As New clsErrorCodeMsg

        Dim fctResult As Boolean = False
        Dim oldProj As clsProjekt = Nothing
        Dim anzFehler As Integer = 0 ' wird als Platzhalter Variable für logFile schreiben benötigt ...
        Dim logArray() As String


        If IsNothing(pname) Then
            fctResult = False
        ElseIf pname.Trim.Length < 2 Then
            fctResult = False
        Else
            If Not isValidPVName(pname) Then
                pname = makeValidProjectName(pname)
            End If

            Dim key As String = calcProjektKey(pname, "")

            ' jetzt muss geprüft werden, ob das Projekt bereits in alleProjekte oder in der Datenbank existiert 
            ' wenn nein, dann wird per KundenProjekt-Nummer gesucht , ansonsten abgebrochen  ... 
            If projektListe.Containskey(key) Then
                oldProj = projektListe.getProject(key)
                fctResult = True

            Else
                ' ist es in der Datenbank? wenn ja, in AlleProjekte holen ... 

                Dim storedAtOrBefore = Date.Now

                If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pname, "", storedAtOrBefore, err) Then
                    oldProj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pname, "", "", storedAtOrBefore, err)
                End If

                If Not IsNothing(oldProj) Then

                    If oldProj.kundenNummer = "" Then
                        ' wenn vorher die Kunden-Nummer noch nicht bekannt war ...
                        oldProj.kundenNummer = projektKDNr
                    End If

                    projektListe.Add(oldProj, updateCurrentConstellation:=False, checkOnConflicts:=False)
                    fctResult = True

                ElseIf projektKDNr <> "" Then
                    ' jetzt wird noch über die Kunden-Projekt-Nummer gesucht ... 
                    ' auuserdem wird dann ggf noch das Projekt angelegt ... 

                    Dim pNames As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(projektKDNr, err)

                    ' greift die LookupTable ? 
                    Dim considerLookUpTable As Boolean = False
                    If Not IsNothing(lookupTable) Then

                        If Not IsNothing(projektKDNr) Then
                            If projektKDNr.Trim.Length > 0 Then
                                considerLookUpTable = lookupTable.ContainsKey(projektKDNr)
                            End If
                        End If

                    End If

                    If pNames.Count = 1 Then
                        ' in der Datenbank gibt es aktuell genau ein Projekt, dem diese Projekt-Kundennummer zugeordnet ist 
                        Dim visboDBname As String = pNames.Item(1)

                        If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(visboDBname, "", storedAtOrBefore, err) Then
                            oldProj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(visboDBname, "", "", storedAtOrBefore, err)
                        End If

                        If Not IsNothing(oldProj) Then
                            projektListe.Add(oldProj, updateCurrentConstellation:=False, checkOnConflicts:=False)
                            fctResult = True

                            ReDim logArray(3)
                            logArray(0) = "erfolgreiches Mapping anhand Plan-View P-Nr -> Visbo DB P-Nr"
                            logArray(1) = projektKDNr
                            logArray(2) = pname
                            logArray(3) = visboDBname

                            Call logger(ptErrLevel.logInfo, "isKnownProject", logArray)

                            ' damit an der aufrufenden Stelle der richtige pName steht ...
                            pname = visboDBname
                        Else
                            fctResult = False
                        End If


                    ElseIf pNames.Count > 1 Then
                        ' in der Datenbank gibt es mehrere Projekte, denen diese Projekt-Kundennummer zugeordnet ist : Fehler ! 
                        ' Eintrag in Log-File 
                        ReDim logArray(1 + pNames.Count)
                        logArray(0) = "kein Import; Mehrfach Zuordnung P-Nr -> Projekt  "
                        logArray(1) = projektKDNr


                        Dim ix As Integer = 1
                        For Each tmpName In pNames
                            logArray(1 + ix) = tmpName
                            ix = ix + 1
                        Next

                        Call logger(ptErrLevel.logError, "isKnownProject", logArray)

                        fctResult = False


                    ElseIf considerLookUpTable Then
                        ' checken ob es in der LoopUpTable eine Projekt-Zuordnung gibt ... 

                        Dim visboDBname As String = lookupTable.Item(projektKDNr)

                        If visboDBname.Trim <> "" Then
                            If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(visboDBname, "", storedAtOrBefore, err) Then

                                Dim rupiProj As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(visboDBname, "", "", storedAtOrBefore, err)
                                If Not IsNothing(rupiProj) Then
                                    projektListe.Add(rupiProj, updateCurrentConstellation:=False, checkOnConflicts:=False)
                                    fctResult = True

                                    ReDim logArray(3)
                                    logArray(0) = "lookupTable erfolgreiches Mapping"
                                    logArray(1) = projektKDNr
                                    logArray(2) = pname
                                    logArray(3) = visboDBname


                                    Call logger(ptErrLevel.logInfo, "isKnownProject", logArray)

                                    ' damit an der aufrufenden Stelle der richtige pName steht ...
                                    pname = visboDBname
                                Else
                                    ' kann eigentlich nicht passieren ...
                                    fctResult = False
                                End If

                            Else
                                ReDim logArray(3)
                                logArray(0) = "lookupTable: Name existiert nicht in initialer Projektliste"
                                logArray(1) = projektKDNr
                                logArray(2) = pname
                                logArray(3) = visboDBname

                                Call logger(ptErrLevel.logError, "isKnownProject", logArray)

                                fctResult = False
                            End If
                        Else

                            ReDim logArray(3)
                            logArray(0) = "lookupTable: kein Name definiert für Projekt-Nummer"
                            logArray(1) = projektKDNr
                            logArray(2) = pname
                            logArray(3) = visboDBname

                            Call logger(ptErrLevel.logError, "isKnownProject", logArray)

                            fctResult = False
                        End If




                    ElseIf createUnknownProject Then
                        ' dann soll das Projekt angelegt werden ...
                        Dim startDate As Date = CDate("01.01.2019")
                        Dim endDate As Date = CDate("31.12.2019")

                        If awinSettings.databaseName.EndsWith("20") Then
                            startDate = CDate("01.01.2020")
                            endDate = CDate("31.12.2020")
                        End If
                        ' es wird kein existierendes Projekt als Vorlage verwendet 
                        Dim myProject As clsProjekt = Nothing
                        oldProj = erstelleProjektAusVorlage(myProject, pname, "Projekt-Platzhalter", startDate, endDate, 0, 2, 5, 5, Nothing, "aus Planview Ist-Daten erzeugtes Projekt", "", kdNr:=projektKDNr)

                        If Not IsNothing(oldProj) Then
                            oldProj.kundenNummer = projektKDNr
                            projektListe.Add(oldProj, updateCurrentConstellation:=False, checkOnConflicts:=False)
                            fctResult = True

                            ReDim logArray(4)
                            logArray(0) = "neu angelegtes Projekt:  "
                            logArray(1) = projektKDNr
                            logArray(2) = pname
                            logArray(3) = ""
                            logArray(4) = pname


                            Call logger(ptErrLevel.logInfo, "isKnownProject", logArray)


                        Else
                            fctResult = False
                            ReDim logArray(2)
                            logArray(0) = "Fehler beim Neu-Anlegen eines Projektes aus Istdaten"
                            logArray(1) = projektKDNr
                            logArray(2) = pname


                            Call logger(ptErrLevel.logError, "isKnownProject", logArray)
                        End If

                    Else

                        fctResult = False

                    End If

                Else
                    fctResult = False
                End If
            End If
        End If

        isKnownProject = fctResult

    End Function

    ''' <summary>
    ''' testet, ob das Summary Projekt einer constellation mit den einzelnen Werten der Projekte übereinstimmt ...
    ''' </summary>
    ''' <param name="current1program"></param>
    ''' <returns></returns>
    Public Function testUProjandSingleProjs(ByVal current1program As clsConstellation,
                                             Optional ByVal considerImportProjekte As Boolean = True) As Boolean

        Dim tmpResult As Boolean = True
        Dim constellationName As String = current1program.constellationName
        Dim projektliste As clsProjekteAlle

        If considerImportProjekte Then
            projektliste = ImportProjekte
        Else
            projektliste = AlleProjekte
        End If

        Dim uProj As clsProjekt = projektliste.getProject(calcProjektKey(constellationName, ""))
        Dim testProjekte As New clsProjekte

        If Not IsNothing(uProj) Then
            Dim uRoles As Collection = uProj.getRoleNameIDs
            Dim GPRoles As Collection = Nothing

            Dim listOfProjectNames As SortedList(Of String, String) = current1program.getProjectNames(considerShowAttribute:=True, showAttribute:=True, fullNameKeys:=True)

            Dim dimension As Integer = listOfProjectNames.Count - 1

            For Each fullName As KeyValuePair(Of String, String) In listOfProjectNames
                Dim hproj As clsProjekt = projektliste.getProject(fullName.Key)
                If IsNothing(GPRoles) Then
                    GPRoles = hproj.getRoleNameIDs
                Else
                    Dim tmpRoleNameIDs As Collection = hproj.getRoleNameIDs
                    For Each tmpRoleNameID As String In tmpRoleNameIDs
                        If GPRoles.Contains(tmpRoleNameID) Then
                            ' alles ok, schon drin 
                        Else
                            GPRoles.Add(tmpRoleNameID, tmpRoleNameID)
                        End If
                    Next

                End If

                If testProjekte.contains(hproj.name) Then
                    ' darf eigentlich n icht sein 
                    Call MsgBox("Fehler ? " & hproj.name)
                Else
                    testProjekte.Add(hproj)
                End If
            Next

            ' 1. Test sind die Collections identisch ? 
            If collectionsAreDifferent(uRoles, GPRoles) Then
                tmpResult = False
            Else
                showRangeLeft = getColumnOfDate(CDate("1.1.2020"))
                showRangeRight = getColumnOfDate(CDate("31.12.2020"))

                For Each tmpRoleNameID As String In uRoles

                    Dim myCollection As New Collection
                    myCollection.Add(tmpRoleNameID)
                    Dim uValues() As Double = uProj.getBedarfeInMonths(mycollection:=myCollection, type:=DiagrammTypen(1))

                    Dim GPvalues() As Double = testProjekte.getRoleValuesInMonth(tmpRoleNameID)

                    If arraysAreDifferent(GPvalues, uValues) Then
                        If System.Math.Abs(GPvalues.Sum - uValues.Sum) > 0.00001 Then
                            tmpResult = False
                        End If

                    End If
                Next

            End If

        End If


        testUProjandSingleProjs = tmpResult

    End Function

    ''' <summary>
    ''' bestimmt, ob es sich um einen gültigen Kapazitäts- bzw Kosten-Input String handelt
    ''' alle Rollen- bzw Kostenart Namen bekannt, alle Werte >= 0 
    ''' </summary>
    ''' <param name="inputStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isValidRoleCostInput(ByVal inputStr As String, ByVal checkRoles As Boolean) As Boolean
        Dim resultValue As Boolean = True
        Dim anzDefinitions As Integer

        If checkRoles Then
            anzDefinitions = RoleDefinitions.Count
        Else
            anzDefinitions = CostDefinitions.Count
        End If


        If Not IsNothing(inputStr) Then
            If inputStr.Trim.Length > 0 Then

                Dim completeStr() As String = inputStr.Split(New Char() {CType("#", Char)}, 100)


                ' jetzt die ganzen Rollen bzw. Kosten abarbeiten 
                Dim i As Integer = 1
                While i <= completeStr.Length And resultValue = True

                    Dim roleCostStr() As String = completeStr(i - 1).Split(New Char() {CType("", Char)}, 2)

                    If roleCostStr.Length = 2 Then

                        Try
                            Dim roleCostName As String = roleCostStr(0).Trim
                            Dim roleCostSum As Double = CDbl(roleCostStr(1).Trim)
                            If checkRoles Then
                                If RoleDefinitions.containsName(roleCostName) And roleCostSum >= 0 Then
                                    ' ok, nichts tun 
                                Else
                                    resultValue = False
                                End If

                            Else
                                If CostDefinitions.containsName(roleCostName) And roleCostSum >= 0 Then
                                    ' ok, nichts tun 
                                Else
                                    resultValue = False
                                End If

                            End If

                        Catch ex As Exception
                            resultValue = False
                        End Try

                    ElseIf roleCostStr.Length = 1 And anzDefinitions >= 1 Then
                        ' es muss sich um eine Zahl größer 0 handeln, Rolle 1 wird angenommen 

                        Try
                            If IsNumeric(roleCostStr(0).Trim) Then
                                If CDbl(roleCostStr(0).Trim) >= 0 Then
                                    ' ok, nichts tun
                                Else
                                    resultValue = False
                                End If

                            ElseIf Not checkRoles And roleCostStr(0) = "filltobudget" Then
                                ' ok , nichts tun 

                            Else
                                resultValue = False
                            End If
                        Catch ex As Exception
                            resultValue = False
                        End Try

                    Else
                        resultValue = False
                    End If

                    i = i + 1

                End While



            Else
                ' leerer String, ok 
                resultValue = True
            End If
        Else
            ' Nothing, ok 
            resultValue = True
        End If

        isValidRoleCostInput = resultValue

    End Function

    ''' <summary>
    ''' prüft, ob es sich um einen zugelassenen Projekt-Namen handelt ....
    ''' nicht zugelassen: #, (, ), Zeilenvorschub 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isValidPVName(ByVal pName As String) As Boolean
        Dim ergebnis As Boolean = False

        ' wenn beides enthalten ist ...
        If (pName.Contains("<") And pName.Contains(">")) Or
            pName.Contains("#") Or
            pName.Contains("(") Or
            pName.Contains(")") Or
            pName.Contains("[") Or
            pName.Contains("]") Or
            pName.Contains(vbCr) Or
            pName.Contains(vbLf) Then
            ergebnis = False
        Else
            ergebnis = True
        End If

        isValidPVName = ergebnis

    End Function

    ''' <summary>
    ''' macht aus einem evtl ungültigen Namen einen gültigen Projekt-NAmen 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function makeValidProjectName(ByVal pName As String) As String

        If (pName.Contains("<") And pName.Contains(">")) Then
            pName = pName.Replace("<", " ")
            pName = pName.Replace(">", " ")
        End If

        If pName.Contains("#") Then
            pName = pName.Replace("#", " ")
        End If
        If pName.Contains("(") Then
            pName = pName.Replace("(", "-")
        End If
        If pName.Contains(")") Then
            pName = pName.Replace(")", "-")
        End If
        If pName.Contains("[") Then
            pName = pName.Replace("[", "-")
        End If
        If pName.Contains("]") Then
            pName = pName.Replace("]", "-")
        End If
        If pName.Contains(vbCr) Then
            pName = pName.Replace(vbCr, " ")
        End If
        If pName.Contains(vbLf) Then
            pName = pName.Replace(vbLf, " ")
        End If

        makeValidProjectName = pName

    End Function

    ''' <summary>
    ''' diese Funktion verarbeitet die Import Projekte 
    ''' wenn sie schon in der Datenbank bzw Session existieren und unterschiedlich sind: es wird eine Variante angelegt, die so heisst wie das Scenario 
    ''' wenn sie bereits existieren und identisch sind: in AlleProjekte holen, wenn nicht schon geschehen
    ''' wenn sie noch nicht existieren: in AlleProjekte anlegen
    ''' in jedem Fall: eine Constellation mit dem Namen cName anlegen
    ''' </summary>
    ''' <param name="cName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function verarbeiteImportProjekte(ByVal cName As String,
                                             Optional ByVal noComparison As Boolean = False,
                                             Optional ByVal noScenarioCreation As Boolean = False,
                                             Optional ByVal considerSummaryProjects As Boolean = False) As clsConstellation

        Dim err As New clsErrorCodeMsg

        ' in der Reihenfolge des Auftretens aufnehmen , Name wie übergeben 
        Dim newC As New clsConstellation(ptSortCriteria.customTF, cName)
        currentSessionConstellation.sortCriteria = ptSortCriteria.customTF

        Dim vglProj As clsProjekt
        Dim lfdZeilenNr As Integer = 2

        Dim outPutCollection As New Collection
        Dim outputLine As String = ""


        Dim logmsg() As String = Nothing

        Dim takeIntoAccount As Boolean = True

        Dim importDate As Date = Date.Now


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste

            Try
                Dim impProjekt As clsProjekt = kvp.Value
                Dim variantName As String = ""

                If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And impProjekt.variantName = "" Then
                    variantName = ptVariantFixNames.pfv.ToString
                Else
                    variantName = impProjekt.variantName
                End If

                If considerSummaryProjects Then
                    takeIntoAccount = (impProjekt.projectType = ptPRPFType.portfolio)
                Else
                    takeIntoAccount = Not (impProjekt.projectType = ptPRPFType.portfolio)
                End If

                If takeIntoAccount Then

                    ' jetzt das Import Datum setzen und dann in PortfolioProjektSummaries verschieben ...
                    impProjekt.timeStamp = importDate

                    Dim importKey As String = calcProjektKey(impProjekt)

                    vglProj = Nothing

                    If noComparison Then
                        ' nicht vergleichen, einfach in AlleProjekte rein machen 
                        If AlleProjekte.Containskey(importKey) Then
                            AlleProjekte.Remove(importKey)
                        End If
                        AlleProjekte.Add(impProjekt)
                    Else
                        ' jetzt muss ggf verglichen werden 
                        If AlleProjekte.Containskey(importKey) Then

                            vglProj = AlleProjekte.getProject(importKey)

                        Else
                            ' nicht in der Session, aber ist es in der Datenbank ?  

                            If Not noDB Then

                                '
                                ' prüfen, ob es in der Datenbank existiert ... wenn ja,  laden und anzeigen
                                ' wenn es sich um einen Portfolio Manager handelt: der Vergleich muss mit der letzten Vorgabe stattfinden, weill nur das kann der Portfolio Manager ja auch speichern ... 

                                If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then


                                    If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(impProjekt.name, variantName, Date.Now, err) Then

                                        ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                                        vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(impProjekt.name, variantName, "", Date.Now, err)

                                        If IsNothing(vglProj) Then
                                            ' kann eigentlich nicht sein 
                                            ReDim logmsg(3)
                                            logmsg(0) = "Projekt existiert scheinbar - lässt sich aber nicht aus DB laden ... "
                                            logmsg(1) = impProjekt.name
                                            logmsg(2) = "bitte kontaktieren Sie ihren System-Administrator! "
                                            logmsg(3) = "kein Import ! "
                                            Call logger(ptErrLevel.logError, "verarbeiteImportProjekte", logmsg)
                                        Else
                                            ' jetzt in AlleProjekte eintragen ... 
                                            AlleProjekte.Add(impProjekt)

                                        End If


                                    ElseIf impProjekt.kundenNummer <> "" Then
                                        ' versuche es darüber zu finden 
                                        Dim nameCollection As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(impProjekt.kundenNummer, err)

                                        If nameCollection.Count = 0 Then
                                            ' es existiert nicht , also eintragen ... 
                                            AlleProjekte.Add(impProjekt)

                                        ElseIf nameCollection.Count = 1 Then
                                            ' es existiert angeblich genau einmal 
                                            vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(nameCollection.Item(1), variantName, "", Date.Now, err)

                                            If Not IsNothing(vglProj) Then

                                                Dim newName As String = impProjekt.name

                                                impProjekt.name = vglProj.name
                                                AlleProjekte.Add(impProjekt)

                                                ReDim logmsg(3)
                                                logmsg(0) = "Projekt existiert unter anderem Namen in VISBO Datenbank - bitte in Rupi-Liste  umbenennen!"
                                                logmsg(1) = impProjekt.kundenNummer
                                                logmsg(2) = "VISBO DB - Name " & vglProj.name
                                                logmsg(3) = "neuer Name " & newName
                                                Call logger(ptErrLevel.logWarning, "verarbeiteImportProjekte", logmsg)
                                            Else
                                                ' kann eigentlich nicht sein 
                                                ReDim logmsg(3)
                                                logmsg(0) = "Projekt existiert scheinbar - lässt sich aber nicht aus DB laden ... "
                                                logmsg(1) = nameCollection.Item(1)
                                                logmsg(2) = "bitte kontaktieren Sie ihren System-Administrator! "
                                                logmsg(3) = "kein Import ! "
                                                Call logger(ptErrLevel.logError, "verarbeiteImportProjekte", logmsg)
                                            End If

                                        ElseIf nameCollection.Count > 1 Then

                                            Dim anz As Integer = nameCollection.Count
                                            ReDim logmsg(anz + 2)
                                            logmsg(0) = "Projekt-Nummer existiert mehrfach; bitte bereinigen - Projekt wurde nicht importiert"
                                            logmsg(1) = impProjekt.kundenNummer
                                            logmsg(2) = impProjekt.name

                                            For ia As Integer = 1 To anz
                                                logmsg(ia + 2) = nameCollection.Item(ia)
                                            Next

                                            Call logger(ptErrLevel.logError, "verarbeiteImportProjekte", logmsg)

                                        End If

                                    Else
                                        ' kann nicht gefunden werden, Kunden-Nummer ist "" 
                                        ' jetzt in AlleProjekte eintragen ... 
                                        AlleProjekte.Add(impProjekt)

                                    End If
                                Else
                                    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & impProjekt.name & "'konnte nicht geladen werden")
                                End If


                            Else
                                ' nicht in der Session, nicht in der Datenbank : es ist bereits in AlleProjekte eingetragen ... 
                                ' jetzt in AlleProjekte eintragen ... 
                                AlleProjekte.Add(impProjekt)
                            End If


                        End If

                        ' wenn jetzt vglProj <> Nothing, dann vergleichen und ggf markieren, wenn unterschiedlich  anlegen ...
                        If Not IsNothing(vglProj) Then

                            ' wenn es sich jetzt um den Portfolio Manager handelt , dann muss kurz das vglProj.variantName auf impProjekt.variantName gesetzt werden 
                            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                vglProj.variantName = impProjekt.variantName
                            End If

                            If Not impProjekt.isIdenticalTo(vglProj) Then
                                ' es gibt Unterschiede, es wird keine Variante mehr angelegt, sondern es wird als verändert markiert
                                impProjekt.marker = True

                                If impProjekt.kundenNummer <> vglProj.kundenNummer Then

                                    ReDim logmsg(3)
                                    logmsg(0) = "Projekt-Nummer hat sich geändert: "
                                    logmsg(1) = impProjekt.name
                                    logmsg(2) = " von " & vglProj.kundenNummer
                                    logmsg(3) = " zu " & impProjekt.kundenNummer

                                    outputLine = impProjekt.name & " :" & " von " & vglProj.kundenNummer & " zu " & impProjekt.kundenNummer
                                    outPutCollection.Add(outputLine)

                                End If

                                'impProjekt.variantName = cName
                                importKey = calcProjektKey(impProjekt)

                                ' wenn die Variante bereits in der Session existiert ..
                                ' wird die bisherige gelöscht , die neue über ImportProjekte neu aufgenommen  
                                If AlleProjekte.Containskey(importKey) Then
                                    AlleProjekte.Remove(importKey)
                                End If

                                ' jetzt das Importierte PRojekt in AlleProjekte aufnehmen 
                                AlleProjekte.Add(impProjekt)
                            End If

                            ' wenn es sich jetzt um den Portfolio Manager handelt , dann muss kurz das vglProj.variantName auf impProjekt.variantName gesetzt werden 
                            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                vglProj.variantName = variantName
                            End If

                        End If

                    End If

                    ' wenn jetzt automatisch setzen des ActualdataDate gemacht werden soll
                    'If awinSettings.autoSetActualDataDate Then
                    '    impProjekt.actualDataUntil = importDate.AddMonths(-1)
                    'End If

                    ' Aufnehmen in Constellation
                    Dim newCItem As New clsConstellationItem
                    newCItem.projectName = impProjekt.name
                    newCItem.variantName = impProjekt.variantName

                    If newC.containsProject(impProjekt.name) Then
                        newCItem.show = False
                    Else
                        newCItem.show = True
                    End If

                    newCItem.start = impProjekt.startDate
                    newCItem.zeile = lfdZeilenNr
                    newCItem.projectTyp = CType(impProjekt.projectType, ptPRPFType).ToString
                    'newCItem.zeile = lfdZeilenNr
                    newC.add(newCItem, sKey:=lfdZeilenNr)

                    lfdZeilenNr = lfdZeilenNr + 1
                Else
                    ' nichts tun ...
                End If

            Catch ex As Exception
                Dim a As Integer = 0
            End Try

        Next

        If outPutCollection.Count > 0 Then
            Call showOutPut(outPutCollection, "es sind Änderungen in der Projekt-Nummer aufgetreten", "siehe Logfile")
        End If

        verarbeiteImportProjekte = newC

    End Function

    ''' <summary>
    ''' ergänzt das übergebene Projekt um die im Ruleset angegebenen Phasen und Meilensteine
    ''' Wenn die Phase mit Namen ruleset.name schon existiert, werden die Elemente hinzugefügt, sofern sie nicht mit demselben Namen in dieser Phase bereits auftreten
    ''' Andernfalls wird bestimmt, wie lange die Phase sein muss
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="addElementSet"></param>
    ''' <remarks></remarks>
    Public Sub awinApplyAddOnRules(ByRef hproj As clsProjekt, ByVal addElementSet As clsAddElements)

        Dim phaseName As String = ""
        Dim topPhaseName As String = ""
        Dim breadCrumb As String = ""
        Dim milestoneName As String = ""
        Dim elemID As String

        Dim topPhase As clsPhase
        Dim cMilestone As clsMeilenstein


        ' erst bestimmen, ob die Phase schon existiert 
        topPhaseName = addElementSet.name
        topPhase = hproj.getPhase(topPhaseName)

        Dim minDate As Date = Date.Now.AddYears(100)
        Dim maxDate As Date = Date.Now.AddYears(-100)

        Dim currentDate As Date
        Dim currentElem As clsAddElementRules
        Dim currentMS As clsMeilenstein
        Dim currentPH As clsPhase
        Dim index As Integer = 1

        ' hier muss bestimmt werden,wie groß die aufnehmende Phase werden soll 
        ' es werden Mindate und MAxdate bestimmt 
        '
        Do While index <= addElementSet.count
            currentElem = addElementSet.getRule(index)

            Dim anzRules As Integer = currentElem.count
            Dim currentRule As clsAddElementRuleItem
            Dim found As Boolean = False

            Dim i As Integer = 1

            Do While i <= currentElem.count And Not found
                currentRule = currentElem.getItem(i)

                With currentRule
                    If .referenceIsPhase Then
                        ' existiert die Phase überhaupt? wenn nicht , weiter zu nächster Regel
                        Dim type As Integer = -1
                        Dim pvName As String = ""
                        Call splitHryFullnameTo2(.referenceName, phaseName, breadCrumb, type, pvName)

                        If type = -1 Or
                            (type = PTItemType.projekt And pvName = calcProjektKey(hproj)) Or
                            (type = PTItemType.vorlage) Then

                            currentPH = hproj.getPhase(name:=phaseName, breadcrumb:=breadCrumb, lfdNr:=1)

                            If Not IsNothing(currentPH) Then
                                found = True
                                If .referenceDateIsStart Then
                                    currentDate = currentPH.getStartDate.AddDays(currentRule.offset)
                                Else
                                    currentDate = currentPH.getEndDate.AddDays(currentRule.offset)
                                End If

                                If DateDiff(DateInterval.Day, minDate, currentDate) < 0 Then
                                    minDate = currentDate
                                End If

                                If currentElem.elemToCreateIsPhase Then
                                    currentDate = currentDate.AddDays(currentElem.duration)
                                End If
                                If DateDiff(DateInterval.Day, maxDate, currentDate) > 0 Then
                                    maxDate = currentDate
                                End If

                            Else
                                i = i + 1
                            End If

                        Else
                            i = i + 1
                        End If


                    Else
                        Dim type As Integer = -1
                        Dim pvName As String = ""
                        Call splitHryFullnameTo2(.referenceName, milestoneName, breadCrumb, type, pvName)

                        If type = -1 Or
                            (type = PTItemType.projekt And pvName = calcProjektKey(hproj)) Or
                            (type = PTItemType.vorlage) Then

                            currentMS = hproj.getMilestone(milestoneName, breadCrumb, 1)

                            If Not IsNothing(currentMS) Then
                                found = True
                                currentDate = currentMS.getDate.AddDays(currentRule.offset)

                                If DateDiff(DateInterval.Day, minDate, currentDate) < 0 Then
                                    minDate = currentDate
                                End If

                                If currentElem.elemToCreateIsPhase Then
                                    currentDate = currentDate.AddDays(currentElem.duration)
                                End If

                                If DateDiff(DateInterval.Day, maxDate, currentDate) > 0 Then
                                    maxDate = currentDate
                                End If
                            Else
                                i = i + 1
                            End If

                        Else
                            i = i + 1
                        End If

                    End If
                End With

            Loop

            index = index + 1

        Loop

        ' jetzt wird die oberste Phase entsprechend aufgenommen 
        '
        Dim startOffset As Integer = DateDiff(DateInterval.Day, hproj.startDate, minDate)
        If startOffset < 0 Then
            minDate = hproj.startDate
            startOffset = 0
        End If

        Dim duration As Integer = DateDiff(DateInterval.Day, minDate, maxDate) + 1

        If IsNothing(topPhase) Then
            ' die Phase existiert noch nicht
            elemID = hproj.hierarchy.findUniqueElemKey(topPhaseName, False)

            topPhase = New clsPhase(parent:=hproj)

            topPhase.nameID = elemID
            topPhase.changeStartandDauer(startOffset, duration)

            ' der Aufbau der Hierarchie erfolgt in addphase
            hproj.AddPhase(topPhase, origName:="",
                           parentID:=rootPhaseName)

        Else

            elemID = topPhase.nameID
            ' die Phase existiert bereits; aber ist sie auch ausreichend dimensioniert ? 
            ' ggf werden Start und Dauer angepasst 
            If startOffset <> topPhase.startOffsetinDays Or duration <> topPhase.dauerInDays Then
                topPhase.changeStartandDauer(startOffset, duration)
            End If

        End If


        ' jetzt müssen die Meilensteine / anderen Plan-Elemente eingetragen werden 
        '
        index = 1
        Do While index <= addElementSet.count

            Dim offs As Integer = 1
            Dim wasSuccessful As Boolean = False
            Dim newItemDate As Date
            Dim referenceMS As clsMeilenstein = Nothing
            Dim referencePH As clsPhase = Nothing
            Dim referenceDate As Date
            Dim currentRule As clsAddElementRuleItem

            currentElem = addElementSet.getRule(index)

            ' soll ein Meilenstein oder eine Phase erzeugt werden ? 
            If currentElem.elemToCreateIsPhase Then
                ' es soll eine Phase erzeugt werden 
            Else
                ' es soll ein Meilenstein erzeugt werden 
                Dim found As Boolean = False

                If IsNothing(topPhase.getMilestone(currentElem.name)) Then
                    ' nur wenn der nicht schon existiert, soll er auch erzeugt werden ... 

                    Do While offs <= currentElem.count And Not found
                        Dim ok As Boolean = False
                        currentRule = currentElem.getItem(offs)

                        If currentRule.referenceIsPhase Then
                            Dim type As Integer = -1
                            Dim pvName As String = ""
                            Call splitHryFullnameTo2(currentRule.referenceName, phaseName, breadCrumb, type, pvName)
                            referencePH = hproj.getPhase(name:=phaseName, breadcrumb:=breadCrumb)

                            If Not IsNothing(referencePH) Then
                                If currentRule.referenceDateIsStart Then
                                    referenceDate = referencePH.getStartDate
                                Else
                                    referenceDate = referencePH.getEndDate
                                End If

                                ok = True
                            Else
                                ok = False
                            End If

                        Else
                            Dim type As Integer = -1
                            Dim pvName As String = ""
                            Call splitHryFullnameTo2(currentRule.referenceName, milestoneName, breadCrumb, type, pvName)
                            referenceMS = hproj.getMilestone(msName:=milestoneName, breadcrumb:=breadCrumb)
                            If Not IsNothing(referenceMS) Then
                                referenceDate = referenceMS.getDate
                                ok = True
                            Else
                                ok = False
                            End If
                        End If

                        ' wenn es ein Referenz-Datum gibt ....
                        If ok Then
                            newItemDate = referenceDate.AddDays(currentRule.offset)
                            cMilestone = New clsMeilenstein(parent:=topPhase)
                            elemID = hproj.hierarchy.findUniqueElemKey(currentRule.newElemName, True)

                            Dim cbewertung As clsBewertung = New clsBewertung

                            With cbewertung
                                '.bewerterName = resultVerantwortlich
                                .colorIndex = 0
                                .datum = Date.Now
                                Dim abstandsText As String = ""
                                If currentRule.offset >= 0 Then
                                    abstandsText = "+" & currentRule.offset.ToString & " Tage"
                                Else
                                    abstandsText = currentRule.offset.ToString & " Tage"
                                End If
                                .description = " = " & currentRule.referenceName & abstandsText
                                ' Änderung tk 29.5.16 deliverables ist jetzt Bestandteil von clsMeilenstein
                                '.deliverables = currentElem.deliverables
                            End With


                            With cMilestone
                                .nameID = elemID
                                .setDate = newItemDate
                                If Not cbewertung Is Nothing Then
                                    .addBewertung(cbewertung)
                                End If
                            End With

                            Try
                                With topPhase
                                    .addMilestone(cMilestone)
                                End With
                            Catch ex As Exception

                            End Try
                            found = True
                        Else
                            offs = offs + 1
                        End If

                    Loop


                End If


            End If

            index = index + 1
        Loop

    End Sub

    ''' <summary>
    ''' lädt die jeweils letzten PName/Variante Projekte aus MongoDB in alleProjekte
    ''' lädt ausserdem alle definierten Konstellationen
    ''' zeigt dann die letzte (last) an 
    ''' </summary>
    ''' <remarks></remarks>
    Sub awinletzteKonstellationLaden(ByVal databaseName As String)

        Dim err As New clsErrorCodeMsg

        'Dim allProjectsList As SortedList(Of String, clsProjekt)
        Dim zeitraumVon As Date = StartofCalendar.AddMonths(showRangeLeft - 1)
        Dim zeitraumbis As Date = StartofCalendar.AddMonths(showRangeRight - 1)
        Dim pname As String = ""
        Dim variantName As String = ""

        Dim lastConstellation As New clsConstellation
        Dim hproj As clsProjekt

        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

            projectConstellations = CType(databaseAcc, DBAccLayer.Request).retrieveConstellationsFromDB(Date.Now, err)

            ' Showprojekte leer machen 
            Try
                'NoShowProjekte.Clear()
                ShowProjekte.Clear()
                lastConstellation = projectConstellations.getConstellation(calcLastSessionScenarioName)
            Catch ex As Exception
                'Call MsgBox("in awinProjekteInitialLaden Fehler ...")
            End Try

            ' jetzt Showprojekte aufbauen - und zwar so, dass Konstellation <Last> wiederhergestellt wird
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In lastConstellation.Liste

                Try
                    hproj = AlleProjekte.getProject(kvp.Key)
                    hproj.startDate = kvp.Value.start
                    hproj.tfZeile = kvp.Value.zeile
                    If kvp.Value.show Then
                        ' nur dann 
                        ShowProjekte.Add(hproj)
                    End If

                Catch ex As Exception
                    Call MsgBox("in ProjekteInitialLaden: " & ex.Message)
                End Try
            Next

        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
        End If

    End Sub

    ''' <summary>
    ''' lädt die Projekte im definierten Zeitraum (nach)
    ''' </summary>
    ''' <param name="databaseName"></param>
    ''' <remarks></remarks>
    Sub awinProjekteImZeitraumLaden(ByVal databaseName As String, ByVal filter As clsFilter)

        Dim err As New clsErrorCodeMsg

        Dim zeitraumVon As Date = StartofCalendar.AddMonths(showRangeLeft - 1)
        Dim zeitraumbis As Date = StartofCalendar.AddMonths(showRangeRight - 1)
        Dim storedHeute As Date = Now
        Dim storedGestern As Date = storedHeute.AddDays(-1)
        Dim pname As String = ""
        Dim variantName As String = ""

        Dim lastConstellation As New clsConstellation
        Dim projekteImZeitraum As New SortedList(Of String, clsProjekt)
        Dim projektHistorie As New clsProjektHistorie


        Dim ok As Boolean = True
        Dim filterIsActive As Boolean
        Dim toShowListe As New SortedList(Of Double, String)


        ' wurde ein definierter Filter mit übergeben ?
        If IsNothing(filter) Then
            filterIsActive = False
        Else
            If filter.isEmpty Then
                filterIsActive = False
            Else
                filterIsActive = True
            End If
        End If

        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

            projekteImZeitraum = CType(databaseAcc, DBAccLayer.Request).retrieveProjectsFromDB(pname, variantName, "", zeitraumVon, zeitraumbis, storedGestern, storedHeute, True, err)
        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen")
        End If

        If AlleProjekte.Count > 0 Then
            ' es sind bereits Projekte geladen 
            Dim atleastOne As Boolean = False

            For Each kvp As KeyValuePair(Of String, clsProjekt) In projekteImZeitraum

                If filterIsActive Then
                    ok = filter.doesNotBlock(kvp.Value)
                Else
                    ok = True
                End If

                If ok Then
                    ' Ist das Projekt bereits in AlleProjekte ? 
                    If AlleProjekte.Containskey(kvp.Key) Then
                        ' das Projekt soll nicht überschrieben werden ...
                        ' also nichts tun 
                    Else
                        ' Workaround: 
                        Dim tmpValue As Integer = kvp.Value.dauerInDays
                        ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                        ' Call awinCreateBudgetWerte(kvp.Value)
                        If Not AlleProjekte.hasAnyConflictsWith(calcProjektKey(kvp.Value), kvp.Value.projectType = ptPRPFType.portfolio) Then
                            AlleProjekte.Add(kvp.Value)
                            If ShowProjekte.contains(kvp.Value.name) Then
                                ' auch hier ist nichts zu tun, dann ist bereits eine andere Variante aktiv ...
                            Else
                                ShowProjekte.Add(kvp.Value)
                                atleastOne = True
                            End If
                        End If

                    End If

                End If

            Next

            ' jetzt ist Showprojekte und AlleProjekte aufgebaut ... 
            ' jetzt muss ClearPlanTafel kommen 
            If atleastOne Then
                Call awinClearPlanTafel()
                Call awinZeichnePlanTafel(True)
            End If

        Else

            ShowProjekte.Clear()
            ' ShowProjekte aufbauen

            For Each kvp As KeyValuePair(Of String, clsProjekt) In projekteImZeitraum

                If filterIsActive Then
                    ok = filter.doesNotBlock(kvp.Value)
                Else
                    ok = True
                End If

                If ok Then

                    Dim tmpValue As Integer = kvp.Value.dauerInDays
                    ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                    ' Call awinCreateBudgetWerte(kvp.Value)

                    If Not AlleProjekte.hasAnyConflictsWith(calcProjektKey(kvp.Value), kvp.Value.projectType = ptPRPFType.portfolio) Then

                        AlleProjekte.Add(kvp.Value)

                        Try
                            ' bei Vorhandensein von mehreren Varianten, immer die Standard Variante laden
                            If ShowProjekte.contains(kvp.Value.name) Then
                                If kvp.Value.variantName = "" Then
                                    ShowProjekte.Remove(kvp.Value.name)
                                    ShowProjekte.Add(kvp.Value)
                                End If
                            Else
                                ShowProjekte.Add(kvp.Value)
                            End If

                        Catch ex As Exception
                            Call MsgBox(ex.Message)
                        End Try
                    End If

                End If

            Next

            Call awinZeichnePlanTafel(True)

        End If


    End Sub

    ''' <summary>
    ''' visualisiert die in der Konstellation aufgeführten Projekte hinzu; 
    ''' wenn Sie bereits geladen sind, wird nachgesehen, ob die richtige Variante aktiviert ist 
    ''' ggf. wird diese Variante dann aktiviert 
    ''' </summary>
    ''' <param name="activeConstellation"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <param name="loadPFV">gibt an, ob die Vorgabe des Portfolio Managers geladen werden soll </param>
    ''' <remarks></remarks>
    Public Sub visualizeConstellation(ByVal activeConstellation As clsConstellation, ByVal storedAtOrBefore As Date,
                                Optional ByVal loadPFV As Boolean = False)


        Dim neErrorMessage As String = " (Datum kann nicht angepasst werden)"
        Dim outPutCollection = New Collection
        Dim outputLine As String = ""


        Dim boardwasEmpty As Boolean = (ShowProjekte.Count = 0)
        ' ab diesem Wert soll neu gezeichnet werden 
        Dim startOfFreeRows As Integer = projectboardShapes.getMaxZeile
        Dim zeilenOffset As Integer = 0

        ' prüfen, ob diese Constellation auch existiert ..
        If IsNothing(activeConstellation) Then
            Call MsgBox(" das Portfolio darf nicht NULL sein ... ")
            Exit Sub
        End If

        ' jetzt muss das Sort-Kriterium übernommen werden 
        If boardwasEmpty And activeConstellation.sortCriteria >= 0 Then
            currentSessionConstellation.sortCriteria = activeConstellation.sortCriteria
        End If



        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In activeConstellation.Liste

            Dim showIT As Boolean = kvp.Value.show
            Try
                Dim realKey As String = kvp.Key
                If loadPFV Then
                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = ptVariantFixNames.pfv.ToString
                    realKey = calcProjektKey(pName, vName)
                End If

                Call putItemOnVisualBoard(realKey, showIT, storedAtOrBefore, boardwasEmpty, activeConstellation, startOfFreeRows, outPutCollection)
            Catch ex As Exception
                Exit For
            End Try

        Next
        'End If


        If outPutCollection.Count > 0 Then

            If outPutCollection.Count > 0 Then
                Call showOutPut(outPutCollection,
                                "Messages when reading Portfolio ",
                                " ")
            End If

        End If


    End Sub

    ''' <summary>
    ''' platziert das (Summary-) Projekt auf dem Board
    ''' </summary>
    ''' <param name="key"></param>
    ''' <param name="showIT"></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <param name="boardwasEmpty"></param>
    ''' <param name="activeConstellation"></param>
    ''' <param name="startOfFreeRows"></param>
    ''' <param name="outPutCollection"></param>
    Private Sub putItemOnVisualBoard(ByVal key As String, ByVal showIT As Boolean, ByVal storedAtOrBefore As Date,
                                     ByVal boardwasEmpty As Boolean, ByVal activeConstellation As clsConstellation, ByVal startOfFreeRows As Integer,
                                     ByRef outPutCollection As Collection)

        Dim err As New clsErrorCodeMsg

        Dim outputLine As String = ""
        Dim pName As String = getPnameFromKey(key)
        Dim vName As String = getVariantnameFromKey(key)
        Dim hproj As clsProjekt
        Dim tryZeile As Integer
        Dim nvErrorMessage As String = " does not exist in DB at " & storedAtOrBefore.ToShortDateString


        If AlleProjekte.Containskey(key) Then
            ' Projekt ist bereits im Hauptspeicher geladen
            hproj = AlleProjekte.getProject(key)

            ' ist es aber auch der richtige TimeStamp ? 
            If DateDiff(DateInterval.Day, hproj.timeStamp, storedAtOrBefore) < 0 Then

                ' in einer Session dürfen keine TimeStamps aktuellen bzw. früheren TimeStamps gemischt werden ... 
                ' Meldung in der , und der Nutzer muss alles neu laden

                outputLine = "es gibt Projekte mit jüngerem TimeStamp in der Session ... "
                outPutCollection.Add(outputLine)
                outputLine = "die Aktion wurde abgebrochen ... "
                outPutCollection.Add(outputLine)
                outputLine = "bitte löschen Sie die Session und laden Sie dann die Szenarien mit dem gewünschten Versions-Datum"
                outPutCollection.Add(outputLine)

                Throw New ArgumentException("Fehler 366734")
            Else
                If showIT Then

                    If ShowProjekte.contains(hproj.name) Then
                        ' dann soll das Projekt da bleiben, wo es ist 
                        Dim shownProject As clsProjekt = ShowProjekte.getProject(hproj.name)
                        If shownProject.variantName = hproj.variantName Then
                            ' es wird bereits gezeigt, nichts machen ...
                        Else
                            tryZeile = shownProject.tfZeile
                            ' jetzt die Variante aktivieren 
                            Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, tryZeile)
                        End If

                    ElseIf boardwasEmpty Then
                        'tryZeile = kvp.Value.zeile
                        tryZeile = activeConstellation.getBoardZeile(hproj.name)
                        Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, tryZeile)
                    Else

                        'tryZeile = kvp.Value.zeile + startOfFreeRows - 1
                        'tryZeile = startOfFreeRows + zeilenOffset
                        tryZeile = startOfFreeRows + activeConstellation.getBoardZeile(hproj.name) - 2
                        Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, tryZeile)
                        'zeilenOffset = zeilenOffset + 1
                    End If


                Else
                    ' gar nichts machen
                End If

            End If




        Else
            If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, storedAtOrBefore, err) Then

                    ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                    hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pName, vName, "", storedAtOrBefore, err)

                    If Not IsNothing(hproj) Then


                        ' tk 4.2.20
                        ' hier muss geprüft werden, ob das Projekt Ressourcen-Zuordnungen für Mitarbeiter enthält, die noch gar nicht da sind bzw. zu dem Zeitpunkt schon weg sind.
                        ' es soll dann aber nur eine Warnung ausgegeben werden, sonst nichts weiter 
                        If DateDiff(DateInterval.Day, Date.Now, storedAtOrBefore) = 0 Then
                            ' nur bei aktuellen Projekten anmeckern ... 

                            Dim invalidNeedNames As Collection = hproj.hasRolesWithInvalidNeeds

                            If invalidNeedNames.Count > 0 Then

                                For Each iVName As String In invalidNeedNames
                                    Dim msgTxt As String = "Projekt " & hproj.getShapeText & " enthält ungültige Ressourcen-Zuordnungen"
                                    msgTxt = msgTxt & vbLf & "Person ist noch nicht oder nicht mehr im Unternehmen: " & iVName
                                    outPutCollection.Add(msgTxt)
                                Next

                            End If

                        End If

                        ' Projekt muss nun in die Liste der geladenen Projekte eingetragen werden
                        Dim newPosition As Integer = -1
                        If currentSessionConstellation.sortCriteria = ptSortCriteria.customTF Then
                            If boardwasEmpty Then
                                ' den gleichen key verwenden wie in der activeConstellation
                                newPosition = activeConstellation.getBoardZeile(hproj.name)
                            Else
                                newPosition = activeConstellation.getBoardZeile(hproj.name) + startOfFreeRows
                            End If
                        End If

                        Try
                            AlleProjekte.Add(hproj, updateCurrentConstellation:=True, sortkey:=newPosition, checkOnConflicts:=True)
                            ' jetzt die Variante aktivieren 
                            ' aber nur wenn es auch das Flag show hat 
                            If showIT Then

                                If boardwasEmpty Then
                                    'tryZeile = kvp.Value.zeile
                                    tryZeile = activeConstellation.getBoardZeile(hproj.name)
                                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, tryZeile)
                                Else
                                    'tryZeile = startOfFreeRows + zeilenOffset
                                    tryZeile = startOfFreeRows + activeConstellation.getBoardZeile(hproj.name) - 2
                                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, False, tryZeile)
                                    'zeilenOffset = zeilenOffset + 1
                                End If

                            End If
                        Catch ex As Exception
                            Call MsgBox("Fehler mit Summary Projekten: " & ex.Message)
                        End Try

                    Else
                        outputLine = pName & "(" & vName & ") Code: 098 " & nvErrorMessage
                        outPutCollection.Add(outputLine)
                    End If

                Else
                    hproj = Nothing

                    outputLine = pName & "(" & vName & ")" & nvErrorMessage
                    outPutCollection.Add(outputLine)

                    'Call MsgBox("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                    'Throw New ArgumentException("Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                End If
            Else
                Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & pName & "'konnte nicht geladen werden")
            End If
        End If

    End Sub

    ''' <summary>
    ''' gibt das hproj zurück, zuerst wird versucht, das aus der AlleProjekte zu holen, dann aus der Datenbank
    ''' zuerst wird über Name gesucht, dan über Kunden-Nummer 
    ''' wenn es noch gar nicht existiert, wird nothing zurückgegeben
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <param name="projektliste"></param>
    ''' <param name="storedAt"></param>
    ''' <returns></returns>
    Public Function getProjektFromSessionOrDB(ByVal pName As String, ByVal vName As String, ByVal projektliste As clsProjekteAlle, ByVal storedAt As Date,
                                              ByVal Optional kdNr As String = "") As clsProjekt

        Dim err As New clsErrorCodeMsg
        Dim hproj As clsProjekt = Nothing

        If pName = "" And kdNr <> "" Then
            Try
                If projektliste.Count > 0 Then
                    hproj = projektliste.getProjectByKDNr(kdNr)
                End If

                If IsNothing(hproj) Then
                    Dim nameCollection As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(kdNr, err)
                    If nameCollection.Count > 0 Then
                        Dim newPname As String = CStr(nameCollection.Item(1))
                        hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(newPname, vName, "", storedAt, err)
                    End If
                End If

            Catch ex As Exception

            End Try
        Else
            Dim key As String = calcProjektKey(pName, vName)

            Try
                hproj = projektliste.getProject(key)
                ' wenn es noch nicht geladen ist, muss das Projekt aus der Datenbank geholt werden ..

                ' stimmt der TimeStamp 
                If Not IsNothing(hproj) Then
                    If hproj.timeStamp >= storedAt Then
                        hproj = Nothing
                    End If
                End If

                If IsNothing(hproj) Then

                    If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, storedAt, err) Then
                        hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pName, vName, "", storedAt, err)
                    Else
                        ' jetzt soll versucht werden, das Projekt über die Kunden-Nummer zu bestimmen, sofern die Kunden-Nummer angegeben wurde 
                        If kdNr <> "" Then
                            Dim nameCollection As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(kdNr, err)
                            If nameCollection.Count > 0 Then
                                Dim newPname As String = CStr(nameCollection.Item(1))
                                hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(newPname, vName, "", storedAt, err)
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try
        End If


        getProjektFromSessionOrDB = hproj

    End Function

    ''' <summary>
    ''' erzeugt das Union Projekt für die Konstellation ; ansonsten wird nichts gemacht 
    ''' wenn die Projekte noch nicht geladen sind, werden sie aus der Datenbank geholt, aber nicht in AlleProjekte geladen ...
    ''' wenn budget = 0 : es hat kein Budget 
    ''' wenn budget = -1: berechne das Budget aus den Budgets der Projekte bzw. Summary Projekte 
    ''' wenn budget > 0 : übernehme diesen Wert als Budget 
    ''' </summary>
    ''' <param name="considerImportProjekte"></param>
    Public Function calcUnionProject(ByVal activeConstellation As clsConstellation,
                                     ByVal considerImportProjekte As Boolean,
                                     ByVal storedAtOrBefore As Date,
                                Optional ByVal budget As Double = -1.0,
                                Optional ByVal description As String = "Summen Projekt eines Programmes / Portfolios",
                                Optional ByVal ampel As Integer = 0,
                                Optional ByVal ampelbeschreibung As String = "",
                                Optional ByVal responsible As String = "") As clsProjekt

        Dim calculateBudget As Boolean = (budget <= -0.99)
        Dim gesamtbudget As Double = budget
        Dim unionProj As clsProjekt = Nothing
        Dim projektListe As clsProjekteAlle = AlleProjekte
        'Dim outPutListe As clsProjekteAlle = AlleProjektSummaries

        ' der Planungs-stand des unionized PRojektes ist das spätesteste Datum, das eines der Projekte im Portfolio hat  
        Dim timeStampOfUnionizedProject As Date = Date.MinValue

        ' das budget 
        If budget >= 0 Then
            gesamtbudget = budget
        Else
            gesamtbudget = 0
        End If

        ' jetzt die Union bilden ... das erste als Default besetzen 
        Dim listOfProjectNames As SortedList(Of String, String) = activeConstellation.getProjectNames(considerShowAttribute:=True,
                                                                                       showAttribute:=True,
                                                                                       fullNameKeys:=True)

        If considerImportProjekte Then
            projektListe = ImportProjekte
        End If

        Try
            If listOfProjectNames.Count > 0 Then
                ' nur, wenn überhaupt Projekte angezeigt würden, muss eine Union gemacht werden 

                ' jetzt mit allen anderen aufsummieren ..
                Dim isFirstProj As Boolean = True
                Dim minActualDate As Date = Date.MinValue
                Dim unionVariantName As String = ""


                For Each kvp As KeyValuePair(Of String, String) In listOfProjectNames

                    Dim hproj As clsProjekt = Nothing

                    If awinSettings.loadPFV Then
                        unionVariantName = ptVariantFixNames.pfv.ToString
                        hproj = getProjektFromSessionOrDB(getPnameFromKey(kvp.Key),
                                                          unionVariantName,
                                                          projektListe, storedAtOrBefore)

                    Else
                        hproj = getProjektFromSessionOrDB(getPnameFromKey(kvp.Key),
                                                          getVariantnameFromKey(kvp.Key),
                                                          projektListe, storedAtOrBefore)
                    End If


                    If Not IsNothing(hproj) Then

                        If timeStampOfUnionizedProject < hproj.timeStamp Then
                            timeStampOfUnionizedProject = hproj.timeStamp
                        End If

                        If isFirstProj Then
                            minActualDate = hproj.actualDataUntil

                            Dim startdate As Date = hproj.startDate
                            Dim endeDate As Date = hproj.endeDate
                            unionProj = New clsProjekt(activeConstellation.constellationName, True, startdate, endeDate)

                            isFirstProj = False
                        End If

                        If budget < 0.0 Then
                            ' das Gesamtbudget soll sich aus der Summe der Einzelbudgets ergeben ... 
                            gesamtbudget = gesamtbudget + hproj.Erloes
                        End If

                        ' hat eines der Projekte bereits eine Actualdata? 
                        ' wenn ja, dann wird das größte auftretende hergenommen 
                        If Not hproj.actualDataUntil = Date.MinValue Then
                            If DateDiff(DateInterval.Month, minActualDate, hproj.actualDataUntil) < 0 Then
                                minActualDate = hproj.actualDataUntil
                            End If
                        End If


                        unionProj.variantName = unionVariantName
                        unionProj = unionProj.unionizeWith(hproj)

                    End If

                Next

                ' jetzt ggf die Attribute noch ergänzen 
                With unionProj
                    .Erloes = gesamtbudget
                    .Status = ProjektStatus(PTProjektStati.beauftragt)
                    .description = description
                    .ampelStatus = ampel
                    .ampelErlaeuterung = ampelbeschreibung
                    .leadPerson = responsible

                    ' gibt es ein ActualDate ? 
                    If Not minActualDate = Date.MinValue Then
                        .actualDataUntil = minActualDate
                    End If

                    ' tk geändert 
                    ' If Date.Now < storedAtOrBefore Then
                    If Date.Now < timeStampOfUnionizedProject Then
                        .timeStamp = Date.Now
                    Else
                        .timeStamp = timeStampOfUnionizedProject
                    End If
                End With


            End If
        Catch ex As Exception

        End Try

        'ProjListe = projektListe ' Liste an Projekte, aus der das SummaryProjekt entstanden ist

        calcUnionProject = unionProj

    End Function

    ''' <summary>
    ''' gibt true zurück, wenn es keine Konflikte zwischen Summary Projekten und Projekten bzw Summary Projekten gibt 
    ''' </summary>
    ''' <param name="constellationName"></param>
    ''' <param name="isSummaryOption"></param>
    ''' <returns></returns>
    Public Function thereAreNoPortfolioProjectConflicts(ByVal constellationName As String,
                                                        ByVal isSummaryOption As Boolean) As Boolean
        Dim tmpResult As Boolean = True

        If Not isSummaryOption Then

        End If

        thereAreNoPortfolioProjectConflicts = tmpResult
    End Function

    ''' <summary>
    ''' zeigt die Konstellation bzw Konstellationen auf der Projekt-Tafel an 
    ''' addToSession gibt an, ob AlleProjekte und ggf ShowProjekte ergänzt wird 
    ''' </summary>
    ''' <param name="constellationsToShow"></param>
    ''' <param name="clearBoard">setzt ShowProjekte zurück, löscht das Zeichenbrett; lässt AlleProjekte unverändert </param>
    ''' <param name="clearSession">setzt alles zurück></param>
    ''' <param name="storedAtOrBefore"></param>
    ''' <remarks></remarks>
    Public Sub showConstellations(ByVal constellationsToShow As clsConstellations,
                                  ByVal clearBoard As Boolean,
                                  ByVal clearSession As Boolean,
                                  ByVal storedAtOrBefore As Date,
                                  Optional ByVal showSummaryProject As Boolean = False,
                                  Optional ByVal onlySessionLoad As Boolean = False)

        Dim err As New clsErrorCodeMsg

        Try
            Dim boardWasEmpty As Boolean = (ShowProjekte.Count = 0)
            Dim sessionWasEmpty As Boolean = (AlleProjekte.Count = 0)

            Dim calculateSummaryProjekt As Boolean = ((myCustomUserRole.customUserRole <> ptCustomUserRoles.PortfolioManager) Or
                                                     (Not awinSettings.loadPFV))
            Dim activeSummaryConstellation As clsConstellation = Nothing

            If clearSession And Not sessionWasEmpty Then
                Call clearCompleteSession()

            ElseIf clearBoard And Not boardWasEmpty Then
                Call clearProjectBoard()

            End If

            ' hier muss eine summaryConstellation gemacht werden 

            ' wird im Falle showSummaryProject benötigt 
            activeSummaryConstellation = New clsConstellation(skey:=ptSortCriteria.customTF)
            Dim zaehler As Integer = 1

            For Each kvp As KeyValuePair(Of String, clsConstellation) In constellationsToShow.Liste

                '' '???ur:8..4.2019: Portfolio-Projekte lesen'' es werden erst mal alle Projekte zu der Constellation kvp geholt
                '' ''Dim projsOfCurConstellation As SortedList(Of String, clsProjekt) =
                '' ''    CType(databaseAcc, DBAccLayer.Request).retrieveProjectsOfOneConstellationFromDB(kvp.Key, err, storedAtOrBefore)

                If kvp.Value.variantName = "" Then    ' StandardVariante der Constellation

                    ' hier wird die Summary Projekt Vorlage erst mal geholt , um das vorgegebene Budget zu ermitteln
                    Dim curSummaryProjVorgabe As clsProjekt = Nothing
                    Dim curSummaryProjToUse As clsProjekt = Nothing

                    Dim vorgabeBudget As Double = -1
                    ' hole die Vorgabe des Summary Projekts, die enthält nämlich die Vorgabe für das Budget 

                    Dim variantName As String = ptVariantFixNames.pfv.ToString
                    ' tk 22.7.19 es muss unterschiedenwerden, ob nur von der Session geladen werden soll 
                    ' das ist z.B wichtig, um nach einem Import von Projekten und den dazugehörigen Projekten die nur in der Session vorhandenen 
                    ' Summary PRojekte, die zu dem Zeitpunkt alle Variante-Name = "" haben zu finden 
                    If onlySessionLoad And Not awinSettings.loadPFV Then
                        variantName = ""
                    End If

                    curSummaryProjVorgabe = getProjektFromSessionOrDB(kvp.Value.constellationName, variantName, AlleProjekte, storedAtOrBefore)
                    If Not IsNothing(curSummaryProjVorgabe) Then
                        vorgabeBudget = curSummaryProjVorgabe.Erloes
                    End If

                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And awinSettings.loadPFV Then
                        ' laden von Datenbank, er ist hier bereits fertig, aber wenn es nothing sein sollte, dann erstelle es .. 

                        If IsNothing(curSummaryProjVorgabe) Then
                            ' hier muss das Budget aus den einzelnen Projekten errechnet werden 
                            curSummaryProjVorgabe = calcUnionProject(kvp.Value, False, storedAtOrBefore, budget:=-1, description:="Summen-Projekt von " & kvp.Key)
                        End If
                        curSummaryProjToUse = curSummaryProjVorgabe
                    Else
                        curSummaryProjToUse = calcUnionProject(kvp.Value, False, storedAtOrBefore, budget:=vorgabeBudget, description:="Summen-Projekt von " & kvp.Key)
                    End If


                    If Not IsNothing(curSummaryProjToUse) Then

                        ' der Summary-Projekt Key ist nicht unbedingt gleich de kvp.key 
                        Dim srKey As String = calcProjektKey(curSummaryProjToUse)

                        If showSummaryProject Then
                            ' dann sollen die Summary Projekte in AlleProjekte eingetragen werden ...
                            If AlleProjekte.Containskey(srKey) Then
                                AlleProjekte.Remove(srKey, True)
                            End If

                            Try
                                AlleProjekte.Add(curSummaryProjToUse, updateCurrentConstellation:=True, checkOnConflicts:=True)
                                Dim cItem As New clsConstellationItem
                                With cItem
                                    .projectName = kvp.Value.constellationName
                                    .variantName = curSummaryProjToUse.variantName
                                    .projectTyp = CType(curSummaryProjToUse.projectType, ptPRPFType).ToString
                                    .zeile = zaehler
                                    .show = True
                                End With
                                zaehler = zaehler + 1
                                activeSummaryConstellation.add(cItem, sKey:=zaehler)
                            Catch ex As Exception

                            End Try
                        Else
                            '' die Summary Projekte können nicht in AlleProjekte eingetragen werden, weil das zu Konflikten mit den dort abgelegten Einzelprojekten führt
                            '' deshalb werden in diesem Fall die SummaryProjekte  in AlleProjektSummaries eingetragen
                            If AlleProjektSummaries.Containskey(srKey) Then
                                AlleProjektSummaries.Remove(srKey, updateCurrentConstellation:=False)
                            End If

                            Try
                                AlleProjektSummaries.Add(curSummaryProjToUse, updateCurrentConstellation:=False, checkOnConflicts:=False)
                            Catch ex As Exception

                            End Try
                        End If

                    End If

                Else
                    ' Variante der Constellation kann kein Summary-Projekt haben
                End If

            Next


            If showSummaryProject Then
                ' damit der Name des einen Portfolios übernommen wird ...
                If constellationsToShow.Count = 1 Then
                    activeSummaryConstellation.constellationName = constellationsToShow.Liste.ElementAt(0).Value.constellationName
                End If
                constellationsToShow.Liste.Clear()
                constellationsToShow.Add(activeSummaryConstellation)
                showSummaryProject = False
            End If



            ' tk 28.10.17, wenn die Anzahl der Constellations < 1 ist, dann muss es immer auf CustomTF gesetzt werden ... 
            Dim anzConstellations As Integer = constellationsToShow.Liste.Count
            For Each kvp As KeyValuePair(Of String, clsConstellation) In constellationsToShow.Liste

                ' tk , jetzt anpassen, wenn es mehr als 1 Constellation sind 
                If anzConstellations > 1 Then
                    currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                    kvp.Value.sortCriteria = ptSortCriteria.customTF
                Else
                    If kvp.Value.sortCriteria <> currentSessionConstellation.sortCriteria Then

                        If kvp.Value.sortCriteria >= 0 Then
                            currentSessionConstellation.sortCriteria = kvp.Value.sortCriteria
                        Else
                            currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                            kvp.Value.sortCriteria = ptSortCriteria.customTF
                        End If
                    Else

                        If kvp.Value.sortCriteria < 0 Then
                            currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                            kvp.Value.sortCriteria = ptSortCriteria.customTF
                        End If

                    End If

                End If

                ' jetzt den Sortier-Modus anpassen 
                Call visualizeConstellation(kvp.Value, storedAtOrBefore, awinSettings.loadPFV)

            Next

            If constellationsToShow.Count = 1 Then


                If clearSession Or sessionWasEmpty Or
                    clearBoard Or boardWasEmpty Then
                    'ur: 2019-07-08: notwendig um die vpid zu retten
                    currentSessionConstellation = constellationsToShow.Liste.ElementAt(0).Value.copy(False)
                    currentConstellationPvName = calcPortfolioKey(constellationsToShow.Liste.ElementAt(0).Value)
                Else
                    currentConstellationPvName = calcLastSessionScenarioName()
                    ' hier muss jetzt der sortType auf CustomTF gesetzt werden 

                    'ur:2019-07-08: es sind mehrere Portfolios in einer currentSessionConstellation
                    currentSessionConstellation.vpID = ""

                    If Not IsNothing(currentSessionConstellation) Then
                        currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                    End If

                End If
            Else
                currentConstellationPvName = calcLastSessionScenarioName()
                ' hier muss jetzt der sortType auf CustomTF gesetzt werden  
                If Not IsNothing(currentSessionConstellation) Then
                    'ur:2019-07-08: es sind mehrere Portfolios in einer currentSessionConstellation
                    currentSessionConstellation.vpID = ""
                    currentSessionConstellation.sortCriteria = ptSortCriteria.customTF
                End If
            End If

            Call awinNeuZeichnenDiagramme(2)

            ' die aktuelle Konstellation in "Last" speichern 
            'Call storeSessionConstellation("Last")

        Catch ex As Exception
            Call MsgBox("Fehler bei Laden : " & vbLf & ex.Message)
        End Try


    End Sub

    ''' <summary>
    ''' speichert eine einzelne Konstellation in die Datenbank
    ''' dabei werden alle Projekte und Projekt-Varianten, die noch nicht oder in anderer Form in der Datenbank gespeichert sind, abgespeichert 
    ''' </summary>
    ''' <param name="currentConstellation"></param>
    ''' <remarks></remarks>
    Public Sub storeSingleConstellationToDB(ByRef outPutCollection As Collection,
                                            ByVal currentConstellation As clsConstellation,
                                            ByVal dbConstellations As SortedList(Of String, String))

        Dim err As New clsErrorCodeMsg

        Dim anzahlNeue As Integer = 0
        Dim anzahlChanged As Integer = 0
        Dim DBtimeStamp As Date = Date.Now
        Dim outputLine As String = ""
        Dim ctimestamp As Date

        ' wenn HistoryMode aktiv ist ... 
        If demoModusHistory Then
            DBtimeStamp = historicDate
        End If

        ' jetzt muss überprüft werden, ob auch alle Projekte bereits in der DB gespeichert sind ...  
        Dim storeISAllowed As Boolean = True

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In currentConstellation.Liste
            If Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(kvp.Value.projectName, kvp.Value.variantName, Date.Now, err) Then
                storeISAllowed = False
                Exit For
            End If
        Next

        If Not storeISAllowed Then
            If awinSettings.englishLanguage Then
                outputLine = "Portfolio contains Projects which aren't existing in DB! Please save projects first!"
                outPutCollection.Add(outputLine)
            Else
                outputLine = "Portfolio enthält Projekte, die noch nicht in der Datenbank enthalten sind. Bitte zuerst Projekte speichern!"
                outPutCollection.Add(outputLine)
            End If
        Else
            ' jetzt muss ggf das Summary Projekt zur Constellation erzeugt und gespeichert werden
            Try
                ' das Summary Project muss auf Basis der geladenen Projekte erstellt werden 
                Dim budget As Double = -1.0
                Dim calculateAndStoreSummaryProjekt As Boolean = False
                Dim mSProj As clsProjekt = Nothing   ' nimmt das gemergte Summary-Projekt aus
                ' TODO: currentConstellation.variantName berücksichtigen
                Dim tmpVariantName As String = getDefaultVariantNameAccordingUserRole()

                Dim oldSummaryP As clsProjekt = getProjektFromSessionOrDB(currentConstellation.constellationName, tmpVariantName, AlleProjekte, Date.Now)

                ' das Portfolio Projekt
                ' tk 5.2.20 das sollte immer (!) neu berechnet werden, schließlich haben sich ja di eProjekte geändert 
                ' und wenn das alles identisch ist, dann wird das durch die spätere Überprüfung rausgefunden ... 
                'calculateAndStoreSummaryProjekt = IsNothing(oldSummaryP) Or myCustomUserRole.customUserRole <> ptCustomUserRoles.PortfolioManager
                If currentConstellation.variantName = "" Then
                    calculateAndStoreSummaryProjekt = True
                Else
                    calculateAndStoreSummaryProjekt = False
                End If

                Dim sproj As clsProjekt = Nothing

                If calculateAndStoreSummaryProjekt Then

                    If Not IsNothing(oldSummaryP) Then
                        'budget = oldSummaryP.budgetWerte.Sum
                        budget = oldSummaryP.Erloes
                        If budget = 0 Then
                            budget = currentConstellation.getBudgetOfShownProjects
                            oldSummaryP.Erloes = budget
                        End If
                        sproj = oldSummaryP
                    Else
                        budget = currentConstellation.getBudgetOfShownProjects
                        sproj = calcUnionProject(currentConstellation, False, Date.Now, budget:=budget)
                        sproj.variantName = tmpVariantName
                    End If




                    If Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(sproj.name, sproj.variantName, Date.Now, err) Then
                        ' speichern des Projektes 

                        sproj.timeStamp = DBtimeStamp

                        ' hier wird kein attrToStore Angabe benötigt, weil das Projekt ja noch gar nicht existiert hat ... 
                        If CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(sproj, dbUsername, mSProj, err) Then

                            If awinSettings.englishLanguage Then
                                outputLine = "Portfolio / Summary Project saved: " & sproj.name & ", " & sproj.variantName
                                outPutCollection.Add(outputLine)
                            Else
                                outputLine = "Portfolio / Summary Projekt gespeichert: " & sproj.name & ", " & sproj.variantName
                                outPutCollection.Add(outputLine)
                            End If

                            anzahlNeue = anzahlNeue + 1

                            Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(sproj.name, sproj.variantName, err)
                            writeProtections.upsert(wpItem)


                        Else
                            ' kann eigentlich gar nicht sein ... wäre nur dann der Fall, wenn ein Projekt komplett gelöscht wurde , aber der Schreibschutz nicht gelöscht wurde 
                            If awinSettings.visboServer Then
                                Select Case err.errorCode
                                    Case 403  'No Permission to Create Visbo Project Version
                                        If awinSettings.englishLanguage Then
                                            outputLine = "!!  No permission to store : " & sproj.name & ", " & sproj.variantName
                                            outPutCollection.Add(outputLine)
                                        Else
                                            outputLine = "!!  Keine Erlaubnis zu speichern : " & sproj.name & ", " & sproj.variantName
                                            outPutCollection.Add(outputLine)
                                        End If

                                    Case 409 ' VisboProjectVersion was already updated in between
                                        If awinSettings.englishLanguage Then
                                            outputLine = "!! Projekt was already updated in between : " & sproj.name & ", " & sproj.variantName
                                            outPutCollection.Add(outputLine)
                                        Else
                                            outputLine = "!!  Projekt wurde inzwischen verändert : " & sproj.name & ", " & sproj.variantName
                                            outPutCollection.Add(outputLine)
                                        End If
                                '' erneut das projekt holen und abändern
                                '' ur: 09.01.2019: wird in storeProjectToDB direkt gemacht
                                'Dim standInDB As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(kvp.Value.name, kvp.Value.variantName, jetzt, err)

                                    Case 423 ' Visbo Project (Portfolio) is locked by another user
                                        If awinSettings.englishLanguage Then
                                            outputLine = err.errorMsg & ": " & sproj.name & ", " & sproj.variantName
                                            outPutCollection.Add(outputLine)
                                        Else
                                            outputLine = "geschüztes Projekt : " & sproj.name & ", " & sproj.variantName
                                            outPutCollection.Add(outputLine)
                                        End If

                                End Select
                            Else

                                ' kann eigentlich gar nicht sein ... wäre nur dann der Fall, wenn ein Projekt komplett gelöscht wurde , aber der Schreibschutz nicht gelöscht wurde 
                                If awinSettings.englishLanguage Then
                                    outputLine = "protected project: " & sproj.name & ", " & sproj.variantName
                                Else
                                    outputLine = "geschütztes Projekt: " & sproj.name & ", " & sproj.variantName
                                End If
                                outPutCollection.Add(outputLine)

                            End If

                            Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(sproj.name, sproj.variantName, err)
                            writeProtections.upsert(wpItem)

                        End If
                    Else
                        ' das Portfolio Projekt wird gespeichert , wenn es Unterschiede gibt 
                        Dim oldProj As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(sproj.name, sproj.variantName, "", Date.Now, err)
                        ' Type = 0: Projekt wird mit Variante bzw. anderem zeitlichen Stand verglichen ...

                        If Not IsNothing(oldProj) Then
                            If Not sproj.isIdenticalTo(oldProj) Then

                                sproj.timeStamp = DBtimeStamp

                                Dim kdNrToStore As Boolean = Not sproj.hasIdenticalKdNr(oldProj)

                                ' abfragen, ob Portfolio MAnager
                                If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                    If sproj.variantName = ptVariantFixNames.pfv.ToString Then
                                        sproj.updatedAt = oldProj.updatedAt
                                    End If
                                End If


                                If CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(sproj, dbUsername, mSProj, err, attrToStore:=kdNrToStore) Then

                                    If awinSettings.englishLanguage Then
                                        outputLine = "saved: " & sproj.name & ", " & sproj.variantName
                                        outPutCollection.Add(outputLine)
                                    Else
                                        outputLine = "gespeichert: " & sproj.name & ", " & sproj.variantName
                                        outPutCollection.Add(outputLine)
                                    End If

                                    ' alles ok
                                    anzahlChanged = anzahlChanged + 1

                                    Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(sproj.name, sproj.variantName, err)
                                    writeProtections.upsert(wpItem)

                                Else
                                    If awinSettings.visboServer Then
                                        Select Case err.errorCode
                                            Case 403  'No Permission to Create Visbo Project Version
                                                If awinSettings.englishLanguage Then
                                                    outputLine = "!!  No permission to store Summary Project : " & sproj.name & ", " & sproj.variantName
                                                    outPutCollection.Add(outputLine)
                                                Else
                                                    outputLine = "!!  Keine Erlaubnis, Summary Projekt  zu speichern : " & sproj.name & ", " & sproj.variantName
                                                    outPutCollection.Add(outputLine)
                                                End If

                                            Case 409 ' VisboProjectVersion was already updated in between
                                                If awinSettings.englishLanguage Then
                                                    outputLine = "!! Summary Project was already updated in between : " & sproj.name & ", " & sproj.variantName
                                                    outPutCollection.Add(outputLine)
                                                Else
                                                    outputLine = "!!  Summary Projekt wurde inzwischen verändert : " & sproj.name & ", " & sproj.variantName
                                                    outPutCollection.Add(outputLine)
                                                End If
                                                '' erneut das projekt holen und abändern
                                                '' ur: 09.01.2019: wird in storeProjectToDB direkt gemacht
                                                'Dim standInDB As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(kvp.Value.name, kvp.Value.variantName, jetzt, err)

                                            Case 423 ' Visbo Project (Portfolio) is locked by another user
                                                If awinSettings.englishLanguage Then
                                                    outputLine = err.errorMsg & ": " & sproj.name & ", " & sproj.variantName
                                                    outPutCollection.Add(outputLine)
                                                Else
                                                    outputLine = "geschüztes Projekt : " & sproj.name & ", " & sproj.variantName
                                                    outPutCollection.Add(outputLine)
                                                End If

                                        End Select
                                    Else

                                        ' kann eigentlich gar nicht sein ... wäre nur dann der Fall, wenn ein Projekt komplett gelöscht wurde , aber der Schreibschutz nicht gelöscht wurde 
                                        If awinSettings.englishLanguage Then
                                            outputLine = "protected project: " & sproj.name & ", " & sproj.variantName
                                        Else
                                            outputLine = "geschütztes Projekt: " & sproj.name & ", " & sproj.variantName
                                        End If
                                        outPutCollection.Add(outputLine)

                                    End If

                                    Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(sproj.name, sproj.variantName, err)
                                    writeProtections.upsert(wpItem)

                                End If
                            End If
                        Else
                            Call MsgBox("Fehler bei Datenbank Lesen (storeSingleConstellationToDB): Projekt existiert , kann aber nicht gelsen werden: " & vbLf &
                                        sproj.name)
                        End If

                    End If

                    ''If Not storeSingleProjectToDB(sproj, outPutCollection, identical:=isIdentical) Then ' wird im 1Click-PPT benötigt
                    ''    Call MsgBox("speichern Summary Projekt mit Fehler ...")
                    ''Else
                    ''    Dim a As Integer = outPutCollection.Count
                    ''End If
                    If Not IsNothing(mSProj) Then
                        ' mergte Summary wurde in die Liste aufgenommen
                        Dim skey As String = calcProjektKey(mSProj.name, mSProj.variantName)
                        If AlleProjektSummaries.Containskey(skey) Then
                            AlleProjektSummaries.Remove(skey, False)
                        End If

                        If Not AlleProjektSummaries.Containskey(skey) Then
                            AlleProjektSummaries.Add(mSProj, False)
                        End If
                    Else
                        ' ungemergtes Summary-Projekt wird in die Liste aufgenommen
                        Dim skey As String = calcProjektKey(sproj.name, sproj.variantName)
                        If AlleProjektSummaries.Containskey(skey) Then
                            AlleProjektSummaries.Remove(skey, False)
                        End If

                        If Not AlleProjektSummaries.Containskey(skey) Then
                            AlleProjektSummaries.Add(sproj, False)
                        End If
                    End If


                End If


            Catch ex As Exception

            End Try


            ' jetzt wird das Portfolio weggeschrieben 
            Try
                Dim storeRequired As Boolean = True
                ' hier wird die Constellation aus der DB geholt , wenn Sie nicht schon geholt wurde ...
                If IsNothing(dbConstellations) Then
                    storeRequired = True
                Else
                    If dbConstellations.Count = 0 Then
                        storeRequired = True
                    Else
                        If dbConstellations.ContainsKey(currentConstellation.constellationName) Then
                            Dim dbConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(currentConstellation.constellationName,
                                                                                                           dbConstellations(currentConstellation.constellationName),
                                                                                                           ctimestamp, err, variantName:=currentConstellation.variantName,
                                                                                                           storedAtOrBefore:=DBtimeStamp)
                            ' dbConstellation ist nothing, wenn z.B. die Variante noch nicht existiert
                            If Not IsNothing(dbConstellation) Then
                                storeRequired = Not currentConstellation.isIdentical(dbConstellation)
                            Else
                                storeRequired = True
                            End If
                        End If
                    End If

                End If


                ' hier wird geprüft, ob die sich überhaupt verändert hat  
                If storeRequired Then

                    ' ur: 26.10.2019: nicht mehr Date.now, da sonst das Summary-Projekt einen Timestamp hat, der vor dem Portfolio liegt, was unlogisch ist

                    currentConstellation.timestamp = DBtimeStamp

                    ' darf das so in der DB gespeichert werden? d.h sind für jedes Projekt genau aine Variante enthalten ? 
                    If currentConstellation.isValidForDBStore Then
                        Dim constellationDB As clsConstellation = currentConstellation.copy(dontConsiderNoShows:=True, prepareForDB:=True)

                        If CType(databaseAcc, DBAccLayer.Request).storeConstellationToDB(constellationDB, err) Then
                            ' alles in Ordnung, Speichern hat geklappt ...
                            Dim tsMessage As String = ""
                            If awinSettings.englishLanguage Then
                                tsMessage = "Zeitstempel: " & DBtimeStamp.ToShortDateString & ", " & DBtimeStamp.ToShortTimeString
                            Else
                                tsMessage = "Timestamp: " & DBtimeStamp.ToShortDateString & ", " & DBtimeStamp.ToShortTimeString
                            End If

                            If awinSettings.englishLanguage Then
                                outputLine = "Saved ... " & vbLf & "Portfolio: " & currentConstellation.constellationName & vbTab &
                                    "Variante: " & currentConstellation.variantName & vbLf & tsMessage

                            Else
                                outputLine = "Gespeichert ... " & vbLf & "Portfolio: " & currentConstellation.constellationName & vbTab &
                                    "Variante: " & currentConstellation.variantName & vbLf & tsMessage
                            End If

                            outPutCollection.Add(outputLine)

                        Else
                            If awinSettings.englishLanguage Then
                                outputLine = "Error when writing scenario: " & currentConstellation.constellationName
                            Else
                                outputLine = "Fehler beim Schreiben Szenario: " & currentConstellation.constellationName
                            End If
                            outPutCollection.Add(outputLine)

                        End If
                    Else
                        If awinSettings.englishLanguage Then
                            outputLine = "Portfolio contains at least one project with more than one variant - please correct: " & currentConstellation.constellationName & "[" & currentConstellation.variantName & "]"
                        Else
                            outputLine = "Portfolio darf pro Projekt nicht mehr als 1 Variante enthalten - bitte korrigieren: " & currentConstellation.constellationName & "[" & currentConstellation.variantName & "]"
                        End If
                        outPutCollection.Add(outputLine)
                    End If
                Else
                    If awinSettings.englishLanguage Then
                        outputLine = "not stored: Portfolio identical to DB-Version : " & currentConstellation.constellationName & "[" & currentConstellation.variantName & "]"
                        outPutCollection.Add(outputLine)
                    Else
                        outputLine = "nicht gespeichert: Portfolio identisch mit Datenbank-Version : " & currentConstellation.constellationName & "[" & currentConstellation.variantName & "]"
                        outPutCollection.Add(outputLine)
                    End If

                End If
            Catch ex As Exception
                If awinSettings.englishLanguage Then
                    outputLine = "Error when writing Portfolio" & vbLf & ex.Message
                Else
                    outputLine = "Fehler beim Schreiben Portfolio" & vbLf & ex.Message
                End If
                outPutCollection.Add(outputLine)
                Exit Sub
            End Try


        End If



    End Sub

    ''' <summary>
    ''' löscht ein bestimmtes Portfolio aus der Datenbank und der Liste der Portfolios im Hauptspeicher
    ''' 
    ''' </summary>
    ''' <param name="constellationName">
    ''' Name, unter dem das Portfolio in der Datenbank gespeichert wurde 
    ''' </param>
    ''' <remarks></remarks>
    ''' 
    Public Sub awinRemoveConstellation(ByVal constellationName As String, ByVal vpid As String, ByVal deleteDB As Boolean)

        Dim err As New clsErrorCodeMsg

        Dim returnValue As Boolean = True
        Dim activeConstellation As New clsConstellation

        ' ur: 12.12.2019 entfernt, da kein readInitConstellations mehr gemacht wird
        ' prüfen, ob diese Constellation überhaupt existiert ..
        'Try
        '    activeConstellation = projectConstellations.getConstellation(constellationName)
        'Catch ex As Exception
        '    Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
        '    Exit Sub
        'End Try

        If deleteDB Then

            If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
                Dim pName As String = getPnameFromKey(constellationName)
                Dim vName As String = getVariantnameFromKey(constellationName)
                ' Konstellation muss aus der Datenbank gelöscht werden.
                returnValue = CType(databaseAcc, DBAccLayer.Request).removeConstellationFromDB(pName, vpid, vName, err)
                If returnValue = False Then
                    Call MsgBox("Fehler bei Löschen Portfolio : " & pName & "[" & vName & "]")
                Else
                    ' jetzt muss die Planung wie die Beauftragung des Portfolio Projekts gelöscht werden ... 
                    'Dim planungsKey As String = calcProjektKey(activeConstellation.constellationName, "")
                    'Dim beauftragungskey As String = calcProjektKey(activeConstellation.constellationName, ptVariantFixNames.pfv.ToString)

                    'Dim returnValue2 As Boolean = CType(databaseAcc, DBAccLayer.Request).deleteProjectTimestampFromDB()

                End If
            Else
                Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & activeConstellation.constellationName & "'konnte nicht gelöscht werden")
                returnValue = False
            End If
        End If

        If returnValue Then
            Try
                If Not IsNothing(constellationName) Then
                    projectConstellations.Remove(constellationName)
                Else
                    Call MsgBox("Es wurde keine Portfolio ausgewählt")
                End If
                If Not IsNothing(activeConstellation) Then
                    ' Konstellation muss aus der Liste aller Portfolios entfernt werden.
                    projectConstellations.Remove(activeConstellation.constellationName)
                Else
                    Call MsgBox("Es wurde kein Portfolio ausgewählt")
                End If

            Catch ex1 As Exception
                Call MsgBox("Fehler in awinRemoveConstellation aufgetreten: " & ex1.Message)
            End Try
        Else
            Call MsgBox("Es ist ein Fehler beim Löschen es Portfolios aus der Datenbank aufgetreten ")
        End If

    End Sub

    ''' <summary>
    ''' lädt die über pName#vName angegebene Variante aus der Datenbank;
    ''' show = true: es wird in Showprojekte eingetragen; sonst nur in AlleProjekte 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <remarks></remarks>
    Public Sub loadProjectfromDB(ByRef outputCollection As Collection,
                                 ByVal pName As String, vName As String, ByVal show As Boolean,
                                 ByVal storedAtORBefore As Date,
                                 ByVal calledFromPPT As Boolean)

        Dim err As New clsErrorCodeMsg

        Dim hproj As clsProjekt
        Dim key As String = calcProjektKey(pName, vName)


        If AlleProjekte.hasAnyConflictsWith(key, False) Then
            Dim outputLine As String = "Projekt " & pName & " kann nicht geladen werden. ist bereits in Summary Projekt enthalten"
            outputCollection.Add(outputLine)
        Else
            ' ab diesem Wert soll neu gezeichnet werden 
            Dim freieZeile As Integer = 2

            If Not calledFromPPT Then
                freieZeile = projectboardShapes.getMaxZeile
            End If


            hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pName, vName, "", storedAtORBefore, err)

            ' tk 4.2.20
            ' hier muss geprüft werden, ob das Projekt Ressourcen-Zuordnungen für Mitarbeiter enthält, die noch gar nicht da sind bzw. zu dem Zeitpunkt schon weg sind.
            ' es soll dann aber nur eine Warnung ausgegeben werden, sonst nichts weiter 
            If Not calledFromPPT And Not IsNothing(hproj) And DateDiff(DateInterval.Day, Date.Now, storedAtORBefore) = 0 Then
                ' nur bei aktuellen Projekten anmeckern ... 

                Dim invalidNeedNames As Collection = hproj.hasRolesWithInvalidNeeds

                If invalidNeedNames.Count > 0 Then

                    For Each iVName As String In invalidNeedNames
                        Dim msgTxt As String = "Projekt " & hproj.getShapeText & " enthält ungültige Ressourcen-Zuordnungen"
                        msgTxt = msgTxt & vbLf & "Person ist noch nicht oder nicht mehr im Unternehmen: " & iVName
                        outputCollection.Add(msgTxt)
                    Next

                End If

            End If



            If Not IsNothing(hproj) Then
                ' prüfen, ob AlleProjekte das Projekt bereits enthält 
                ' danach ist sichergestellt, daß AlleProjekte das Projekt bereit enthält 

                ' wenn jetzt gefiltert wurde und der Varianten-Name pfv ist, dann umsetzen 
                ' das muss aber nur gemacht werden, wenn nicht von Powerpoint , nur lesend aufgerufen ...
                If awinSettings.filterPFV And hproj.variantName = ptVariantFixNames.pfv.ToString And Not calledFromPPT Then
                    hproj.variantName = ""
                    vName = ""
                    key = calcProjektKey(pName, "")

                    ' tk um nachher auch speichern zu können , muss die Planungs-Variante jetzt auch gelesen werden ... 
                    Dim dummyProj As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pName, vName, "", storedAtORBefore, err)
                    hproj.updatedAt = dummyProj.updatedAt

                    If Not hproj.isIdenticalTo(dummyProj) Then
                        hproj.marker = True
                    End If

                End If

                If AlleProjekte.Containskey(key) Then
                    AlleProjekte.Remove(key, updateCurrentConstellation:=True)
                End If

                AlleProjekte.Add(hproj, updateCurrentConstellation:=True)

                ' nur machen, wenn nicht von PPT aufgerufen 
                If Not calledFromPPT Then
                    ' jetzt die writeProtections aktualisieren 
                    Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                    writeProtections.upsert(wpItem)

                    If show Then
                        ' prüfen, ob es bereits in der Showprojekt enthalten ist
                        ' diese Prüfung und die entsprechenden Aktionen erfolgen im 
                        ' replaceProjectVariant

                        Call replaceProjectVariant(pName, vName, False, True, freieZeile)

                    End If
                Else
                    ' wenn es aus PPT aus aufgerufen wird, muss das Projekt auch in ShowPRojekte eingetragen werden, 
                    ' sofern nicht schon ein PRojekt gleichen Namens drin ist. 
                    If Not ShowProjekte.contains(hproj.name) Then
                        ShowProjekte.Add(hproj)
                    End If
                End If


            Else
                Dim outputLine As String = "existiert nicht: " & pName & ", " & vName & " @ " & storedAtORBefore.ToString
                outputCollection.Add(outputLine)
            End If
        End If

    End Sub

    ''' <summary>
    ''' löscht in der Datenbank das ganze Projekt mit allen Varianten und Timestamps
    ''' nur: wenn es nicht in einem Portfolio referenziert ist
    ''' </summary>
    ''' <param name="outputCollection"></param>
    ''' <param name="pname"></param>
    Public Sub deleteCompleteProjectFromDB(ByRef outputCollection As Collection,
                                           ByVal pname As String)

        Dim deleteIsAllowed As Boolean = True
        Dim err As New clsErrorCodeMsg
        Dim outputline As String = ""


        ' Liste der Scenarios, die irgendeine Variante referenzieren ... 
        ' leerer Sting, wenn es keine Referenzen gibt .. 
        outputline = projectConstellations.getSzenarioNamesWith(pname, "$ALL", False)

        ' wenn es keine Referenzen gibt, ist der Delete erlaubt 
        deleteIsAllowed = (outputline = "")

        ''Dim variantListe As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveVariantNamesFromDB(pname, err)
        ''hinzufügen der Standardvariante
        ''variantListe.Add("", "")

        ''If Not IsNothing(variantListe) Then

        ''    For Each vname In variantListe
        ''        If notReferencedByAnyPortfolio(pname, vname) Then
        ''            deleteIsAllowed = deleteIsAllowed And True
        ''        Else
        ''            outputline = ("Projekt  '" & pname & "'  : nicht gelöscht - es wird in einem Portfolio referenziert")
        ''            outputCollection.Add(outputline)
        ''            deleteIsAllowed = False
        ''            Exit For
        ''        End If
        ''    Next
        ''Else
        ''    deleteIsAllowed = True

        ''End If

        If deleteIsAllowed Then
            If CType(databaseAcc, DBAccLayer.Request).removeCompleteProjectFromDB(pname, err) Then
                If awinSettings.englishLanguage Then
                    outputline = ("Project  '" & pname & "'  : deleted ")
                Else
                    outputline = ("Projekt  '" & pname & "'  : gelöscht ")
                End If
                outputCollection.Add(outputline)
            End If
        Else
            If awinSettings.englishLanguage Then
                outputline = "Delete denied: " & pname & " referenced by portfolios:" & vbLf & "   " & outputline
            Else
                outputline = "Delete nicht möglich: " & pname & " enthalten in Portfolios:" & vbLf & "   " & outputline
            End If
            outputCollection.Add(outputline)
        End If

    End Sub

    ''' <summary>
    ''' löscht in der Datenbank alle Timestamps der Projekt-Variante pname, variantname
    ''' die Timestamps werden zudem alle im Papierkorb gesichert 
    ''' </summary>
    ''' <param name="pname">Projektname</param>
    ''' <param name="variantName">Variantenname</param>
    ''' <remarks></remarks>
    Public Sub deleteCompleteProjectVariant(ByRef outputCollection As Collection,
                                            ByVal pname As String, ByVal variantName As String, ByVal kennung As Integer,
                                            Optional ByVal keepAnzVersions As Integer = 100)

        Dim err As New clsErrorCodeMsg

        Dim outputLine As String = ""

        ' tk 7.10.19 calledFromPPT nur true, wenn kennung = PTtvActions.loadPVinPPT
        Dim calledFromPPT As Boolean = (kennung = PTTvActions.loadPVInPPT)

        Dim anzTests As Integer = 0
        Dim anzDeleted As Integer = 0
        If kennung = PTTvActions.delFromDB Or
            kennung = PTTvActions.delAllExceptFromDB Then


            If kennung = PTTvActions.delAllExceptFromDB Then

                ' an dieser Stelle wird gecheckt
                ' 1. ist es eine echte Variante und hat sie keine customFields? 
                ' 2. wenn ja, dann hole die Basis-Variante , hat sie CustomFields
                ' 3. wenn ja, dann kopiere die Custom-Fields und speichere die Variante 
                ' mach dann den den Rest 
                ' Start Sonderbehandlung 
                If variantName <> "" Then
                    Dim anzCorrected As Integer = 0
                    Dim variantProject As clsProjekt
                    Dim baseProject As clsProjekt
                    Dim vExisted As Boolean = False
                    Dim bExisted As Boolean = False
                    Dim oCollection As New Collection
                    Dim keyV As String = calcProjektKey(pname, variantName)
                    Dim keyB As String = calcProjektKey(pname, "")

                    If Not AlleProjekte.Containskey(keyV) Then
                        Call loadProjectfromDB(oCollection, pname, variantName, False, Date.Now, calledFromPPT)
                    Else
                        vExisted = True
                    End If

                    If Not AlleProjekte.Containskey(keyB) Then
                        Call loadProjectfromDB(oCollection, pname, "", False, Date.Now, calledFromPPT)
                    Else
                        bExisted = True
                    End If

                    variantProject = AlleProjekte.getProject(keyV)
                    baseProject = AlleProjekte.getProject(keyB)
                    '
                    '' Sonderbehandlung alter Fehler bei Variantenbildung: Custom-Fields wurde nicht aus Base-Variant übernommen 
                    '  tk 19.1.19 diese Sonderbehandlung ist jetzt nicht mehr nötig 
                    'If Not IsNothing(variantProject) And Not IsNothing(baseProject) Then

                    '    ' Sonderbehandlung wegen ehemaligem Fehler, wo bei Varianten-Bildung die Custom-fields aus der Base-Variant nicht übernommen wurden 
                    '    If variantProject.getCustomFieldsCount = 0 And baseProject.getCustomFieldsCount > 0 Then
                    '        variantProject.copyCustomFieldsFrom(baseProject)
                    '        Dim zeitStempel As Date = variantProject.timeStamp

                    '        ' jetzt löschen, dann speichern ; wenn das löschen schiefgeht aufgrund Schreibschutz, dann geht auch das Speichern schief ... 
                    '        If writeProtections.isProtected(keyV, dbUsername) Then
                    '            ' kann nichts machen ...
                    '        Else
                    '            If CType(databaseAcc, DBAccLayer.Request).deleteProjectTimestampFromDB(pname, variantName, zeitStempel, dbUsername, err) Then
                    '                ' all ok 
                    '                If CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(variantProject, dbUsername, err) Then
                    '                    ' alles ok; jetzt  
                    '                Else

                    '                End If
                    '            End If
                    '        End If


                    '    End If

                    'End If

                    If Not vExisted Then
                        If AlleProjekte.Containskey(keyV) Then
                            AlleProjekte.Remove(keyV, updateCurrentConstellation:=True)
                        End If
                    End If
                    If Not bExisted Then
                        If AlleProjekte.Containskey(keyB) Then
                            AlleProjekte.Remove(keyB, updateCurrentConstellation:=True)
                        End If
                    End If
                End If

                ' Ende Sonderbehandlung  
                ' 
                '

                Dim timeStampsToDelete As Collection = identifyTimeStampsToDelete(pname, variantName, keepAnzVersions)


                If timeStampsToDelete.Count >= 1 Then

                    For Each singleTimeStamp As Date In timeStampsToDelete


                        If CType(databaseAcc, DBAccLayer.Request).deleteProjectTimestampFromDB(pname, variantName, singleTimeStamp, dbUsername, err) Then
                            ' all ok 
                            anzDeleted = anzDeleted + 1
                        Else
                            If awinSettings.englishLanguage Then
                                outputLine = "-->Error deleting (protected?): " & pname & ", " & variantName & ", " & singleTimeStamp.ToShortDateString
                            Else
                                outputLine = "-->Fehler beim Löschen (geschützt?): " & pname & ", " & variantName & ", " & singleTimeStamp.ToShortDateString
                            End If

                            outputCollection.Add(outputLine)
                        End If

                    Next

                    If awinSettings.englishLanguage Then
                        outputLine = pname & " (" & variantName & "): " & anzDeleted & " timestamps deleted"
                    Else
                        outputLine = pname & " (" & variantName & "): " & anzDeleted & " TimeStamps gelöscht"
                    End If

                    outputCollection.Add(outputLine)

                Else
                    If awinSettings.englishLanguage Then
                        outputLine = outputLine = pname & " (" & variantName & "): 0 timestamps deleted"
                    Else
                        outputLine = pname & " (" & variantName & "): 0 TimeStamps gelöscht"
                    End If

                    outputCollection.Add(outputLine)
                End If


            Else
                ' jetzt alle Timestamps in der Datenbank löschen 

                ' das darf aber nur passieren, wenn das Projekt, die Variante in keinem Szenario mehr referenziert wird ... 
                ' das hier ist eine doppelte Schranke sozusagen - in der Aufruf Schnittstelle wird das auch schon überprüft  
                If notReferencedByAnyPortfolio(pname, variantName) Then
                    Try

                        If Not IsNothing(projekthistorie) Then
                            projekthistorie.clear() ' alte Historie löschen
                        End If

                        projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB _
                                                (projectname:=pname, variantName:=variantName,
                                                 storedEarliest:=Date.MinValue, storedLatest:=Date.Now.AddDays(1), err:=err)


                        ' jetzt über alle Elemente der Projekthistorie ..
                        For Each kvp As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste

                            If CType(databaseAcc, DBAccLayer.Request).deleteProjectTimestampFromDB(pname, variantName, kvp.Key, dbUsername, err) Then
                                ' all ok 
                                anzDeleted = anzDeleted + 1
                            Else
                                If awinSettings.englishLanguage Then
                                    outputLine = "-->Error deleting (protected?): " & pname & ", " & variantName & ", " & kvp.Key.ToShortDateString
                                Else
                                    outputLine = "-->Fehler beim Löschen (geschützt?): " & pname & ", " & variantName & ", " & kvp.Key.ToShortDateString
                                End If

                                outputCollection.Add(outputLine)

                            End If

                        Next

                    Catch ex As Exception

                    End Try

                Else
                    If variantName = "" Then
                        If awinSettings.englishLanguage Then
                            outputLine = "delete denied: " & pname & " - Scenarios: "
                        Else
                            outputLine = "Löschen verweigert:  " & pname & " - Szenarien: "
                        End If

                    Else
                        If awinSettings.englishLanguage Then
                            outputLine = "delete denied: " & pname & " (" & variantName & ") " & " - Scenarios: "
                        Else
                            outputLine = "Löschen verweigert:  " & pname & " (" & variantName & ") " & " - Szenarien: "
                        End If

                    End If
                    ' false: ohne den Zusatztext : referenced by Portfolio(s: 
                    outputLine = outputLine & projectConstellations.getSzenarioNamesWith(pname, variantName, False)
                    outputCollection.Add(outputLine)
                End If


            End If



        ElseIf kennung = PTTvActions.delFromSession Or
            kennung = PTTvActions.deleteV Then

            ' eine einzelne Variante kann nur gelöscht werden, wenn 
            ' es sich weder um die variantName = "" noch um die aktuell gezeigte Variante handelt 

            Dim hproj As clsProjekt
            Try
                hproj = ShowProjekte.getProject(pname)
            Catch ex As Exception
                hproj = Nothing
            End Try

            If IsNothing(hproj) Then
                Dim key As String = calcProjektKey(pname, variantName)
                AlleProjekte.Remove(key, updateCurrentConstellation:=True)

            ElseIf hproj.variantName <> variantName Then
                Dim key As String = calcProjektKey(pname, variantName)
                AlleProjekte.Remove(key, updateCurrentConstellation:=True)

            Else
                ' es wird in Showprojekte und in AlleProjekte gelöscht, ausserdem auch auf der Projekt-Tafel 

                Dim key As String = calcProjektKey(pname, variantName)

                Try

                    ' jetzt muss die bisherige Variante aus Showprojekte rausgenommen werden ..
                    ShowProjekte.Remove(hproj.name)

                    ' die gewählte Variante wird rausgenommen
                    AlleProjekte.Remove(key)

                    Call clearProjektinPlantafel(pname)

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        outputLine = "delete denied: " & pname & " (" & variantName & ") " & " - Scenarios: ""Error when deleting: " & pname & " (" & variantName & ")"
                    Else
                        outputLine = "delete denied: " & pname & " (" & variantName & ") " & " - Scenarios: ""Fehler beim Löschen: " & pname & " (" & variantName & ")"
                    End If
                    outputCollection.Add(outputLine)
                End Try


            End If



        End If


    End Sub

    ''' <summary>
    ''' bestimmt die Time-Stamps, die gelöscht werden sollen 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <param name="keepAnzVersions"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function identifyTimeStampsToDelete(ByVal pName As String, vName As String, Optional ByVal keepAnzVersions As Integer = -1) As Collection

        Dim err As New clsErrorCodeMsg

        Dim tsToDelete As New Collection


        If Not IsNothing(projekthistorie) Then
            projekthistorie.clear() ' alte Historie löschen
        End If

        projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB _
                                (projectname:=pName, variantName:=vName,
                                 storedEarliest:=Date.MinValue, storedLatest:=Date.Now.AddDays(1), err:=err)

        If projekthistorie.Count <= keepAnzVersions Or projekthistorie.Count < 2 Then
            ' es muss nix gelöscht werden ... 

        Else
            Dim listToKeep As New SortedList(Of Date, String)
            Dim anzahlTS As Integer = projekthistorie.Count


            ' das erste Projekt merken 
            If Not listToKeep.ContainsKey(projekthistorie.ElementAt(0).timeStamp) Then
                listToKeep.Add(projekthistorie.ElementAt(0).timeStamp, "")
            End If

            ' das letzte Projekt merken 
            If Not listToKeep.ContainsKey(projekthistorie.ElementAt(anzahlTS - 1).timeStamp) Then
                listToKeep.Add(projekthistorie.ElementAt(anzahlTS - 1).timeStamp, "")
            End If



            ' das letzte Projekt merken, das im Vergleich zum ersten verändert ist ... 
            Dim cIX As Integer = anzahlTS - 1
            Dim lastKeptProjekt As clsProjekt = projekthistorie.ElementAt(cIX)



            Dim vIX As Integer = cIX - 1
            Dim vglProjekt As clsProjekt = projekthistorie.ElementAt(cIX)

            If vIX >= 0 Then
                vglProjekt = projekthistorie.ElementAt(vIX)
            End If


            Dim finished As Boolean = (vIX <= 0)
            Dim anzKept As Integer = listToKeep.Count

            Do While Not finished And anzKept < keepAnzVersions

                Do While vglProjekt.isIdenticalTo(lastKeptProjekt) And vIX >= 1
                    vIX = vIX - 1
                    vglProjekt = projekthistorie.ElementAt(vIX)
                Loop

                ' jetzt ist das vglProjekt ungleich dem lastkeptProjekt oder das Ende ist erreicht 
                If vIX <= 0 Then
                    ' end of operation 
                    finished = True
                Else
                    ' falls es Duplikate gibt: das früheste Projekt finden, das identisch zu vglProjekt ist  
                    vIX = vIX - 1
                    Dim memorizeProjekt As clsProjekt = projekthistorie.ElementAt(vIX)

                    Do Until Not memorizeProjekt.isIdenticalTo(vglProjekt) Or vIX = 0
                        vIX = vIX - 1
                        memorizeProjekt = projekthistorie.ElementAt(vIX)
                    Loop

                    If Not memorizeProjekt.isIdenticalTo(vglProjekt) Then
                        ' es wurde ein Unterschied festgestellt 
                        vIX = vIX + 1
                        lastKeptProjekt = projekthistorie.ElementAt(vIX)

                        If Not listToKeep.ContainsKey(projekthistorie.ElementAt(vIX).timeStamp) Then
                            listToKeep.Add(projekthistorie.ElementAt(vIX).timeStamp, "")
                        End If

                        vIX = vIX - 1
                        If vIX = 0 Then
                            finished = True
                        End If

                        vglProjekt = projekthistorie.ElementAt(vIX)
                    Else
                        finished = True
                    End If


                End If

                anzKept = listToKeep.Count
            Loop


            ' jetzt wird die ProjektHistorie um die toKeepVersions erleichtert ...
            Dim errorOccurred As Boolean = False
            For Each kvp As KeyValuePair(Of Date, String) In listToKeep

                Try
                    If projekthistorie.contains(kvp.Key) Then
                        projekthistorie.remove(kvp.Key)
                    Else
                        errorOccurred = True
                    End If
                Catch ex As Exception
                    errorOccurred = True
                End Try


            Next

            If listToKeep.Count < 1 Then
                ' nichts tun ... 
            Else
                ' jetzt wird gelöscht ... 
                ' hier nur bei diesem Projekt weitermachen, wenn kein Fehler aufgetreten ist; das ist sonst zu kritisch 

                If Not errorOccurred Then

                    For Each kvp As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste

                        If Not tsToDelete.Contains(kvp.Key) Then
                            tsToDelete.Add(kvp.Key, kvp.Key)
                        End If

                    Next

                End If
            End If

        End If

        identifyTimeStampsToDelete = tsToDelete

    End Function

    ''' <summary>
    ''' gibt true zurück, wenn diese Projekt-Variante in keinem Portfolio enthalten ist ... 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="variantName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function notReferencedByAnyPortfolio(ByVal pname As String, ByVal variantName As String) As Boolean

        Dim atleastOneReference As Boolean = False

        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            'If kvp.Key = calcLastSessionScenarioName() Or kvp.Key = calcLastEditorScenarioName() Then
            If kvp.Key = calcLastSessionScenarioName() Then
                ' nichts tun , die zählen nicht 
            Else
                Dim pvName As String = calcProjektKey(pname, variantName)
                atleastOneReference = atleastOneReference Or kvp.Value.contains(pvName, False)

                ' wenn es sich um den variantenNAme pfv handelt, dann noch checken, ob die Basis Variante enthalten ist
                ' eine pfv Vorgabe darf nicht gelöscht werden, solange die Basis Variante noch Teil eines Portfolios ist .. 
                If variantName = ptVariantFixNames.pfv.ToString Then
                    pvName = calcProjektKey(pname, "")
                    atleastOneReference = atleastOneReference Or kvp.Value.contains(pvName, False)
                End If
            End If


        Next

        notReferencedByAnyPortfolio = Not atleastOneReference

    End Function



    ''' <summary>
    ''' löscht den angegebenen timestamp von pname#variantname aus der Datenbank
    ''' speichert den timestamp im Papierkorb
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="variantName"></param>
    ''' <param name="timeStamp"></param>
    ''' <param name="first"></param>
    ''' <remarks></remarks>
    Public Sub deleteProjectVariantTimeStamp(ByRef outputCollection As Collection,
                                             ByVal pname As String, ByVal variantName As String,
                                                  ByVal timeStamp As Date, ByRef first As Boolean)

        Dim err As New clsErrorCodeMsg

        Dim outputLine As String = ""


        Dim hproj As clsProjekt

        If first Then
            projekthistorie.clear() ' alte Historie löschen
            projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB _
                                   (projectname:=pname, variantName:=variantName,
                                    storedEarliest:=Date.MinValue, storedLatest:=Date.Now, err:=err)
            first = False
        End If



        hproj = projekthistorie.ElementAtorBefore(timeStamp)

        If DateDiff(DateInterval.Second, timeStamp, hproj.timeStamp) <> 0 Then
            outputLine = "Fehler:" & timeStamp.ToShortDateString & vbLf &
            hproj.timeStamp.ToShortDateString
            outputCollection.Add(outputLine)
            'Call MsgBox("hier ist was faul" & timeStamp.ToShortDateString & vbLf & _
            '             hproj.timeStamp.ToShortDateString)
        End If
        timeStamp = hproj.timeStamp

        If IsNothing(hproj) Then
            outputLine = "Timestamp " & timeStamp.ToShortDateString & vbLf &
                        "zu Projekt " & projekthistorie.First.getShapeText & " nicht gefunden"
            outputCollection.Add(outputLine)
            'Call MsgBox("Timestamp " & timeStamp.ToShortDateString & vbLf & _
            '            "zu Projekt " & projekthistorie.First.getShapeText & " nicht gefunden")

        Else
            ' Speichern im Papierkorb, dann löschen

            If CType(databaseAcc, DBAccLayer.Request).deleteProjectTimestampFromDB(projectname:=pname, variantName:=variantName,
                                  stored:=timeStamp, userName:=dbUsername, err:=err) Then
                'Call MsgBox("ok, gelöscht")
            Else
                outputLine = "Fehler beim Löschen von " & pname & ", " & variantName & ", " &
                              timeStamp.ToShortDateString
                outputCollection.Add(outputLine)
                'Call MsgBox("Fehler beim Löschen von " & pname & ", " & variantName & ", " & _
                '            timeStamp.ToShortDateString)
            End If
            '    Else
            '    ' es ging etwas schief


            '    Call MsgBox("Fehler beim Speichern im Papierkorb:" & vbLf & _
            '                hproj.name & ", " & hproj.timeStamp.ToShortDateString)
            'End If

        End If

    End Sub



    ''' <summary>
    ''' liefert den Status der Basis-Variante zurück; 
    ''' wenn die nicht existiert, wird der übergebene Default Wert zurückgegeben 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getStatusOfBaseVariant(ByVal pName As String, ByVal defaultStatus As String) As String

        Dim err As New clsErrorCodeMsg

        Dim baseVariantProj As clsProjekt = Nothing
        ' wenn es keine bisherige Basis Variante gibt, bleibt der status der Variante erhalten
        Dim baseVariantStatus As String = defaultStatus

        ' den Status der bisherigen Basis-Variante ermitteln

        If Not noDB Then
            baseVariantProj = AlleProjekte.getProject(pName, "")

            If IsNothing(baseVariantProj) Then

                If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, "", Date.Now, err) Then
                    baseVariantProj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(pName, "", "", Date.Now, err)
                    If Not IsNothing(baseVariantProj) Then
                        baseVariantStatus = baseVariantProj.Status
                    Else

                        If awinSettings.visboDebug Then
                            Call MsgBox("BasisVariante kann nicht gefunden werden")
                        End If

                    End If

                End If
            Else
                baseVariantStatus = baseVariantProj.Status
            End If
        End If

        getStatusOfBaseVariant = baseVariantStatus
    End Function


    ''' <summary>
    ''' liefert die Namen aller Projekte im Show, die nicht zum angegebenen Filter passen ...
    ''' </summary>
    ''' <param name="filterName"></param>
    ''' <remarks></remarks>
    Friend Function getProjectNamesNotFittingToFilter(ByVal filterName As String) As Collection

        Dim nameCollection As New Collection
        Dim filter As New clsFilter
        Dim ok As Boolean = False

        Dim todoListe As New Collection


        filter = filterDefinitions.retrieveFilter("Last")

        If IsNothing(filter) Then

            ' nichts tun und Showprojekte bleibt unverändert ... 
        Else

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                If Not filter.isEmpty Then
                    ok = filter.doesNotBlock(kvp.Value)
                Else
                    ok = True
                End If

                If Not ok Then
                    ' aus Showprojekte rausnehmen und Projekt-Tafel aktualisieren 
                    Try
                        nameCollection.Add(kvp.Value.name)
                    Catch ex As Exception

                    End Try
                Else

                End If

            Next

            ' Liste gefüllt mit Projekte, die auf den aktuellen Filter passen

        End If

        getProjectNamesNotFittingToFilter = nameCollection

    End Function

    ''' <summary>
    ''' baut aus der Datenbank die Projekt-Varianten Liste auf, die zu dem gegeb. Zeitpunkt bereits in der Datenbank existiert haben 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Function buildPvNamesList(ByVal storedAtOrBefore As Date, Optional ByVal fromReST As Boolean = False) As SortedList(Of String, String)

        Dim err As New clsErrorCodeMsg

        Dim zeitraumVon As Date = StartofCalendar
        Dim zeitraumbis As Date = StartofCalendar.AddYears(50)

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            ' es ist ein Zeitraum definiert 
            zeitraumVon = getDateofColumn(showRangeLeft, False)
            zeitraumbis = getDateofColumn(showRangeRight, True)
        End If


        buildPvNamesList = CType(databaseAcc, DBAccLayer.Request).retrieveProjectVariantNamesFromDB(zeitraumVon, zeitraumbis, storedAtOrBefore, err, )

    End Function




    ''' <summary>
    ''' wird hauptsächlich benötigt in Verbindung mit updateTreeView und frmProjPortfolioAdmin 
    ''' liefert eine Liste von Varianten-Namen, eingeschlossen in Klammern, die es zu Projekt pName gibt 
    ''' (), (v1), etc..
    ''' </summary>
    ''' <param name="pvNames"></param>
    ''' <param name="pName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getVariantListeFromPVNames(ByVal pvNames As SortedList(Of String, String), ByVal pName As String) As Collection
        Dim tmpResult As New Collection
        Dim vglName As String
        Dim variantName As String

        For Each kvp As KeyValuePair(Of String, String) In pvNames
            Dim tmpStr() As String = kvp.Key.Split(New Char() {CChar("#")})
            vglName = tmpStr(0)
            variantName = "()"

            If vglName = pName Then
                If tmpStr.Length = 1 Then
                    variantName = "()"
                ElseIf tmpStr.Length > 1 Then
                    variantName = "(" & tmpStr(1) & ")"
                End If

                If Not tmpResult.Contains(variantName) Then
                    tmpResult.Add(variantName, variantName)
                End If
            End If


        Next
        getVariantListeFromPVNames = tmpResult

    End Function




    ''' <summary>
    ''' 
    ''' liefert die Liste von Varianten-Namen, die es zu einem vp  mit Name pName oder vpid gibt 
    ''' 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vpid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getVariantListeFromPName(ByVal pName As String, Optional ByVal vpid As String = "",
                                             Optional ByVal vpType As Integer = ptPRPFType.project,
                                             Optional ByVal portfolioliste As SortedList(Of String, String) = Nothing) As Collection
        Dim tmpResult As New Collection
        Dim err As New clsErrorCodeMsg

        If Not IsNothing(portfolioliste) Then
            For Each kvp As KeyValuePair(Of String, String) In portfolioliste
                If kvp.Key.Contains(pName) Then
                    tmpResult.Add(kvp.Value)
                End If
            Next
        Else
            tmpResult = CType(databaseAcc, DBAccLayer.Request).retrieveVariantNamesFromDB(pName, err, vpType)
        End If


        getVariantListeFromPName = tmpResult

    End Function



    ''' <summary>
    ''' Prozedur um Username und Passwort für die Datenbank-Benutzung abzufragen und auch zu testen.
    ''' </summary>
    ''' <remarks></remarks>
    Function loginProzedur() As Boolean


        ' tk, 17.11.16 das wird nicht benötigt, rausgenommen, damit die 
        ' Login Prozedur auch von Powerpoint aus aufgerufen werden kann 
        ' appInstance.EnableEvents = False
        ' enableOnUpdate = False

        Dim loginDialog As New frmAuthentication
        Dim returnValue As DialogResult
        Dim i As Integer = 0



        returnValue = DialogResult.Retry

        ' ur: 30.6.2016: Login-Versuche auf fünf limitiert
        While returnValue = DialogResult.Retry And i < 5

            returnValue = loginDialog.ShowDialog
            i = i + 1

        End While

        If returnValue = DialogResult.Abort Or i >= 5 Then

            Return False
        Else

            Return True
        End If

    End Function


    ''' <summary>
    ''' es wird der LoginProzess angestoßen. Bei erfolgreichem Login wird in den Settings verschlüsselt
    ''' userNamePWD gemerkt. Damit ist es möglich den nächsten Login zu automatisieren
    ''' </summary>
    ''' <param name="noDBAccess"></param>
    ''' <returns>true = erfolgreich</returns>
    Public Function logInToMongoDB(ByVal noDBAccess As Boolean) As Boolean
        ' jetzt die Login Maske aufrufen, aber nur wenn nicht schon ein Login erfolgt ist .. ... 

        If noDBAccess Then
            If awinSettings.databaseURL <> "" Then
                '' ur: 23.03.2020: Angabe von VC nicht mehr nötig, es findet Auswahl statt
                '' If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

                ' jetzt prüfen , ob es bereits gespeicherte User-Credentials gibt 
                If IsNothing(awinSettings.userNamePWD) Then
                    ' tk: 17.11.16: Einloggen in Datenbank 
                    noDBAccess = Not loginProzedur()
                    If Not noDBAccess Then
                        ' in diesem Fall das mySettings setzen 
                        Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
                        awinSettings.userNamePWD = visboCrypto.verschluessleUserPwd(dbUsername, dbPasswort)
                    End If
                Else
                    If awinSettings.userNamePWD = "" Then
                        ' tk: 17.11.16: Einloggen in Datenbank 
                        noDBAccess = Not loginProzedur()

                        If Not noDBAccess Then
                            ' in diesem Fall das mySettings setzen 
                            Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
                            awinSettings.userNamePWD = visboCrypto.verschluessleUserPwd(dbUsername, dbPasswort)
                        End If

                    Else
                        ' die gespeicherten User-Credentials hernehmen, um sich einzuloggen 
                        ' noDBAccess = Not autoVisboLogin(awinSettings.userNamePWD)

                        ' wenn das jetzt nicht geklappt hat, soll wieder das login Fenster kommen ..
                        If noDBAccess Then
                            noDBAccess = Not loginProzedur()

                            If Not noDBAccess Then
                                ' in diesem Fall das mySettings setzen 
                                Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)
                                awinSettings.userNamePWD = visboCrypto.verschluessleUserPwd(dbUsername, dbPasswort)

                            End If

                        End If
                    End If
                End If
            End If
        End If

        logInToMongoDB = Not noDBAccess

    End Function

    ''' <summary>
    ''' Funktion testet die vorhandene Datenbank-authorisierungsinfog
    ''' </summary>
    ''' <param name="user"></param>
    ''' <param name="pwd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function testLoginInfo_OK(ByVal user As String, ByVal pwd As String) As Boolean


        'Dim ok As Boolean = CType(databaseAcc, DBAccLayer.Request).createIndicesOnce()
        Dim ok As Boolean = CType(databaseAcc, DBAccLayer.Request).pingMongoDb()

        testLoginInfo_OK = ok
    End Function



    ''' <summary>
    ''' übergebenene ProjektListe wird um die Projekte reduziert, die nicht zu dem Filter passen
    ''' das wird nur aufgerufen, wenn der Filter angewendet werden soll 
    ''' </summary>
    ''' <param name="projektListe"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function reduzierenWgFilter(ByVal projektListe As clsProjekteAlle) As clsProjekteAlle
        Dim filter As New clsFilter
        Dim ok As Boolean = False
        Dim newProjektliste As New clsProjekteAlle



        ' wenn applyFilter = true, dann soll  unter Anwendung 
        ' des Filters "Last" nachgeladen werden

        filter = filterDefinitions.retrieveFilter("Last")

        If IsNothing(filter) Then

            ' Liste unverändert zurückgeben
            reduzierenWgFilter = projektListe
        Else

            For Each kvp As KeyValuePair(Of String, clsProjekt) In projektListe.liste

                If Not filter.isEmpty Then
                    ok = filter.doesNotBlock(kvp.Value)
                Else
                    ok = True
                End If

                If ok Then
                    Try
                        newProjektliste.Add(kvp.Value, False)
                    Catch ex As Exception
                        Call MsgBox("Fehler in reduzierenWgFilter" & kvp.Key)
                    End Try
                Else

                End If

            Next

            ' Liste gefüllt mit Projekte, die auf den aktuellen Filter passen
            reduzierenWgFilter = newProjektliste
        End If

    End Function




    ''' <summary>
    ''' ruft das Formular auf, um Filter zu definieren
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub defineFilterDB()
        Dim auswahlFormular As New frmNameSelection
        Dim returnValue As DialogResult

        With auswahlFormular

            '.showModePortfolio = True
            .menuOption = PTmenue.filterdefinieren

            '.Show()
            returnValue = .ShowDialog
        End With

    End Sub


    ''' <summary>
    ''' zeichnet das Leistbarkeits-Chart 
    ''' </summary>
    ''' <param name="selCollection">Collection mit den Phasne-, Meilenstein, Rollen- oder Kostenarten</param>
    ''' <param name="chTyp">Typ: es handelt sich um Phasen, rollen, etc. </param>
    ''' <param name="chtop">auf welcher Höhe soll das Chart gezeichnet werden</param>
    ''' <param name="chleft">auf welcher x-Koordinate soll das Chart gezeichnet werden</param>
    ''' <remarks></remarks>
    Friend Sub zeichneLeistbarkeitsChart(ByVal selCollection As Collection, ByVal chTyp As String, ByVal oneChart As Boolean,
                                              ByRef chtop As Double, ByRef chleft As Double, ByVal chwidth As Double, ByVal chHeight As Double)


        Dim repObj As Excel.ChartObject
        Dim myCollection As Collection


        '' Window Position festlegen 
        'chHeight = maxScreenHeight / 4 - 3
        'chWidth = maxScreenWidth / 5 - 3

        'chWidth = 265 + (showRangeRight - showRangeLeft - 12 + 1) * boxWidth + (showRangeRight - showRangeLeft) * screen_correct
        'chHeight = awinSettings.ChartHoehe1


        If oneChart = True Then


            ' alles in einem Chart anzeigen
            myCollection = New Collection
            For Each element As String In selCollection
                myCollection.Add(element, element)
            Next

            repObj = Nothing
            Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                              chwidth, chHeight, False, chTyp, False)


            'chtop = chtop + 7 + chHeight
            chtop = chtop + 2 + chHeight
            'chleft = chleft + 7
        Else
            ' für jedes ITEM ein eigenes Chart machen
            For Each element As String In selCollection
                ' es muss jedesmal eine neue Collection erzeugt werden - die Collection wird in DiagramList gemerkt
                ' wenn die mit Clear leer gemacht wird, funktioniert der Diagram Update nicht mehr ....
                myCollection = New Collection
                myCollection.Add(element, element)
                repObj = Nothing

                Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                   chwidth, chHeight, False, chTyp, False)

                'chtop = chtop + 5
                'chleft = chleft + 7

                chtop = chtop + 2 + chHeight
            Next

        End If

    End Sub

    ''' <summary>
    ''' wird aus Formular NameSelection bzw. HrySelection aufgerufen
    ''' besetzt die Filter-Auswahl Dropbox mit Filternamen aus Datenbank
    ''' </summary>
    ''' <param name="menuOption"></param>
    ''' <param name="filterDropbox"></param>
    ''' <remarks></remarks>
    Public Sub frmHryNameReadFilterVorlagen(ByVal menuOption As Integer, ByRef filterDropbox As System.Windows.Forms.ComboBox)


        ' einlesen und anzeigen der in der Datenbank definierten Filter
        If menuOption = PTmenue.filterdefinieren Then

            If Not noDB Then
                ' Filter mit Namen "fName" in DB speichern


                ' Datenbank ist gestartet
                If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                    Dim listofDBFilter As SortedList(Of String, clsFilter) = CType(databaseAcc, DBAccLayer.Request).retrieveAllFilterFromDB(False)
                    For Each kvp As KeyValuePair(Of String, clsFilter) In listofDBFilter
                        If Not filterDefinitions.Liste.ContainsKey(kvp.Key) Then
                            filterDefinitions.Liste.Add(kvp.Key, kvp.Value)
                        End If
                    Next
                Else
                    Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
                End If
            Else

            End If

        Else
            If menuOption = PTmenue.visualisieren Or
                menuOption = PTmenue.multiprojektReport Or
                menuOption = PTmenue.einzelprojektReport Or
                menuOption = PTmenue.leistbarkeitsAnalyse Then

                If Not noDB Then

                    ' allee Filter aus DB lesen

                    ' Datenbank ist gestartet
                    If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                        Dim listofDBFilter As SortedList(Of String, clsFilter) = CType(databaseAcc, DBAccLayer.Request).retrieveAllFilterFromDB(True)
                        For Each kvp As KeyValuePair(Of String, clsFilter) In listofDBFilter

                            If Not selFilterDefinitions.Liste.ContainsKey(kvp.Key) Then
                                selFilterDefinitions.Liste.Add(kvp.Key, kvp.Value)
                            End If

                        Next
                    Else
                        Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
                    End If

                End If
            End If
        End If


    End Sub

    ''' <summary>
    ''' führt die Aktionen Visualisieren, Leistbarkeit, Meilenstein Trendanalyse aus dem Hierarchie bzw. Namen-Auswahl Fenster durch 
    ''' 
    ''' </summary>
    ''' <param name="menueOption"></param>
    ''' <remarks></remarks>
    Public Sub frmHryNameActions(ByVal menueOption As Integer,
                                 ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection,
                                 ByVal selectedRoles As Collection, ByVal selectedCosts As Collection,
                                 ByVal oneChart As Boolean, ByVal filtername As String)

        Dim chTyp As String
        Dim validOption As Boolean

        If menueOption = PTmenue.visualisieren Or menueOption = PTmenue.einzelprojektReport Or
            menueOption = PTmenue.excelExport Or menueOption = PTmenue.multiprojektReport Or
            menueOption = PTmenue.sessionFilterDefinieren Or menueOption = PTmenue.filterdefinieren Or
            menueOption = PTmenue.vorlageErstellen Or menueOption = PTmenue.meilensteinTrendanalyse Then
            validOption = True
        ElseIf showRangeRight - showRangeLeft >= minColumns - 1 Then
            validOption = True
        Else
            validOption = False
        End If

        If menueOption = PTmenue.leistbarkeitsAnalyse Then

            Dim myCollection As New Collection

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And validOption Then

                Dim formerSU As Boolean = appInstance.ScreenUpdating
                appInstance.ScreenUpdating = False

                ' Window Position festlegen
                'Dim chtop As Double = 50.0 + awinSettings.ChartHoehe1
                'Dim chleft As Double = (showRangeRight - 1) * boxWidth + 4
                Dim chtop As Double
                Dim chleft As Double
                Dim chwidth As Double
                Dim chHeight As Double


                'If visboZustaende.projectBoardMode = ptModus.graficboard Then
                '    chleft = (showRangeRight - 1) * boxWidth + 4
                'Else
                '    chleft = 5
                'End If

                '' um es im neuen Portfolio Chart Window anzuzeigen ... 
                'chtop = 3
                'chleft = 3




                If selectedPhases.Count > 0 Then
                    If awinSettings.considerCategories Then
                        chTyp = DiagrammTypen(7)
                    Else
                        chTyp = DiagrammTypen(0)
                    End If

                    If oneChart Then
                        Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, selectedPhases.Count, chtop, chleft, chwidth, chHeight)
                    Else
                        Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 1, chtop, chleft, chwidth, chHeight)
                    End If
                    Call zeichneLeistbarkeitsChart(selectedPhases, chTyp, oneChart,
                                                   chtop, chleft, chwidth, chHeight)
                End If

                If selectedMilestones.Count > 0 Then
                    If awinSettings.considerCategories Then
                        chTyp = DiagrammTypen(8)
                    Else
                        chTyp = DiagrammTypen(5)
                    End If

                    If oneChart Then
                        Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, selectedMilestones.Count, chtop, chleft, chwidth, chHeight)
                    Else
                        Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 1, chtop, chleft, chwidth, chHeight)
                    End If
                    Call zeichneLeistbarkeitsChart(selectedMilestones, chTyp, oneChart,
                                                   chtop, chleft, chwidth, chHeight)
                End If

                If selectedRoles.Count > 0 Then
                    chTyp = DiagrammTypen(1)

                    If oneChart Then
                        Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, selectedRoles.Count, chtop, chleft, chwidth, chHeight)
                    Else
                        Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 1, chtop, chleft, chwidth, chHeight)
                    End If
                    Call zeichneLeistbarkeitsChart(selectedRoles, chTyp, oneChart,
                                                   chtop, chleft, chwidth, chHeight)
                End If

                If selectedCosts.Count > 0 Then
                    chTyp = DiagrammTypen(2)

                    If oneChart Then
                        Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, selectedCosts.Count, chtop, chleft, chwidth, chHeight)
                    Else
                        Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 1, chtop, chleft, chwidth, chHeight)
                    End If

                    Call zeichneLeistbarkeitsChart(selectedCosts, chTyp, oneChart,
                                                   chtop, chleft, chwidth, chHeight)
                End If


                appInstance.ScreenUpdating = formerSU

            Else

            End If

        ElseIf menueOption = PTmenue.visualisieren Then


            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And validOption Then

                If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0) And
                    (selectedRoles.Count > 0 Or selectedCosts.Count > 0) Then
                    Call MsgBox("es können nur entweder Phasen / Meilensteine oder Rollen oder Kosten angezeigt werden")

                ElseIf selectedPhases.Count > 0 Or selectedMilestones.Count > 0 Then

                    If selectedPhases.Count > 0 Then
                        Call deleteBeschriftungen()
                        Call awinZeichnePhasen(selectedPhases, False, True)

                        ' Selektion der selektierten Projekte wieder sichtbar machen
                        If selectedProjekte.Count > 0 Then
                            Call awinSelect()
                        End If
                    End If

                    If selectedMilestones.Count > 0 Then
                        ' Phasen anzeigen 
                        Dim farbID As Integer = 4
                        Call deleteBeschriftungen()
                        Call awinZeichneMilestones(selectedMilestones, farbID, False, True)

                    End If

                ElseIf selectedRoles.Count > 0 Then

                    Call awinDeleteProjectChildShapes(0)
                    Call deleteBeschriftungen()
                    'Call awinZeichneBedarfe(selectedRoles, DiagrammTypen(1))

                ElseIf selectedCosts.Count > 0 Then

                    Call awinDeleteProjectChildShapes(0)
                    Call deleteBeschriftungen()
                    'Call awinZeichneBedarfe(selectedCosts, DiagrammTypen(2))

                Else
                    Call MsgBox("noch nicht implementiert")
                End If

            Else
                Call MsgBox("bitte mindestens ein Element aus einer der Kategorien selektieren  ")
            End If

            ' selektierte Projekte weiterhin als selektiert darstellen
            If selectedProjekte.Count > 0 Then
                Call awinSelect()
            End If

        ElseIf menueOption = PTmenue.filterdefinieren Then

            'Call MsgBox("ok, Filter gespeichert")

        ElseIf menueOption = PTmenue.sessionFilterDefinieren Then
            ' keine Message ausgeben ...

        ElseIf menueOption = PTmenue.excelExport Or menueOption = PTmenue.vorlageErstellen Then

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0) _
                    And validOption Then

                Try
                    Call createDateiFromSelection(filtername, menueOption)
                    If menueOption = PTmenue.excelExport Then
                        Call MsgBox("ok, Excel File in " & exportOrdnerNames(PTImpExp.rplan) & " erzeugt")
                    Else
                        Call MsgBox("ok, Excel File in " & exportOrdnerNames(PTImpExp.modulScen) & " erzeugt")
                    End If

                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try

            Else
                Call MsgBox("bitte mindestens ein Element aus einer der Kategorien Phasen / Meilensteine selektieren  ")
            End If
        ElseIf menueOption = PTmenue.meilensteinTrendanalyse Then


            If selectedMilestones.Count > 0 Then
                ' Window Position festlegen

                Call awinShowMilestoneTrend(selectedMilestones)
            Else
                Call MsgBox("Bitte Meilensteine auswählen! ")

            End If

        Else

            Call MsgBox("noch nicht unterstützt")

        End If

    End Sub


    Sub awinShowMilestoneTrend(ByVal selectedMilestones As Collection)

        Dim err As New clsErrorCodeMsg

        Dim singleShp As Excel.Shape
        Dim listOfItems As New Collection
        Dim nameList As New SortedList(Of Date, String)
        Dim title As String = "Meilensteine auswählen"
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim selektierteProjekte As New clsProjekte
        Dim top As Double, left As Double, height As Double, width As Double
        Dim repObj As Excel.ChartObject = Nothing

        Dim pName As String, vglName As String = " "
        Dim variantName As String

        Call projektTafelInit()

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                ' eingangs-prüfung, ob auch nur ein Element selektiert wurde ...
                If awinSelection.Count = 1 Then

                    ' Aktion durchführen ...

                    singleShp = awinSelection.Item(1)

                    Try

                        hproj = ShowProjekte.getProject(singleShp.Name, True)
                        nameList = hproj.getMilestones
                        listOfItems = hproj.getElemIdsOf(selectedMilestones, True)


                        ' jetzt muss die ProjektHistorie aufgebaut werden 
                        With hproj
                            pName = .name
                            variantName = .variantName
                        End With

                        If Not projekthistorie Is Nothing Then
                            If projekthistorie.Count > 0 Then
                                vglName = projekthistorie.First.getShapeText
                            End If
                        Else
                            projekthistorie = New clsProjektHistorie
                        End If

                        If vglName <> hproj.getShapeText Then

                            ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                            projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName,
                                                                                storedEarliest:=StartofCalendar, storedLatest:=Date.Now, err:=err)
                            projekthistorie.Add(Date.Now, hproj)


                        Else
                            ' der aktuelle Stand hproj muss hinzugefügt werden 
                            Dim lastElem As Integer = projekthistorie.Count - 1
                            projekthistorie.RemoveAt(lastElem)
                            projekthistorie.Add(Date.Now, hproj)
                        End If



                        With singleShp
                            top = .Top + boxHeight + 5
                            left = .Left - 5
                        End With

                        height = 2 * ((nameList.Count - 1) * 20 + 110)
                        width = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 10, 24 * boxWidth + 10)


                        Call createMsTrendAnalysisOfProject(hproj, repObj, listOfItems, top, left, height, width)


                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                Else
                    Call MsgBox("bitte nur ein Projekt selektieren ...")
                End If
            Else
                Call MsgBox("vorher ein Projekt selektieren ...")
            End If

        Else
            Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
            'projekthistorie.clear()
        End If
        enableOnUpdate = True
        appInstance.EnableEvents = True





    End Sub


    ''' <summary>
    ''' speichert den letzten Filter unter "fname" und setzt die temporären Collections wieder zurück 
    ''' </summary>
    ''' <remarks></remarks>
    '''
    Public Sub storeFilter(ByVal fName As String, ByVal menuOption As Integer,
                                              ByVal fBU As Collection, ByVal fTyp As Collection,
                                              ByVal fPhase As Collection, ByVal fMilestone As Collection,
                                              ByVal fRole As Collection, ByVal fCost As Collection,
                                              ByVal calledFromHry As Boolean)

        Dim lastFilter As clsFilter

        If menuOption = PTmenue.filterdefinieren Or
            menuOption = PTmenue.sessionFilterDefinieren Or
            menuOption = PTmenue.filterAuswahl Then

            ' tk 10.9.18 nicht mehr notwednig 
            ''If calledFromHry Then
            ''    Dim nameLastFilter As clsFilter = filterDefinitions.retrieveFilter("Last")

            ''    If Not IsNothing(nameLastFilter) Then
            ''        With nameLastFilter
            ''            lastFilter = New clsFilter(fName, .BUs, .Typs, fPhase, fMilestone, fRole, fCost)
            ''        End With
            ''    Else
            ''        lastFilter = New clsFilter(fName, fBU, fTyp,
            ''                          fPhase, fMilestone,
            ''                         fRole, fCost)
            ''    End If


            ''Else
            ''    lastFilter = New clsFilter(fName, fBU, fTyp,
            ''                          fPhase, fMilestone,
            ''                         fRole, fCost)
            ''End If

            lastFilter = New clsFilter(fName, fBU, fTyp,
                                      fPhase, fMilestone,
                                     fRole, fCost)

            filterDefinitions.storeFilter(fName, lastFilter)

            If Not noDB Then


                ' Filter mit Namen "fName" in DB speichern

                ' Datenbank ist gestartet
                If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                    Dim filterToStoreInDB As clsFilter = filterDefinitions.retrieveFilter(fName)
                    Dim returnvalue As Boolean = CType(databaseAcc, DBAccLayer.Request).storeFilterToDB(filterToStoreInDB, False)
                    If returnvalue = False Then
                        Call MsgBox("Fehler bei Schreiben Filter: " & fName)
                    End If
                Else
                    Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
                End If


            End If

        Else        ' nicht menuOption = PTmenue.filterdefinieren

            ' tk 10.9.18 nicht mehr notwendig 
            ''If calledFromHry Then
            ''    Dim nameLastFilter As clsFilter = selFilterDefinitions.retrieveFilter("Last")

            ''    If Not IsNothing(nameLastFilter) Then
            ''        With nameLastFilter
            ''            lastFilter = New clsFilter(fName, .BUs, .Typs, fPhase, fMilestone, fRole, fCost)
            ''        End With
            ''    Else
            ''        lastFilter = New clsFilter(fName, fBU, fTyp,
            ''                          fPhase, fMilestone,
            ''                         fRole, fCost)
            ''    End If


            ''Else
            ''    lastFilter = New clsFilter(fName, fBU, fTyp,
            ''                          fPhase, fMilestone,
            ''                         fRole, fCost)
            ''End If

            lastFilter = New clsFilter(fName, fBU, fTyp,
                                      fPhase, fMilestone,
                                     fRole, fCost)
            selFilterDefinitions.storeFilter(fName, lastFilter)

            If Not noDB Then

                ' Filter mit Namen "fName" in DB speichern

                ' Datenbank ist gestartet
                If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                    Dim filterToStoreInDB As clsFilter = selFilterDefinitions.retrieveFilter(fName)
                    Dim returnvalue As Boolean = CType(databaseAcc, DBAccLayer.Request).storeFilterToDB(filterToStoreInDB, True)
                Else
                    Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Filter kann nicht in DB gespeichert werden")
                End If

            End If

        End If

    End Sub

    ''' <summary>
    ''' löscht das angegebene Projekt mit Name pName inkl all seiner Varianten 
    ''' </summary>
    ''' <param name="pName">
    ''' gibt an , ob es der erste Aufruf war
    ''' wenn ja, kommt erst der Bestätigungs-Dialog 
    ''' wenn nein, wird ohne Aufforderung zur Bestätigung gelöscht 
    ''' </param>
    ''' <remarks></remarks>
    Public Sub awinDeleteProjectInSession(ByVal pName As String,
                                          Optional ByVal considerDependencies As Boolean = False,
                                          Optional ByVal upDateDiagrams As Boolean = False,
                                          Optional ByVal vName As String = Nothing)


        Dim hproj As clsProjekt

        Dim tmpCollection As New Collection

        Dim formerEOU As Boolean = enableOnUpdate
        enableOnUpdate = False


        If ShowProjekte.contains(pName) Then

            ' Aktuelle Konstellation ändert sich dadurch
            If currentConstellationPvName <> calcLastSessionScenarioName() Then
                currentConstellationPvName = calcLastSessionScenarioName()
            End If

            hproj = ShowProjekte.getProject(pName)
            If IsNothing(vName) Or vName = hproj.variantName Then
                Call putProjectInNoShow(hproj.name, considerDependencies, upDateDiagrams)
            End If


        End If

        ' jetzt müssen alle oder die ausgewählte Variante aus AlleProjekte gelöscht werden 
        If IsNothing(vName) Then
            AlleProjekte.RemoveAllVariantsOf(pName)
        Else
            Dim key As String = calcProjektKey(pName, vName)
            If AlleProjekte.Containskey(key) Then
                AlleProjekte.Remove(key)
            End If
        End If

        enableOnUpdate = formerEOU

    End Sub


    ''' <summary>
    ''' nimmt das angegebene Projekt aus ShowProjekte heraus
    ''' löscht das Projekt auf der Plan-Tafel und schicbt die restlichen Projekte weiter nach oben 
    ''' wenn considerDependencies=true: dann werden alle abhängigen Projekte, die ebenfalls im ShowProjekte sind, auch rausgenommen
    ''' wenn upDateDiagrams=true: alle Diagramme werde neu gezeichnet  
    ''' 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="considerDependencies"></param>
    ''' <param name="upDateDiagrams"></param>
    ''' <remarks></remarks>
    Public Sub putProjectInNoShow(ByVal pName As String, ByVal considerDependencies As Boolean, ByVal upDateDiagrams As Boolean)

        Dim pZeile As Integer
        Dim tmpCollection As New Collection
        Dim anzahlZeilen As Integer = 1

        If ShowProjekte.contains(pName) Then

            Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
            pZeile = calcYCoordToZeile(projectboardShapes.getCoord(pName)(0))

            If hproj.extendedView Then
                anzahlZeilen =
                    hproj.calcNeededLines(tmpCollection, tmpCollection, awinSettings.drawphases Or hproj.extendedView, False)
            End If

            'pZeile = ShowProjekte.getPTZeile(selectedProjectName)
            'Call MsgBox("Zeile: " & pZeile.ToString)

            Call clearProjektinPlantafel(pName)

            ShowProjekte.Remove(pName)

            Call moveShapesUp(pZeile + 1, anzahlZeilen, True)

        End If

        ' jetzt muss noch geprüft werden , ob considerDependencies true ist 
        If considerDependencies Then
            ' ggf. die Projekte einblenden, von denen dieses Projekt abhängt 
            Dim toDoListe As Collection = allDependencies.activeListe(pName, PTdpndncyType.inhalt)
            If toDoListe.Count > 0 Then
                For Each dprojectName As String In toDoListe
                    Call putProjectInNoShow(dprojectName, considerDependencies, False)
                Next

            End If
        Else
            ' nichts tun 
        End If

        If upDateDiagrams Then
            ' jetzt müssen die Portfolio Diagramme neu gezeichnet werden 
            Call awinNeuZeichnenDiagramme(2)
        End If


    End Sub

    ''' <summary>
    ''' bringt die angegebene Projekt-Variante ins Show ... 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vNAme"></param>
    ''' <param name="considerDependencies"></param>
    ''' <param name="upDateDiagrams"></param>
    ''' <remarks></remarks>
    Public Sub putProjectInShow(ByVal pName As String, ByVal vName As String,
                                    ByVal considerDependencies As Boolean,
                                    ByVal upDateDiagrams As Boolean,
                                    ByVal myConstellation As clsConstellation,
                                    Optional ByVal parentChoice As Boolean = False,
                                    Optional pZeile As Integer = -1)

        Dim key As String = calcProjektKey(pName, vName)
        Dim hproj As clsProjekt = AlleProjekte.getProject(key)


        If IsNothing(hproj) And parentChoice Then
            Dim variantNames As Collection = AlleProjekte.getVariantNames(pName, False)
            vName = CStr(variantNames.Item(1))
            key = calcProjektKey(pName, vName)
            hproj = AlleProjekte.getProject(key)
        End If

        ' wenn immer noch Nothing, nichts tun ... 
        If IsNothing(hproj) Then
            Exit Sub
        End If

        Dim anzahlZeilen As Integer = 1

        If Not ShowProjekte.contains(pName) Then
            ShowProjekte.Add(hproj)
            If pZeile < 2 Then
                'pZeile = ShowProjekte.getPTZeile(pName)
                pZeile = myConstellation.getBoardZeile(pName)
            End If

            Dim tmpCollection As New Collection

            If hproj.extendedView Then
                anzahlZeilen =
                    hproj.calcNeededLines(tmpCollection, tmpCollection, awinSettings.drawphases Or hproj.extendedView, False)
            End If

            If pZeile > 0 Then
                Call moveShapesDown(tmpCollection, pZeile, anzahlZeilen, 0)
                Call ZeichneProjektinPlanTafel(tmpCollection, pName, pZeile, tmpCollection, tmpCollection)
            End If
        End If

        ' jetzt muss das Projekt neu gezeichnet werden ; 
        ' dazu muss die Einfügestelle bestimmt werden, dann alle anderen Shapes nach unten verschoben werden 
        ' hier muss die Zeile über Showprojekte bestimmt werden, einfach nach der Sortier-Reihenfolge 
        ' das kann später dann noch angepasst werden 

        'Dim pZeile2 As Integer = node.Index
        'Call MsgBox("Zeile: " & pZeile.ToString)



        ' jetzt muss noch geprüft werden , ob considerDependencies true ist 
        If considerDependencies Then
            ' ggf. die Projekte einblenden, von denen dieses Projekt abhängt 
            Dim toDoListe As Collection = allDependencies.passiveListe(pName, PTdpndncyType.inhalt)
            If toDoListe.Count > 0 Then
                For Each mprojectName As String In toDoListe
                    Call putProjectInShow(pName:=mprojectName,
                                          vName:="", considerDependencies:=considerDependencies,
                                          upDateDiagrams:=False,
                                          myConstellation:=myConstellation, parentChoice:=True)
                Next

            End If
        Else
            ' nichts tun 
        End If

        If upDateDiagrams Then
            ' jetzt müssen die Portfolio Diagramme neu gezeichnet werden 
            Call awinNeuZeichnenDiagramme(2)
        End If

    End Sub


    Public Sub retrieveProfilSelection(ByVal profilName As String, ByVal menuOption As Integer,
                                     ByRef selectedBUs As Collection, ByRef selectedTyps As Collection,
                                     ByRef selectedPhases As Collection, ByRef selectedMilestones As Collection,
                                     ByRef selectedRoles As Collection, ByRef selectedCosts As Collection, ByRef reportProfil As clsReportAll)
        Try
            If menuOption = PTmenue.reportBHTC Then

                ' Datumsangaben sichern
                Dim vondate_sav As Date = reportProfil.VonDate
                Dim bisdate_sav As Date = reportProfil.BisDate
                Dim PPTvondate_sav As Date = reportProfil.CalendarVonDate
                Dim PPTbisdate_sav As Date = reportProfil.CalendarBisDate

                ' Projekte sichern
                Dim projects_sav As New SortedList(Of Double, String)
                For Each kvp As KeyValuePair(Of Double, String) In reportProfil.Projects
                    projects_sav.Add(kvp.Key, kvp.Value)
                Next

                ' Datumsangaben zurücksichern
                reportProfil.CalendarVonDate = PPTvondate_sav
                reportProfil.CalendarBisDate = PPTbisdate_sav
                reportProfil.calcRepVonBis(vondate_sav, bisdate_sav)



                ' für BHTC immer true
                reportProfil.ExtendedMode = True
                ' für BHTC immer false
                reportProfil.Ampeln = False
                reportProfil.AllIfOne = False
                reportProfil.FullyContained = False
                reportProfil.SortedDauer = False
                reportProfil.ProjectLine = False
                reportProfil.UseOriginalNames = False

                ' Projekte zurücksichern
                reportProfil.Projects.Clear()
                For Each kvp As KeyValuePair(Of Double, String) In projects_sav
                    reportProfil.Projects.Add(kvp.Key, kvp.Value)
                Next

            Else
                '  menuOption = PTmenue.reportMultiprojektTafel


                ' Einlesen des ausgewählten ReportProfils
                reportProfil = XMLImportReportProfil(profilName)


            End If


            '  und bereitstellen der Auswahl für Hierarchieselection
            selectedPhases = copySortedListtoColl(reportProfil.Phases)
            selectedMilestones = copySortedListtoColl(reportProfil.Milestones)
            selectedRoles = copySortedListtoColl(reportProfil.Roles)
            selectedCosts = copySortedListtoColl(reportProfil.Costs)
            selectedBUs = copySortedListtoColl(reportProfil.BUs)
            selectedTyps = copySortedListtoColl(reportProfil.Typs)

        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Lesen des ReportProfils: retrieveProfilSelection")
        End Try

    End Sub


    Public Sub storeReportProfil(ByVal menuOption As Integer,
                                     ByVal selectedBUs As Collection, ByVal selectedTyps As Collection,
                                     ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection,
                                     ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, ByVal reportProfil As clsReportAll)



        '  und bereitstellen der Auswahl für Hierarchieselection
        reportProfil.Phases = copyColltoSortedList(selectedPhases)
        reportProfil.Milestones = copyColltoSortedList(selectedMilestones)
        reportProfil.Roles = copyColltoSortedList(selectedRoles)
        reportProfil.Costs = copyColltoSortedList(selectedCosts)
        reportProfil.BUs = copyColltoSortedList(selectedBUs)
        reportProfil.Typs = copyColltoSortedList(selectedTyps)


        With awinSettings

            ' tk : wird für Darstellung Projekt auf Multiprojekt Tafel verwendet; hier nicht setzen ! 
            '.drawProjectLine = True
            reportProfil.ExtendedMode = .mppExtendedMode
            reportProfil.OnePage = .mppOnePage
            reportProfil.AllIfOne = .mppShowAllIfOne
            reportProfil.Ampeln = .mppShowAmpel
            reportProfil.Legend = .mppShowLegend
            reportProfil.MSDate = .mppShowMsDate
            reportProfil.MSName = .mppShowMsName
            reportProfil.PhDate = .mppShowPhDate
            reportProfil.PhName = .mppShowPhName
            reportProfil.ProjectLine = .mppShowProjectLine
            reportProfil.SortedDauer = .mppSortiertDauer
            reportProfil.VLinien = .mppVertikalesRaster
            reportProfil.FullyContained = .mppFullyContained
            reportProfil.ShowHorizontals = .mppShowHorizontals
            reportProfil.UseAbbreviation = .mppUseAbbreviation
            reportProfil.UseOriginalNames = .mppUseOriginalNames
            reportProfil.KwInMilestone = .mppKwInMilestone

            If menuOption = PTmenue.reportMultiprojektTafel Then
                reportProfil.projectsWithNoMPmayPass = .mppProjectsWithNoMPmayPass
            Else
                ' dann gilt: menuOption = PTmenue.reportBHTC
                reportProfil.projectsWithNoMPmayPass = Nothing
                reportProfil.description = ""
            End If

        End With


        ' Schreiben des ausgewählten ReportProfils
        Call XMLExportReportProfil(reportProfil)

    End Sub
    ''' <summary>
    ''' versucht das Projekt für mich zu schützen 
    ''' gibt false zurück , wenn das Projekt durch andere geschützt ist 
    ''' </summary>
    ''' <param name="pname"></param>
    ''' <param name="vName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function tryToprotectProjectforMe(ByVal pName As String, ByVal vName As String) As Boolean

        Dim err As New clsErrorCodeMsg

        Dim wpItem As clsWriteProtectionItem
        Dim isProtectedbyOthers As Boolean

        If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, Date.Now, err) Then

            ' es existiert in der Datenbank ...
            If CType(databaseAcc, DBAccLayer.Request).checkChgPermission(pName, vName, dbUsername, err) Then

                isProtectedbyOthers = False
                ' jetzt prüfen, ob es Null ist, von mir permanent/nicht permanent geschützt wurde .. 
                wpItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(pName, vName, err)

                Dim notYetDone As Boolean = False

                If IsNothing(wpItem) Then
                    ' wpitem kann NULL sein
                    notYetDone = True

                ElseIf wpItem.permanent Then
                    notYetDone = False
                    ' meinen permanenten Schutz einbauen 
                    writeProtections.upsert(wpItem)

                Else
                    notYetDone = True
                End If

                If notYetDone Then
                    wpItem = New clsWriteProtectionItem(calcProjektKey(pName, vName),
                                                              ptWriteProtectionType.project,
                                                              dbUsername,
                                                              False,
                                                              True)

                    If CType(databaseAcc, DBAccLayer.Request).setWriteProtection(wpItem, err) Then
                        ' erfolgreich ...
                        writeProtections.upsert(wpItem)
                    Else
                        ' in diesem Fall wurde es in der Zwischenzeit von jdn anders geschützt  
                        isProtectedbyOthers = True
                    End If

                End If

            Else
                isProtectedbyOthers = True
            End If
        Else
            ' das Projekt existiert bisher nur in der Session des Nutzers 
            isProtectedbyOthers = False
        End If


        tryToprotectProjectforMe = Not isProtectedbyOthers

    End Function
    ''' <summary>
    ''' löscht die bedingte Farb-Codierung 
    ''' </summary>
    Public Sub deleteColorFormatMassEdit()
        If CType(appInstance.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.meRC) Then

            Try
                Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
                                                        .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
                Dim mahleRange As Excel.Range = meWS.Range("MahleInfo")

                If Not IsNothing(mahleRange) Then

                    ' die bedingte Farb-Codierung ausschalten 
                    With mahleRange
                        Do While .FormatConditions.Count > 0
                            .FormatConditions.Item(1).delete
                        Loop
                    End With

                End If

            Catch ex As Exception

            End Try

        Else
            ' einfach nichts machen ..
        End If
    End Sub

    ''' <summary>
    ''' stellt sicher, dass die Prozent- bzw. Frei-Tage Werte im Falle meExtendedView = true mit einer entsprechenden Color-Codierung dargestellt werden 
    ''' 
    ''' </summary>
    Public Sub colorFormatMassEditRC()
        If CType(appInstance.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.meRC) Then

            Try
                Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
                                                        .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
                Dim mahleRange As Excel.Range = meWS.Range("MahleInfo")

                If Not IsNothing(mahleRange) Then
                    ' die bedingte Farb-Codierung einschalten 
                    If awinSettings.mePrzAuslastung Then
                        With mahleRange

                            Do While .FormatConditions.Count > 0
                                .FormatConditions.Item(1).delete
                            Loop

                            Dim przColorScale As Excel.ColorScale = .FormatConditions.AddColorScale(3)

                            CType(przColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                            CType(przColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).Value = 0
                            CType(przColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeGreen

                            CType(przColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                            CType(przColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).Value = 1.1
                            CType(przColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeYellow

                            CType(przColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                            CType(przColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).Value = 1.5
                            CType(przColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeRed

                        End With
                    Else
                        With mahleRange

                            Do While .FormatConditions.Count > 0
                                .FormatConditions.Item(1).delete
                            Loop

                            Dim przColorScale As Excel.ColorScale = .FormatConditions.AddColorScale(3)

                            CType(przColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                            CType(przColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).Value = -5
                            CType(przColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeRed

                            CType(przColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                            CType(przColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).Value = 0
                            CType(przColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeYellow

                            CType(przColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                            CType(przColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).Value = 5
                            CType(przColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeGreen

                        End With
                    End If
                End If

            Catch ex As Exception

            End Try

        Else
            ' einfach nichts machen ..
        End If

    End Sub
    ''' <summary>
    ''' nur zu verwenden, wenn AutoReduce falsch 
    ''' dann aber wesentlich schneller als mit autoReduce 
    ''' aktualisiert nur in der aktuellen Zeile die Summe 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="cphase"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="rcNameID"></param>
    ''' <param name="isRole"></param>
    ''' <param name="zeile"></param>
    Public Sub updateMassEditSummenValue(ByVal hproj As clsProjekt, ByVal cphase As clsPhase,
                                              ByVal von As Integer, ByVal bis As Integer,
                                              ByVal rcNameID As String,
                                              ByVal isRole As Boolean,
                                              ByVal zeile As Integer)

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        Dim columnSummen As Integer = visboZustaende.meColRC + 2
        Dim columnRC As Integer = visboZustaende.meColRC
        Dim tmpSum As Double = 0.0

        If CType(appInstance.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.meRC) Then
            ' nur dann befindet sich das Programm im MassEdit Sheet 

            'Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
            Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

            If rcNameID.Trim.Length = 0 Then
                ' nichts tun 

            Else
                ' Update Lauf der Summen 

                If Not IsNothing(hproj) And Not IsNothing(cphase) Then

                    Dim xWerte() As Double
                    Dim ixZeitraum As Integer
                    Dim ix As Integer
                    Dim anzLoops As Integer

                    ' diese MEthode definiert, wo der Zeitraum sich mit den Werte überlappt ... 
                    ' Anzloops sind die Anzahl Überlappungen 
                    Call awinIntersectZeitraum(getColumnOfDate(cphase.getStartDate), getColumnOfDate(cphase.getEndDate),
                                       ixZeitraum, ix, anzLoops)

                    If isRole Then
                        If RoleDefinitions.isValidCombination(rcNameID) Then
                            Dim tmpRole As clsRolle = cphase.getRoleByRoleNameID(rcNameID)

                            If Not IsNothing(tmpRole) Then
                                xWerte = tmpRole.Xwerte

                                ' jetzt werden die Werte summiert ...
                                Try
                                    For al As Integer = 1 To anzLoops
                                        tmpSum = tmpSum + xWerte(ix + al - 1)
                                    Next
                                Catch ex As Exception
                                    Call MsgBox("Fehler bei Summenbildung ...")
                                    tmpSum = 0
                                End Try


                            Else
                                ' Summe löschen
                            End If
                        Else
                            ' Summe löschen
                        End If

                    Else
                        Dim costName As String = rcNameID
                        If CostDefinitions.containsName(costName) Then
                            Dim tmpCost As clsKostenart = cphase.getCost(costName)

                            If Not IsNothing(tmpCost) Then
                                xWerte = tmpCost.Xwerte

                                ' jetzt werden die Werte summiert ...
                                Try
                                    For al As Integer = 1 To anzLoops
                                        tmpSum = tmpSum + xWerte(ix + al - 1)
                                    Next
                                Catch ex As Exception
                                    Call MsgBox("Fehler bei Summenbildung ...")
                                    tmpSum = 0
                                End Try

                            Else
                                ' Summe löschen
                            End If
                        Else
                            ' Summe löschen
                        End If

                    End If

                Else
                    ' Summe löschen 
                End If
            End If

            ' jetzt den Wert in die Zelle schreiben
            If tmpSum > 0 Then
                CType(meWS.Cells(zeile, columnSummen), Excel.Range).Value = tmpSum
                '.ToString("#,##0")
            Else
                CType(meWS.Cells(zeile, columnSummen), Excel.Range).Value = ""
            End If

        Else
            Call MsgBox("Mass-Edit Sheet nicht aktiv ...")
        End If

        appInstance.EnableEvents = formerEE


    End Sub

    ' tk 5.5. wird nicht mehr benötigt 
    '''' <summary>
    '''' aktualisiert die Summen-Werte im Massen-Edit Sheet der Ressourcen-/Kostenzuordnungen  
    '''' </summary>
    '''' <param name="pname"></param>
    '''' <param name="von"></param>
    '''' <param name="bis"></param>
    '''' <param name="roleCostNames"></param>
    '''' <remarks></remarks>
    'Public Sub updateMassEditSummenValues(ByVal pname As String, ByVal phaseNameID As String,
    '                                          ByVal von As Integer, ByVal bis As Integer,
    '                                          ByVal roleCostNames As Collection)


    '    Dim formerEE As Boolean = appInstance.EnableEvents
    '    appInstance.EnableEvents = False

    '    If CType(appInstance.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.meRC) Then
    '        ' nur dann befindet sich das Programm im MassEdit Sheet 

    '        'Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
    '        Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
    '        .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

    '        If IsNothing(roleCostNames) Then
    '            ' nichts tun 
    '        ElseIf roleCostNames.Count = 0 Then
    '            ' nichts tun 
    '        Else
    '            ' Update Lauf der Summen 
    '            Dim columnSummen As Integer = visboZustaende.meColRC + 1
    '            Dim columnRC As Integer = visboZustaende.meColRC

    '            ' jetzt muss einfach jede Zeile im Mass-Edit Sheet durchgegangen werden 
    '            For zeile As Integer = 2 To visboZustaende.meMaxZeile

    '                Dim curpName As String = CStr(meWS.Cells(zeile, 2).value)
    '                Dim curphaseName As String = CStr(meWS.Cells(zeile, 4).value)
    '                Dim curphaseNameID As String = calcHryElemKey(curphaseName, False)
    '                Dim curComment As Excel.Comment = CType(meWS.Cells(zeile, 4), Excel.Range).Comment
    '                If Not IsNothing(curComment) Then
    '                    curphaseNameID = curComment.Text
    '                End If

    '                ' es soll auf jeden Fall auch die Rootphase geupdated werden ..., da ja die evtl auch als secondbest geändert wurde ...
    '                If curpName = pname And ((curphaseNameID = phaseNameID) Or (curphaseNameID = rootPhaseName)) Then

    '                    Dim curRCName As String = CStr(meWS.Cells(zeile, columnRC).value)

    '                    If Not IsNothing(curRCName) Then
    '                        If curRCName.Trim.Length > 0 Then
    '                            If roleCostNames.Contains(curRCName) Then
    '                                Dim tmpSum As Double = 0.0
    '                                ' jetzt muss die Summe aktualisiert werden 
    '                                Dim hproj As clsProjekt = ShowProjekte.getProject(pname)
    '                                If Not IsNothing(hproj) Then
    '                                    Dim cphase As clsPhase = hproj.getPhaseByID(curphaseNameID)

    '                                    If Not IsNothing(cphase) Then

    '                                        Dim xWerte() As Double
    '                                        Dim ixZeitraum As Integer
    '                                        Dim ix As Integer
    '                                        Dim anzLoops As Integer

    '                                        ' diese MEthode definiert, wo der Zeitraum sich mit den Werte überlappt ... 
    '                                        ' Anzloops sind die Anzahl Überlappungen 
    '                                        Call awinIntersectZeitraum(getColumnOfDate(cphase.getStartDate), getColumnOfDate(cphase.getEndDate),
    '                                                           ixZeitraum, ix, anzLoops)

    '                                        If RoleDefinitions.containsName(curRCName) Then

    '                                            Dim tmpRole As clsRolle = cphase.getRole(curRCName)

    '                                            If Not IsNothing(tmpRole) Then
    '                                                xWerte = tmpRole.Xwerte

    '                                                ' jetzt werden die Werte summiert ...
    '                                                Try
    '                                                    For al As Integer = 1 To anzLoops
    '                                                        tmpSum = tmpSum + xWerte(ix + al - 1)
    '                                                    Next
    '                                                Catch ex As Exception
    '                                                    Call MsgBox("Fehler bei Summenbildung ...")
    '                                                    tmpSum = 0
    '                                                End Try


    '                                            Else
    '                                                ' Summe löschen
    '                                            End If

    '                                        ElseIf CostDefinitions.containsName(curRCName) Then

    '                                            Dim tmpCost As clsKostenart = cphase.getCost(curRCName)

    '                                            If Not IsNothing(tmpCost) Then
    '                                                xWerte = tmpCost.Xwerte

    '                                                ' jetzt werden die Werte summiert ...
    '                                                Try
    '                                                    For al As Integer = 1 To anzLoops
    '                                                        tmpSum = tmpSum + xWerte(ix + al - 1)
    '                                                    Next
    '                                                Catch ex As Exception
    '                                                    Call MsgBox("Fehler bei Summenbildung ...")
    '                                                    tmpSum = 0
    '                                                End Try

    '                                            Else
    '                                                ' Summe löschen
    '                                            End If
    '                                        Else
    '                                            ' Summe löschen 
    '                                        End If

    '                                    Else
    '                                        ' Summe löschen  
    '                                    End If
    '                                Else
    '                                    ' Summe löschen 
    '                                End If

    '                                ' jetzt den Wert in die Zelle schreiben
    '                                If tmpSum > 0 Then
    '                                    CType(meWS.Cells(zeile, columnSummen), Excel.Range).Value = tmpSum
    '                                    '.ToString("#,##0")
    '                                Else
    '                                    CType(meWS.Cells(zeile, columnSummen), Excel.Range).Value = ""
    '                                End If

    '                            End If
    '                        End If
    '                    End If

    '                End If

    '            Next

    '        End If


    '    Else
    '        Call MsgBox("Mass-Edit Sheet nicht aktiv ...")
    '    End If

    '    appInstance.EnableEvents = formerEE


    'End Sub

    ''' <summary>
    ''' aktualisiert für das angegebene Projekt die Validation Strings aller leeren / empty RoleCost Felder gemäß dem übergebenen 
    ''' dient dazu, um die Validation an der rootPhaseName Setzung zu orientieren 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="validationString"></param>
    ''' <remarks></remarks>
    Public Sub updateEmptyRcCellValidations(ByVal pName As String, ByVal validationString As String)

        ' tk 16.9.18 wird nicht mehr benötigt 
        'Dim formerEE As Boolean = appInstance.EnableEvents
        'appInstance.EnableEvents = False

        'If CType(appInstance.ActiveSheet, Excel.Worksheet).Name = arrWsNames(ptTables.meRC) Then
        '    ' nur dann befindet sich das Programm im MassEdit Sheet 

        '    Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
        '    .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

        '    If IsNothing(pName) Or IsNothing(validationString) Then
        '        ' nichts tun 
        '    ElseIf pName.Trim.Length = 0 Or validationString.Trim.Length = 0 Then
        '        ' nichts tun 
        '    Else
        '        ' Update der Validations der leeren RoleCost Zuordnungen  
        '        Dim columnRC As Integer = visboZustaende.meColRC

        '        ' jetzt muss einfach jede Zeile im Mass-Edit Sheet durchgegangen werden 
        '        For zeile As Integer = 2 To visboZustaende.meMaxZeile

        '            Dim curpName As String = CStr(meWS.Cells(zeile, 2).value)
        '            Dim curphaseName As String = CStr(meWS.Cells(zeile, 4).value)
        '            Dim needsUpdate As Boolean = False

        '            ' es soll auf jeden Fall auch die Rootphase geupdated werden ..., da ja die evtl auch als secondbest geändert wurde ...
        '            If curpName = pName And curphaseName <> "." Then

        '                Dim curRCName As String = CStr(meWS.Cells(zeile, columnRC).value)

        '                If IsNothing(curRCName) Then
        '                    needsUpdate = True
        '                ElseIf curRCName.Trim.Length = 0 Then
        '                    needsUpdate = True
        '                End If

        '            End If

        '            If needsUpdate Then
        '                Try
        '                    With CType(meWS.Cells(zeile, columnRC), Excel.Range)
        '                        If Not IsNothing(.Validation) Then
        '                            .Validation.Delete()
        '                        End If
        '                        '' ur: 28.09.2017

        '                        ' '' jetzt wird die ValidationList aufgebaut 

        '                        ''.Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop, _
        '                        ''                               Formula1:=validationString)

        '                    End With
        '                Catch ex As Exception

        '                End Try

        '            End If

        '        Next

        '    End If


        'Else
        '    Call MsgBox("Mass-Edit Sheet nicht aktiv ...")
        'End If

        'appInstance.EnableEvents = formerEE

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="phaseNameID"></param>
    ''' <param name="rcNameID"></param>
    ''' <returns>0, wenn nichts gefunden wird</returns>
    Public Function findeZeileInMeRC(ByVal ws As Excel.Worksheet,
                                     ByVal pName As String,
                                     ByVal phaseNameID As String,
                                     ByVal rcNameID As String) As Integer

        Dim tmpResult As Integer = 0
        Dim zeile As Integer = 2

        Dim colPName As Integer = visboZustaende.meColpName
        Dim colPhaseName As Integer = visboZustaende.meColRC - 1
        Dim colRcName As Integer = visboZustaende.meColRC
        Dim found As Boolean = False

        Dim vglPname As String = CStr(CType(ws.Cells(zeile, colPName), Excel.Range).Value)
        Dim vglRcNameID As String = getRCNameIDfromExcelRange(CType(ws.Range(ws.Cells(zeile, colRcName), ws.Cells(zeile, colRcName + 1)), Excel.Range))
        Dim vglPhaseNameID As String = getPhaseNameIDfromExcelCell(CType(ws.Cells(zeile, colPhaseName), Excel.Range))


        Do While zeile <= visboZustaende.meMaxZeile And Not found

            If rcNameID = "*" Then
                ' Wildcard für rcNameID
                found = vglPname = pName And
                        vglPhaseNameID = phaseNameID
            Else
                ' rcNameID ist mit entscheidend
                found = vglPname = pName And
                        vglPhaseNameID = phaseNameID And
                        vglRcNameID = rcNameID
            End If

            If Not found Then
                zeile = zeile + 1
                vglPname = CStr(CType(ws.Cells(zeile, colPName), Excel.Range).Value)
                vglRcNameID = getRCNameIDfromExcelRange(CType(ws.Range(ws.Cells(zeile, colRcName), ws.Cells(zeile, colRcName + 1)), Excel.Range))
                vglPhaseNameID = getPhaseNameIDfromExcelCell(CType(ws.Cells(zeile, colPhaseName), Excel.Range))
            End If

        Loop

        If found Then
            tmpResult = zeile
        End If

        findeZeileInMeRC = tmpResult

    End Function

    ''' <summary>
    ''' gibt eine Zeile zurück, die zu dem angegebenen Projekt, der Phase und dem rcName eine Sammelrolle zurückgibt
    ''' Wenn mehrere mögliche Sammelrollen existieren, dann wird die erste auftretende zurückgegeben
    ''' 0, wenn in diesem Projekt zu dieser Rolle keine Sammelrolle definiert ist  
    ''' ausserdem wird erst in der gleichen Phase gesucht, oder aber in der RootPhase
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="phaseNameID"></param>
    ''' <param name="rcName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function findeSammelRollenZeile(ByVal pName As String, ByVal phaseNameID As String, ByVal rcName As String) As Integer
        Dim found As Boolean = False
        Dim curZeile As Integer = 2

        Dim chckName As String
        Dim chckPhNameID As String
        Dim chckRCName As String
        Dim bestName As String = ""
        Dim secondBestzeile As Integer = 0

        Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoledef(rcName)

        If Not IsNothing(tmpRole) Then

            Dim istSammelRolle As Boolean = tmpRole.isCombinedRole

            If istSammelRolle Then
                curZeile = 0
            Else
                ' nur für echte Rollen durchführen ...
                Dim potentialParentRoles As Collection = RoleDefinitions.getSummaryRoles(rcName)

                If potentialParentRoles.Count = 0 Then
                    curZeile = 0

                Else
                    '
                    ' auf die Suche gehen ... 
                    Dim meWS As Excel.Worksheet = CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
                    .Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

                    With meWS
                        chckName = CStr(meWS.Cells(curZeile, 2).value)
                        If IsNothing(chckName) Then
                            chckName = ""
                        End If

                        Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
                        If IsNothing(phaseName) Then
                            phaseName = ""
                        End If

                        chckPhNameID = calcHryElemKey(phaseName, False)
                        Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
                        If Not IsNothing(curComment) Then
                            chckPhNameID = curComment.Text
                        End If

                        chckRCName = CStr(meWS.Cells(curZeile, 5).value)
                        If IsNothing(chckRCName) Then
                            chckRCName = ""
                        End If

                    End With
                    ' 
                    ' jetzt wird erst geprüft, ob es eine Sammelrolle in der gleichen Phase gibt 
                    ' dann wird geprüft , ob es eine Sammelrolle in der rootphase gibt 



                    Do While Not found And curZeile <= visboZustaende.meMaxZeile


                        If ((chckName = pName) And
                            ((phaseNameID = chckPhNameID) Or (rootPhaseName = chckPhNameID))) Then

                            If potentialParentRoles.Contains(chckRCName) Then
                                ' nimm jetzt einfach mal den ersten, der auftritt  ... 
                                If phaseNameID = chckPhNameID Then
                                    found = True
                                Else
                                    secondBestzeile = curZeile
                                    ' noch weitersuchen, ob nicht noch das found-Kriterium greift ... 
                                End If

                            End If

                        End If

                        If Not found Then

                            curZeile = curZeile + 1

                            With meWS
                                chckName = CStr(meWS.Cells(curZeile, 2).value)
                                If IsNothing(chckName) Then
                                    chckName = ""
                                End If

                                Dim phaseName As String = CStr(meWS.Cells(curZeile, 4).value)
                                If IsNothing(phaseName) Then
                                    phaseName = ""
                                End If

                                chckPhNameID = calcHryElemKey(phaseName, False)
                                Dim curComment As Excel.Comment = CType(meWS.Cells(curZeile, 4), Excel.Range).Comment
                                If Not IsNothing(curComment) Then
                                    chckPhNameID = curComment.Text
                                End If

                                chckRCName = CStr(meWS.Cells(curZeile, 5).value)
                                If IsNothing(chckRCName) Then
                                    chckRCName = ""
                                End If

                            End With

                        End If

                    Loop

                End If

            End If

        End If




        If found Then
            findeSammelRollenZeile = curZeile
        ElseIf secondBestzeile > 0 Then
            findeSammelRollenZeile = secondBestzeile
        Else
            findeSammelRollenZeile = 0
        End If

    End Function



    ''' <summary>
    ''' Erstellen eines Powerpoint-Reports auf Grund von einem ReportProfil, TimeRange, DB Zugriff, und ausgewählte EinzelProjekte oder Konstellationen
    ''' </summary>
    ''' <param name="projekte">Projektname oder Konstellationsname</param>
    ''' <param name="variante">Variante eines Projektes</param>
    ''' <param name="profilname">Name des Reportprofils</param>
    ''' <param name="vonDate">von Zeit</param>
    ''' <param name="bisDate">bis Zeit</param>
    ''' <param name="reportname">Name des Report (wie abgespeichert werden soll)</param>
    ''' <param name="dbUsername">DB User</param>
    ''' <param name="dbPassword">DB pwd</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function reportErstellen(ByVal projekte As String, ByVal variante As String, ByVal profilname As String, ByVal timestamp As Date,
                                        ByVal vonDate As Date, ByVal bisDate As Date, ByVal reportname As String, ByVal append As Boolean,
                                        ByVal dbUsername As String, ByVal dbPassword As String) As Boolean

        Dim err As New clsErrorCodeMsg

        Dim currentPresentationName As String = ""

        Dim reportProfil As clsReportAll = XMLImportReportProfil(profilname)
        Dim zeilenhoehe As Double = 0.0     ' zeilenhöhe muss für alle Projekte gleich sein, daher mit übergeben
        Dim legendFontSize As Single = 0.0  ' FontSize der Legenden der Schriftgröße des Projektnamens angepasst

        Dim selectedPhases As New Collection
        Dim selectedMilestones As New Collection
        Dim selectedRoles As New Collection
        Dim selectedCosts As New Collection
        Dim selectedBUs As New Collection
        Dim selectedTypes As New Collection

        reportErstellen = False

        selectedPhases = copySortedListtoColl(reportProfil.Phases)
        selectedMilestones = copySortedListtoColl(reportProfil.Milestones)
        selectedRoles = copySortedListtoColl(reportProfil.Roles)
        selectedCosts = copySortedListtoColl(reportProfil.Costs)
        selectedBUs = copySortedListtoColl(reportProfil.BUs)
        selectedTypes = copySortedListtoColl(reportProfil.Typs)

        With awinSettings

            .mppExtendedMode = reportProfil.ExtendedMode
            .mppOnePage = reportProfil.OnePage
            .mppShowAllIfOne = reportProfil.AllIfOne
            .mppShowAmpel = reportProfil.Ampeln
            .mppShowLegend = reportProfil.Legend
            .mppShowMsDate = reportProfil.MSDate
            .mppShowMsName = reportProfil.MSName
            .mppShowPhDate = reportProfil.PhDate
            .mppShowPhName = reportProfil.PhName
            .mppShowProjectLine = reportProfil.ProjectLine
            .mppSortiertDauer = reportProfil.SortedDauer
            .mppVertikalesRaster = reportProfil.VLinien
            .mppFullyContained = reportProfil.FullyContained
            .mppShowHorizontals = reportProfil.ShowHorizontals
            .mppUseAbbreviation = reportProfil.UseAbbreviation
            .mppUseOriginalNames = reportProfil.UseOriginalNames
            .mppKwInMilestone = reportProfil.KwInMilestone
            .mppProjectsWithNoMPmayPass = reportProfil.projectsWithNoMPmayPass

        End With

        If Not (IsNothing(vonDate) Or vonDate = Date.MinValue) Then
            showRangeLeft = getColumnOfDate(vonDate)
        Else
            showRangeLeft = 0
        End If
        If Not (IsNothing(bisDate) Or bisDate = Date.MinValue) Then
            showRangeRight = getColumnOfDate(bisDate)
        Else
            showRangeRight = 0
        End If


        Try
            If Not reportProfil.isMpp Then

                Try


                    Dim vorlagendateiname As String = awinPath & RepProjectVorOrdner & "\" & reportProfil.PPTTemplate
                    If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

                        'Das gewählte Projekt reporten

                        Dim hproj As New clsProjekt
                        hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(projekte, variante, "", timestamp, err)

                        If Not IsNothing(hproj) Then

                            Dim key As String = calcProjektKey(hproj)

                            ' diese Liste wird benötigt, damit in zeichneMultiprojektSicht die Routine bestimmeProjekteAndMinMaxDates funktioniert
                            If Not AlleProjekte.Containskey(calcProjektKey(hproj)) Then
                                AlleProjekte.Add(hproj)
                            End If


                            If Not ShowProjekte.contains(hproj.name) Then  ' akt. Projekt nicht in ShowProjekte
                                ShowProjekte.Add(hproj)

                            Else   ' es ist eventuell nicht die richtige Variante enthalten
                                If ShowProjekte.getProject(hproj.name).variantName <> hproj.variantName Then
                                    ShowProjekte.Remove(hproj.name)
                                    ShowProjekte.Add(hproj)
                                End If
                            End If

                            Call createPPTSlidesFromProject(hproj, vorlagendateiname,
                                                        selectedPhases, selectedMilestones,
                                                        selectedRoles, selectedCosts,
                                                        selectedBUs, selectedTypes, True,
                                                        True, zeilenhoehe, legendFontSize,
                                                        Nothing, Nothing)


                            Dim pptApp As Microsoft.Office.Interop.PowerPoint.Application = Nothing
                            Try
                                ' prüft, ob bereits Powerpoint geöffnet ist 
                                pptApp = CType(GetObject(, "PowerPoint.Application"), Microsoft.Office.Interop.PowerPoint.Application)
                            Catch ex As Exception
                                Try
                                    pptApp = CType(CreateObject("PowerPoint.Application"), Microsoft.Office.Interop.PowerPoint.Application)

                                Catch ex1 As Exception
                                    Call MsgBox("Powerpoint konnte nicht gestartet werden ..." & ex1.Message)
                                    reportErstellen = False
                                    Exit Function
                                End Try

                            End Try
                            ' aktive Präsentation unter angegebenem Namen "reportname" abspeichern
                            Dim currentPraesi As Microsoft.Office.Interop.PowerPoint.Presentation = pptApp.ActivePresentation

                            If reportname = "" Then
                                Dim aktDate As String = Date.Now.ToString
                                reportname = aktDate & "Report.pptx"
                                Call logger(ptErrLevel.logInfo, "EinzelprojektReport mit ' " & projekte & "/" & variante & "/" &
                                                      profilname & "/ ... wurde in " & reportname & "ersatzweise gespeichert", "reportErstellen", anzFehler)
                            Else
                                reportname = reportname & ".pptx"
                            End If

                            If My.Computer.FileSystem.FileExists(reportOrdnerName & reportname) And append Then

                                ' die Seiten 2 - ende der vorhandenen Powerpoint-Datei müssen in das currentPraesi eingefügt werden
                                Dim oldPraesi As Microsoft.Office.Interop.PowerPoint.Presentation = pptApp.Presentations.Open(reportOrdnerName & reportname)
                                Dim anzoldSlides As Integer = oldPraesi.Slides.Count
                                oldPraesi.Close()

                                currentPraesi.Slides.InsertFromFile(FileName:=reportOrdnerName & reportname, Index:=1, SlideStart:=2, SlideEnd:=anzoldSlides)
                                currentPraesi.SaveAs(reportOrdnerName & reportname)
                                currentPraesi.Close()
                            Else
                                'If My.Computer.FileSystem.FileExists(reportOrdnerName & reportname & ".pptx") Then
                                '    My.Computer.FileSystem.DeleteFile(reportOrdnerName & reportname & ".pptx")
                                'End If
                                currentPraesi.SaveAs(reportOrdnerName & reportname)
                                currentPraesi.Close()


                            End If

                            reportErstellen = True
                        Else

                            Call logger(ptErrLevel.logError, "reportErstellen", "Projekt '" & projekte & "' existiert nicht in DB!", anzFehler)

                        End If
                    Else
                        Call logger(ptErrLevel.logError, "reportErstellen", "Vorlagendatei " & vorlagendateiname & " existiert nicht!", anzFehler)
                    End If

                Catch ex As Exception

                End Try

            Else    ' isMPP

                Try

                    If Not (showRangeLeft > 0 And showRangeRight > showRangeLeft) Then

                        showRangeLeft = getColumnOfDate(reportProfil.VonDate)
                        showRangeRight = getColumnOfDate(reportProfil.BisDate)

                    End If

                    Dim hproj As New clsProjekt
                    Dim constellations As New clsConstellations
                    '' '????ur
                    constellations = CType(databaseAcc, DBAccLayer.Request).retrieveConstellationsFromDB(timestamp, err)
                    If Not IsNothing(constellations) Then

                        Dim curconstellation As clsConstellation = constellations.getConstellation(projekte)

                        If Not IsNothing(curconstellation) Then

                            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In curconstellation.Liste

                                hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(kvp.Value.projectName, kvp.Value.variantName, "", timestamp, err)

                                If Not IsNothing(hproj) Then

                                    Dim key As String = calcProjektKey(hproj)

                                    ' diese Liste wird benötigt, damit in zeichneMultiprojektSicht die Routine bestimmeProjekteAndMinMaxDates funktioniert
                                    If Not AlleProjekte.Containskey(calcProjektKey(hproj)) Then
                                        AlleProjekte.Add(hproj)
                                    End If

                                    If kvp.Value.show Then

                                        If Not ShowProjekte.contains(hproj.name) Then
                                            ShowProjekte.Add(hproj)
                                        Else
                                            If ShowProjekte.getProject(hproj.name).variantName <> kvp.Value.variantName Then
                                                ShowProjekte.Remove(kvp.Value.projectName)
                                                ShowProjekte.Add(hproj)
                                            End If
                                        End If

                                    End If
                                Else

                                    Call logger(ptErrLevel.logError, "reportErstellen", "Projekt '" & kvp.Value.projectName & " mit TimeStamp '" & timestamp.ToString & "' existiert nicht in DB!", anzFehler)

                                End If  ' if hproj existiert
                            Next


                            Dim vorlagendateiname As String = awinPath & RepPortfolioVorOrdner & "\" & reportProfil.PPTTemplate
                            If My.Computer.FileSystem.FileExists(vorlagendateiname) Then

                                Call createPPTSlidesFromConstellation(vorlagendateiname,
                                                                      selectedPhases, selectedMilestones,
                                                                      selectedRoles, selectedCosts,
                                                                      selectedBUs, selectedTypes, True,
                                                                      Nothing, Nothing)

                                Dim pptApp As Microsoft.Office.Interop.PowerPoint.Application = Nothing
                                Try
                                    ' prüft, ob bereits Powerpoint geöffnet ist 
                                    pptApp = CType(GetObject(, "PowerPoint.Application"), Microsoft.Office.Interop.PowerPoint.Application)
                                Catch ex As Exception
                                    Try
                                        pptApp = CType(CreateObject("PowerPoint.Application"), Microsoft.Office.Interop.PowerPoint.Application)

                                    Catch ex1 As Exception
                                        Call MsgBox("Powerpoint konnte nicht gestartet werden ..." & ex1.Message)
                                        reportErstellen = False
                                        Exit Function
                                    End Try

                                End Try

                                ' aktive Präsentation unter angegebenem Namen "reportname" abspeichern
                                Dim currentPraesi As Microsoft.Office.Interop.PowerPoint.Presentation = pptApp.ActivePresentation

                                If reportname = "" Then
                                    Dim aktDate As String = Date.Now.ToString
                                    reportname = aktDate & "MP Report.pptx"
                                    Call logger(ptErrLevel.logInfo, "MulitprojektReport mit ' " & projekte & "/" &
                                                          profilname & "/ ... wurde in " & reportname & "ersatzweise gespeichert", "reportErstellen", anzFehler)
                                Else
                                    reportname = reportname & ".pptx"
                                End If

                                If My.Computer.FileSystem.FileExists(reportOrdnerName & reportname) And append Then

                                    ' die Seiten 2 - ende der vorhandenen Powerpoint-Datei müssen in das currentPraesi eingefügt werden
                                    Dim oldPraesi As Microsoft.Office.Interop.PowerPoint.Presentation = pptApp.Presentations.Open(reportOrdnerName & reportname)
                                    Dim anzoldSlides As Integer = oldPraesi.Slides.Count
                                    oldPraesi.Close()

                                    currentPraesi.Slides.InsertFromFile(FileName:=reportOrdnerName & reportname, Index:=1, SlideStart:=2, SlideEnd:=anzoldSlides)
                                    currentPraesi.SaveAs(reportOrdnerName & reportname)
                                    currentPraesi.Close()
                                Else
                                    'If My.Computer.FileSystem.FileExists(reportOrdnerName & reportname & ".pptx") Then
                                    '    My.Computer.FileSystem.DeleteFile(reportOrdnerName & reportname & ".pptx")
                                    'End If
                                    currentPraesi.SaveAs(reportOrdnerName & reportname)
                                    currentPraesi.Close()


                                End If

                                reportErstellen = True

                            End If
                        Else
                            Call logger(ptErrLevel.logError, "reportErstellen", "angegebene Constellation nicht in der DB", anzFehler)

                        End If

                    Else
                        Call logger(ptErrLevel.logError, "reportErstellen", "keine Constellations in der DB vorhanden", anzFehler)
                    End If

                Catch ex As Exception

                End Try

            End If



        Catch ex As Exception

            Call MsgBox("Fehler: " & vbLf & ex.Message)

            reportErstellen = False
        End Try

        'pptApp.Quit()

    End Function


    ''' <summary>
    ''' behandelt die Missing Definitions, nimmt ggf in 
    ''' </summary>
    ''' <param name="definitionName"></param>
    ''' <param name="isVorlage"></param>
    ''' <param name="isMilestone"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isMissingDefinitionOK(ByVal definitionName As String, ByVal isVorlage As Boolean, ByVal isMilestone As Boolean) As Boolean

        Dim checkResult As Boolean = True

        If isMilestone And Not MilestoneDefinitions.Contains(definitionName) Then
            ' Behandlung Meilenstein Definition, aber nur wenn nicht enthalten ... 

            Dim hMilestoneDef As New clsMeilensteinDefinition

            With hMilestoneDef
                .name = definitionName
                .belongsTo = ""
                .shortName = ""
                .darstellungsKlasse = ""
                .UID = MilestoneDefinitions.Count + 1
            End With

            If (isVorlage And awinSettings.alwaysAcceptTemplateNames) Or
                awinSettings.addMissingPhaseMilestoneDef Then
                ' in die Milestone-Definitions aufnehmen 
                checkResult = True
                Try
                    If Not MilestoneDefinitions.Contains(hMilestoneDef.name) Then
                        MilestoneDefinitions.Add(hMilestoneDef)
                        ' wird gesetzt, damit am Ende klar ist, ob irgendwelche Phasen hinzugefügt wurden
                        ' somit kann der OrgaAdmin gefragt werden, ob die in der VCSetting-Customization in der DB gespeichert werden sollen
                        MilestoneDefsAndPhaseDefsAdded = True
                    End If
                Catch ex As Exception
                    checkResult = False
                End Try

            Else


                ' in die Missing Milestone-Definitions aufnehmen 
                Try
                    ' das Element aufnehmen, in Abhängigkeit vom Setting 
                    If awinSettings.importUnknownNames Then
                        checkResult = True
                    Else
                        checkResult = False
                    End If

                    If Not missingMilestoneDefinitions.Contains(hMilestoneDef.name) Then
                        missingMilestoneDefinitions.Add(hMilestoneDef)
                    End If

                Catch ex As Exception
                End Try
            End If

        ElseIf Not isMilestone And Not (PhaseDefinitions.Contains(definitionName)) Then

            ' Behandlung Phasen 
            Dim hphaseDef As clsPhasenDefinition
            hphaseDef = New clsPhasenDefinition

            hphaseDef.darstellungsKlasse = ""
            hphaseDef.shortName = ""
            hphaseDef.name = definitionName
            hphaseDef.UID = PhaseDefinitions.Count + 1



            If (isVorlage And awinSettings.alwaysAcceptTemplateNames) Or
                awinSettings.addMissingPhaseMilestoneDef Then
                ' in die Phase-Definitions aufnehmen 
                checkResult = True
                Try
                    If Not PhaseDefinitions.Contains(hphaseDef.name) Then
                        PhaseDefinitions.Add(hphaseDef)
                        ' wird gesetzt, damit am Ende klar ist, ob irgendwelche Phasen hinzugefügt wurden
                        ' somit kann der OrgaAdmin gefragt werden, ob die in der VCSetting-Customization in der DB gespeichert werden sollen
                        MilestoneDefsAndPhaseDefsAdded = True
                    End If
                Catch ex As Exception
                    checkResult = False
                End Try
            Else
                ' in Abhängigkeit vom Setting die Elemente aufnehmen oder nicht 
                Try
                    If awinSettings.importUnknownNames Then
                        checkResult = True
                    Else
                        checkResult = False
                    End If

                    If Not missingPhaseDefinitions.Contains(hphaseDef.name) Then
                        missingPhaseDefinitions.Add(hphaseDef)
                    End If

                Catch ex As Exception
                    checkResult = False
                End Try


            End If
        End If

        isMissingDefinitionOK = checkResult

    End Function


    ''' <summary>
    ''' setzt die Projekt-Historie für das angegebene Projekt
    ''' wenn nicht existiert, wird Projekt-Historie auf Nothing gesetzt 
    ''' </summary>
    ''' <param name="pName"></param>
    ''' <param name="vName"></param>
    ''' <remarks></remarks>
    Public Sub setProjektHistorie(ByVal pName As String, ByVal vName As String)

        Dim err As New clsErrorCodeMsg

        Dim holeHistory As Boolean = True
        Dim vglProj As clsProjekt = Nothing

        If Not projekthistorie Is Nothing Then
            If projekthistorie.Count > 0 Then
                vglProj = projekthistorie.First
            End If
        End If

        If Not noDB Then

            If Not IsNothing(vglProj) Then
                If vglProj.name = pName And vglProj.variantName = vName Then
                    holeHistory = False
                End If
            End If


            If holeHistory Then

                If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
                    Try
                        If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, Date.Now, err) Then
                            projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:=vName,
                                                                        storedEarliest:=Date.MinValue, storedLatest:=Date.Now, err:=err)
                        Else
                            projekthistorie.clear()
                        End If

                    Catch ex As Exception
                        projekthistorie.clear()
                    End Try
                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Database Connection failed ...")
                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
                    End If

                End If
            End If

        End If
    End Sub

    ''' <summary>
    ''' aktualisiert mit dem selektierten Projekt die evtl angezeigten Projekt-Info Charts
    ''' replaceProj = false, wenn die Skalierung nicht angepasst werden soll; also z.Bsp bei Aufruf aus Time-Machine 
    ''' </summary>
    ''' <param name="hproj">das selektierte Projekt</param>
    ''' <remarks></remarks>
    Public Sub aktualisiereCharts(ByVal hproj As clsProjekt, ByVal replaceProj As Boolean,
                                  Optional ByVal calledFromMassEdit As Boolean = False,
                                  Optional ByVal currentRCName As String = "")

        ' Validieren ...
        If IsNothing(currentRCName) Then
            currentRCName = ""
        End If

        ' tk neu ab 18.1.20
        Dim currentRoleNameID As String = ""
        Dim rcID As Integer = -1
        Dim teamID As Integer = -1
        Dim isRole As Boolean = False
        Dim isCost As Boolean = False

        Dim potentialParents() As Integer = Nothing
        Dim q2StillNeedsToBeDefined As Boolean = False

        ' wenn currentRCName = "" dann soll Gesamtkosten gezeigt werden ... 
        If currentRCName <> "" Then
            rcID = RoleDefinitions.parseRoleNameID(currentRCName, teamID)
            currentRoleNameID = RoleDefinitions.bestimmeRoleNameID(rcID, teamID)

            If rcID < 0 Then
                ' vorauss ist es eine Kostenart 
                If CostDefinitions.containsName(currentRCName) Then
                    isCost = True
                    rcID = CostDefinitions.getCostdef(currentRCName).UID
                Else
                    ' jetzt werdne einfach dei Gesamtkosten gezeigt 
                    currentRCName = ""
                End If
            Else
                isRole = True
            End If
        End If



        Dim err As New clsErrorCodeMsg

        Dim chtobj As Excel.ChartObject

        Dim vglName As String = hproj.name.Trim
        Dim founddiagram As New clsDiagramm
        ' ''Dim IDkennung As String

        Dim currentWsName As String
        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentWsName = arrWsNames(ptTables.mptPrCharts)
        Else
            currentWsName = arrWsNames(ptTables.meCharts)
        End If

        ' aktualisieren der Window Caption ...
        Try
            If visboWindowExists(PTwindows.mptpr) Then
                Dim tmpmsg As String = hproj.getShapeText & " (" & hproj.timeStamp.ToString & ")"
                projectboardWindows(PTwindows.mptpr).Caption = bestimmeWindowCaption(PTwindows.mptpr, addOnMsg:=tmpmsg)
            End If
        Catch ex As Exception

        End Try


        If Not (hproj Is Nothing) Then

            ' bei Projekten, egal ob standard Projekt oder Portfolio Projekt wird immer mit der Vorlagen-Variante verglichen
            Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString
            'If hproj.projectType = ptPRPFType.portfolio Then
            '    tmpVariantName = portfolioVName
            'End If

            With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(currentWsName), Excel.Worksheet)
                Dim tmpArray() As String
                Dim anzDiagrams As Integer
                anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count

                If anzDiagrams > 0 Then
                    For i = 1 To anzDiagrams
                        chtobj = CType(.ChartObjects(i), Excel.ChartObject)
                        If chtobj.Name <> "" Then
                            tmpArray = chtobj.Name.Split(New Char() {CType("#", Char)}, 5)
                            ' chtobj name ist aufgebaut: pr#PTprdk.kennung#pName#Auswahl
                            If tmpArray(0) = "pr" Then

                                Dim chartTyp As String = ""
                                Dim typID As Integer = -1
                                Dim auswahl As Integer = -1
                                Dim chartPname As String = ""
                                Dim roleCostName As String = ""
                                Call getChartKennungen(chtobj.Name, chartTyp, typID, auswahl, chartPname, roleCostName)

                                If calledFromMassEdit And typID <> PTprdk.Ergebnis Then
                                    roleCostName = currentRCName

                                    ' sicherstellen, dass es bei massEdit nur Kostenbalken2 sein kann 
                                    If typID = PTprdk.KostenBalken Then
                                        typID = PTprdk.KostenBalken2
                                    End If

                                    '' mit welchem  soll verglichen werden ?  
                                    'If awinSettings.meCompareWithLastVersion Then
                                    '    typID = PTprdk.KostenBalken2
                                    'Else
                                    '    typID = PTprdk.KostenBalken
                                    'End If

                                    If roleCostName = "" Then
                                        ' damit werden die Gesamtkosten gezeigt ..
                                        auswahl = 2
                                    End If
                                End If

                                Dim scInfo As New clsSmartPPTChartInfo
                                With scInfo
                                    .hproj = hproj
                                    .detailID = typID

                                    ' Setzung von .q2 wird ggf in Kostenbalken, Kostenbalken2, Personabalken und Personalbalken2 noch mal revidiert ..
                                    .q2 = roleCostName

                                    If visboZustaende.projectBoardMode = ptModus.graficboard Then

                                        If typID = PTprdk.KostenBalken Or typID = PTprdk.KostenBalken2 Or
                                            typID = PTprdk.KostenPie Then

                                            If typID = PTprdk.KostenBalken Then
                                                .vergleichsTyp = PTVergleichsTyp.erster
                                            End If

                                            If isCost Then
                                                .elementTyp = ptElementTypen.costs
                                            Else
                                                If auswahl = 1 And roleCostName = "" Then
                                                    .elementTyp = ptElementTypen.costs

                                                ElseIf auswahl = 2 And roleCostName = "" Then
                                                    .elementTyp = ptElementTypen.rolesAndCost
                                                End If
                                            End If



                                            .einheit = PTEinheiten.euro


                                        ElseIf typID = PTprdk.PersonalBalken Or typID = PTprdk.PersonalBalken2 Or
                                            typID = PTprdk.PersonalPie Then

                                            If typID = PTprdk.PersonalBalken Then
                                                .vergleichsTyp = PTVergleichsTyp.erster
                                            End If

                                            .elementTyp = ptElementTypen.roles

                                            If auswahl = 1 Then
                                                .einheit = PTEinheiten.personentage
                                            Else
                                                .einheit = PTEinheiten.euro
                                            End If


                                        ElseIf typID = PTprdk.Ergebnis Then
                                            .elementTyp = ptElementTypen.ergebnis
                                        End If

                                    Else
                                        If visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                                            .einheit = PTEinheiten.euro
                                            .elementTyp = ptElementTypen.costs

                                        ElseIf visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
                                            .einheit = PTEinheiten.personentage
                                            .elementTyp = ptElementTypen.roles

                                        Else
                                            .einheit = PTEinheiten.euro
                                            .elementTyp = ptElementTypen.rolesAndCost
                                        End If

                                        If awinSettings.meCompareVsLastPlan Then
                                            .vergleichsArt = PTVergleichsArt.planungsstand
                                            .vergleichsTyp = PTVergleichsTyp.standVom
                                            .vergleichsDatum = awinSettings.meDateForLastPlan
                                        Else
                                            .vergleichsArt = PTVergleichsArt.beauftragung
                                            .vergleichsTyp = PTVergleichsTyp.letzter
                                            .vergleichsDatum = Date.Now
                                        End If

                                    End If


                                End With

                                If replaceProj Or (chartPname.Trim = vglName) Then
                                    Select Case typID


                                        ' replaceProj sorgt in den nachfolgenden Sequenzen dafür, daß das Chart im Falle eines Aufrufes aus der 
                                        ' Time-Machine (replaceProj = false) nicht in der Skalierung angepasst wird; das geschieht initial beim Laden der Time-Machine
                                        ' wenn es aus dem Selektieren von Projekten aus aufgerufen wird, dann wird die optimal passende Skalierung schon jedesmal berechnet 

                                        Case PTprdk.Phasen
                                            ' Update Phasen Diagramm

                                            If CInt(tmpArray(3)) = PThis.current Then
                                                ' nur dann muss aktualisiert werden ...
                                                Call updatePhasesBalken(hproj, chtobj, auswahl, replaceProj)
                                            End If

                                        Case PTprdk.KostenBalken
                                            Dim vglProj As clsProjekt = Nothing


                                            Try
                                                vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)
                                            Catch ex As Exception
                                                vglProj = Nothing
                                            End Try



                                            scInfo.vergleichsTyp = PTVergleichsTyp.erster
                                            scInfo.vglProj = vglProj

                                            ' tk neu 18.1.2020
                                            ' 
                                            If myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektleitungRestricted Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                                ' 
                                                potentialParents = RoleDefinitions.getIDArray(myCustomUserRole.specifics)



                                            ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then

                                                Try
                                                    Dim allRoleIDsOfP As SortedList(Of Integer, Boolean) = vglProj.getRoleIDs()
                                                    Dim checkID As Integer

                                                    If teamID > 0 Then
                                                        checkID = teamID
                                                    Else
                                                        ' die potentiellen Parents sind die Parents der Kostenstelle / des Teams
                                                        checkID = rcID
                                                    End If

                                                    ' die potentiellen Parents sind  die Parents des Teams
                                                    Dim tmpList As Integer() = RoleDefinitions.getParentArray(RoleDefinitions.getRoleDefByID(checkID), includingMySelf:=True)
                                                    Dim ergListe As New List(Of Integer)
                                                    Dim found As Boolean = False
                                                    Dim ix As Integer = 1

                                                    Do While Not found And ix <= tmpList.Length
                                                        If allRoleIDsOfP.ContainsKey(tmpList(ix - 1)) Then
                                                            found = True
                                                        Else
                                                            ix = ix + 1
                                                        End If
                                                    Loop

                                                    If found Then
                                                        scInfo.q2 = RoleDefinitions.getRoleDefByID(tmpList(ix - 1)).name
                                                    Else
                                                        scInfo.q2 = RoleDefinitions.getRoleDefByID(rcID).name
                                                    End If

                                                Catch ex As Exception
                                                    scInfo.q2 = RoleDefinitions.getRoleDefByID(rcID).name
                                                End Try


                                                potentialParents = Nothing

                                            End If


                                            If Not IsNothing(potentialParents) Then

                                                Dim tmpParentName As String = ""

                                                If teamID = -1 Then
                                                    tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                Else
                                                    Dim tmpTeamName As String = RoleDefinitions.getRoleDefByID(teamID).name
                                                    tmpParentName = RoleDefinitions.chooseParentFromList(tmpTeamName, potentialParents)
                                                    If tmpParentName = "" Then
                                                        tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                    Else
                                                        Dim tmpParentNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpParentName, "")
                                                        If Not IsNothing(vglProj) Then
                                                            If vglProj.containsRoleNameID(tmpParentNameID) Then
                                                                ' passt bereits 
                                                            Else
                                                                tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                            End If
                                                        Else
                                                            tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                        End If
                                                    End If
                                                End If

                                                If tmpParentName <> "" Then
                                                    scInfo.q2 = tmpParentName
                                                End If

                                            End If

                                            ' Ende tk neu 18.1.20

                                            Call updateExcelChartOfProject(scInfo, chtobj, replaceProj, calledFromMassEdit)

                                        Case PTprdk.KostenBalken2
                                            Dim vglProj As clsProjekt = Nothing

                                            Try
                                                If (visboZustaende.projectBoardMode = ptModus.massEditRessSkills Or visboZustaende.projectBoardMode = ptModus.massEditCosts) And awinSettings.meCompareVsLastPlan Then
                                                    Dim vpID As String = ""
                                                    vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, vpID, awinSettings.meDateForLastPlan, err)
                                                Else
                                                    If scInfo.vergleichsTyp = PTVergleichsTyp.erster Then
                                                        scInfo.vergleichsTyp = PTVergleichsTyp.erster
                                                        vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)
                                                    Else
                                                        scInfo.vergleichsTyp = PTVergleichsTyp.letzter
                                                        vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, Date.Now, err)
                                                    End If

                                                End If

                                            Catch ex As Exception
                                                vglProj = Nothing
                                            End Try

                                            scInfo.vglProj = vglProj

                                            ' now define scinfo.q2 bestimmen
                                            ' if it is lastPlan , then scinfo.q2 just is rcNAme
                                            If awinSettings.meCompareVsLastPlan And (visboZustaende.projectBoardMode = ptModus.massEditRessSkills Or visboZustaende.projectBoardMode = ptModus.massEditCosts) Then
                                                scInfo.q2 = roleCostName
                                            Else

                                                ' tk neu 18.1.2020
                                                If myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Or
                                                   myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektleitungRestricted Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                                    ' 
                                                    potentialParents = RoleDefinitions.getIDArray(myCustomUserRole.specifics)

                                                ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then

                                                    Try
                                                        Dim allRoleIDsOfP As SortedList(Of Integer, Boolean) = vglProj.getRoleIDs()
                                                        Dim checkID As Integer

                                                        If teamID > 0 Then
                                                            checkID = teamID
                                                        Else
                                                            ' die potentiellen Parents sind die Parents der Kostenstelle / des Teams
                                                            checkID = rcID
                                                        End If

                                                        ' die potentiellen Parents sind  die Parents des Teams
                                                        Dim tmpList As Integer() = RoleDefinitions.getParentArray(RoleDefinitions.getRoleDefByID(checkID), includingMySelf:=True)
                                                        Dim ergListe As New List(Of Integer)
                                                        Dim found As Boolean = False
                                                        Dim ix As Integer = 1

                                                        Do While Not found And ix <= tmpList.Length
                                                            If allRoleIDsOfP.ContainsKey(tmpList(ix - 1)) Then
                                                                found = True
                                                            Else
                                                                ix = ix + 1
                                                            End If
                                                        Loop

                                                        If found Then
                                                            scInfo.q2 = RoleDefinitions.getRoleDefByID(tmpList(ix - 1)).name
                                                        Else
                                                            scInfo.q2 = RoleDefinitions.getRoleDefByID(rcID).name
                                                        End If

                                                    Catch ex As Exception
                                                        scInfo.q2 = RoleDefinitions.getRoleDefByID(rcID).name
                                                    End Try


                                                    potentialParents = Nothing
                                                End If


                                                If Not IsNothing(potentialParents) Then

                                                    Dim tmpParentName As String = ""

                                                    If teamID = -1 Then
                                                        tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                    Else
                                                        Dim tmpTeamName As String = RoleDefinitions.getRoleDefByID(teamID).name
                                                        tmpParentName = RoleDefinitions.chooseParentFromList(tmpTeamName, potentialParents)
                                                        If tmpParentName = "" Then
                                                            tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                        Else
                                                            Dim tmpParentNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpParentName, "")
                                                            If Not IsNothing(vglProj) Then
                                                                If vglProj.containsRoleNameID(tmpParentNameID) Then
                                                                    ' passt bereits 
                                                                Else
                                                                    tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                                End If
                                                            Else
                                                                tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                            End If
                                                        End If
                                                    End If

                                                    If tmpParentName <> "" Then
                                                        scInfo.q2 = tmpParentName
                                                    End If

                                                End If

                                            End If
                                            ' Ende tk neu 18.1.20

                                            'Call updateRessBalkenOfProject(hproj, vglProj, chtobj, auswahl, replaceProj, chartPname)
                                            ' an der letzten Stelle stelle steht wenn dann die Rolle 
                                            'Call updateRessBalkenOfProject(hproj, vglProj, chtobj, auswahl, replaceProj, roleCostName)
                                            Call updateExcelChartOfProject(scInfo, chtobj, replaceProj, calledFromMassEdit)

                                        Case PTprdk.PersonalBalken
                                            Dim vglProj As clsProjekt = Nothing


                                            Try
                                                vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)
                                            Catch ex As Exception
                                                vglProj = Nothing
                                            End Try

                                            scInfo.vergleichsTyp = PTVergleichsTyp.erster
                                            scInfo.vglProj = vglProj

                                            ' tk neu 18.1.2020
                                            If myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektleitungRestricted Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                                ' 
                                                potentialParents = RoleDefinitions.getIDArray(myCustomUserRole.specifics)

                                            ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then

                                                Try
                                                    Dim allRoleIDsOfP As SortedList(Of Integer, Boolean) = vglProj.getRoleIDs()
                                                    Dim checkID As Integer

                                                    If teamID > 0 Then
                                                        checkID = teamID
                                                    Else
                                                        ' die potentiellen Parents sind die Parents der Kostenstelle / des Teams
                                                        checkID = rcID
                                                    End If

                                                    ' die potentiellen Parents sind  die Parents des Teams
                                                    Dim tmpList As Integer() = RoleDefinitions.getParentArray(RoleDefinitions.getRoleDefByID(checkID), includingMySelf:=True)
                                                    Dim ergListe As New List(Of Integer)
                                                    Dim found As Boolean = False
                                                    Dim ix As Integer = 1

                                                    Do While Not found And ix <= tmpList.Length
                                                        If allRoleIDsOfP.ContainsKey(tmpList(ix - 1)) Then
                                                            found = True
                                                        Else
                                                            ix = ix + 1
                                                        End If
                                                    Loop

                                                    If found Then
                                                        scInfo.q2 = RoleDefinitions.getRoleDefByID(tmpList(ix - 1)).name
                                                    Else
                                                        scInfo.q2 = RoleDefinitions.getRoleDefByID(rcID).name
                                                    End If

                                                Catch ex As Exception
                                                    scInfo.q2 = RoleDefinitions.getRoleDefByID(rcID).name
                                                End Try


                                                potentialParents = Nothing
                                            End If

                                            If Not IsNothing(potentialParents) Then

                                                Dim tmpParentName As String = ""

                                                If teamID = -1 Then
                                                    tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                Else
                                                    Dim tmpTeamName As String = RoleDefinitions.getRoleDefByID(teamID).name
                                                    tmpParentName = RoleDefinitions.chooseParentFromList(tmpTeamName, potentialParents)
                                                    If tmpParentName = "" Then
                                                        tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                    Else
                                                        Dim tmpParentNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpParentName, "")
                                                        If Not IsNothing(vglProj) Then
                                                            If vglProj.containsRoleNameID(tmpParentNameID) Then
                                                                ' passt bereits 
                                                            Else
                                                                tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                            End If
                                                        Else
                                                            tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                        End If


                                                    End If
                                                End If

                                                If tmpParentName <> "" Then
                                                    scInfo.q2 = tmpParentName
                                                End If

                                            End If

                                            ' Ende tk neu 18.1.20

                                            'Call updateRessBalkenOfProject(hproj, vglProj, chtobj, auswahl, replaceProj, chartPname)
                                            ' an der letzten Stelle stelle steht wenn dann die Rolle 
                                            'Call updateRessBalkenOfProject(hproj, vglProj, chtobj, auswahl, replaceProj, roleCostName)
                                            Call updateExcelChartOfProject(scInfo, chtobj, replaceProj, calledFromMassEdit)

                                        Case PTprdk.PersonalBalken2
                                            Dim vglProj As clsProjekt = Nothing

                                            Try
                                                vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, Date.Now, err)
                                            Catch ex As Exception
                                                vglProj = Nothing
                                            End Try

                                            scInfo.vergleichsTyp = PTVergleichsTyp.letzter
                                            scInfo.vglProj = vglProj

                                            ' tk neu 18.1.2020
                                            If myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektleitungRestricted Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                                ' 
                                                potentialParents = RoleDefinitions.getIDArray(myCustomUserRole.specifics)

                                            ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or
                                                    myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then

                                                Try
                                                    Dim allRoleIDsOfP As SortedList(Of Integer, Boolean) = vglProj.getRoleIDs()
                                                    Dim checkID As Integer

                                                    If teamID > 0 Then
                                                        checkID = teamID
                                                    Else
                                                        ' die potentiellen Parents sind die Parents der Kostenstelle / des Teams
                                                        checkID = rcID
                                                    End If

                                                    ' die potentiellen Parents sind  die Parents des Teams
                                                    Dim tmpList As Integer() = RoleDefinitions.getParentArray(RoleDefinitions.getRoleDefByID(checkID), includingMySelf:=True)
                                                    Dim ergListe As New List(Of Integer)
                                                    Dim found As Boolean = False
                                                    Dim ix As Integer = 1

                                                    Do While Not found And ix <= tmpList.Length
                                                        If allRoleIDsOfP.ContainsKey(tmpList(ix - 1)) Then
                                                            found = True
                                                        Else
                                                            ix = ix + 1
                                                        End If
                                                    Loop

                                                    If found Then
                                                        scInfo.q2 = RoleDefinitions.getRoleDefByID(tmpList(ix - 1)).name
                                                    Else
                                                        scInfo.q2 = RoleDefinitions.getRoleDefByID(rcID).name
                                                    End If

                                                Catch ex As Exception
                                                    scInfo.q2 = RoleDefinitions.getRoleDefByID(rcID).name
                                                End Try


                                                potentialParents = Nothing
                                            End If

                                            If Not IsNothing(potentialParents) Then

                                                Dim tmpParentName As String = ""

                                                If teamID = -1 Then
                                                    tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                Else
                                                    Dim tmpTeamName As String = RoleDefinitions.getRoleDefByID(teamID).name
                                                    tmpParentName = RoleDefinitions.chooseParentFromList(tmpTeamName, potentialParents)
                                                    If tmpParentName = "" Then
                                                        tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                    Else
                                                        Dim tmpParentNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpParentName, "")
                                                        If Not IsNothing(vglProj) Then
                                                            If vglProj.containsRoleNameID(tmpParentNameID) Then
                                                                ' passt bereits 
                                                            Else
                                                                tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                            End If
                                                        Else
                                                            tmpParentName = RoleDefinitions.chooseParentFromList(currentRCName, potentialParents)
                                                        End If



                                                    End If
                                                End If

                                                If tmpParentName <> "" Then
                                                    scInfo.q2 = tmpParentName
                                                End If

                                            End If

                                            ' Ende tk neu 18.1.20

                                            'Call updateRessBalkenOfProject(hproj, vglProj, chtobj, auswahl, replaceProj, chartPname)
                                            ' an der letzten Stelle stelle steht wenn dann die Rolle 
                                            'Call updateRessBalkenOfProject(hproj, vglProj, chtobj, auswahl, replaceProj, roleCostName)
                                            Call updateExcelChartOfProject(scInfo, chtobj, replaceProj, calledFromMassEdit)

                                        Case PTprdk.PersonalPie


                                            ' Update Pie-Diagramm
                                            Call updateRessPieOfProject(hproj, chtobj, auswahl)



                                        Case PTprdk.KostenPie


                                            Call updateCostPieOfProject(hproj, chtobj, auswahl)


                                        Case PTprdk.StrategieRisiko

                                            Call updateProjectPfDiagram(hproj, chtobj, auswahl)

                                        Case PTprdk.FitRisikoVol

                                            Call updateProjectPfDiagram(hproj, chtobj, auswahl)

                                        Case PTprdk.ComplexRisiko

                                            Call updateProjectPfDiagram(hproj, chtobj, auswahl)

                                        Case PTprdk.Ergebnis
                                            ' Update Ergebnis Diagramm
                                            Call updateProjektErgebnisCharakteristik2(hproj, chtobj, auswahl, replaceProj, calledFromMassEdit)

                                        Case PTprdk.SollIstGesamtkosten

                                            Call setProjektHistorie(hproj.name, hproj.variantName)
                                            Call updateSollIstOfProject(hproj, chtobj, Date.Now, auswahl, "", True, False)

                                        Case PTprdk.SollIstPersonalkosten

                                            Call setProjektHistorie(hproj.name, hproj.variantName)
                                            Call updateSollIstOfProject(hproj, chtobj, Date.Now, auswahl, "", True, False)

                                        Case PTprdk.SollIstSonstKosten

                                            Call setProjektHistorie(hproj.name, hproj.variantName)
                                            Call updateSollIstOfProject(hproj, chtobj, Date.Now, auswahl, "", True, False)

                                        Case Else


                                    End Select

                                End If

                            End If

                        End If

                    Next
                End If

            End With

        End If

    End Sub

    ''' <summary>
    ''' speichert ein einzelnes Projekt in der Datenbank , gibt in outputCollection die Meldungen zurück 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="outputCollection"></param>
    ''' <returns></returns>
    Public Function storeSingleProjectToDB(ByVal hproj As clsProjekt, ByRef outputCollection As Collection,
                                           Optional ByRef identical As Boolean = False) As Boolean


        Dim err As New clsErrorCodeMsg

        Dim tmpResult As Boolean = False

        Dim jetzt As Date = Date.Now

        enableOnUpdate = False


        Dim outputline As String = ""

        Try
            Dim formerVName As String = ""
            If Not (hproj.projectType = ptPRPFType.projectTemplate) Then

                ' die aktuelle WriteProtection holen 
                writeProtections.adjustListe(False) = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)

            End If

            ' die aktuelle Konstellation wird unter dem Namen <Last> gespeichert ..
            'Call storeSessionConstellation("Last")

            If CType(DatabaseAcc, DBAccLayer.Request).pingMongoDb() And Not noDB Then

                '' hier wird der Wert für kvp.Value.timeStamp = heute gesetzt 

                If demoModusHistory Then
                    hproj.timeStamp = historicDate
                Else
                    hproj.timeStamp = jetzt
                End If

                ' wenn es sich jetzt um einen Portfolio Manager handelt 
                ' er kann und darf nur mit Varianten-Name pfv speichern; es sei denn er hat selber eine Variante erzeugt bzw 
                ' es handelt sich bereits um die pfv Variante 
                ' prüfen auf Rolle 
                formerVName = hproj.variantName

                If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                    If hproj.variantName = "" Then
                        hproj.variantName = ptVariantFixNames.pfv.ToString
                    End If

                    ' tk 16.5.20 - immer wenn der Portfolio Manager speichert, wird das Projekt beauftragt 
                    'hproj.Status = ProjektStatus(PTProjektStati.beauftragt)
                End If

                ' das wurde rausgenommen, weil darin in AlleProjekte.Add die UpdateConstellation geändetr wurde , so dass dort die pfv-Variante referneziert war
                'Call changeVariantNameAccordingUserRole(hproj)


                Dim storeNeeded As Boolean = False
                Dim kdNrToStore As Boolean = False

                If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(hproj.name, hproj.variantName, hproj.timeStamp, err) Then
                    ' prüfen, ob es Unterschied gibt 
                    Dim standInDB As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, "", hproj.timeStamp, err)

                    If Not IsNothing(standInDB) Then
                        ' prüfe, ob es Unterschiede gibt
                        storeNeeded = Not hproj.isIdenticalTo(standInDB)
                        kdNrToStore = Not hproj.hasIdenticalKdNr(standInDB)

                        ' abfragen, ob Portfolio MAnager
                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                            If hproj.variantName = ptVariantFixNames.pfv.ToString Then
                                hproj.updatedAt = standInDB.updatedAt
                            End If
                        End If

                    Else
                        ' existiert nicht in der DB, also speichern; eigentlich darf dieser Zweig nie betreten werden !? 
                        storeNeeded = True
                    End If

                Else
                    storeNeeded = True
                End If


                If storeNeeded Then

                    ' nimmt das ggf zu mergende Projekt auf
                    Dim mProj As clsProjekt = Nothing


                    If CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(hproj, dbUsername, mProj, err, attrToStore:=kdNrToStore) Then

                        If awinSettings.englishLanguage Then

                            outputline = "saved: " & hproj.name & ", " & hproj.variantName
                            outputCollection.Add(outputline)

                        Else
                            outputline = "gespeichert: " & hproj.name & ", " & hproj.variantName
                            outputCollection.Add(outputline)
                        End If

                        If Not IsNothing(mProj) Then

                            'mProj statt hproj in AlleProjekte und ShowProjekte eintragen
                            Dim hProjKey As String = calcProjektKey(hproj.name, hproj.variantName)

                            If AlleProjekte.Containskey(hProjKey) Then
                                AlleProjekte.Remove(hProjKey, False)
                                AlleProjekte.Add(mProj, False)
                                ShowProjekte.Remove(hproj.name)
                                ShowProjekte.Add(mProj)
                            Else
                                AlleProjekte.Add(mProj, False)
                                ShowProjekte.Add(mProj)
                            End If

                            Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(mProj.name, mProj.variantName, err)
                            writeProtections.upsert(wpItem)

                        Else

                            Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                            writeProtections.upsert(wpItem, False)

                        End If

                        tmpResult = True
                        'Call MsgBox("ok, Projekt '" & hproj.name & "' gespeichert!" & vbLf & hproj.timeStamp.ToShortDateString)
                    Else
                        If awinSettings.visboServer Then
                            Select Case err.errorCode
                                Case 403  'No Permission to Create Visbo Project Version
                                    If awinSettings.englishLanguage Then
                                        outputline = "!!  No permission to store : " & hproj.name & ", " & hproj.variantName
                                        outputCollection.Add(outputline)
                                    Else
                                        outputline = "!!  Keine Erlaubnis zu speichern : " & hproj.name & ", " & hproj.variantName
                                        outputCollection.Add(outputline)
                                    End If

                                Case 409 ' VisboProjectVersion was already updated in between
                                    If awinSettings.englishLanguage Then
                                        outputline = "!! Projekt was already updated in between : " & hproj.name & ", " & hproj.variantName
                                        outputCollection.Add(outputline)
                                    Else
                                        outputline = "!!  Projekt wurde inzwischen verändert : " & hproj.name & ", " & hproj.variantName
                                        outputCollection.Add(outputline)
                                    End If
                                                '' erneut das projekt holen und abändern
                                                '' ur: 09.01.2019: wird in storeProjectToDB direkt gemacht
                                                'Dim standInDB As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(kvp.Value.name, kvp.Value.variantName, jetzt, err)

                                Case 423 ' Visbo Project (Portfolio) is locked by another user
                                    If awinSettings.englishLanguage Then
                                        outputline = err.errorMsg & ": " & hproj.name & ", " & hproj.variantName
                                        outputCollection.Add(outputline)
                                    Else
                                        outputline = "geschüztes Projekt : " & hproj.name & ", " & hproj.variantName
                                        outputCollection.Add(outputline)
                                    End If

                            End Select
                        Else

                            ' kann eigentlich gar nicht sein ... wäre nur dann der Fall, wenn ein Projekt komplett gelöscht wurde , aber der Schreibschutz nicht gelöscht wurde 
                            If awinSettings.englishLanguage Then
                                outputline = "protected project: " & hproj.name & ", " & hproj.variantName
                            Else
                                outputline = "geschütztes Projekt: " & hproj.name & ", " & hproj.variantName
                            End If
                            outputCollection.Add(outputline)

                        End If

                        Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                        writeProtections.upsert(wpItem, False)

                        tmpResult = False

                    End If
                Else
                    ' storeNeeded ist false, Kein Speichern erforderlich
                    identical = True
                    tmpResult = True
                End If
            Else

                tmpResult = False
                If awinSettings.englishLanguage Then
                    Throw New ArgumentException("No Database reachable!")
                Else
                    Throw New ArgumentException("Datenbank ist nicht aktiviert!")
                End If
            End If

            If Not IsNothing(hproj) Then
                hproj.variantName = formerVName
            End If


        Catch ex As Exception

            tmpResult = False
            ' Call MsgBox("Fehler beim Speichern der Projekte in die Datenbank. Datenbank nicht aktiviert?")
            If awinSettings.englishLanguage Then
                outputline = "Error when saving: " & hproj.name & ", " & hproj.variantName & ", " & ex.Message
            Else
                outputline = "Fehler beim Speichern: " & hproj.name & ", " & hproj.variantName & ", " & ex.Message
            End If

            outputCollection.Add(outputline)


        End Try


        storeSingleProjectToDB = tmpResult

    End Function



    ''' <summary>
    ''' speichert alle Projekte, die geladen sind bzw. in der Liste AlleProjekte enthalten sind
    ''' </summary>
    ''' <param name="everythingElse">true = auch Rollen und Kosten werden gespeichert in der DB
    '''                              false = nur Projekte werden gespeichert</param>
    Public Sub StoreAllProjectsinDB(Optional everythingElse As Boolean = False)

        Dim err As New clsErrorCodeMsg

        Dim jetzt As Date = Now
        Dim zeitStempel As Date

        enableOnUpdate = False

        Dim outPutCollection As New Collection
        Dim outputline As String = ""

        ' die aktuelle WriteProtection holen 
        writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)

        ' die aktuelle Konstellation wird unter dem Namen <Last> gespeichert ..
        'Call storeSessionConstellation("Last")

        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() And Not noDB Then

            Try
                Dim formerVName As String = ""

                ' jetzt werden die gezeigten Projekte in die Datenbank geschrieben 
                Dim anzahlStores As Integer = 0

                Dim pvNameListe As Collection = AlleProjekte.getPvNameListe

                ' jetzt werden alle Projekte gespeichert, alle Varianten 
                For Each curPVName As String In pvNameListe


                    Try
                        Dim hproj As clsProjekt = AlleProjekte.getProject(curPVName)
                        ' wenn es sich jetzt um einen Portfolio Manager handelt 
                        ' er kann und darf nur mit Varianten-Name pfv speichern; es sei denn er hat selber eine Variante erzeugt bzw 
                        ' es handelt sich bereits um die pfv Variante 
                        ' prüfen auf Rolle 

                        ' nur speichern, wenn es sich um ein Projekt, nicht um ein Portfolio handelt ...
                        If hproj.projectType = ptPRPFType.project Then

                            formerVName = hproj.variantName

                            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                If hproj.variantName = "" Then
                                    hproj.variantName = ptVariantFixNames.pfv.ToString
                                End If

                                ' tk 16.5.20 - immer wenn der Portfolio Manager speichert, wird das Projekt beauftragt 
                                'hproj.Status = ProjektStatus(PTProjektStati.beauftragt)
                            End If

                            'Call changeVariantNameAccordingUserRole(hproj)

                            Dim pvName As String = calcProjektKey(hproj.name, hproj.variantName)
                            If Not writeProtections.isProtected(pvName, dbUsername) Then

                                'hier wird der Wert für kvp.Value.timeStamp = heute gesetzt 
                                If demoModusHistory Then
                                    hproj.timeStamp = historicDate
                                Else
                                    hproj.timeStamp = jetzt
                                End If

                                Dim storeNeeded As Boolean = False
                                Dim kdNrToStore As Boolean = False
                                Dim standInDB As clsProjekt = Nothing

                                If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(hproj.name, hproj.variantName, hproj.timeStamp, err) Then
                                    ' prüfen, ob es Unterschied gibt 
                                    standInDB = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, "", hproj.timeStamp, err)
                                    If Not IsNothing(standInDB) Then
                                        ' prüfe, ob es Unterschiede gibt
                                        storeNeeded = Not hproj.isIdenticalTo(standInDB)
                                        kdNrToStore = Not hproj.hasIdenticalKdNr(standInDB)

                                        ' abfragen, ob Portfolio MAnager
                                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                            If hproj.variantName = ptVariantFixNames.pfv.ToString Then
                                                hproj.updatedAt = standInDB.updatedAt
                                            End If
                                        End If
                                    Else
                                        ' existiert nicht in der DB, also speichern; eigentlich darf dieser Zweig nie betreten werden !? 
                                        storeNeeded = True
                                    End If
                                Else
                                    storeNeeded = True
                                End If

                                If storeNeeded Then

                                    If kdNrToStore Then
                                        If Not IsNothing(standInDB) Then
                                            outputline = "Kunden-Nummer wurde geändert: von " & standInDB.kundenNummer & " zu " & hproj.kundenNummer
                                            outPutCollection.Add(outputline)
                                        End If
                                    End If

                                    Dim mproj As clsProjekt = Nothing
                                    Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString

                                    If CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(hproj, dbUsername, mproj, err, attrToStore:=kdNrToStore) Then

                                        If awinSettings.englishLanguage Then
                                            outputline = "saved : " & hproj.name & ", " & hproj.variantName
                                            outPutCollection.Add(outputline)
                                        Else
                                            outputline = "gespeichert : " & hproj.name & ", " & hproj.variantName
                                            outPutCollection.Add(outputline)
                                        End If

                                        anzahlStores = anzahlStores + 1

                                        ' jetzt die writeProtections aktualisieren 
                                        If Not IsNothing(mproj) Then

                                            'mProj statt hproj in AlleProjekte und ShowProjekte eintragen
                                            Dim hProjKey As String = calcProjektKey(hproj.name, hproj.variantName)

                                            If AlleProjekte.Containskey(hProjKey) Then
                                                AlleProjekte.Remove(hProjKey, False)
                                                AlleProjekte.Add(mproj, False)
                                                ShowProjekte.Remove(hproj.name)
                                                ShowProjekte.Add(mproj)
                                            Else
                                                AlleProjekte.Add(mproj, False)
                                                ShowProjekte.Add(mproj)
                                            End If

                                            Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(mproj.name, mproj.variantName, err)
                                            writeProtections.upsert(wpItem)

                                        Else

                                            Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                                            writeProtections.upsert(wpItem, False)

                                        End If

                                    Else
                                        If awinSettings.visboServer Then
                                            Select Case err.errorCode
                                                Case 403  'No Permission to Create Visbo Project Version
                                                    If awinSettings.englishLanguage Then
                                                        outputline = "!!  No permission to store : " & hproj.name & ", " & hproj.variantName
                                                        outPutCollection.Add(outputline)
                                                    Else
                                                        outputline = "!!  Keine Erlaubnis zu speichern : " & hproj.name & ", " & hproj.variantName
                                                        outPutCollection.Add(outputline)
                                                    End If

                                                Case 409 ' VisboProjectVersion was already updated in between
                                                    If awinSettings.englishLanguage Then
                                                        outputline = "!! Projekt was already updated in between : " & hproj.name & ", " & hproj.variantName
                                                        outPutCollection.Add(outputline)
                                                    Else
                                                        outputline = "!!  Projekt wurde inzwischen verändert : " & hproj.name & ", " & hproj.variantName
                                                        outPutCollection.Add(outputline)
                                                    End If

                                                Case 423 ' Visbo Project (Portfolio) is locked by another user
                                                    If awinSettings.englishLanguage Then
                                                        outputline = err.errorMsg & ": " & hproj.name & ", " & hproj.variantName
                                                        outPutCollection.Add(outputline)
                                                    Else
                                                        outputline = "geschüztes Projekt : " & hproj.name & ", " & hproj.variantName
                                                        outPutCollection.Add(outputline)
                                                    End If

                                            End Select
                                        Else
                                            If awinSettings.englishLanguage Then
                                                outputline = "protected project : " & hproj.name & ", " & hproj.variantName
                                                outPutCollection.Add(outputline)
                                            Else
                                                outputline = "geschütztes Projekt : " & hproj.name & ", " & hproj.variantName
                                                outPutCollection.Add(outputline)
                                            End If
                                        End If


                                        Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                                        writeProtections.upsert(wpItem)

                                    End If
                                End If
                            Else
                                ' nicht mehr rausschreiben - das ist ohnehin erwartet ... 
                            End If

                            '  den Varianten-Namen zurücksetzen
                            hproj.variantName = formerVName


                        End If



                    Catch ex As Exception

                        If awinSettings.englishLanguage Then
                            outputline = "!! Error when writing to database ..." & vbLf & ex.Message
                            outPutCollection.Add(outputline)
                        Else
                            outputline = "!! Fehler beim Speichern der Projekte in die Datenbank." & vbLf & ex.Message
                            outPutCollection.Add(outputline)
                        End If
                        ' Call MsgBox("Fehler beim Speichern der Projekte in die Datenbank. Datenbank nicht aktiviert?")
                        Throw New ArgumentException(outputline)
                        'Exit Sub
                    End Try

                Next


                historicDate = historicDate.AddMinutes(5)
                If historicDate > Date.Now Then
                    historicDate = Date.Now
                End If



                If everythingElse Then
                    ' jetzt werden alle definierten Constellations weggeschrieben
                    Dim errMsg As New clsErrorCodeMsg

                    'ur: 13.12.2019: nur noch Portfolio Namen holen
                    'Dim dbConstellations As clsConstellations = CType(databaseAcc, DBAccLayer.Request).retrieveConstellationsFromDB(Date.Now, errMsg)
                    Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)


                    For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

                        If kvp.Key <> "Sort Result" And kvp.Key <> "Filter Result" Then
                            Try
                                ' ur:13.12.2019: 
                                Call storeSingleConstellationToDB(outPutCollection, kvp.Value, dbPortfolioNames)
                                'Call storeSingleConstellationToDB(outPutCollection, kvp.Value, dbConstellations)

                            Catch ex As Exception

                                If awinSettings.englishLanguage Then
                                    outputline = "Error when writing portfolio " & kvp.Key
                                Else
                                    outputline = "Fehler in Schreiben Portfolio " & kvp.Key
                                End If

                            End Try
                        End If

                    Next

                End If


                zeitStempel = AlleProjekte.First.timeStamp

                ' Leerzeile einfügen
                outputline = "  "
                outPutCollection.Add(outputline)

                If anzahlStores > 0 Then
                    If anzahlStores = 1 Then
                        If awinSettings.englishLanguage Then
                            outputline = "1 project/project-variant stored: " &
                                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                        Else
                            outputline = "1 Projekt/Projekt-Variante gespeichert " &
                                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                        End If

                    Else
                        If awinSettings.englishLanguage Then
                            outputline = anzahlStores & " projects/project-variants stored: " &
                                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                        Else
                            outputline = anzahlStores & " Projekte/Projekt-Varianten gespeichert: " &
                                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                        End If

                    End If
                Else
                    If awinSettings.englishLanguage Then
                        outputline = "no projects stored ( no changes / no permission ) "
                    Else
                        outputline = "keine Projekte gespeichert ( keine Änderungen / keine Erlaubnis )"
                    End If

                End If

                outPutCollection.Add(outputline)

                ' tk 1.5.19 wird nicht mehr benötigt ...
                'If everythingElse Then

                '    If anzahlStores > 0 Then
                '        If anzahlStores = 1 Then
                '            If awinSettings.englishLanguage Then
                '                outputline = "ok, portfolios stored!" & vbLf & vbLf &
                '                        "1 project/project-variant stored " & vbLf &
                '                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                '            Else
                '                outputline = "ok, Portfolios gespeichert!" & vbLf & vbLf &
                '                        "es wurde 1 Projekt bzw. Projekt-Variante gespeichert" & vbLf &
                '                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                '            End If

                '        Else
                '            If awinSettings.englishLanguage Then
                '                outputline = "ok, portfolios stored!" & vbLf & vbLf &
                '                        anzahlStores & " projects/project-variants stored" & vbLf &
                '                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                '            Else
                '                outputline = "ok, Portfolios gespeichert!" & vbLf & vbLf &
                '                        "es wurden " & anzahlStores & " Projekte bzw. Projekt-Varianten gespeichert " & vbLf &
                '                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                '            End If

                '        End If
                '    Else
                '        If awinSettings.englishLanguage Then
                '            outputline = "ok, portfolios stored!" & vbLf &
                '                "no projects stored ( no changes / no permission ) "
                '        Else
                '            outputline = "ok, Portfolios gespeichert!" & vbLf &
                '                "keine Projekte gespeichert ( keine Änderungen / keine Erlaubnis )"
                '        End If

                '    End If


                'Else

                '    If anzahlStores > 0 Then
                '        If anzahlStores = 1 Then
                '            If awinSettings.englishLanguage Then
                '                outputline = "1 project/project-variant stored: " &
                '                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                '            Else
                '                outputline = "1 Projekt/Projekt-Variante gespeichert " &
                '                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                '            End If

                '        Else
                '            If awinSettings.englishLanguage Then
                '                outputline = anzahlStores & " projects/project-variants stored: " &
                '                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                '            Else
                '                outputline = anzahlStores & " Projekte/Projekt-Varianten gespeichert: " &
                '                        zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString
                '            End If

                '        End If
                '    Else
                '        If awinSettings.englishLanguage Then
                '            outputline = "no projects stored ( no changes / no permission ) "
                '        Else
                '            outputline = "keine Projekte gespeichert ( keine Änderungen / keine Erlaubnis )"
                '        End If

                '    End If
                'End If



                If outPutCollection.Count > 0 Then
                    Dim msgH As String, msgE As String
                    If awinSettings.englishLanguage Then
                        If everythingElse Then
                            msgH = "Save Everything (Projects, Portfolios)"
                        Else
                            msgH = "Save Projects"
                        End If

                        msgE = "following results:"
                    Else
                        If everythingElse Then
                            msgH = "Alles speichern (Projekte, Portfolios)"
                        Else
                            msgH = "Projekte speichern"
                        End If

                        msgE = "Rückmeldungen"
                    End If

                    Call showOutPut(outPutCollection, msgH, msgE)

                End If

                ' Änderung 18.6 - wenn gespeichert wird, soll die Projekthistorie zurückgesetzt werden 
                Try
                    If projekthistorie.Count > 0 Then
                        projekthistorie.clear()
                    End If
                Catch ex As Exception

                End Try

            Catch ex As Exception
                Throw New ArgumentException("Fehler beim Speichern der Projekte in die Datenbank." & vbLf & ex.Message)
                'Call MsgBox(" Fehler beim Speichern in die Datenbank")
            End Try
        Else

            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen")

        End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' speichert die selektierten Projekte
    ''' </summary>
    ''' <returns>Anzahl der erfolgreich gespeicherten Projekte</returns>
    Public Function StoreSelectedProjectsinDB() As Integer

        Dim err As New clsErrorCodeMsg

        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt
        Dim hilfshproj As clsProjekt
        Dim jetzt As Date = Now
        Dim anzSelectedProj As Integer = 0
        Dim anzStoredProj As Integer = 0
        Dim variantCollection As Collection

        Dim outputCollection As New Collection
        Dim outputline As String = ""
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        Dim awinSelection As Excel.ShapeRange

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                anzSelectedProj = awinSelection.Count

                For i = 1 To awinSelection.Count

                    ' jetzt die Aktion durchführen ...
                    singleShp1 = awinSelection.Item(i)

                    Try

                        hilfshproj = ShowProjekte.getProject(singleShp1.Name, True)

                    Catch ex As Exception
                        Throw New ArgumentException("Projekt nicht gefunden ...")
                        enableOnUpdate = True
                    End Try

                    ' alle geladenen Variante in variantCollection holen
                    variantCollection = AlleProjekte.getVariantNames(hilfshproj.name, False)

                    For vi = 1 To variantCollection.Count

                        Dim hVname As String = variantCollection.Item(vi)
                        'Dim tmpStr(5) As String
                        'Dim trennzeichen1 As String = "("
                        'Dim trennzeichen2 As String = ")"

                        '' VariantenNamen von den () befreien
                        'tmpStr = variantCollection(vi).Split(New Char() {CChar(trennzeichen1)}, 4)
                        'tmpStr = tmpStr(1).Split(New Char() {CChar(trennzeichen2)}, 4)
                        'hVname = tmpStr(0)

                        ' gesamte ProjektInfo der Variante aus Liste AlleProjekte lesen
                        hproj = AlleProjekte.getProject(calcProjektKey(hilfshproj.name, hVname))


                        Try
                            ' wenn es sich jetzt um einen Portfolio Manager handelt 
                            ' er kann und darf nur mit Varianten-Name pfv speichern; es sei denn er hat selber eine Variante erzeugt bzw 
                            ' es handelt sich bereits um die pfv Variante 
                            ' prüfen auf Rolle 

                            Dim formerVName As String = hproj.variantName
                            If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                If hproj.variantName = "" Then
                                    hproj.variantName = ptVariantFixNames.pfv.ToString
                                End If

                                ' tk 16.5.20 - immer wenn der Portfolio Manager speichert, wird das Projekt beauftragt 
                                'hproj.Status = ProjektStatus(PTProjektStati.beauftragt)
                            End If
                            'Call changeVariantNameAccordingUserRole(hproj)

                            ' ur: 31.1.2019: hier ist es zu früh, den neuen Timestamp zu setzen, 
                            ' denn es muss evt. das Projekt mit dem alten Timestamp nochmals aus DB geholt werden
                            ' ur: 14.2.2019: wieder zurückgeändert
                            '' hier wird der Wert für kvp.Value.timeStamp = heute gesetzt 
                            If demoModusHistory Then
                                hproj.timeStamp = historicDate
                            Else
                                hproj.timeStamp = jetzt
                            End If

                            Dim storeNeeded As Boolean = False
                            Dim kdNrToStore As Boolean = False
                            If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(hproj.name, hproj.variantName, hproj.timeStamp, err) Then
                                ' prüfen, ob es Unterschied gibt 
                                Dim standInDB As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, "", hproj.timeStamp, err)


                                If Not IsNothing(standInDB) Then
                                    ' prüfe, ob es Unterschiede gibt
                                    storeNeeded = Not hproj.isIdenticalTo(standInDB)
                                    kdNrToStore = Not hproj.hasIdenticalKdNr(standInDB)

                                    ' abfragen, ob Portfolio MAnager
                                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                                        If hproj.variantName = ptVariantFixNames.pfv.ToString Then
                                            hproj.updatedAt = standInDB.updatedAt
                                        End If
                                    End If
                                Else
                                    ' existiert nicht in der DB, also speichern; eigentlich darf dieser Zweig nie betreten werden !? 
                                    storeNeeded = True
                                End If
                            Else
                                storeNeeded = True
                            End If

                            If storeNeeded Then

                                Dim mproj As clsProjekt = Nothing

                                'Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString

                                ''ur: 15.1.2020: wird nun ja im Server erledigt

                                'If hproj.variantName <> vorgabeVariantName Then

                                '    '
                                '    ' hier muss die Berechnung der keyMetrics-Daten erfolgen
                                '    '
                                '    hproj.keyMetrics = calcKeyMetricsOfProject(hproj)


                                'Else
                                '    ' hier ist noch zu überlegen, was zu tun ist.
                                '    ' z.B.  leere keyMetrics
                                '    hproj.keyMetrics = New clsKeyMetrics
                                'End If


                                If CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(hproj, dbUsername, mproj, err, attrToStore:=kdNrToStore) Then

                                    If awinSettings.englishLanguage Then
                                        outputline = "saved : " & hproj.name & ", " & hproj.variantName
                                        outputCollection.Add(outputline)
                                    Else
                                        outputline = "gespeichert : " & hproj.name & ", " & hproj.variantName
                                        outputCollection.Add(outputline)
                                    End If

                                    anzStoredProj = anzStoredProj + 1

                                    If Not IsNothing(mproj) Then

                                        Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(mproj.name, mproj.variantName, err)
                                        writeProtections.upsert(wpItem)

                                        ' gemergte Projekt nun in AlleProjekte und ShowProjekte ersetzen
                                        Call replaceProjectVariant(mproj.name, mproj.variantName, False, True, mproj.tfZeile)

                                    Else

                                        Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                                        writeProtections.upsert(wpItem, False)

                                    End If
                                    'Call MsgBox("ok, Projekt '" & hproj.name & "' gespeichert!" & vbLf & hproj.timeStamp.ToShortDateString)
                                Else

                                    If awinSettings.visboServer Then
                                        Select Case err.errorCode
                                            Case 403  'No Permission to Create Visbo Project Version
                                                If awinSettings.englishLanguage Then
                                                    outputline = "!!  No permission to store : " & hproj.name & ", " & hproj.variantName
                                                    outputCollection.Add(outputline)
                                                Else
                                                    outputline = "!!  Keine Erlaubnis zu speichern : " & hproj.name & ", " & hproj.variantName
                                                    outputCollection.Add(outputline)
                                                End If

                                            Case 409 ' VisboProjectVersion was already updated in between
                                                If awinSettings.englishLanguage Then
                                                    outputline = "!! Projekt was already updated in between : " & hproj.name & ", " & hproj.variantName
                                                    outputCollection.Add(outputline)
                                                Else
                                                    outputline = "!!  Projekt wurde inzwischen verändert : " & hproj.name & ", " & hproj.variantName
                                                    outputCollection.Add(outputline)
                                                End If
                                                '' erneut das projekt holen und abändern
                                                '' ur: 09.01.2019: wird in storeProjectToDB direkt gemacht
                                                'Dim standInDB As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(kvp.Value.name, kvp.Value.variantName, jetzt, err)

                                            Case 423 ' Visbo Project (Portfolio) is locked by another user
                                                If awinSettings.englishLanguage Then
                                                    outputline = err.errorMsg & ": " & hproj.name & ", " & hproj.variantName
                                                    outputCollection.Add(outputline)
                                                Else
                                                    outputline = "geschüztes Projekt : " & hproj.name & ", " & hproj.variantName
                                                    outputCollection.Add(outputline)
                                                End If

                                        End Select
                                    Else

                                        ' kann eigentlich gar nicht sein ... wäre nur dann der Fall, wenn ein Projekt komplett gelöscht wurde , aber der Schreibschutz nicht gelöscht wurde 
                                        If awinSettings.englishLanguage Then
                                            outputline = "protected project: " & hproj.name & ", " & hproj.variantName
                                        Else
                                            outputline = "geschütztes Projekt: " & hproj.name & ", " & hproj.variantName
                                        End If
                                        outputCollection.Add(outputline)

                                    End If

                                    Dim wpItem As clsWriteProtectionItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                                    writeProtections.upsert(wpItem)

                                End If
                            End If

                            If IsNothing(hproj) Then
                                hproj.variantName = formerVName
                            End If

                        Catch ex As Exception

                            If awinSettings.englishLanguage Then
                                outputline = "Error when saving: " & hproj.name & ", " & hproj.variantName & vbLf & ex.Message
                            Else
                                outputline = "Fehler beim Speichern: " & hproj.name & ", " & hproj.variantName & vbLf & ex.Message
                            End If

                            outputCollection.Add(outputline)
                            'Throw New ArgumentException("Fehler beim Speichern der Projekte in die Datenbank." & vbLf & ex.Message)
                            'Exit Sub
                        End Try

                    Next vi

                Next i

            Else
                'Call MsgBox("Es wurde kein Projekt selektiert")
                ' die Anzahl selektierter und auch gespeicherter Projekte ist damit = 0
                anzStoredProj = anzSelectedProj
                Return anzSelectedProj
            End If


        Else

            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen")

        End If


        enableOnUpdate = True
        'If demoModusHistory Then
        '    Call MsgBox("ok, " & anzStoredProj & " Projekte und Varianten gespeichert!" & vbLf & historicDate.ToShortDateString & ", " & historicDate.ToShortTimeString)
        'Else
        '    Call MsgBox("ok, " & anzStoredProj & " Projekte und Varianten gespeichert!" & vbLf & jetzt.ToShortDateString & ", " & jetzt.ToShortTimeString)
        'End If

        If demoModusHistory Then
            Call MsgBox("ok, " & anzStoredProj & " Projekte und Varianten gespeichert!" & vbLf & historicDate.ToShortDateString & ", " & historicDate.ToShortTimeString)
        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("ok, " & anzStoredProj & " projects and variants are stored now!" & vbLf & jetzt.ToShortDateString & ", " & jetzt.ToShortTimeString)
            Else
                Call MsgBox("ok, " & anzStoredProj & " Projekte und Varianten gespeichert!" & vbLf & jetzt.ToShortDateString & ", " & jetzt.ToShortTimeString)
            End If
        End If

        If outputCollection.Count > 0 Then
            Dim msgH As String = ""
            Dim msgE As String
            If awinSettings.englishLanguage Then
                msgE = "following results:"
            Else
                msgE = "Rückmeldungen"
            End If

            Call showOutPut(outputCollection, msgH, msgE)

        End If

        Return anzStoredProj

    End Function


    ''' <summary>
    ''' Es wird die keyMetrics des Projekte berechnet und als result zurückgegeben
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <returns></returns>
    Public Function calcKeyMetricsOfProject(ByVal hproj As clsProjekt) As clsKeyMetrics


        Dim result As New clsKeyMetrics
        Dim err As New clsErrorCodeMsg

        Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString
        Dim index As Integer = getColumnOfDate(hproj.timeStamp) - hproj.Start

        Dim lproj As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, Date.Now, err)
        '
        ' hier muss die Berechnung der keyMetrics-Daten erfolgen        '
        '
        If Not IsNothing(lproj) Then

            result.costBaseLastTotal = lproj.getSummeKosten()
            result.costCurrentTotal = hproj.getSummeKosten()

            result.costBaseLastActual = lproj.getSummeKosten(index)
            result.costCurrentActual = hproj.getSummeKosten(index)

            result.endDateBaseLast = lproj.endeDate
            result.endDateCurrent = hproj.endeDate

            Dim baseMs As SortedList(Of Date, String) = lproj.getMilestones
            Dim basePhases As SortedList(Of Date, String) = lproj.getPhases
            result.timeCompletionBaseLastActual = lproj.getTimeCompletionMetric(baseMs, basePhases, hproj.timeStamp).Sum
            result.timeCompletionBaseLastTotal = lproj.getTimeCompletionMetric(baseMs, basePhases, hproj.timeStamp, True).Sum

            result.timeCompletionCurrentActual = hproj.getTimeCompletionMetric(baseMs, basePhases, hproj.timeStamp).Sum
            result.timeCompletionCurrentTotal = hproj.getTimeCompletionMetric(baseMs, basePhases, hproj.timeStamp, True).Sum

            Dim baseDeliverables As SortedList(Of String, String) = lproj.getDeliverables
            result.deliverableCompletionBaseLastActual = lproj.getDeliverableCompletionMetric(baseDeliverables, hproj.timeStamp).Sum
            result.deliverableCompletionBaseLastTotal = lproj.getDeliverableCompletionMetric(baseDeliverables, hproj.timeStamp, True).Sum

            result.deliverableCompletionCurrentActual = hproj.getDeliverableCompletionMetric(baseDeliverables, hproj.timeStamp).Sum
            result.deliverableCompletionCurrentTotal = hproj.getDeliverableCompletionMetric(baseDeliverables, hproj.timeStamp, True).Sum

        Else

            ' result bleibt nahezu leer, d.h. es werden nur costCurrentActual und costCurrentTotal und endDateCurrent besetzt
            result.costCurrentTotal = hproj.getSummeKosten()
            result.costCurrentActual = hproj.getSummeKosten(index)
            result.endDateCurrent = hproj.endeDate
        End If

        calcKeyMetricsOfProject = result

    End Function


End Module
