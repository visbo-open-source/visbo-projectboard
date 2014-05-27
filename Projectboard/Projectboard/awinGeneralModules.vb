Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ClassLibrary1
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Windows


Public Module awinGeneralModules


   

    ''' <summary>
    ''' schreibt evtl neu durch Inventur hinzugekommene Phasen in 
    ''' das Customization File 
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub awinWritePhaseDefinitions()

        Dim phaseDefs As Range
        Dim foundRow As Integer
        Dim phName As String, phColor As Long
        Dim lastrow As Excel.Range

        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False



        ' hier muss jetzt das File Projekt Tafel Definitions.xlsx aufgemacht werden ...
        Try
            appInstance.Workbooks.Open(awinPath & customizationFile)

        Catch ex As Exception
            Call MsgBox("Customization File nicht gefunden - Abbruch")
            Throw New ArgumentException("Customization File nicht gefunden - Abbruch")
        End Try


        Dim wsName4 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(4)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

        phaseDefs = wsName4.Range("awin_Phasen_Definition")
        lastrow = phaseDefs.Rows(phaseDefs.Rows.Count)


        ' jetzt muss getestet werden, ob jede Phase in PhaseDefinitions bereits in der Customization vorkommt 

        For i = 1 To PhaseDefinitions.Count
            phName = PhaseDefinitions.getPhaseDef(i).name
            phColor = CLng(PhaseDefinitions.getPhaseDef(i).farbe)

            Try

                foundRow = phaseDefs.Find(What:=phName).Row
                ' wenn es gefunden wurde - keine weitere Aktion nötig 

            Catch ex As Exception
                ' andernfalls eintragen 

                lastrow = phaseDefs.Rows(phaseDefs.Rows.Count)
                CType(lastrow.EntireRow, Excel.Range).Insert(XlInsertShiftDirection.xlShiftDown)
                lastrow.Cells(1, 1).offset(-1, 0).value = phName
                lastrow.Cells(1, 1).offset(-1, 0).interior.color = phColor

            End Try


        Next


        appInstance.ActiveWorkbook.Close(SaveChanges:=True)
        appInstance.ScreenUpdating = True
        appInstance.EnableEvents = True

    End Sub

    Friend Sub awinsetTypen()

        Dim i As Integer
        'Dim Start As Integer, Dauer As Integer
        'Dim startdate As Date
        'Dim zeile As Integer
        'Dim pname As String
        Dim c As Excel.Range
        'Dim marge As Double
        'Dim sfit As Double, risk As Double
        'Dim vorlagenName As String
        'Dim projStatus As String

        'Dim hproj As clsProjekt
        'Dim hpv As clsProjektvorlage
        Dim hrole As clsRollenDefinition
        Dim hcost As clsKostenartDefinition
        Dim hphase As clsPhasenDefinition
        'Dim DifferenceInMonths As Long
        Dim dateiListe As New Collection
        Dim dateiName As String
        Dim tmpStr As String



        awinPath = appInstance.ActiveWorkbook.Path & "\"


        ProjektStatus(0) = "geplant"
        ProjektStatus(1) = "beauftragt"
        ProjektStatus(2) = "beauftragt, Änderung noch nicht freigegeben"
        ProjektStatus(3) = "beendet" ' ein Projekt wurde in seinem Verlauf beendet, ohne es plangemäß abzuschliessen
        ProjektStatus(4) = "abgeschlossen"


        DiagrammTypen(0) = "Phase"
        DiagrammTypen(1) = "Rolle"
        DiagrammTypen(2) = "Kostenart"
        DiagrammTypen(3) = "Portfolio"
        DiagrammTypen(4) = "Ergebnis"
        DiagrammTypen(5) = "Meilenstein"
        DiagrammTypen(6) = "Meilenstein Trendanalyse"

        ergebnisChartName(0) = "Earned Value"
        ergebnisChartName(1) = "Earned Value - gewichtet"
        ergebnisChartName(2) = "Verbesserungs-Potential"
        ergebnisChartName(3) = "Risiko-Abschlag"

        ReDim portfolioDiagrammtitel(20)
        portfolioDiagrammtitel(PTpfdk.Phasen) = "Phasen - Übersicht"
        portfolioDiagrammtitel(PTpfdk.Rollen) = "Rollen - Übersicht"
        portfolioDiagrammtitel(PTpfdk.Kosten) = "Kosten - Übersicht"
        portfolioDiagrammtitel(PTpfdk.ErgebnisWasserfall) = summentitel1
        portfolioDiagrammtitel(PTpfdk.FitRisiko) = summentitel2
        portfolioDiagrammtitel(PTpfdk.Auslastung) = summentitel9
        portfolioDiagrammtitel(PTpfdk.UeberAuslastung) = summentitel10
        portfolioDiagrammtitel(PTpfdk.Unterauslastung) = summentitel11
        portfolioDiagrammtitel(PTpfdk.ZieleV) = summentitel6
        portfolioDiagrammtitel(PTpfdk.ZieleF) = summentitel7
        portfolioDiagrammtitel(PTpfdk.ComplexRisiko) = "Komplexität, Risiko und Volumen"
        portfolioDiagrammtitel(PTpfdk.ZeitRisiko) = "Zeit, Risiko und Volumen"
        portfolioDiagrammtitel(PTpfdk.AmpelFarbe) = ""
        portfolioDiagrammtitel(PTpfdk.ProjektFarbe) = ""
        portfolioDiagrammtitel(PTpfdk.Meilenstein) = "Meilenstein - Übersicht"
        portfolioDiagrammtitel(PTpfdk.FitRisikoVol) = "strategischer Fit, Risiko & Volumen"
        portfolioDiagrammtitel(PTpfdk.Dependencies) = "Abhängigkeiten: Aktive bzw passive Beeinflussung"
        portfolioDiagrammtitel(PTpfdk.betterWorseL) = "Abweichungen zum letztem Stand"
        portfolioDiagrammtitel(PTpfdk.betterWorseB) = "Abweichungen zur Beauftragung"
        portfolioDiagrammtitel(PTpfdk.Budget) = "Budget Übersicht"


        windowNames(0) = "Cockpit Phasen"
        windowNames(1) = "Cockpit Rollen"
        windowNames(2) = "Cockpit Kosten"
        windowNames(3) = "Cockpit Wertigkeit"
        windowNames(4) = "Cockpit Ergebnisse"
        windowNames(5) = "Projekt Tafel"

        '
        ' die Namen der Worksheets Ressourcen und Portfolio verfügbar machen
        '
        arrWsNames(1) = "Portfolio"
        arrWsNames(2) = "Vorlage"                          ' dient als Hilfs-Sheet für Anzeige in Plantafel 
        arrWsNames(3) = "Tabelle1"
        arrWsNames(4) = "Einstellungen"
        arrWsNames(5) = "Tabelle2"
        arrWsNames(6) = "Edit Allgemein"
        arrWsNames(7) = ""                          ' war Kosten ; ist nicht mehr notwendig
        arrWsNames(8) = "Projekt iRessourcen"
        arrWsNames(9) = "Projekt iKosten"
        arrWsNames(10) = "Portfolio Übersicht"
        arrWsNames(11) = "Projekt editieren"
        arrWsNames(12) = "Projektdefinition Erloese"
        arrWsNames(13) = "Projekt iErloese"
        arrWsNames(14) = "Objekte"
        arrWsNames(15) = "Portfolio Vorlage"


        ProjectBoardDefinitions.My.Settings.loadProjectsOnChange = False

        showRangeLeft = 0
        showRangeRight = 0

        'selectedRoleNeeds = 0
        'selectedCostNeeds = 0

        ' bestimmen der maximalen Breite und Höhe 
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False


        With appInstance.ActiveWindow


            If .WindowState = Excel.XlWindowState.xlMaximized Then
                maxScreenHeight = .Height
                maxScreenWidth = .Width
            Else
                Dim formerState As Excel.XlWindowState = .WindowState
                .WindowState = Excel.XlWindowState.xlMaximized
                maxScreenHeight = .Height
                maxScreenWidth = .Width
                .WindowState = formerState
            End If


        End With

        miniHeight = maxScreenHeight / 6
        miniWidth = maxScreenWidth / 10



        Dim oGrenze As Integer = UBound(frmCoord, 1)
        ' hier werden die Top- & Left- Default Positionen der Formulare gesetzt 
        For i = 0 To oGrenze
            frmCoord(i, PTpinfo.top) = maxScreenHeight * 0.3
            frmCoord(i, PTpinfo.left) = maxScreenWidth * 0.4
        Next

        ' jetzt setzen der Werte für Status-Information und Milestone-Information
        frmCoord(PTfrm.projInfo, PTpinfo.top) = 125
        frmCoord(PTfrm.projInfo, PTpinfo.left) = My.Computer.Screen.WorkingArea.Width - 500

        frmCoord(PTfrm.msInfo, PTpinfo.top) = 125 + 280
        frmCoord(PTfrm.msInfo, PTpinfo.left) = My.Computer.Screen.WorkingArea.Width - 500

        ' With listOfWorkSheets(arrWsNames(4))


        ' hier muss jetzt das File Projekt Tafel Definitions.xlsx aufgemacht werden ...
        Try
            appInstance.Workbooks.Open(awinPath & customizationFile)

        Catch ex As Exception
            Call MsgBox("Customization File nicht gefunden - Abbruch")
            appInstance.ScreenUpdating = formerSU
            Throw New ArgumentException("Customization File nicht gefunden - Abbruch")
            Exit Sub
        End Try


        Dim wsName4 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(4)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

        With wsName4

            '
            ' Phasen Definitionen auslesen - im bereich awin_Phasen_Definition
            '
            i = 0

            For Each c In .Range("awin_Phasen_Definition")

                If CStr(c.Value) <> "" Then
                    i = i + 1
                    tmpStr = CType(c.Value, String)
                    ' das neue ...
                    hphase = New clsPhasenDefinition
                    With hphase
                        .farbe = c.Interior.Color
                        .name = tmpStr.Trim
                        .UID = i

                        ' hat die Phase einen Schwellwert ? 
                        Try
                            If CInt(c.Offset(0, 1).Value) > 0 Then
                                .schwellWert = CInt(c.Offset(0, 1).Value)
                            End If
                        Catch ex As Exception

                        End Try

                        ' ist die Phase eine special Phase ? 
                        Try
                            If CStr(c.Offset(0, 2).Value).Trim = "LeLe" Then
                                specialListofPhases.Add(hphase.name, hphase.name)
                            End If
                        Catch ex As Exception

                        End Try
                    End With

                    Try
                        PhaseDefinitions.Add(hphase)
                    Catch ex As Exception

                    End Try


                End If

            Next c


            '
            ' Rollen Definitionen auslesen - im bereich awin_Rollen_Definition
            '
            i = 0
            For Each c In .Range("awin_Rollen_Definition")
                If CStr(c.Value) <> "" Then
                    i = i + 1
                    tmpStr = CType(c.Value, String)
                    If i = 1 Then
                        rollenKapaFarbe = c.Offset(0, 1).Interior.Color
                    End If


                    ' jetzt kommt die Rollen Definition 
                    hrole = New clsRollenDefinition
                    Dim cp As Integer
                    With hrole
                        .name = tmpStr.Trim
                        .Startkapa = CDbl(c.Offset(0, 1).Value)
                        .tagessatzIntern = CDbl(c.Offset(0, 2).Value)
                        If CDbl(c.Offset(0, 3).Value) = 0.0 Then
                            .tagessatzExtern = CDbl(c.Offset(0, 2).Value) * 1.35
                        Else
                            .tagessatzExtern = CDbl(c.Offset(0, 3).Value)
                        End If
                        ' Auslesen der zukünftigen Kapazität
                        For cp = 1 To 120
                            .kapazitaet(cp) = CType(c.Offset(0, 3 + cp).Value, Double)
                            If .kapazitaet(cp) <= 0 Then
                                ' Kapa kann nicht negative sein
                                ' wenn nichts angegeben wird, soll die Startkapa verwendet werden 
                                .kapazitaet(cp) = .Startkapa
                            End If
                        Next
                        .farbe = c.Interior.Color
                        .UID = i
                    End With
                    RoleDefinitions.Add(hrole)
                    'hrole = Nothing

                End If

            Next c



            i = 0
            For Each c In .Range("awin_Kosten_Definition")

                If CStr(c.Value) <> "" Or i > 0 Then
                    i = i + 1


                    ' jetzt kommt die Kostenarten Definition
                    hcost = New clsKostenartDefinition
                    With hcost
                        If CStr(c.Value) <> "" Then
                            tmpStr = CType(c.Value, String)
                            .name = tmpStr.Trim
                        Else
                            .name = "Personalkosten"
                        End If
                        .farbe = c.Interior.Color
                        .UID = i
                    End With

                    CostDefinitions.Add(hcost)
                    'hcost = Nothing
                End If

            Next c


            '
            ' max Projektdauer auslesen
            '
            'maxProjektdauer = .Range("Max_Dauer_eines_Projektes").Value

            '
            ' linker und rechter Rand für Diagramme auslesen
            '

            Try
                'showRangeLeft = CInt(.Range("Linker_Rand_Ressourcen_Diagramme").Value)
                'showRangeRight = CInt(.Range("Rechter_Rand_Ressourcen_Diagramme").Value)
                showtimezone_color = .Range("Show_Time_Zone_Color").Interior.Color
                noshowtimezone_color = .Range("NoShow_Time_Zone_Color").Interior.Color
                nrOfDaysMonth = CDbl(.Range("Arbeitstage_pro_Monat").Value)
                farbeInternOP = .Range("Farbe_intern_ohne_Projekte").Interior.Color
                farbeExterne = .Range("Farbe_externe_Ressourcen").Interior.Color
                iProjektFarbe = .Range("Farbe_für_Projekte_ohne_Vorlage").Interior.Color
                iWertFarbe = .Range("Farbe_Ress_Kost_Werte").Interior.Color
                vergleichsfarbe0 = .Range("Vergleichsfarbe1").Interior.Color
                vergleichsfarbe1 = .Range("Vergleichsfarbe2").Interior.Color
                vergleichsfarbe2 = .Range("Vergleichsfarbe3").Interior.Color

                'Dim tmpcolor As Microsoft.Office.Interop.Excel.ColorFormat

                Try
                    awinSettings.SollIstFarbeB = CLng(.Range("Soll_Ist_Farbe_Beauftragung").Interior.Color)
                    awinSettings.SollIstFarbeL = CLng(.Range("Soll_Ist_Farbe_letzte_Freigabe").Interior.Color)
                    awinSettings.SollIstFarbeC = CLng(.Range("Soll_Ist_Farbe_Aktuell").Interior.Color)
                    awinSettings.AmpelGruen = CLng(.Range("AmpelGruen").Interior.Color)
                    'tmpcolor = CType(.Range("AmpelGruen").Interior.Color, Microsoft.Office.Interop.Excel.ColorFormat)
                    awinSettings.AmpelGelb = CLng(.Range("AmpelGelb").Interior.Color)
                    awinSettings.AmpelRot = CLng(.Range("AmpelRot").Interior.Color)
                    awinSettings.AmpelNichtBewertet = CLng(.Range("AmpelNichtBewertet").Interior.Color)
                    awinSettings.glowColor = CLng(.Range("GlowColor").Interior.Color)

                Catch ex As Exception
                    appInstance.ScreenUpdating = formerSU
                    Throw New ArgumentException("Customization File fehlerhaft - Farben fehlen ... " & vbLf & ex.Message)
                End Try

                ergebnisfarbe1 = .Range("Ergebnisfarbe1").Interior.Color
                ergebnisfarbe2 = .Range("Ergebnisfarbe2").Interior.Color
                weightStrategicFit = CDbl(.Range("WeightStrategicFit").Value)
                ' jetzt wird KalenderStart, Zeiteinheit und Datenbank Name ausgelesen 
                awinSettings.kalenderStart = CDate(.Range("Start_Kalender").Value)
                awinSettings.zeitEinheit = CStr(.Range("Zeiteinheit").Value)
                awinSettings.kapaEinheit = CStr(.Range("kapaEinheit").Value)
                awinSettings.offsetEinheit = CStr(.Range("offsetEinheit").Value)
                awinSettings.databaseName = CStr(.Range("Datenbank").Value)
                awinSettings.EinzelRessExport = CInt(.Range("EinzelRessourcenExport").Value)
                awinSettings.zeilenhoehe1 = CDbl(.Range("Zeilenhoehe1").Value)
                awinSettings.zeilenhoehe2 = CDbl(.Range("Zeilenhoehe2").Value)
                awinSettings.spaltenbreite = CDbl(.Range("Spaltenbreite").Value)
                awinSettings.autoCorrectBedarfe = True
                awinSettings.propAnpassRess = False
            Catch ex As Exception
                appInstance.ScreenUpdating = formerSU
                Throw New ArgumentException("korrupte Einstellungen ... Abbruch " & ex.Message)
            End Try

            StartofCalendar = awinSettings.kalenderStart
            '
            ' ende Auslesen Einstellungen in Sheet "Einstellungen"
            '
        End With



        ' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....
        appInstance.EnableEvents = False
        appInstance.ActiveWorkbook.Close(SaveChanges:=False) ' ur: 6.5.2014 savechanges hinzugefügt
        appInstance.EnableEvents = True

        showtimezone = True

        ' jetzt werden die Projekt-Vorlagen ausgelesen 
        Dim dirName As String = awinPath & projektVorlagenOrdner
        Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirName)

        For i = 1 To listOfFiles.Count
            dateiName = listOfFiles.Item(i - 1)

            Try
                appInstance.Workbooks.Open(dateiName)
                Dim hproj As New clsProjektvorlage
                Call awinImportProject(Nothing, hproj, True, Date.Now)
                ' Auslesen der Projektvorlage wird wie das Importieren eines Projekts behandelt, nur am Ende in die Liste der Projektvorlagen eingehängt
                ' Kennzeichen für Projektvorlage ist der 3.Parameter im Aufruf (isTemplate)

                Projektvorlagen.Add(hproj)
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)

            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox(ex.Message)
            End Try


        Next

        Try
            awinSettings.lastProjektTyp = Projektvorlagen.Liste.ElementAt(0).Value.VorlagenName
        Catch ex As Exception
            awinSettings.lastProjektTyp = ""
        End Try




        ' jetzt ist wieder das Excel, das initial aufgerufen wurde - das ActiveWorkbook 
        ' hier wird die Farbe der Zeitleiste bestimmt
        ' ausserdem werden hier die Bezeichnungen der Spalten eingetragen
        appInstance.EnableEvents = False


        ' bestimmen der Spaltenbreite und Spaltenhöhe ...
        Dim testCase As String = appInstance.ActiveWorkbook.Name
        Dim wsName3 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(3)), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

        'wsName3 = CType(CType(appInstance.ActiveWorkbook, Excel.Workbook).Worksheets(arrWsNames(3)), Excel.Worksheet)

        Dim tmpRange As Excel.Range

        'With wsName3
        ''    .Activate()
        'End With

        Dim tempWSName As String = CType(appInstance.ActiveSheet, Excel.Worksheet).Name

        Dim tmpStart As Date
        Try
            With wsName3
                Dim rng As Excel.Range
                'Dim colDate As date
                If awinSettings.zeitEinheit = "PM" Then


                    ' Änderung am 16.7.2013
                    ' Eintrag des Kalenders ...
                    'Dim htxt As String = StartofCalendar.ToShortDateString
                    '.cells(zeile, spalte + 2).FormulaR1C1 = hproj.startDate.ToString("MMM yy")
                    '.cells(zeile, spalte + 3).FormulaR1C1 = hproj.startDate.AddMonths(1).ToString("MMM yy")
                    CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar
                    CType(.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(1)
                    rng = .Range(.cells(1, 1), .cells(1, 2))
                    rng.NumberFormat = "mmm-yy"

                    Dim destinationRange As Excel.Range = .Range(.Cells(1, 1), .Cells(1, 480))
                    With destinationRange
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .NumberFormat = "mmm-yy"
                        .WrapText = False
                        .Orientation = 90
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .Interior.color = noshowtimezone_color
                    End With

                    rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)


                    ' Ende Änderung 16.7 
                    'With rng
                    '
                    '    .NumberFormat = "mmm-yy"
                    'End With
                    'For i = 1 To 210
                    '    colDate = StartofCalendar.AddMonths(i - 1)
                    '    .cells(1, i).value = coldate
                    'Next
                ElseIf awinSettings.zeitEinheit = "PW" Then
                    For i = 1 To 210
                        CType(.cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).value = StartofCalendar.AddDays((i - 1) * 7)
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
                    i = 1
                    For w = 1 To 30
                        For d = 0 To 4
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


                ' hier werden jetzt die Spaltenbreiten und Zeilenhöhen gesetzt 

                Dim maxRows As Integer = .Rows.Count
                Dim maxColumns As Integer = .Columns.Count

                tmpRange = CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range)
                CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe1
                CType(.Range(.Cells(2, 1), .Cells(maxRows, maxColumns)), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = awinSettings.zeilenhoehe2
                CType(.Columns, Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = awinSettings.spaltenbreite


                .Range(.Cells(2, 1), .Cells(maxRows, maxColumns)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                .Range(.Cells(2, 1), .Cells(maxRows, maxColumns)).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter


                boxWidth = CDbl(CType(.cells(3, 3), Global.Microsoft.Office.Interop.Excel.Range).width)
                boxHeight = CDbl(CType(.cells(3, 3), Global.Microsoft.Office.Interop.Excel.Range).height)

                topOfMagicBoard = CDbl(CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range).Height) + 0.1 * boxHeight
                screen_correct = 0.1 * 19.3 / boxWidth


                Dim laenge As Integer
                laenge = showRangeRight - showRangeLeft

                If laenge > 0 And showRangeLeft > 0 Then
                    .Range(.Cells(1, showRangeLeft), .Cells(1, showRangeLeft + laenge)).Interior.Color = showtimezone_color
                End If

            End With
        Catch ex As Exception
            'Call MsgBox("oops - unerwarteter Fehler ...")
        End Try




        ' hier werden die neuen bzw. geänderten Projekte in ImportProjekte eingelesen ...
        'Call awinImportProjects()
        appInstance.EnableEvents = True


        Dim request As New Request(awinSettings.databaseName)

        ' Datenbank ist gestartet
        If request.pingMongoDb() Then

            ' alle Konstellationen laden 
            projectConstellations = request.retrieveConstellationsFromDB()


            ' hier werden jetzt auch alle Abhängigkeiten geladen 
            allDependencies = request.retrieveDependenciesFromDB()

            Dim axt As Integer = 9

            'hier wird die Start-Konfiguration gespeichert
            '5.11. ausblenden
            'Call awinStoreConstellation("Start")

            'hier werden die Projekte in die Plantafel gezeichnet 
            '5.11. ausblenden
            'Call awinZeichnePlanTafel() ' an der alten Stelle 

            'appInstance.ScreenUpdating = True
        Else
            Throw New ArgumentException("Datenbank - Verbindung ist unterbrochen ...")
        End If


    End Sub

    '
    '
    '
    Public Sub awinChangeTimeSpan(ByVal von As Integer, ByVal bis As Integer)

        'Dim k As Integer

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False


        If von < 1 Then
            von = 1
        End If

        If bis < von + 6 Then
            bis = von + 6
        End If



        If showRangeLeft <> von Or showRangeRight <> bis Or _
            AlleProjekte.Count = 0 Or _
            DeletedProjekte.Count > 0 Then


            '
            ' wenn roentgenblick.ison , werden Bedarfe angezeigt - die müssen hier ausgeblendet werden - nachher mit den neuen Werten eingeblendet werden
            '
            If roentgenBlick.isOn Then
                Call awinNoshowProjectNeeds()
            End If

            If showtimezone Then
                '
                ' aktualisieren der Showtime zone, erst die alte ausblenden , dann die neue einblenden
                '
                Call awinShowtimezone(showRangeLeft, showRangeRight, "False")
                Call awinShowtimezone(von, bis, "True")
            End If

            showRangeLeft = von
            showRangeRight = bis


            ' jetzt werden - falls nötig die Projekte nachgeladen ... 
            Try
                If ProjectBoardDefinitions.My.Settings.loadProjectsOnChange Then

                    Call awinProjekteImZeitraumLaden(awinSettings.databaseName)

                    ' jetzt sind wieder alle Projekte des Zeitraums da - deswegen muss nicht ggf nachgeladen werden 
                    DeletedProjekte.Clear()

                    '
                    '   wenn "selectedRoleNeeds" ungleich Null ist, werden Bedarfe angezeigt - die müssen hier wieder - mit den neuen Daten für show_range_lefet, .._right eingeblendet werden
                    '
                    If roentgenBlick.isOn Then
                        With roentgenBlick
                            Call awinShowProjectNeeds1(mycollection:=.myCollection, type:=.type)
                        End With
                    End If



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



        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    ''' <summary>
    '''speziell auf BMW Rplan Outpunt angepasstes Inventur Import File 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub bmwImportProjektInventur(ByRef myCollection As Collection)

        Dim zeile As Integer, spalte As Integer
        Dim pName As String = " "

        Dim lastRow As Integer

        Dim hproj As clsProjekt
        Dim vproj As clsProjektvorlage
        Dim geleseneProjekte As Integer
        Dim projektFarbe As Object
        Dim anfang As Integer, ende As Integer
        Dim cphase As clsPhase
        Dim cresult As clsResult
        Dim cbewertung As clsBewertung
        Dim ix As Integer
        Dim tmpStr(20) As String
        Dim aktuelleZeile As String
        Dim nameSopTyp As String = " "
        Dim nameBU As String
        Dim sopDate As Date
        Dim tmpStartSop As Date ' wird benutzt , um eine Hilfsphase zu machen 
        Dim startDate As Date, endDate As Date
        Dim startoffset As Integer, duration As Integer
        Dim vorlagenName As String
        Dim phaseName As String
        Dim itemName As String
        Dim zufall As New Random(10)
        Dim farbKennung As Integer
        Dim responsible As String



        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1
        geleseneProjekte = 0


        Try
            'Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Tabelle1"), _
            '                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet, _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                Dim tstStr As String
                Try
                    tstStr = activeWSListe.Cells(2, 1).value
                    projektFarbe = activeWSListe.Cells(2, 1).Interior.Color
                Catch ex As Exception
                    projektFarbe = activeWSListe.Cells(2, 1).Interior.ColorIndex
                End Try





                lastRow = System.Math.Max(CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row, _
                                          CType(.Cells(2000, 2), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row)

                While zeile <= lastRow

                    anfang = zeile + 1
                    ix = anfang

                    Do While CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color <> projektFarbe And ix <= lastRow
                        ix = ix + 1
                    Loop

                    ende = ix - 1

                    ' hier wird Name, Typ, SOP, Business Unit, vname, Start-Datum, Dauer der Phase(1) ausgelesen  
                    aktuelleZeile = activeWSListe.Cells(zeile, 2).value.Trim
                    startDate = CDate(activeWSListe.Cells(zeile, 3).value)
                    endDate = CDate(activeWSListe.Cells(zeile, 4).value)
                    farbKennung = CInt(activeWSListe.Cells(zeile, 12).value)
                    responsible = CStr(activeWSListe.Cells(zeile, 9).value)


                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    If duration < 0 Then
                        startDate = endDate
                        duration = -1 * duration
                        endDate = startDate.AddDays(duration)
                    End If

                    tmpStr = aktuelleZeile.Trim.Split(New Char() {"[", "]"}, 5)

                    Try
                        nameSopTyp = tmpStr(0).Trim
                        pName = nameSopTyp
                        Try
                            nameBU = tmpStr(1)
                            tmpStr = nameBU.Split(New Char() {" "}, 3)
                            nameBU = tmpStr(0)
                        Catch ex1 As Exception
                            nameBU = ""
                        End Try


                    Catch ex As Exception
                        Throw New Exception("Name, SOP, Typ kann nicht bestimmt werden " & vbLf & nameSopTyp)
                    End Try

                    Dim foundIX As Integer = -1

                    tmpStr = nameSopTyp.Trim.Split(New Char() {" "}, 15)
                    Dim k As Integer = 0

                    Do While foundIX < 0 And k <= tmpStr.Length - 2
                        If tmpStr(k).Trim = "SOP" And k < tmpStr.Length - 1 Then
                            Try
                                sopDate = CDate(tmpStr(k + 1)).AddMonths(1).AddDays(-1)
                                tmpStartSop = CDate(tmpStr(k + 1))
                            Catch ex As Exception
                                Dim tmp1Str(3) As String
                                tmp1Str = tmpStr(k + 1).Split(New Char() {"/"}, 8)

                                If CInt(tmp1Str(1)) < 50 Then
                                    tmp1Str(1) = CStr(2000 + CInt(tmp1Str(1)))
                                End If
                                tmpStr(k + 1) = tmp1Str(0) & "-" & tmp1Str(1)
                                sopDate = CDate(tmpStr(k + 1)).AddMonths(1).AddDays(-1)
                                tmpStartSop = CDate(tmpStr(k + 1))
                            End Try

                            foundIX = k + 2
                        Else
                            k = k + 1
                        End If
                    Loop

                    If foundIX < 0 Then
                        ' SOP Date konnte nicht bestimmt werden 
                        sopDate = endDate
                        tmpStartSop = sopDate.AddDays(-28)
                        foundIX = tmpStr.Length - 1
                    End If

                    Select Case tmpStr(foundIX).Trim
                        Case "eA"
                            vorlagenName = "Enge Ableitung"
                        Case "wA"
                            vorlagenName = "Weite Ableitung"
                        Case "E"
                            vorlagenName = "Erstanläufer"
                        Case Else
                            vorlagenName = "Erstanläufer"
                    End Select

                    '
                    ' jetzt wird das Projekt angelegt 
                    '
                    hproj = New clsProjekt

                    Try
                        vproj = Projektvorlagen.getProject(vorlagenName)


                        hproj.farbe = vproj.farbe
                        hproj.Schrift = vproj.Schrift
                        hproj.Schriftfarbe = vproj.Schriftfarbe
                        hproj.name = ""
                        hproj.VorlagenName = vorlagenName
                        hproj.earliestStart = vproj.earliestStart
                        hproj.latestStart = vproj.latestStart
                        hproj.ampelStatus = farbKennung
                        hproj.leadPerson = responsible

                    Catch ex As Exception
                        Throw New Exception("es gibt keine entsprechende Vorlage ..  " & vbLf & ex.Message)
                    End Try


                    Try

                        hproj.name = pName
                        hproj.startDate = startDate
                        hproj.earliestStartDate = hproj.startDate.AddMonths(hproj.earliestStart)
                        hproj.latestStartDate = hproj.startDate.AddMonths(hproj.latestStart)
                        If DateDiff(DateInterval.Month, startDate, Date.Now) <= 0 Then
                            hproj.Status = ProjektStatus(0)
                        Else
                            hproj.Status = ProjektStatus(1)
                        End If

                        hproj.StrategicFit = zufall.NextDouble * 10
                        hproj.Risiko = zufall.NextDouble * 10
                        hproj.volume = zufall.NextDouble * 1000000
                        hproj.complexity = zufall.NextDouble
                        hproj.businessUnit = nameBU
                        hproj.description = nameSopTyp

                        hproj.Erloes = 0.0


                    Catch ex As Exception
                        Throw New Exception("in erstelle InventurProjekte: " & vbLf & ex.Message)
                    End Try

                    ' jetzt werden all die Phasen angelegt , beginnend mit der ersten 
                    cphase = New clsPhase(parent:=hproj)
                    cphase.name = pName
                    startoffset = 0
                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    cphase.changeStartandDauer(startoffset, duration)

                    cresult = New clsResult(parent:=cphase)
                    cresult.name = "SOP"
                    cresult.setDate = sopDate

                    cbewertung = New clsBewertung
                    cbewertung.colorIndex = farbKennung
                    cbewertung.description = " .. es wurde  keine Erläuterung abgegeben .. "
                    cresult.addBewertung(cbewertung)

                    cphase.AddResult(cresult)

                    hproj.AddPhase(cphase)


                    Dim phaseIX As Integer = PhaseDefinitions.Count + 1


                    Dim pStartDate As Date
                    Dim pEndDate As Date
                    Dim ok As Boolean = True
                    Dim lastPhaseName As String = cphase.name

                    For i = anfang To ende

                        Try
                            itemName = .Cells(i, 2).value.trim
                        Catch ex As Exception
                            itemName = ""
                            ok = False
                        End Try

                        If ok Then

                            pStartDate = CDate(.Cells(i, 3).value)
                            pEndDate = CDate(.Cells(i, 4).value)
                            startoffset = DateDiff(DateInterval.Day, hproj.startDate, pStartDate)
                            duration = DateDiff(DateInterval.Day, pStartDate, pEndDate) + 1

                            If duration > 1 Then
                                ' es handelt sich um eine Phase 
                                phaseName = itemName
                                cphase = New clsPhase(parent:=hproj)
                                cphase.name = phaseName

                                If PhaseDefinitions.Contains(phaseName) Then
                                    ' nichts tun 
                                Else
                                    ' in die Phase-Definitions aufnehmen 

                                    Dim hphase As clsPhasenDefinition
                                    hphase = New clsPhasenDefinition

                                    hphase.farbe = .Cells(i, 1).Interior.Color
                                    hphase.name = phaseName
                                    hphase.UID = phaseIX
                                    phaseIX = phaseIX + 1

                                    Try
                                        PhaseDefinitions.Add(hphase)
                                    Catch ex As Exception

                                    End Try

                                End If

                                cphase.changeStartandDauer(startoffset, duration)
                                hproj.AddPhase(cphase)
                                lastPhaseName = cphase.name

                            ElseIf duration = 1 Then

                                Try
                                    ' es handelt sich um einen Meilenstein 

                                    Dim bewertungsAmpel As Integer
                                    Dim explanation As String

                                    bewertungsAmpel = CInt(.Cells(i, 12).value)
                                    explanation = CStr(.Cells(i, 1).value)

                                    cphase = hproj.getPhase(lastPhaseName)
                                    cresult = New clsResult(parent:=cphase)
                                    cbewertung = New clsBewertung



                                    If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                                        ' es gibt keine Bewertung
                                        bewertungsAmpel = 0
                                    End If

                                    ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                    With cbewertung
                                        '.bewerterName = resultVerantwortlich
                                        .colorIndex = bewertungsAmpel
                                        .datum = Date.Now
                                        .description = explanation
                                    End With

                                    With cresult
                                        .name = itemName
                                        .setDate = pEndDate
                                        If Not cbewertung Is Nothing Then
                                            .addBewertung(cbewertung)
                                        End If
                                    End With

                                    With cphase
                                        .addresult(cresult)
                                    End With
                                Catch ex As Exception

                                End Try

                                

                                
                            End If
                            

                           

                            ' handelt es sich um eine Phase oder um einen Meilenstein ? 

                           
                        End If


                    Next


                    ' jetzt muss das Projekt eingetragen werden 
                    ImportProjekte.Add(hproj)
                    myCollection.Add(hproj.name)


                    zeile = ende + 1

                    Do While CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color <> projektFarbe And zeile <= lastRow
                        zeile = zeile + 1
                    Loop

                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei BMW Projekt-Inventur " & vbLf & ex.Message & vbLf & pName)
        End Try



    End Sub



    Public Sub awinImportProjektInventur(ByRef myCollection As Collection)
        Dim zeile As Integer, spalte As Integer
        Dim pName As String
        Dim vName As String
        Dim start As Date
        Dim ende As Date
        Dim budget As Double
        Dim sfit As Double, risk As Double
        Dim description As String
        Dim businessUnit As String
        Dim lastRow As Integer
        Dim startSpalte As Integer
        Dim vglName As String
        Dim hproj As clsProjekt
        Dim vproj As clsProjektvorlage
        Dim geleseneProjekte As Integer
        Dim ProjektdauerIndays As Integer = 0

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1
        geleseneProjekte = 0


        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Liste"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                While zeile <= lastRow

                    pName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    vName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)

                    If Projektvorlagen.Liste.ContainsKey(vName) Then

                        vproj = Projektvorlagen.getProject(vName)

                        start = CDate(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        ende = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        If start <> Date.MinValue And ende <> Date.MinValue Then
                            ProjektdauerIndays = calcDauerIndays(start, ende)
                        ElseIf start <> Date.MinValue Then
                            ProjektdauerIndays = vproj.dauerInDays
                            ende = calcDatum(start, vproj.dauerInDays)
                        ElseIf ende <> Date.MinValue Then
                            ProjektdauerIndays = vproj.dauerInDays
                            start = calcDatum(ende, -vproj.dauerInDays)
                        End If

                        'startSpalte = CInt(DateDiff(DateInterval.Month, StartofCalendar, start) + 1)
                        If startSpalte < 1 Then
                            startSpalte = 1
                        End If

                        budget = CDbl(CType(.Cells(zeile, spalte + 4), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        risk = CDbl(CType(.Cells(zeile, spalte + 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        sfit = CDbl(CType(.Cells(zeile, spalte + 6), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        'volume = CDbl(CType(.Cells(zeile, spalte + 6), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        'complexity = CDbl(CType(.Cells(zeile, spalte + 7), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        businessUnit = CStr(CType(.Cells(zeile, spalte + 7), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        description = CStr(CType(.Cells(zeile, spalte + 8), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        vglName = pName.Trim & "#" & ""

                        If AlleProjekte.ContainsKey(vglName) Then
                            ' nichts tun ...
                            Call MsgBox("Projekt aus Inventur Liste existiert bereits - keine Neuanlage")
                        Else
                            'Projekt anlegen ,Verschiebung um 
                            hproj = New clsProjekt(start, start.AddMonths(-3), start.AddMonths(3))

                            Call erstelleInventurProjekt(hproj, pName, vName, start, ende, budget, zeile, sfit, risk, _
                                                         0, 0, businessUnit, description)
                            If Not hproj Is Nothing Then
                                Try
                                    ImportProjekte.Add(hproj)
                                    myCollection.Add(hproj.name)
                                Catch ex As Exception

                                End Try

                            End If

                        End If

                    Else
                        CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."

                    End If

                        zeile = zeile + 1

                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei Projekt-Inventur")
        End Try



    End Sub


    Public Sub awinImportProject(ByRef hprojekt As clsProjekt, ByRef hprojTemp As clsProjektvorlage, ByVal isTemplate As Boolean, ByVal importDatum As Date)

        Dim zeile As Integer, spalte As Integer
        Dim hproj As New clsProjekt
        Dim hwert As Integer
        Dim anzFehler As Integer = 0
        Dim ProjektdauerIndays As Integer = 0
        Dim endedateProjekt As Date


        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 1
        spalte = 1
        ' ------------------------------------------------------------------------------------------------------
        ' Einlesen der Stammdaten
        ' ------------------------------------------------------------------------------------------------------

        Try
            Dim wsGeneralInformation As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Stammdaten"), _
                Global.Microsoft.Office.Interop.Excel.Worksheet)
            With wsGeneralInformation

                .Unprotect(Password:="x")       ' Blattschutz aufheben

                ' Projekt-Name auslesen
                hproj.name = CType(.Range("Projekt_Name").Value, String)
                hproj.farbe = .Range("Projekt_Name").Interior.Color
                hproj.Schriftfarbe = .Range("Projekt_Name").Font.Color
                hproj.Schrift = CInt(.Range("Projekt_Name").Font.Size)


                ' Kurzbeschreibung, kein Problem, wenn nicht da ...
                Try
                    hproj.description = CType(.Range("ProjektBeschreibung").Value, String)
                Catch ex As Exception

                End Try


                ' Verantwortlich - kein Problem wenn nicht da 
                Try
                    hproj.leadPerson = CType(.Range("Projektleiter").Value, String)
                Catch ex As Exception

                End Try


                ' Start
                hproj.startDate = CType(.Range("StartDatum").Value, Date)

                ' Ende

                endedateProjekt = CType(.Range("EndeDatum").Value, Date)  ' Projekt-Ende für spätere Verwendung merken
                ProjektdauerIndays = calcDauerIndays(hproj.startDate, endedateProjekt)
                Dim startOffset As Integer = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(0))

                ' Budget
                Try
                    hproj.Erloes = CType(.Range("Budget").Value, Double)
                Catch ex1 As Exception

                End Try


                ' Ampel-Farbe
                hwert = CType(.Range("Bewertung").Value, Integer)

                If hwert >= 0 And hwert <= 3 Then
                    hproj.ampelStatus = hwert
                End If

                ' Ampel-Bewertung 
                hproj.ampelErlaeuterung = CType(.Range("BewertgErläuterung").Value, String)


            End With
        Catch ex As Exception
            Throw New ArgumentException("Fehler in awinImportProject, Lesen Stammdaten")
        End Try

        ' ------------------------------------------------------------------------------------------------------
        ' Einlesen der Attribute
        ' ------------------------------------------------------------------------------------------------------

        Try
            Dim wsAttribute As Excel.Worksheet
            Try
                wsAttribute = CType(appInstance.ActiveWorkbook.Worksheets("Attribute"), _
                   Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                wsAttribute = Nothing
            End Try

            If Not IsNothing(wsAttribute) Then

                With wsAttribute

                    .Unprotect(Password:="x")       ' Blattschutz aufheben


                    '   Varianten-Name
                    Try
                        hproj.variantName = CType(.Range("Variant_Name").Value, String)
                        hproj.variantName = hproj.variantName.Trim
                        If hproj.variantName.Length = 0 Then
                            hproj.variantName = Nothing
                        End If
                    Catch ex1 As Exception
                        hproj.variantName = Nothing
                    End Try


                    ' Business Unit - kein Problem wenn nicht da   
                    Try
                        hproj.businessUnit = CType(.Range("Business_Unit").Value, String)
                    Catch ex As Exception

                    End Try

                    ' Status    ist ein read-only Feld
                    ' hproj.Status = .Range("Status").Value

                    ' Risiko
                    hproj.Risiko = .Range("Risiko").Value


                    ' Strategic Fit
                    hproj.StrategicFit = .Range("Strategischer_Fit").Value


                    '' Komplexitätszahl - kein Problem, wenn nicht da  --- BMW---
                    'Try
                    '    hproj.complexity = CType(.Range("Complexity").Value, Double)
                    'Catch ex As Exception
                    '    hproj.complexity = 0.5 ' Default
                    'End Try

                    '' Volumen - kein Problem, wenn nicht da    --- BMW ---
                    'Try
                    '    hproj.volume = CType(.Range("Volume").Value, Double)
                    'Catch ex As Exception
                    '    hproj.volume = 10 ' Default
                    'End Try



                End With
            End If
        Catch ex As Exception
            Throw New ArgumentException("Fehler in awinImportProject, Lesen Attribute")
        End Try

     
        ' ------------------------------------------------------------------------------------------------------
        ' Einlesen der Ressourcen
        ' ------------------------------------------------------------------------------------------------------
        Dim wsRessourcen As Excel.Worksheet
        Try
            wsRessourcen = CType(appInstance.ActiveWorkbook.Worksheets("Ressourcen"), _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
        Catch ex As Exception
            wsRessourcen = Nothing
            ' ------------------------------------------------------------------------------------------------------
            ' Erzeugen und eintragen der Projekt-Phase (= erste Phase mit Dauer des Projekts)
            ' ------------------------------------------------------------------------------------------------------
            Try
                Dim cphase As New clsPhase(hproj)

                ' ProjektPhase wird erzeugt
                cphase = New clsPhase(parent:=hproj)
                cphase.name = hproj.name

                ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                With cphase
                    .name = hproj.name
                    Dim startOffset As Integer = 0
                    .changeStartandDauer(startOffset, ProjektdauerIndays)
                End With
                ' ProjektPhase wird hinzugefügt
                hproj.AddPhase(cphase)

            Catch ex1 As Exception
                Throw New ArgumentException("Fehler in awinImportProject, Erzeugen ProjektPhase")
            End Try

        End Try

        If Not IsNothing(wsRessourcen) Then

            Try
                With wsRessourcen
                    Dim rng As Excel.Range
                    Dim zelle As Excel.Range
                    Dim chkPhase As Boolean = True
                    Dim chkRolle As Boolean = True
                    Dim firsttime As Boolean = False
                    Dim added As Boolean = True
                    Dim Xwerte As Double()
                    Dim crole As clsRolle
                    Dim cphase As New clsPhase(hproj)
                    Dim ccost As clsKostenart
                    Dim phaseName As String = ""

                    Dim anfang As Integer, ende As Integer  ', projDauer As Integer

                    Dim farbeAktuell As Object
                    Dim r As Integer, k As Integer


                    .Unprotect(Password:="x")       ' Blattschutz aufheben


                    Dim tmpws As Excel.Range = CType(wsRessourcen.Range("Phasen_des_Projekts"), Excel.Range)

                    rng = .Range("Phasen_des_Projekts")

                    If CStr(.Range("Phasen_des_Projekts").Cells(1).value) <> hproj.name Then

                        ' ProjektPhase wird hinzugefügt
                        cphase = New clsPhase(parent:=hproj)
                        added = False
                        phaseName = hproj.name

                        ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                        With cphase
                            .name = phaseName
                            Dim startOffset As Integer = 0
                            .changeStartandDauer(startOffset, ProjektdauerIndays)
                            Dim phaseStartdate As Date = .getStartDate
                            Dim phaseEnddate As Date = .getEndDate
                            'projDauer = calcDauerIndays(phaseStartdate, phaseEnddate)
                            firsttime = True
                        End With
                        'Call MsgBox("Projektnamen/Phasen Konflikt in awinImportProjekt" & vbLf & "Problem wurde behoben")

                    End If

                    zeile = 0

                    For Each zelle In rng

                        zeile = zeile + 1

                        ' nachsehen, ob Phase angegeben oder Rolle/Kosten
                        If Len(CType(zelle.Value, String)) > 1 Then
                            phaseName = CType(zelle.Value, String).Trim
                        Else
                            phaseName = ""
                        End If

                        ' hier wird die Rollen bzw Kosten Information ausgelesen
                        Dim hname As String
                        Try
                            hname = CType(zelle.Offset(0, 1).Value, String).Trim
                        Catch ex1 As Exception
                            hname = ""
                        End Try

                        If Len(phaseName) > 0 And Len(hname) <= 0 Then
                            chkPhase = True
                            chkRolle = False
                            If Not firsttime Then
                                firsttime = True
                            End If
                        End If

                        If Len(phaseName) <= 0 And Len(hname) > 0 Then
                            If zeile = 1 Then
                                Call MsgBox(" es fehlt die ProjektPhase")
                            Else
                                chkPhase = False
                                chkRolle = True
                            End If
                        Else
                        End If

                        If Len(phaseName) > 0 And Len(hname) > 0 Then
                            chkPhase = True
                            chkRolle = True
                        End If

                        If Len(phaseName) <= 0 And Len(hname) <= 0 Then
                            chkPhase = False
                            chkRolle = False
                            ' beim 1.mal: abspeichern der letzten Phase mit Ihren Rollen
                            ' beim 2.mal: for - Schleife abbrechen
                        End If

                        Select Case chkPhase
                            Case True
                                If Not added Then
                                    hproj.AddPhase(cphase)
                                End If

                                cphase = New clsPhase(parent:=hproj)
                                added = False

                                ' Auslesen der Phasen Dauer
                                anfang = 1  ' anfang enthält den rel.Anfang einer Phase
                                Try
                                    While CInt(zelle.Offset(0, anfang + 1).Interior.ColorIndex) = -4142 And
                                        Not (CType(zelle.Offset(0, anfang + 1).Value, String) = "x")
                                        anfang = anfang + 1
                                    End While
                                Catch ex As Exception
                                    Throw New ArgumentException("Es wurden keine oder falsche Angaben zur Phasendauer der Phase '" & phaseName & "' gemacht." & vbLf &
                                                                "Bitte überprüfen Sie dies.")
                                End Try

                                ende = anfang + 1

                                If CInt(zelle.Offset(0, anfang + 1).Interior.ColorIndex) = -4142 Then
                                    While CType(zelle.Offset(0, ende + 1).Value, String) = "x"
                                        ende = ende + 1
                                    End While
                                    ende = ende - 1
                                Else
                                    farbeAktuell = zelle.Offset(0, anfang + 1).Interior.Color
                                    While CInt(zelle.Offset(0, ende + 1).Interior.Color) = CInt(farbeAktuell)

                                        ende = ende + 1
                                    End While
                                    ende = ende - 1
                                End If

                                With cphase
                                    .name = phaseName
                                    ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                                    Dim startOffset As Integer
                                    Dim dauerIndays As Integer
                                    startOffset = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(anfang - 1))
                                    dauerIndays = calcDauerIndays(hproj.startDate.AddDays(startOffset), ende - anfang + 1, True)

                                    .changeStartandDauer(startOffset, dauerIndays)
                                    .Offset = 0

                                    ' hier muss eine Routine aufgerufen werden, die die Dauer in Tagen berechnet !!!!!!
                                    Dim phaseStartdate As Date = .getStartDate
                                    Dim phaseEnddate As Date = .getEndDate

                                End With
                                Select Case chkRolle
                                    Case True
                                        Throw New ArgumentException("Rollen/Kosten-Bedarfe zur Phase '" & phaseName & "' bitte in die darauffolgenden Zeilen eintragen")
                                    Case False  ' es wurde nur eine Phase angegeben: korrekt

                                End Select

                            Case False ' auslesen Rollen- bzw. Kosten-Information

                                Select Case chkRolle
                                    Case True
                                        ' hier wird die Rollen bzw Kosten Information ausgelesen
                                        '
                                        ' entweder nun Rollen/Kostendefinition oder Ende der Phasen
                                        '
                                        If RoleDefinitions.Contains(hname) Then
                                            Try
                                                r = CInt(RoleDefinitions.getRoledef(hname).UID)

                                                ReDim Xwerte(ende - anfang)


                                                For m = anfang To ende

                                                    Try
                                                        Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
                                                    Catch ex As Exception
                                                        Xwerte(m - anfang) = 0.0
                                                    End Try

                                                Next m

                                                crole = New clsRolle(ende - anfang + 1)
                                                With crole
                                                    .RollenTyp = r
                                                    .Xwerte = Xwerte
                                                End With

                                                With cphase
                                                    .AddRole(crole)
                                                End With
                                            Catch ex As Exception
                                                '
                                                ' handelt es sich um die Kostenart Definition?
                                                '
                                            End Try

                                        ElseIf CostDefinitions.Contains(hname) Then

                                            Try

                                                k = CInt(CostDefinitions.getCostdef(hname).UID)

                                                ReDim Xwerte(ende - anfang)

                                                For m = anfang To ende
                                                    Try
                                                        Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + 1).Value)
                                                    Catch ex As Exception
                                                        Xwerte(m - anfang) = 0.0
                                                    End Try

                                                Next m

                                                ccost = New clsKostenart(ende - anfang + 1)
                                                With ccost
                                                    .KostenTyp = k
                                                    .Xwerte = Xwerte
                                                End With


                                                With cphase
                                                    .AddCost(ccost)
                                                End With

                                            Catch ex As Exception

                                            End Try

                                        End If

                                    Case False  ' es wurde weder Phase noch Rolle angegeben. 
                                        If firsttime Then
                                            firsttime = False
                                        Else 'beim 2. mal: letzte Phase hinzufügen; ENDE von For-Schleife for each Zelle
                                            hproj.AddPhase(cphase)
                                            Exit For
                                        End If

                                End Select

                        End Select

                    Next zelle


                End With
            Catch ex As Exception
                Throw New ArgumentException("Fehler in awinImportProject, Lesen Ressourcen von '" & hproj.name & "' " & vbLf & ex.Message)
            End Try

        End If

        '' hier wurde jetzt die Reihenfolge geändert - erst werden die Phasen Definitionen eingelesen ..

        '' jetzt werden die Daten für die Phasen sowie die Termine/Deliverables eingelesen 

        Try
            Dim wsTermine As Excel.Worksheet
            Try
                wsTermine = CType(appInstance.ActiveWorkbook.Worksheets("Termine"), _
                                                             Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                wsTermine = Nothing
            End Try

            If Not IsNothing(wsTermine) Then
                Try
                    With wsTermine
                        Dim lastrow As Integer
                        Dim lastcolumn As Integer
                        Dim phaseName As String
                        Dim resultName As String
                        Dim resultDate As Date
                        Dim resultVerantwortlich As String = ""
                        Dim bewertungsAmpel As Integer
                        Dim explanation As String
                        Dim bewertungsdatum As Date = importDatum
                        Dim Nummer As String
                        Dim tbl As Excel.Range
                        Dim sortBereich As Excel.Range
                        Dim sortKey As Excel.Range
                        Dim rowOffset As Integer
                        Dim columnOffset As Integer


                        .Unprotect(Password:="x")       ' Blattschutz aufheben

                        tbl = .ListObjects("ErgebnTabelle").Range
                        rowOffset = tbl.Row             ' ist die erste Zeile der ErgebnTabelle = Überschriftszeile
                        columnOffset = tbl.Column

                        ' hiermit soll die Tabelle der Termine nach der laufenden Nummer sortiert werden

                        lastrow = CInt(.Cells(2000, columnOffset).End(XlDirection.xlUp).row)
                        lastcolumn = CInt(.Cells(rowOffset, 2000).End(XlDirection.xlToLeft).column)

                        'sortBereich ist der Inhalt der ErgebnTabelle
                        sortBereich = .Range(.Cells(rowOffset + 1, columnOffset), .Cells(lastrow, lastcolumn))
                        ' sortKey ist die erste Spalte der ErgebnTabelle
                        sortKey = .Range(.Cells(rowOffset + 1, columnOffset), .Cells(lastrow, columnOffset))

                        With .Sort
                            ' Bestehende Sortierebenen löschen
                            .SortFields.Clear()
                            ' Sortierung nach der laufenden Nummer in der ErgebnTabelle also erste Spalte 
                            .SortFields.Add(Key:=sortKey, Order:=XlSortOrder.xlAscending)
                            .SetRange(sortBereich)
                            .Apply()
                        End With

                        For zeile = rowOffset + 1 To lastrow


                            Dim cResult As clsResult
                            Dim cBewertung As clsBewertung
                            Dim cphase As clsPhase
                            Dim objectName As String
                            Dim startDate As Date, endeDate As Date
                            Dim bezug As String


                            Dim isPhase As Boolean = False
                            Dim isMeilenstein As Boolean = False
                            Dim cphaseExisted As Boolean = True

                            Try
                                ' Wenn es keine Phasen gibt in diesem Projekt, so wird trotzdem die Phase1, die ProjektPhase erzeugt.

                                If hproj.AllPhases.Count = 0 Then
                                    Dim duration As Integer
                                    Dim offset As Integer

                                    ' Erzeuge ProjektPhase mit Länge des Projekts
                                    cphase = New clsPhase(parent:=hproj)
                                    cphase.name = hproj.name
                                    'cphaseExisted = False       ' Phase existiert noch nicht

                                    offset = 0

                                    If ProjektdauerIndays < 1 Or offset < 0 Then
                                        Throw New Exception("unzulässige Angaben für Offset und Dauer: " & _
                                                            offset.ToString & ", " & duration.ToString)
                                    End If

                                    cphase.changeStartandDauer(offset, ProjektdauerIndays)
                                    hproj.AddPhase(cphase)

                                End If                            'Phase 1 ist nun angelegt


                                Try
                                    Nummer = CType(.Cells(zeile, columnOffset).value, String).Trim
                                Catch ex As Exception
                                    Nummer = Nothing
                                    Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                End Try

                                Try
                                    ' bestimme, worum es sich handelt: Phase oder Meilenstein
                                    objectName = CType(.Cells(zeile, columnOffset + 1).value, String).Trim
                                Catch ex As Exception
                                    objectName = Nothing
                                    Throw New Exception("In Tabelle 'Termine' ist der PhasenName nicht angegeben ")
                                    Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                End Try


                                If PhaseDefinitions.Contains(objectName) Then
                                    isPhase = True
                                    isMeilenstein = False
                                Else
                                    If hproj.name = objectName Then
                                        isPhase = True
                                        isMeilenstein = False
                                    Else
                                        isPhase = False
                                        isMeilenstein = True
                                    End If
                                End If


                                Try
                                    bezug = CType(.Cells(zeile, columnOffset + 2).value, String).Trim
                                Catch ex As Exception
                                    bezug = Nothing
                                End Try

                                Try
                                    startDate = CDate(.Cells(zeile, columnOffset + 3).value)
                                Catch ex As Exception
                                    startDate = Date.MinValue
                                End Try

                                endeDate = CDate(.Cells(zeile, columnOffset + 4).value)

                                If DateDiff(DateInterval.Month, hproj.startDate, startDate) < 0 Then
                                    ' kein Startdatum angegeben

                                    If startDate <> Date.MinValue Then
                                        cphase = Nothing
                                        Throw New Exception("Die Phase '" & objectName & "' beginnt vor dem Projekt !" & vbLf &
                                                     "Bitte korrigieren Sie dies in der Datei'" & hproj.name & ".xlsx'")
                                    Else
                                        ' objectName ist ein Meilenstein
                                        cphase = hproj.getPhase(bezug)
                                        If IsNothing(cphase) Then
                                            If hproj.AllPhases.Count > 0 Then
                                                cphase = hproj.getPhase(1)
                                            Else
                                                ' Erzeuge ProjektPhase mit Länge des Projekts


                                            End If

                                        End If
                                    End If

                                    'isPhase = False

                                Else
                                    'objectName ist eine Phase
                                    'isPhase = True

                                    ' ist der Phasen Name in der Liste der definitionen überhaupt bekannt ? 
                                    If Not PhaseDefinitions.Contains(objectName) Then

                                        ' jetzt noch prüfen, ob es sich um die Phase (1) handelt, dann kann sie ja nicht in der PhaseDefinitions enthalten sein  ..
                                        If hproj.name <> objectName Then
                                            Throw New Exception("Phase '" & objectName & "' ist nicht definiert!" & vbLf &
                                                           "Bitte löschen Sie diese Phase aus '" & hproj.name & "'.xlsx, Tabellenblatt 'Termine'")
                                        End If

                                    End If

                                    ' an dieser stelle ist sichergestellt, daß der Phasen Name bekannt ist
                                    ' Prüfen, ob diese Phase bereits in hproj über das ressourcen Sheet angelegt wurde 
                                    cphase = hproj.getPhase(objectName)
                                    If IsNothing(cphase) Then
                                        cphase = New clsPhase(parent:=hproj)
                                        cphase.name = objectName
                                        cphaseExisted = False       ' Phase existiert noch nicht
                                    End If
                                End If

                                If isPhase Then  'xxxx Phase
                                    Try

                                        Dim duration As Integer
                                        Dim offset As Integer



                                        duration = calcDauerIndays(startDate, endeDate)
                                        offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)


                                        If duration < 1 Or offset < 0 Then
                                            Throw New Exception("unzulässige Angaben für Offset und Dauer: " & _
                                                                offset.ToString & ", " & duration.ToString)
                                        End If

                                        cphase.changeStartandDauer(offset, duration)

                                        ' jetzt wird auf Inkonsistenz geprüft 
                                        Dim inkonsistent As Boolean = False

                                        If cphase.CountRoles > 0 Or cphase.CountCosts > 0 Then
                                            ' prüfen , ob es Inkonsistenzen gibt ? 
                                            For r = 1 To cphase.CountRoles
                                                If cphase.getRole(r).Xwerte.Length <> cphase.relEnde - cphase.relStart + 1 Then
                                                    inkonsistent = True
                                                End If
                                            Next

                                            For k = 1 To cphase.CountCosts
                                                If cphase.getCost(k).Xwerte.Length <> cphase.relEnde - cphase.relStart + 1 Then
                                                    inkonsistent = True
                                                End If
                                            Next
                                        End If

                                    If inkonsistent Then
                                        anzFehler = anzFehler + 1
                                        Throw New Exception("Der Import konnte nicht fertiggestellt werden. " & vbLf & "Die Dauer der Phase '" & cphase.name & "'  in 'Termine' ist ungleich der in 'Ressourcen' " & vbLf &
                                                             "Korrigieren Sie bitte gegebenenfalls diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
                                    End If
                                    If Not cphaseExisted Then
                                        hproj.AddPhase(cphase)
                                    End If


                                    Catch ex As Exception
                                        Throw New Exception(ex.Message)
                                    End Try

                                Else


                                    Try
                                        ' wenn kein Datum angegeben wurde, soll das Ende der Phase als Datum angenommen werden 
                                        If DateDiff(DateInterval.Month, hproj.startDate, resultDate) < -1 Then
                                            resultDate = hproj.startDate.AddDays(cphase.startOffsetinDays + cphase.dauerInDays - 1)
                                        End If
                                        'xxxxxx
                                        phaseName = cphase.name
                                        cResult = New clsResult(parent:=cphase)
                                        cBewertung = New clsBewertung

                                        resultName = objectName.Trim
                                        resultDate = endeDate

                                        ' wenn kein Datum angegeben wurde, soll das Ende der Phase als Datum angenommen werden 
                                        If DateDiff(DateInterval.Month, hproj.startDate, resultDate) < -1 Then
                                            resultDate = hproj.startDate.AddDays(cphase.startOffsetinDays + cphase.dauerInDays)
                                        Else
                                            If DateDiff(DateInterval.Day, endedateProjekt, endeDate) > 0 Then
                                                Call MsgBox("der Meilenstein '" & resultName & "' liegt später als das Ende des gesamten Projekts" & vbLf &
                                                            "Bitte korrigieren Sie dies im Tabellenblatt Ressourcen der Datei '" & hproj.name & ".xlsx")
                                            End If

                                        End If

                                        ' resultVerantwortlich = CType(.Cells(zeile, 5).value, String)
                                        bewertungsAmpel = CType(.Cells(zeile, columnOffset + 5).value, Integer)
                                        explanation = CType(.Cells(zeile, columnOffset + 6).value, String)


                                        If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                                            ' es gibt keine Bewertung
                                            bewertungsAmpel = 0
                                        End If
                                        ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                        With cBewertung
                                            '.bewerterName = resultVerantwortlich
                                            .colorIndex = bewertungsAmpel
                                            .datum = importDatum
                                            .description = explanation
                                        End With



                                        With cResult
                                            .setDate = resultDate
                                            '.verantwortlich = resultVerantwortlich
                                            .name = resultName
                                            If Not cBewertung Is Nothing Then
                                                .addBewertung(cBewertung)
                                            End If
                                        End With

                                        With hproj.getPhase(phaseName)
                                            .addresult(cResult)
                                        End With


                                    Catch ex As Exception
                                        ' Schreiben des Fehlers in das Fehlerprotokoll - muss noch ergänzt werden 
                                        anzFehler = anzFehler + 1
                                    End Try

                                End If

                            Catch ex As Exception
                                ' letzte belegte Zeile wurde bereits bearbeitet.
                                zeile = lastrow + 1 ' erzwingt das Ende der For - Schleife
                                Nummer = Nothing
                                Throw New Exception(ex.Message)
                            End Try

                        Next

                    End With
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try

            End If
            If anzFehler > 0 Then
                Call MsgBox("Anzahl Fehler bei Import der Termine von " & hproj.name & " : " & anzFehler)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

            If isTemplate Then
                ' hier müssen die Werte für die Vorlage übergeben werden.
                Dim projVorlage As New clsProjektvorlage
                projVorlage.VorlagenName = hproj.name
                projVorlage.Schrift = hproj.Schrift
                projVorlage.Schriftfarbe = hproj.Schriftfarbe
                projVorlage.farbe = hproj.farbe
                projVorlage.earliestStart = -6
                projVorlage.latestStart = 6
                projVorlage.AllPhases = hproj.AllPhases
                hprojTemp = projVorlage

            Else
                hprojekt = hproj
            End If

    End Sub


    ''' <summary>
    ''' lädt die jeweils letzten PName/Variante Projekte aus MongoDB in alleProjekte
    ''' lädt ausserdem alle definierten Konstellationen
    ''' zeigt dann die letzte (last) an 
    ''' </summary>
    ''' <remarks></remarks>
    Sub awinletzteKonstellationLaden(ByVal databaseName As String)

        'Dim allProjectsList As SortedList(Of String, clsProjekt)
        Dim zeitraumVon As Date = StartofCalendar.AddMonths(showRangeLeft - 1)
        Dim zeitraumbis As Date = StartofCalendar.AddMonths(showRangeRight - 1)
        Dim storedHeute As Date = Now
        Dim storedGestern As Date = storedHeute.AddDays(-1)
        Dim pname As String = ""
        Dim variantName As String = ""
        Dim request As New Request(databaseName)
        Dim lastConstellation As New clsConstellation
        Dim hproj As clsProjekt

        If request.pingMongoDb() Then

            projectConstellations = request.retrieveConstellationsFromDB()

            ' Showprojekte leer machen 
            Try
                'NoShowProjekte.Clear()
                ShowProjekte.Clear()
                lastConstellation = projectConstellations.getConstellation("Last")
            Catch ex As Exception
                'Call MsgBox("in awinProjekteInitialLaden Fehler ...")
            End Try

            ' jetzt Showprojekte aufbauen - und zwar so, dass Konstellation <Last> wiederhergestellt wird
            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In lastConstellation.Liste

                Try
                    hproj = AlleProjekte(kvp.Key)
                    hproj.startDate = kvp.Value.Start
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
    Sub awinProjekteImZeitraumLaden(ByVal databaseName As String)

        Dim zeitraumVon As Date = StartofCalendar.AddMonths(showRangeLeft - 1)
        Dim zeitraumbis As Date = StartofCalendar.AddMonths(showRangeRight - 1)
        Dim storedHeute As Date = Now
        Dim storedGestern As Date = storedHeute.AddDays(-1)
        Dim pname As String = ""
        Dim variantName As String = ""
        Dim request As New Request(databaseName)
        Dim lastConstellation As New clsConstellation
        Dim projekteImZeitraum As New SortedList(Of String, clsProjekt)
        Dim projektHistorie As New clsProjektHistorie
        Dim laengeInTagen As Integer

        If request.pingMongoDb() Then

            projekteImZeitraum = request.retrieveProjectsFromDB(pname, variantName, zeitraumVon, zeitraumbis, storedGestern, storedHeute, True)
        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen")
        End If

        If AlleProjekte.Count > 0 Then
            ' prüfen, welche bereits geladen sind, welche nicht ...

            For Each kvp As KeyValuePair(Of String, clsProjekt) In projekteImZeitraum

                Try
                    laengeInTagen = kvp.Value.dauerInDays
                    Dim keyStr As String = kvp.Value.name & "#" & kvp.Value.variantName
                    AlleProjekte.Add(keyStr, kvp.Value)

                    ShowProjekte.Add(kvp.Value)

                    Call awinCreateBudgetWerte(kvp.Value)
                    Call ZeichneProjektinPlanTafel(kvp.Value.name, kvp.Value.tfZeile)

                Catch ex As Exception
                    ' nichts tun - das Projekt ist einfach nur schon da .... 

                End Try

            Next

        Else
            AlleProjekte = projekteImZeitraum
            ShowProjekte.Clear()
            ' ShowProjekte aufbauen

            For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte

                Try
                    laengeInTagen = kvp.Value.dauerInDays
                    ShowProjekte.Add(kvp.Value)

                    Call awinCreateBudgetWerte(kvp.Value)

                    Call ZeichneProjektinPlanTafel(kvp.Value.name, kvp.Value.tfZeile)

                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try

            Next

        End If


    End Sub
    ''' <summary>
    ''' lädt ein bestimmtes Portfolio von der Datenbank und zeigt es  
    ''' in der Projekttafel an.
    ''' 
    ''' </summary>
    ''' <param name="constellationName">
    ''' Name, unter dem das Portfolio in der Datenbank gespeichert wurde 
    ''' </param>
    ''' <remarks></remarks>
    ''' 
    Public Sub awinLoadConstellation(ByVal constellationName As String, ByRef successMessage As String)
        Dim activeConstellation As New clsConstellation
        Dim hproj As New clsProjekt
        Dim request As New Request(awinSettings.databaseName)

        


        ' prüfen, ob diese Constellation bereits existiert ..
        Try
            activeConstellation = projectConstellations.getConstellation(constellationName)
        Catch ex As Exception
            Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
            Exit Sub
        End Try

        ' die aktuelle Konstellation in "Last" speichern 
        Call awinStoreConstellation("Last")

        ' jetzt wird die activeConstellation in ShowProjekte bzw. NoShowProjekte umgesetzt 
        ' dazu werden erst mal alle Projekte in Showprojekte in Noshowprojekte verschoben ...

        'For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
        '    NoShowProjekte.Add(kvp.Value)
        'Next
        ShowProjekte.Clear()
        ' jetzt werden die Start-Values entsprechend gesetzt ..

        For Each kvp As KeyValuePair(Of String, clsConstellationItem) In activeConstellation.Liste

            If AlleProjekte.ContainsKey(kvp.Key) Then
                ' Projekt ist bereits im Hauptspeicher geladen
                hproj = AlleProjekte(kvp.Key)
            Else
                If request.pingMongoDb() Then

                    ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                    hproj = request.retrieveOneProjectfromDB(kvp.Value.projectName, kvp.Value.variantName)

                    ' Projekt muss nun in die Liste der geladenen Projekte eingetragen werden
                    AlleProjekte.Add(kvp.Key, hproj)

                Else
                    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & kvp.Value.projectName & "'konnte nicht geladen werden")
                End If
            End If

            With hproj

                ' Änderung THOMAS Start 
                If .Status = ProjektStatus(0) Then
                    .startDate = kvp.Value.Start
                ElseIf .startDate <> kvp.Value.Start Then
                    ' wenn das Datum nicht angepasst werden kann, weil das Projekt bereits beauftragt wurde  
                    successMessage = successMessage & vbLf & hproj.name & ": " & kvp.Value.Start.ToShortDateString
                End If
                ' Änderung THOMAS Ende 

                .StartOffset = 0
                .tfZeile = kvp.Value.zeile
            End With

            If kvp.Value.show Then

                Try
                    ShowProjekte.Add(hproj)

                    Dim pname As String
                    Dim tryzeile As Integer
                    With hproj
                        pname = .name
                        tryzeile = .tfZeile
                    End With
                    ' nicht zeichnen - das wird nachher alles auf einen Schlag erledigt ..
                    'Call ZeichneProjektinPlanTafel(pname, tryzeile)

                    'NoShowProjekte.Remove(hproj.name)
                Catch ex1 As Exception
                    Call MsgBox("Fehler in awinLoadConstellation aufgetreten: " & ex1.Message)
                End Try

            End If

        Next

        
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
    Public Sub awinRemoveConstellation(ByVal constellationName As String)

        Dim activeConstellation As New clsConstellation
        Dim request As New Request(awinSettings.databaseName)

        ' prüfen, ob diese Constellation überhaupt existiert ..
        Try
            activeConstellation = projectConstellations.getConstellation(constellationName)
        Catch ex As Exception
            Call MsgBox(" Projekt-Konstellation " & constellationName & " existiert nicht ")
            Exit Sub
        End Try

        If request.pingMongoDb() Then

            ' Konstellation muss aus der Datenbank gelöscht werden.

            If request.removeConstellationFromDB(activeConstellation) Then

                Try
                    ' Konstellation muss aus der Liste aller Portfolios entfernt werden.
                    projectConstellations.Remove(activeConstellation.constellationName)
                Catch ex1 As Exception
                    Call MsgBox("Fehler in awinRemoveConstellation aufgetreten: " & ex1.Message)
                End Try
            Else
                Call MsgBox("Es ist ein Fehler beim Löschen es Portfolios aus der Datenbank aufgetreten ")
            End If

        Else
            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekt '" & activeConstellation.constellationName & "'konnte nicht geladen werden")
        End If

        'Try
        '    ' Konstellation muss aus der Liste aller Portfolios entfernt werden.
        '    projectConstellations.Remove(activeConstellation.constellationName)
        'Catch ex1 As Exception
        '    Call MsgBox("Fehler in awinRemoveConstellation aufgetreten: " & ex1.Message)
        'End Try


    End Sub
    ' ''' <summary>
    ' ''' 
    ' ''' </summary>
    ' ''' <param name="constellationName"></param>
    ' ''' <remarks></remarks>
    'Public Sub awinStoreConstellation(ByVal constellationName As String)

    '    Dim request As New Request(awinSettings.databaseName)
    '    ' prüfen, ob diese Constellation bereits existiert ..
    '    If projectConstellations.Contains(constellationName) Then

    '        Try
    '            projectConstellations.Remove(constellationName)
    '        Catch ex As Exception

    '        End Try

    '    End If

    '    Dim newC As New clsConstellation
    '    With newC
    '        .constellationName = constellationName
    '    End With

    '    Dim newConstellationItem As clsConstellationItem
    '    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
    '        newConstellationItem = New clsConstellationItem
    '        With newConstellationItem
    '            .projectName = kvp.Key
    '            .show = True
    '            .Start = kvp.Value.startDate
    '            .variantName = kvp.Value.variantName
    '            .zeile = kvp.Value.tfZeile
    '        End With
    '        newC.Add(newConstellationItem)
    '    Next


    '    Try
    '        projectConstellations.Add(newC)

    '    Catch ex As Exception
    '        Call MsgBox("Fehler bei Add projectConstellations in awinStoreConstellations")
    '    End Try

    '    ' Portfolio in die Datenbank speichern
    '    If request.pingMongoDb() Then
    '        If Not request.storeConstellationToDB(newC) Then
    '            Call MsgBox("Fehler beim Speichern der projektConstellation '" & newC.constellationName & "' in die Datenbank")
    '        End If
    '    Else
    '        Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!")
    '    End If

    'End Sub


    

    ''' <summary>
    ''' erzeugt die Excel Datei mit den Projekt-Ressourcen Zuordnungen 
    ''' Vorbedingung Ressourcen Datei ist bereits geöffnet
    ''' 
    ''' </summary>
    ''' <param name="typus">
    ''' 0: alle Ressourcen in einer Datei ; 1: pro Rolle eine Datei ; 2: pro Kostenart eine Datei 
    ''' </param>
    ''' <param name="qualifier">
    ''' gibt den Bezeichner der Rolle / Kostenart an 
    ''' </param>
    ''' <remarks></remarks>
    Public Sub awinExportRessZuordnung(ByVal typus As Integer, ByVal qualifier As String)

        Dim anzRollen As Integer
        Dim i As Integer, m As Integer
        Dim heute As Date = Date.Now
        Dim heuteColumn As Integer
        Dim currentRole As String = " "
        Dim kapaValues() As Double
        Dim currentColor As Long
        Dim zeile As Integer = 1
        Dim zeitSpanne As Integer = 6
        Dim rng As Excel.Range, destinationRange As Excel.Range
        Dim bedarfsWerte() As Double
        Dim projWerte() As Double
        Dim mycollection As New Collection
        Dim statusColor As Long = awinSettings.AmpelNichtBewertet
        Dim statusValue As Double = 0.0
        Dim xlsBlattname(2) As String
        Dim colPointer As Integer = 2
        Dim loopi As Integer = 1
        'Dim currentColumn As Integer = 1
        Dim vorausschau As Integer = 3
        Dim cellFormula As String
        Dim personalrange As Excel.Range
        Dim rngSource As Excel.Range
        Dim rngTarget As Excel.Range
        Dim rcol As Integer
        Dim anzPeople As Integer


        Dim startZeile As Integer, endZeile As Integer
        Dim tmpDate As Date


        xlsBlattname(0) = "Summary"
        xlsBlattname(1) = "Zuordnung"
        xlsBlattname(2) = "Kapazität"

        ReDim bedarfsWerte(zeitSpanne - 1)
        ReDim kapaValues(zeitSpanne - 1)

        If typus = 0 Then
            anzRollen = RoleDefinitions.Count
        Else
            anzRollen = 1
        End If

        heuteColumn = getColumnOfDate(heute) + 1


        ' ----------------------------------------------------------
        ' Schreiben Summary 
        '-----------------------------------------------------------
        Try

            With CType(appInstance.Worksheets(xlsBlattname(0)), Global.Microsoft.Office.Interop.Excel.Worksheet)


                ' Löschen der alten Werte 
                rng = .Range(.Cells(2, 1), .Cells(2002, 21))
                rng.Clear()

                ' Schreiben der betrachteten Monate in zeile 1
                If typus = 0 Then
                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Rolle"
                    CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = "Projekt"
                Else
                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = " "
                    CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = "Projekt"
                End If


                If awinSettings.zeitEinheit = "PM" Then
                    m = 1
                    CType(.Cells(zeile, 5), Global.Microsoft.Office.Interop.Excel.Range).Value = heute.AddMonths(m)
                    CType(.Cells(zeile, 6), Global.Microsoft.Office.Interop.Excel.Range).Value = heute.AddMonths(m + 1)
                    rng = .Range(.Cells(zeile, 5), .Cells(zeile, 6))


                    With rng
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        .NumberFormat = "mmm-yy"
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .ColumnWidth = 10
                    End With


                    destinationRange = .Range(.Cells(zeile, 5), .Cells(zeile, 5 + zeitSpanne - 1))

                    With destinationRange
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        .NumberFormat = "mmm-yy"
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = Excel.Constants.xlContext
                        .MergeCells = False
                        .ColumnWidth = 10
                    End With

                    rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)



                ElseIf awinSettings.zeitEinheit = "PW" Then
                ElseIf awinSettings.zeitEinheit = "PT" Then

                End If
                zeile = 2

                Dim sumrangeAnfang As Integer, sumrangeEnde As Integer

                For i = 1 To anzRollen

                    zeile = zeile + 1
                    currentRole = " "
                    If typus = 0 Then
                        With RoleDefinitions.getRoledef(i)
                            currentRole = .name
                            For m = 0 To zeitSpanne - 1
                                kapaValues(m) = .kapazitaet(m + heuteColumn)
                            Next
                            currentColor = CLng(.farbe)
                        End With
                    ElseIf typus = 1 Then
                        Try
                            With RoleDefinitions.getRoledef(qualifier)
                                currentRole = .name
                                For m = 0 To zeitSpanne - 1
                                    kapaValues(m) = .kapazitaet(m + heuteColumn)
                                Next
                                currentColor = CLng(.farbe)
                            End With
                        Catch ex As Exception

                        End Try
                    Else
                        Call MsgBox("Kostenarten noch nicht definiert ...")
                        Exit Sub
                    End If


                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = currentRole
                    ' jetzt wird der Bereich mit hellgrauer Farbe abgesetzt 
                    rng = .Range(.Cells(zeile - 1, 1), .Cells(zeile, 1 + 5 + zeitSpanne - 2))
                    rng.Interior.Color = awinSettings.AmpelNichtBewertet

                    zeile = zeile + 1

                    mycollection.Add(currentRole)
                    sumrangeAnfang = zeile

                    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                        ' benötigt dieses Projekt in den nächsten Monaten die Rolle <currentRole>


                        If kvp.Value.Start <= getColumnOfDate(Date.Now) + zeitSpanne And _
                            kvp.Value.Start + kvp.Value.Dauer - 1 >= getColumnOfDate(Date.Now) + 1 And _
                            kvp.Value.Status <> ProjektStatus(3) And _
                            kvp.Value.Status <> ProjektStatus(4) Then

                            With kvp.Value
                                'statusValue = 
                                ReDim bedarfsWerte(zeitSpanne - 1)
                                ReDim projWerte(.Dauer - 1)
                                projWerte = .getBedarfeInMonths(mycollection, DiagrammTypen(1))

                                Dim aix As Integer
                                aix = heuteColumn - .Start

                                If aix >= 0 Then
                                    For m = 0 To zeitSpanne - 1
                                        If m + aix <= .Dauer - 1 Then
                                            bedarfsWerte(m) = projWerte(m + aix)
                                        End If
                                    Next
                                Else
                                    For m = 0 To zeitSpanne - 1
                                        If m + aix >= 0 Then
                                            bedarfsWerte(m) = projWerte(m + aix)
                                        End If
                                    Next
                                End If


                            End With

                            ' wenn die Summe größer Null ist , wird eine Zeile in das Excel File eingetragen
                            If bedarfsWerte.Sum > 0 Then
                                ' jetzt werden der Status und die Ampelbewertung errechnet ...
                                Call getStatusColorProject(kvp.Value, 1, 1, " ", statusValue, statusColor)

                                CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value = kvp.Value.name
                                CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = statusColor
                                CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).Value = statusValue
                                CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).NumberFormat = "0.00"
                                rng = .Range(.Cells(zeile, 5), .Cells(zeile, 5 + zeitSpanne - 1))
                                rng.Value = bedarfsWerte
                                zeile = zeile + 1
                            End If


                        End If

                    Next

                    sumrangeEnde = zeile - 1

                    ' jetzt wird der Bereich der Projekte mit der entsprechenden Farbe gekennzeichnet 
                    rng = .Range(.Cells(sumrangeAnfang, 1), .Cells(sumrangeEnde, 1))
                    rng.Interior.Color = currentColor

                    ' jetzt wird die Summenformel eingesetzt 
                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Summe"
                    cellFormula = "=SUM(R[" & sumrangeAnfang - zeile & "]C:R[" & sumrangeEnde - zeile & "]C)"
                    For m = 0 To 5
                        CType(.Cells(zeile, 5 + m), Global.Microsoft.Office.Interop.Excel.Range).FormulaR1C1 = cellFormula
                    Next

                    zeile = zeile + 1

                    ' jetzt wird die Kapa eingetragen 
                    CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Kapazität"
                    For m = 0 To 5
                        CType(.Cells(zeile, 5 + m), Global.Microsoft.Office.Interop.Excel.Range).Value = kapaValues(m)
                    Next

                    ' jetzt wird der Bereich mit hellgrauer Farbe abgesetzt 
                    rng = .Range(.Cells(zeile - 1, 1), .Cells(zeile, 1 + 5 + zeitSpanne - 2))
                    rng.Interior.Color = awinSettings.AmpelNichtBewertet

                    zeile = zeile + 3

                    Try
                        mycollection.Clear()
                    Catch ex As Exception
                        mycollection = New Collection
                    End Try

                Next

            End With

        Catch ex As Exception
            Call MsgBox("Register " & xlsBlattname(0) & " existiert nicht ")
        End Try

        '
        ' wenn nur die Zusammenfassung gefragt war: dann wird jetzt die Routine verlassen 
        '
        If typus = 0 Then
            Exit Sub
        End If

        ' ----------------------------------------------------------
        ' Schreiben der Feinplanungs-Sheet - für jeden Monat der Vorausschau eines 
        '-----------------------------------------------------------

        Dim projekttitelZeile As Integer = 3
        Dim bewertungszeile As Integer = 4
        mycollection.Add(currentRole)

        Dim oldName As String
        Dim oldWS As Excel.Worksheet
        Dim currentWS As Excel.Worksheet

        For loopi = 1 To vorausschau
            Dim blattName As String = xlsBlattname(1) & " " & Date.Now.AddMonths(loopi).ToString("MMM yy")

            Try


                currentWS = CType(appInstance.Worksheets(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)
                ' kopieren unter Name blattname & old before:=blattname
                oldName = blattName & " (old)"

                Try
                    oldWS = CType(appInstance.Worksheets(oldName), Global.Microsoft.Office.Interop.Excel.Worksheet)
                    oldWS.Delete()
                    currentWS.Copy(Before:=currentWS)
                    With CType(appInstance.ActiveSheet, Global.Microsoft.Office.Interop.Excel.Worksheet)
                        .Name = oldName
                    End With


                Catch ex1 As Exception
                    ' oldWS existiert nicht - also erzeugen durch Kopie 
                    currentWS.Copy(Before:=currentWS)
                    With CType(appInstance.ActiveSheet, Global.Microsoft.Office.Interop.Excel.Worksheet)
                        .Name = oldName
                    End With
                End Try



            Catch ex As Exception

                Try
                    With CType(appInstance.Worksheets.Add(Before:=appInstance.Worksheets(xlsBlattname(0))), _
                                    Global.Microsoft.Office.Interop.Excel.Worksheet)
                        .Name = blattName
                    End With
                Catch ex2 As Exception
                    Call MsgBox("Tabelle Summary nicht vorhanden ... " & vbLf & ex2.Message)
                    Exit Sub
                End Try


            End Try

            Dim wsBlattname As Excel.Worksheet = CType(appInstance.Worksheets(blattName), _
                                                                    Global.Microsoft.Office.Interop.Excel.Worksheet)


            Try
                ' jetzt zurücksetzen der Planungs-Unterstütung, Register Zuordnung 

                With wsBlattname

                    ' -------------------------------------------------------
                    ' Inhalt leeren ...
                    ' -------------------------------------------------------
                    With .Range(.Cells(1, 1), .Cells(500, 500))
                        .ClearContents()
                        .Interior.Color = awinSettings.AmpelNichtBewertet
                    End With

                    ' -------------------------------------------------------
                    ' Überschrift schreiben 
                    ' -------------------------------------------------------
                    With CType(.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range)
                        .Value = xlsBlattname(1) & " " & qualifier & " " & Date.Now.AddMonths(loopi).ToString("MMM yy")
                        .Font.Size = 20
                        .Font.Bold = True
                    End With

                    CType(.Rows(1), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = 48


                    CType(.Rows("2:100"), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = 16
                    CType(.Rows(5), Global.Microsoft.Office.Interop.Excel.Range).RowHeight = 1
                    CType(.Columns(1), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = 25
                    '.columns(2).columnwidth = 12
                    CType(.Columns("B:CV"), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = 8

                    ' -------------------------------------------------------
                    ' 3. Zeile schreiben: Kapazität und Projekt-Titel    
                    ' -------------------------------------------------------
                    With CType(.Rows(projekttitelZeile), Global.Microsoft.Office.Interop.Excel.Range)
                        .RowHeight = 96
                        .Font.Size = 12
                        .Font.Bold = True
                    End With

                    With CType(.Cells(projekttitelZeile, 2), Global.Microsoft.Office.Interop.Excel.Range)
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .Value = "Kapazität"
                        .Orientation = 90
                        '.AddIndent = True
                        '.IndentLevel = 1
                    End With

                    With .Range(.Cells(projekttitelZeile, 4), .Cells(projekttitelZeile, 100))
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                        .WrapText = True
                        .Orientation = 90
                        '.AddIndent = True
                        '.IndentLevel = 1
                    End With


                    ' -------------------------------------------------------
                    ' 4. Zeile schreiben: Projekt-Bewertung  
                    ' -------------------------------------------------------

                    'With .cells(4, 2)
                    '    .value = "Projekt-Bewertung"
                    '    .font.size = 12
                    '    .font.bold = True
                    'End With
                    With .Range(.Cells(bewertungszeile, 3), .Cells(bewertungszeile, 100))
                        .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                        .Font.Size() = 8
                        .Font.Bold = False
                        .NumberFormat = "0.0#"
                    End With





                End With
            Catch ex As Exception
                Call MsgBox("Fehler mit Tabellenblatt " & blattName)
                Exit Sub
            End Try


            ' jetzt werden die Zuordnungs-Werte ausgelesen 
            ' erweitert 26.7.13 - Planungshilfe für den Ressourcen Manager erstellen
            Try
                ' jetzt werden erst mal die Personen ausgelesen und deren Kapa im besagten Monat 
                With CType(appInstance.Worksheets(xlsBlattname(2)), Global.Microsoft.Office.Interop.Excel.Worksheet)

                    ' wenn noch keine Personen angelegt sind -> Exit 
                    personalrange = .Range("Personenliste")
                    startZeile = personalrange.Row
                    rcol = personalrange.Column
                    ' die letzte Zeile ist die Summenzeile , deshalb ...  
                    endZeile = startZeile + personalrange.Rows.Count - 2
                    anzPeople = endZeile - startZeile + 1
                    rngSource = .Range(.Cells(startZeile, rcol), .Cells(endZeile, rcol))
                    rngTarget = CType(CType(appInstance.Worksheets(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet).Cells(6, 1), _
                                                Global.Microsoft.Office.Interop.Excel.Range)

                    rngSource.Copy(rngTarget)


                    ' jetzt wird die Spalte gesucht, wo die Werte für den nächsten Monat stehen 
                    rcol = 2
                    Dim found As Boolean
                    tmpDate = CDate(CType(.Cells(1, rcol), Global.Microsoft.Office.Interop.Excel.Range).Value)

                    If DateDiff(DateInterval.Month, heute, tmpDate) > 0 Then
                        found = True
                    Else
                        found = False
                    End If

                    Do While Not found And rcol < 200
                        rcol = rcol + 1
                        tmpDate = CDate(CType(.Cells(1, rcol), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        If DateDiff(DateInterval.Month, heute, tmpDate) > loopi - 1 Then
                            found = True
                        Else
                            found = False
                        End If
                    Loop

                    If found Then
                        ' jetzt müssen die Werte referenziert werden 


                        For k = startZeile To endZeile
                            cellFormula = "=" & xlsBlattname(2).Trim & "!R[-4]C[" & rcol - 2 & "]"
                            CType(wsBlattname.Cells(k - startZeile + 6, 2), _
                                    Global.Microsoft.Office.Interop.Excel.Range).FormulaR1C1 = cellFormula
                        Next

                        CType(wsBlattname.Cells(endZeile - startZeile + 7, 1), _
                                Global.Microsoft.Office.Interop.Excel.Range).Value = "Extern"
                        'rngSource = .range(.cells(startZeile, rcol), .cells(endZeile, rcol))
                        'rngTarget = appInstance.Worksheets(blattName).cells(5, 2)
                        'rngSource.Copy(rngTarget)
                    Else
                        Call MsgBox("keine Werte für Folge-Monate von " & heute.ToShortDateString & " gefunden ...")
                        Exit Sub
                    End If


                End With
            Catch ex As Exception
                Call MsgBox("es sind keine Mitarbeiter im Register " & xlsBlattname(2) & " angelegt")
                Exit Sub
            End Try


            'currentColumn = currentColumn + 3
            ' jetzt wird der Rest der Zuordnungs-Datei geschrieben : die Projekt-Daten

            With wsBlattname

                Dim anzProjekte As Integer = 0
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    Dim tmpWert As Double = 0.0
                    If kvp.Value.Start <= getColumnOfDate(Date.Now) + loopi And _
                              kvp.Value.Status <> ProjektStatus(3) And _
                              kvp.Value.Status <> ProjektStatus(4) Then

                        With kvp.Value

                            tmpWert = .getBedarfeInMonth(mycollection, DiagrammTypen(1), heuteColumn + loopi - 1)

                        End With

                        ' wenn die Summe größer Null ist , wird eine Zeile in das Excel File eingetragen
                        If tmpWert > 0 Then
                            ' jetzt werden der Status und die Ampelbewertung errechnet ...
                            Call getStatusColorProject(kvp.Value, 1, 1, " ", statusValue, statusColor)

                            CType(.Cells(projekttitelZeile, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = kvp.Value.name

                            ' Schreiben der Summen-Formel 

                            cellFormula = "=SUM(R[-" & anzPeople + 1 & "]C:R[-1]C)"
                            CType(.Cells(6 + anzPeople + 1, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).FormulaR1C1 = cellFormula

                            ' Schreiben des Bedarfs
                            CType(.Cells(6 + anzPeople + 2, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = tmpWert

                            ' Schreiben der Farbe
                            CType(.Cells(bewertungszeile, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = statusColor

                            ' Schreiben des Wertes 
                            CType(.Cells(bewertungszeile, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = statusValue

                            'currentColumn = currentColumn + 1
                            anzProjekte = anzProjekte + 1


                        End If


                    End If

                Next
                ' Schreiben der Zeilen-Summen: Summe Zuordnung pro MA
                cellFormula = "=SUM(RC[1]:RC[" & anzProjekte & "])"

                For k = 6 To 6 + anzPeople
                    CType(.Cells(k, 3), Global.Microsoft.Office.Interop.Excel.Range).FormulaR1C1 = cellFormula
                Next

                CType(.Cells(7 + anzPeople, 5 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = ""
                CType(.Cells(8 + anzPeople, 4 + anzProjekte), Global.Microsoft.Office.Interop.Excel.Range).Value = "Projekt-Bedarf"
                CType(.Columns(3 + anzProjekte + 1), Global.Microsoft.Office.Interop.Excel.Range).ColumnWidth = 25

                ' Schreiben des Prozentsatzes wieviel des Projektbedarfes wird durch interne abgedeckt 
                cellFormula = "=SUM(R[-" & anzPeople + 2 & "]C:R[-3]C)/SUM(RC[2]:RC[" & anzProjekte + 1 & "])"

                With CType(.Cells(8 + anzPeople, 2), Global.Microsoft.Office.Interop.Excel.Range)
                    .FormulaR1C1 = cellFormula
                    .NumberFormat = "0%"
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                End With

                ' die Eingabe Felder farblos machen  
                .Range(.Cells(6, 4), .Cells(6 + anzPeople, 4 + anzProjekte - 1)).Interior.ColorIndex = Excel.Constants.xlNone




                With .Range(.Cells(6, 1), .Cells(6 + anzPeople, 1))
                    .Interior.Color = awinSettings.AmpelNichtBewertet
                End With

                With .Range(.Cells(6, 2), .Cells(6 + anzPeople, 2))
                    .Font.Size = 12
                    .Font.Bold = True
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .AddIndent = True
                    .IndentLevel = 2
                End With

                ' die vertikalen Summen-Felder etwas einrücken ..
                With .Range(.Cells(6, 3), .Cells(6 + anzPeople, 3))
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    .AddIndent = True
                    .IndentLevel = 2
                End With


                With .Range(.Cells(8 + anzPeople, 4), .Cells(8 + anzPeople, 4 + anzProjekte))
                    .Font.Size = 12
                    .Font.Bold = True
                End With


                ' jetzt wird das Zuordnungs-Blatt geschützt 
                .Range(.Cells(1, 1), .Cells(1 + 8, 1 + 4 + anzProjekte)).Locked = True
                ' Freigeben des Eingabe Bereiches 
                .Range(.Cells(6, 4), .Cells(6 + anzPeople, 4 + anzProjekte - 1)).Locked = False
                CType(.Cells(6, 4), Global.Microsoft.Office.Interop.Excel.Range).Activate()
                .Protect()

            End With

        Next








    End Sub

    ''' <summary>
    ''' zeichnet die Plantafel mit den Projekten neu; 
    ''' versucht dabei immer die alte Position der Projekte zu übernehmen 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinZeichnePlanTafel()

        Dim todoListe As New SortedList(Of Double, String)
        Dim key As Double
        Dim pname As String
        Dim zeile As Integer, lastZeile As Integer, curZeile As Integer, max As Integer
        Dim lastZeileOld As Integer
        Dim hproj As clsProjekt




        ' aufbauen der todoListe, so daß nachher die Projekte von oben nach unten gezeichnet werden können 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            With kvp.Value
                key = 10000 * .tfZeile + kvp.Value.Start
                todoListe.Add(key, .name)
            End With

        Next

        zeile = 2
        lastZeile = 0


        If ProjectBoardDefinitions.My.Settings.drawPhases = True Then
            ' dann sollen die Projekte im extended mode gezeichnet werden 
            ' jetzt erst mal die Konstellation "last" speichern
            Call awinStoreConstellation("Last")

            ' jetzt die todoListe abarbeiten
            For i = 1 To todoListe.Count
                pname = todoListe.ElementAt(i - 1).Value
                hproj = ShowProjekte.getProject(pname)

                If i = 1 Then
                    curZeile = hproj.tfZeile
                    lastZeileOld = hproj.tfZeile
                    lastZeile = curZeile
                    max = curZeile
                Else
                    If lastZeileOld = hproj.tfZeile Then
                        curZeile = lastZeile
                    Else
                        lastZeile = max
                        lastZeileOld = hproj.tfZeile
                    End If

                End If

                hproj.tfZeile = curZeile
                lastZeile = curZeile
                'Call ZeichneProjektinPlanTafel2(pname, curZeile)
                Call ZeichneProjektinPlanTafel(pname, curZeile)
                curZeile = lastZeile + getNeededSpace(hproj)


                If curZeile > max Then
                    max = curZeile
                End If


            Next

        Else


            Dim tryzeile As Integer

            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                pname = kvp.Key
                tryzeile = kvp.Value.tfZeile
                If tryzeile <= 1 Then
                    tryzeile = -1
                End If
                Call ZeichneProjektinPlanTafel(pname, tryzeile) ' es wird versucht, an der alten Stelle zu zeichnen 
            Next


        End If





    End Sub
 
End Module
