Imports ProjectBoardDefinitions
Imports DBAccLayer
Imports ClassLibrary1
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Forms

Imports System
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports System.Drawing
Imports System.Globalization

Imports Microsoft.VisualBasic
Imports System.Security.Principal


''' <summary>
''' hier sind alle Import-/Export Related MEthoden drin
''' </summary>
Public Module agm2
    ' benötigteEnumerationen
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
    ''' wird aktuell nirgends verwendet  
    ''' </summary>
    ''' <param name="myCollection"></param>
    Public Sub awinImportModule(ByRef myCollection As Collection)

        Dim zeile As Integer, spalte As Integer
        Dim pName As String = ""
        Dim vorlagenName As String = ""
        Dim start As Date
        Dim ende As Date
        Dim budget As Double
        Dim dauer As Integer = 0
        Dim sfit As Double, risk As Double
        Dim volume As Double, complexity As Double
        Dim description As String = ""
        Dim businessUnit As String = ""
        Dim lastRow As Integer
        Dim lastColumn As Integer
        'Dim startSpalte As Integer
        Dim vglName As String = ""
        Dim hproj As New clsProjekt
        Dim vproj As clsProjektvorlage
        Dim geleseneProjekte As Integer
        Dim ProjektdauerIndays As Integer = 0
        Dim ok As Boolean = False

        Dim fullProjectNames As New SortedList(Of String, String)
        Dim firstZeile As Excel.Range

        Dim scenarioName As String = appInstance.ActiveWorkbook.Name
        Dim tmpName As String = ""

        ' bestimme den Namen des Szenarios - das ist gleich der NAme der Excel Datei 
        Dim positionIX As Integer = scenarioName.IndexOf(".xls") - 1
        tmpName = ""
        For ih As Integer = 0 To positionIX
            tmpName = tmpName & scenarioName.Chars(ih)
        Next
        scenarioName = tmpName.Trim

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1
        geleseneProjekte = 0

        Dim suchstr(7) As String
        suchstr(ptModuleSpalten.produktlinie) = "Produktlinie"
        suchstr(ptModuleSpalten.name) = "Name"
        suchstr(ptModuleSpalten.projektTyp) = "Projekt-Typ"
        suchstr(ptModuleSpalten.abhaengigVon) = "ist abhängig von"
        suchstr(ptModuleSpalten.strategicFit) = "strat. Bedeutung"
        suchstr(ptModuleSpalten.risiko) = "Risiko der Umsetzung"
        suchstr(ptModuleSpalten.volume) = "Produktions-Volumen"
        suchstr(ptModuleSpalten.budget) = "Budget"


        Dim inputColumns(7) As Integer



        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Tabelle1"),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)

                ' jetzt werden die Spalten bestimmt 
                Try
                    For i As Integer = 0 To 7
                        inputColumns(i) = firstZeile.Find(What:=suchstr(i)).Column
                    Next
                Catch ex As Exception

                End Try

                lastColumn = firstZeile.End(XlDirection.xlToLeft).Column
                lastColumn = CType(.Cells(1, 10000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row





                While zeile <= lastRow
                    ok = False

                    pName = CStr(CType(.Cells(zeile, inputColumns(ptModuleSpalten.name)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    vorlagenName = "Projekt-Platzhalter"

                    ' jetzt muss das Start bzw. Ende Date für das Projekt bestimmt werden
                    ' es ist bestimmt durch das erste auftretende Datum bzw. das letzte auftretende Datum
                    Dim projectStartDate As Date = StartofCalendar.AddYears(100)
                    Dim projectEndDate As Date = StartofCalendar.AddYears(-100)

                    Dim firstC As Integer = inputColumns.Max + 1
                    Dim lastC As Integer = lastColumn
                    Dim anzahlPhasenToAdd As Integer = CInt((lastC - firstC + 1) / 5)
                    Dim allesOK As Boolean
                    Dim ignore As Boolean

                    For i As Integer = 1 To anzahlPhasenToAdd
                        Dim tmpDate As Date
                        Dim chkName As String
                        tmpDate = CDate(CType(.Cells(zeile, firstC + 1 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)

                        Try
                            chkName = CStr(CType(.Cells(zeile, firstC + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                        Catch ex As Exception
                            ignore = True
                            chkName = ""
                        End Try


                        If DateDiff(DateInterval.Day, StartofCalendar, tmpDate) < 0 Or chkName = "-" Then
                            ignore = True
                        Else
                            ignore = False
                        End If

                        If Not ignore Then
                            If DateDiff(DateInterval.Day, projectStartDate, tmpDate) < 0 Then
                                projectStartDate = tmpDate
                            End If

                            tmpDate = CDate(CType(.Cells(zeile, firstC + 2 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            If DateDiff(DateInterval.Day, projectEndDate, tmpDate) > 0 Then
                                projectEndDate = tmpDate
                            End If
                        End If

                    Next


                    If Projektvorlagen.Liste.ContainsKey(vorlagenName) Then

                        vproj = Projektvorlagen.getProject(vorlagenName)
                        Try

                            start = projectStartDate
                            ende = projectEndDate
                            dauer = calcDauerIndays(start, ende)
                            budget = CDbl(CType(.Cells(zeile, inputColumns(ptModuleSpalten.budget)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            risk = CDbl(CType(.Cells(zeile, inputColumns(ptModuleSpalten.risiko)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            sfit = CDbl(CType(.Cells(zeile, inputColumns(ptModuleSpalten.strategicFit)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            volume = CDbl(CType(.Cells(zeile, inputColumns(ptModuleSpalten.volume)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            complexity = 0.2
                            businessUnit = CStr(CType(.Cells(zeile, inputColumns(ptModuleSpalten.produktlinie)), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            description = ""
                            'vglName = pName.Trim & "#" & ""
                            vglName = calcProjektKey(pName.Trim, scenarioName)


                            If DateDiff(DateInterval.Day, StartofCalendar, start) >= 0 Then

                                If DateDiff(DateInterval.Day, start, ende) > 0 Then
                                    ' nichts tun , Ende-Datum ist ein gültiges Datum
                                    ok = True
                                ElseIf DateDiff(DateInterval.Day, StartofCalendar, ende) >= 0 Then
                                    ' auch Ende ist ein gültiges Datum , liegt nur vor Start
                                    ' also vertauschen der beiden 
                                    Dim tmpDate As Date = ende
                                    ende = start
                                    start = tmpDate
                                    ok = True
                                Else
                                    ' Ende Datum wird anhand der Laufzeit der Vorlage oder der Dauer berechnet
                                    If dauer > 0 Then
                                        ProjektdauerIndays = dauer
                                    Else
                                        ProjektdauerIndays = vproj.dauerInDays
                                    End If
                                    ende = calcDatum(start, ProjektdauerIndays)
                                    ok = True
                                End If

                            ElseIf DateDiff(DateInterval.Day, StartofCalendar, ende) >= 0 Then
                                ' hier ist Start kein gültiges Datum innerhalb der Projekt-Tafel 
                                ' Start Datum wird anhand der Laufzeit der Vorlage berechnet
                                If dauer > 0 Then
                                    ProjektdauerIndays = -1 * dauer
                                Else
                                    ProjektdauerIndays = -1 * vproj.dauerInDays
                                End If

                                start = calcDatum(ende, ProjektdauerIndays)

                                If DateDiff(DateInterval.Day, StartofCalendar, start) >= 0 Then
                                    ' Start ist ein korrektes Datum 
                                    ok = True
                                Else
                                    CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Start liegt vor Kalender-Start "
                                    ok = False
                                End If

                            Else
                                CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "ungültiges Start- und Ende-Datum"
                                ok = False
                            End If

                        Catch ex As Exception
                            CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."
                            ok = False
                        End Try


                    Else
                        CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."
                        ok = False
                    End If

                    ' jetzt die Aktion durchführen, wenn alles ok 
                    If ok Then
                        If AlleProjekte.Containskey(vglName) Then
                            ' nichts tun ...
                            Call MsgBox("Projekt aus Inventur Liste existiert bereits - keine Neuanlage")
                        Else
                            Try
                                fullProjectNames.Add(vglName, vglName)
                                'Projekt anlegen ,Verschiebung um 
                                hproj = New clsProjekt(start, start.AddMonths(-1), start.AddMonths(1))

                                Dim capacityNeeded As String = ""
                                hproj = erstelleInventurProjekt(pName, vorlagenName, scenarioName,
                                                             start, ende, budget, zeile, sfit, risk,
                                                             capacityNeeded, Nothing, businessUnit, description, Nothing, "", 0.0)

                                If Not IsNothing(hproj) Then
                                    projectStartDate = start
                                    projectEndDate = ende
                                Else
                                    ok = False
                                End If

                            Catch ex As Exception
                                ok = False
                            End Try


                        End If
                    End If

                    If ok Then

                        Dim phaseName As String = ""
                        Dim scaleRule As Integer
                        Dim moduleNames() As String
                        Dim moduleName As String
                        Dim allNames As String
                        Dim planModul As clsProjektvorlage

                        ' jetzt müssen die Module ergänzt werden 
                        For i As Integer = 1 To anzahlPhasenToAdd

                            start = CDate(CType(.Cells(zeile, firstC + 1 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            ende = CDate(CType(.Cells(zeile, firstC + 2 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)


                            Dim startOffset As Integer = CInt(DateDiff(DateInterval.Day, projectStartDate, start))
                            Dim endOffset As Integer = CInt(DateDiff(DateInterval.Day, projectStartDate, ende))


                            Try
                                phaseName = CStr(CType(.Cells(zeile, firstC + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                If phaseName = "-" Or endOffset - startOffset = 0 Then
                                    allesOK = False
                                    phaseName = "-"
                                Else
                                    allesOK = True
                                End If
                            Catch ex As Exception
                                allesOK = False
                            End Try

                            Dim parentPhase As clsPhase = Nothing



                            If allesOK Then

                                '
                                ' jetzt muss die aufnehmende Phase erstmal angelegt werden 
                                '
                                If Not IsNothing(phaseName) Then

                                    If phaseName.Length > 0 Then

                                        parentPhase = New clsPhase(parent:=hproj)
                                        parentPhase.nameID = hproj.hierarchy.findUniqueElemKey(phaseName, False)
                                        parentPhase.changeStartandDauer(startOffset, calcDauerIndays(start, ende))

                                        hproj.AddPhase(parentPhase, origName:=phaseName,
                                               parentID:=rootPhaseName)

                                    End If

                                End If


                                scaleRule = CInt(CType(.Cells(zeile, firstC + 3 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                allNames = CStr(CType(.Cells(zeile, firstC + 4 + (i - 1) * 5), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                ' jetzt müssen die einzelnen Module ausgelesen werden 
                                ' aber nur, wenn überhaupt was drin steht und das auch als Modul existiert ...
                                '
                                If Not IsNothing(allNames) Then

                                    If Not allNames.Trim.Length = 0 Then

                                        moduleNames = allNames.Split(New Char() {CChar("#")}, 20)
                                        Dim anzahl As Integer = moduleNames.Length

                                        For ix As Integer = 1 To anzahl
                                            moduleName = moduleNames(ix - 1)
                                            If ModulVorlagen.Contains(moduleName) Then
                                                planModul = ModulVorlagen.getProject(moduleName)

                                                If Not IsNothing(parentPhase) Then

                                                    planModul.moduleCopyTo(hproj, parentPhase.nameID, moduleName, startOffset, endOffset, True)

                                                End If
                                            End If
                                        Next

                                    End If

                                End If

                            End If
                        Next

                        ' jetzt die Projekt eintragen 
                        If Not hproj Is Nothing Then
                            Try
                                ImportProjekte.Add(hproj, False)
                                myCollection.Add(calcProjektKey(hproj))
                            Catch ex As Exception

                            End Try

                        End If

                    End If

                    zeile = zeile + 1

                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei Module Import ...")
        End Try


        ' jetzt noch ein Szenario anlegen, wenn ImportProjekte was enthält 
        If ImportProjekte.Count > 0 Then
            Call storeSessionConstellation(scenarioName, fullProjectNames)
        End If

        currentConstellationName = scenarioName

    End Sub

    ''' <summary>
    ''' baut die Liste der Darstellungsklassen auf 
    ''' übergeben wird das Excel Worksheet 
    ''' </summary>
    ''' <param name="ws"></param>
    ''' <remarks></remarks>
    Friend Sub aufbauenAppearanceDefinitions(ByVal ws As Excel.Worksheet)

        Dim appDefinition As clsAppearance
        Dim errMsg As String = ""
        Dim firstMilestone As Boolean = True
        Dim firstPhase As Boolean = True

        With ws

            For Each shp As Excel.Shape In .Shapes
                appDefinition = New clsAppearance
                With appDefinition

                    If shp.Title <> "" Then

                        .name = shp.Title
                        If shp.AlternativeText = "1" Then
                            .isMilestone = True
                        Else
                            .isMilestone = False
                        End If
                        .form = shp

                        Try
                            appearanceDefinitions.Add(.name, appDefinition)

                            If .isMilestone And firstMilestone Then
                                awinSettings.defaultMilestoneClass = .name
                                firstMilestone = False

                            ElseIf Not .isMilestone And firstPhase Then
                                awinSettings.defaultPhaseClass = .name
                                firstPhase = False
                            End If
                        Catch ex As Exception
                            errMsg = "Mehrfach Definition in den Darstellungsklassen ... " & vbLf &
                                         "bitte korrigieren"
                            Throw New Exception(errMsg)
                        End Try


                    End If

                End With


            Next

        End With

    End Sub

    ''' <summary>
    ''' erstellt das Vorlagen File aus der Liste der Phasen 
    ''' aktuell wird nur die Übergabe von Phasen unterstützt
    ''' </summary>
    ''' <param name="phaseList"></param>
    ''' <param name="milestoneList"></param>
    ''' <remarks></remarks>
    Public Sub createVorlageFromSelection(ByVal phaseList As SortedList(Of String, String),
                                              ByVal milestoneList As SortedList(Of String, String))

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim elemName As String = ""
        Dim breadcrumb As String = ""
        Dim lfdNr As Integer = 1
        Dim fullName As String
        Dim ext As String = ""

        appInstance.EnableEvents = False
        enableOnUpdate = False


        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try
            appInstance.Workbooks.Add()


        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            enableOnUpdate = True
            Throw New ArgumentException("Excel Export nicht gefunden - Abbruch")
        End Try

        'appInstance.Workbooks(myCustomizationFile).Activate()
        Dim wsName As Excel.Worksheet = CType(appInstance.ActiveSheet,
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)


        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        Dim startDate As Date, endDate As Date
        Dim tmpRange As Excel.Range
        Dim anzahlProjekte As Integer = ShowProjekte.Count

        With wsName
            ' jetzt werden alle Spalten auf Breite 25 gesetzt 
            tmpRange = CType(.Range(.Cells(zeile, spalte), .Cells(zeile, spalte).offset(0, 500)), Excel.Range)
            tmpRange.ColumnWidth = 25

            ' jetzt wird der Header geschrieben 
            CType(.Cells(zeile, spalte), Excel.Range).Value = "Produktlinie"
            CType(.Cells(zeile, spalte + 1), Excel.Range).Value = "Name"
            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "Projekt-Typ"
            CType(.Cells(zeile, spalte + 3), Excel.Range).Value = "ist abhängig von"
            CType(.Cells(zeile, spalte + 4), Excel.Range).Value = "strat. Bedeutung"
            CType(.Cells(zeile, spalte + 5), Excel.Range).Value = "Risiko der Umsetzung"
            CType(.Cells(zeile, spalte + 6), Excel.Range).Value = "Produktions-Volumen"
            CType(.Cells(zeile, spalte + 7), Excel.Range).Value = "Budget"


            spalte = spalte + 8


            ' hier muss noch korrigiert werden: wenn es bei einem oder mehreren Projekten mehrere Elemente dieses Namens und Breadcrumbs gibt, so 
            ' muss das in dieser Liste auch vorgesehen werden 
            For ix As Integer = 1 To phaseList.Count

                Dim phaseName As String = ""
                CType(.Cells(zeile, spalte), Excel.Range).Value = "Phasen-Name"
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = "Start-Datum"
                CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "Ende-Datum"
                CType(.Cells(zeile, spalte + 3), Excel.Range).Value = "Skalierungs-Regel"
                CType(.Cells(zeile, spalte + 4), Excel.Range).Value = "Modul-Namen(n"

                tmpRange = CType(.Range(.Cells(zeile, spalte + 1), .Cells(zeile, spalte + 1).offset(anzahlProjekte, 1)), Excel.Range)
                tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                tmpRange.NumberFormat = "dd/mm/yy;@"

                spalte = spalte + 5

            Next


        End With


        'es geht von vorne los 
        spalte = 1
        zeile = 2

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            With wsName
                ' Produktlinie schreiben 

                Try
                    If kvp.Value.businessUnit.Length > 0 Then
                        CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.businessUnit
                    Else
                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                    End If
                Catch ex As Exception
                    CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                End Try


                ' Name schreiben 
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = kvp.Value.name

                ' Projekt-Typ schreiben 
                Try
                    If kvp.Value.VorlagenName.Length > 0 Then
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = kvp.Value.VorlagenName
                    Else
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                    End If
                Catch ex As Exception
                    CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                End Try

                ' ist abhängig von schreiben
                CType(.Cells(zeile, spalte + 3), Excel.Range).Value = ""


                ' strategische Bedeutung schreiben 
                CType(.Cells(zeile, spalte + 4), Excel.Range).Value = kvp.Value.StrategicFit

                ' risiko Kennzahl schreiben 
                CType(.Cells(zeile, spalte + 5), Excel.Range).Value = kvp.Value.Risiko


                ' Produktions-Volumen schreiben 
                CType(.Cells(zeile, spalte + 6), Excel.Range).Value = kvp.Value.volume

                ' Budget schreiben 
                CType(.Cells(zeile, spalte + 7), Excel.Range).Value = ""

                ' Phasen Information schreiben

                spalte = spalte + 8


                ' hier muss noch korrigiert werden: wenn es bei einem oder mehreren Projekten mehrere Elemente dieses Namens und Breadcrumbs gibt, so 
                ' muss das in dieser Liste auch vorgesehen werden 

                Dim cphase As clsPhase
                For ix As Integer = phaseList.Count To 1 Step -1

                    fullName = CStr(phaseList.ElementAt(ix - 1).Value)
                    elemName = ""
                    breadcrumb = ""
                    lfdNr = 0
                    Dim type As Integer = -1
                    Dim pvName As String = ""
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)


                    cphase = kvp.Value.getPhase(elemName, breadcrumb, lfdNr)
                    Dim phaseName As String

                    If Not IsNothing(cphase) Then
                        Try

                            phaseName = kvp.Value.getBestNameOfID(cphase.nameID, True, False)
                            startDate = cphase.getStartDate
                            endDate = cphase.getEndDate

                            CType(.Cells(zeile, spalte), Excel.Range).Value = phaseName.Replace("#", "-")
                            CType(.Cells(zeile, spalte + 1), Excel.Range).Value = startDate
                            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = endDate
                            CType(.Cells(zeile, spalte + 3), Excel.Range).Value = "1"
                            CType(.Cells(zeile, spalte + 4), Excel.Range).Value = ""

                        Catch ex As Exception


                        End Try
                    Else

                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"


                    End If

                    spalte = spalte + 5

                Next


            End With

            zeile = zeile + 1
            spalte = 1

        Next

        'Dim expFName As String = awinPath & exportFilesOrdner & _
        '    "\Vorlage_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Dim expFName As String = exportOrdnerNames(PTImpExp.modulScen) &
            "\Vorlage_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Try
            appInstance.ActiveWorkbook.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True



    End Sub


    ''' <summary>
    ''' schreibt die übergebenen Phasen und Meilensteine in eine Excel Datei 
    ''' </summary>
    ''' <param name="phaseList"></param>
    ''' <param name="milestoneList"></param>
    ''' <remarks></remarks>
    Public Sub exportSelectionToExcel(ByVal phaseList As SortedList(Of String, String),
                                            ByVal milestoneList As SortedList(Of String, String))

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim elemName As String = ""
        Dim breadcrumb As String = ""
        Dim lfdNr As Integer = 1
        Dim fullName As String
        Dim ext As String = ""

        appInstance.EnableEvents = False
        enableOnUpdate = False


        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try
            'appInstance.Workbooks.Open(awinPath & requirementsOrdner & excelExportVorlage)
            appInstance.Workbooks.Add()


        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            enableOnUpdate = True
            Throw New ArgumentException("Excel Export nicht gefunden - Abbruch")
        End Try

        'appInstance.Workbooks(myCustomizationFile).Activate()
        Dim wsName As Excel.Worksheet = CType(appInstance.ActiveSheet,
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)


        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        Dim startDate As Date, endDate As Date
        Dim earliestDate As Date, latestDate As Date
        Dim tmpRange As Excel.Range
        Dim anzahlProjekte As Integer = ShowProjekte.Count

        With wsName
            ' jetzt werden alle Spalten auf Breite 25 gesetzt 
            tmpRange = CType(.Range(.Cells(zeile, spalte), .Cells(zeile, spalte).offset(0, 200)), Excel.Range)
            tmpRange.ColumnWidth = 25

            ' jetzt wird der Header geschrieben 
            CType(.Cells(zeile, spalte), Excel.Range).Value = "Produktlinie"
            CType(.Cells(zeile, spalte + 1), Excel.Range).Value = "Name"
            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "Projekt-Typ"

            spalte = spalte + 2


            ' hier muss noch orrigiert werden: wenn es bei einem oder mehreren Projekten mehrere Elemente dieses Namens und Breadcrumbs gibt, so 
            ' muss das in dieser Liste auch vorgesehen werden 
            For ix As Integer = 1 To phaseList.Count

                Try
                    fullName = CStr(phaseList.ElementAt(ix - 1).Value)
                Catch ex As Exception
                    fullName = ""
                End Try

                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)

                If lfdNr > 1 Then
                    ext = " " & lfdNr.ToString
                Else
                    ext = ""
                End If
                If breadcrumb = "" Then
                    CType(.Cells(zeile, spalte + ix), Excel.Range).Value = elemName & ext
                Else
                    CType(.Cells(zeile, spalte + ix), Excel.Range).Value = breadcrumb.Replace("#", "-") & "-" & elemName & ext
                End If

            Next

            spalte = spalte + phaseList.Count

            ' hier muss noch orrigiert werden: wenn es bei einem oder mehreren Projekten mehrere Elemente dieses NAmens gibt, so 
            ' muss das in dieser Liste auch vorgesehen werden 

            For ix As Integer = 1 To milestoneList.Count

                Try
                    fullName = CStr(milestoneList.ElementAt(ix - 1).Value)
                Catch ex As Exception
                    fullName = ""
                End Try

                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)

                If lfdNr > 1 Then
                    ext = " " & lfdNr.ToString
                Else
                    ext = ""
                End If
                If breadcrumb = "" Then
                    CType(.Cells(zeile, spalte + ix), Excel.Range).Value = elemName & ext
                Else
                    CType(.Cells(zeile, spalte + ix), Excel.Range).Value = breadcrumb.Replace("#", "-") & "-" & elemName & ext
                End If


            Next


            ' Datumsformat einstellen 
            Dim s1 As Integer = 4 + phaseList.Count
            Dim o1 As Integer = milestoneList.Count - 1

            ' mittig darstellen 
            tmpRange = CType(.Range(.Cells(zeile, 4), .Cells(zeile, 4).offset(anzahlProjekte, s1 + o1 - 4)), Excel.Range)
            tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            tmpRange = CType(.Range(.Cells(zeile + 1, s1), .Cells(zeile + 1, s1).offset(anzahlProjekte - 1, o1)), Excel.Range)
            tmpRange.NumberFormat = "dd/mm/yy;@"

            spalte = spalte + milestoneList.Count

            CType(.Cells(zeile, spalte + 1), Excel.Range).Value = "Dauer (T)"
            CType(.Range(.Cells(zeile + 1, spalte + 1), .Cells(zeile + 1, spalte + 1).offset(anzahlProjekte - 1, 0)), Excel.Range).NumberFormat = "0"

            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "Dauer (M)"
            CType(.Range(.Cells(zeile + 1, spalte + 2), .Cells(zeile + 1, spalte + 2).offset(anzahlProjekte - 1, 0)), Excel.Range).NumberFormat = "0.0"

        End With


        zeile = 2
        spalte = 1
        Dim minCol As Integer
        Dim maxCol As Integer

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            earliestDate = kvp.Value.endeDate
            latestDate = kvp.Value.startDate

            ' wird benötigt, um festzustellen, ob überhaupt eines der Elemente im aktuell 
            ' betrachteten Projekt vorkommt 
            Dim atleastOne As Boolean = False

            With wsName
                ' Produktlinie schreiben 

                Try
                    If kvp.Value.businessUnit.Length > 0 Then
                        CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.businessUnit
                    Else
                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                    End If
                Catch ex As Exception
                    CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                End Try


                ' Name schreiben 
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = kvp.Value.name

                ' Projekt-Typ schreiben 
                Try
                    If kvp.Value.VorlagenName.Length > 0 Then
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = kvp.Value.VorlagenName
                    Else
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                    End If
                Catch ex As Exception
                    CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                End Try


                ' Phasen Information schreiben

                Dim cphase As clsPhase
                spalte = spalte + 3

                For ix As Integer = 1 To phaseList.Count

                    fullName = CStr(phaseList.ElementAt(ix - 1).Value)
                    elemName = ""
                    breadcrumb = ""
                    lfdNr = 0
                    Dim type As Integer = -1
                    Dim pvName As String = ""
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)


                    cphase = kvp.Value.getPhase(elemName, breadcrumb, lfdNr)

                    If Not IsNothing(cphase) Then
                        Try
                            startDate = cphase.getStartDate
                            endDate = cphase.getEndDate

                            atleastOne = True

                            If DateDiff(DateInterval.Day, startDate, earliestDate) > 0 Then
                                earliestDate = startDate
                                minCol = spalte
                            End If

                            If DateDiff(DateInterval.Day, latestDate, endDate) > 0 Then
                                latestDate = endDate
                                maxCol = spalte
                            End If

                            CType(.Cells(zeile, spalte), Excel.Range).Value = startDate.ToShortDateString & " - " & endDate.ToShortDateString

                        Catch ex As Exception
                            CType(.Cells(zeile, spalte), Excel.Range).Value = "?"

                        End Try
                    Else

                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"


                    End If

                    spalte = spalte + 1



                Next


                ' Meilensteine schreiben 

                Dim milestone As clsMeilenstein = Nothing

                For ix As Integer = 1 To milestoneList.Count

                    fullName = CStr(milestoneList.ElementAt(ix - 1).Value)
                    elemName = ""
                    breadcrumb = ""
                    lfdNr = 0
                    Dim type As Integer = -1
                    Dim pvName As String = ""
                    Call splitBreadCrumbFullnameTo3(fullName, elemName, breadcrumb, lfdNr, type, pvName)

                    milestone = kvp.Value.getMilestone(elemName, breadcrumb, lfdNr)

                    If Not IsNothing(milestone) Then
                        Try
                            startDate = milestone.getDate

                            atleastOne = True

                            If DateDiff(DateInterval.Day, startDate, earliestDate) > 0 Then
                                earliestDate = startDate
                                minCol = spalte
                            End If

                            If DateDiff(DateInterval.Day, latestDate, startDate) > 0 Then
                                latestDate = startDate
                                maxCol = spalte
                            End If

                            CType(.Cells(zeile, spalte), Excel.Range).Value = startDate


                        Catch ex As Exception
                            CType(.Cells(zeile, spalte), Excel.Range).Value = "?"
                            CType(.Cells(zeile, spalte), Excel.Range).Value = "?"
                        End Try
                    Else

                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                        CType(.Cells(zeile, spalte), Excel.Range).Value = "-"

                    End If

                    spalte = spalte + 1

                Next

                Dim dauerT As Long
                Dim dauerM As Double

                ' Dauer in Tagen schreiben 

                Try
                    If atleastOne Then
                        dauerT = DateDiff(DateInterval.Day, earliestDate, latestDate)
                        dauerM = 12 * dauerT / 365
                    Else
                        dauerT = 0
                        dauerM = 0.0
                    End If
                Catch ex As Exception
                    dauerT = 0
                    dauerM = 0.0
                End Try


                CType(.Cells(zeile, spalte), Excel.Range).Value = dauerT
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = dauerM

                ' jetzt einfärben, welche Daten zu der Dauer geführt haben 
                If minCol = maxCol And minCol > 0 Then
                    CType(.Cells(zeile, minCol), Excel.Range).Interior.Color = awinSettings.AmpelGruen
                Else
                    If minCol > 0 Then
                        CType(.Cells(zeile, minCol), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                    End If
                    If maxCol > 0 Then
                        CType(.Cells(zeile, maxCol), Excel.Range).Interior.Color = awinSettings.AmpelGelb
                    End If

                End If



            End With

            zeile = zeile + 1
            spalte = 1

        Next

        'Dim expFName As String = awinPath & exportFilesOrdner & _
        '    "\Report_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Dim expFName As String = exportOrdnerNames(PTImpExp.rplan) &
            "\Report_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Try
            appInstance.ActiveWorkbook.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True



    End Sub
    ''' <summary>
    ''' erstellt die Vorlage für die InputDatei des Batch-Report
    ''' Input-Tabelle wird erzeugt, wie vom VISBO ReportGen erwartet
    ''' ReportProfile - Tabelle wird bestückt aus den vorhandenen ReportProfilen in Directory ReportProfile
    ''' ProjekteSzenarien - Tabelle wird bestückt aus Liste AlleProjekte (d.h. es müssen Projekte oder Szenarien geladen sein
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub createReportGenTemplate()

        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim tmpRange As Excel.Range

        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        appInstance.EnableEvents = False
        enableOnUpdate = False


        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try
            'appInstance.Workbooks.Open(awinPath & requirementsOrdner & excelExportVorlage)
            appInstance.Workbooks.Add()


        Catch ex As Exception
            appInstance.EnableEvents = formerEE
            enableOnUpdate = True
            Throw New ArgumentException("Excel Export nicht gefunden - Abbruch")
        End Try



        Dim wsName As Excel.Worksheet
        appInstance.Worksheets.Add()
        wsName = CType(appInstance.ActiveSheet,
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
        wsName.Name = "ProjekteSzenarien"

        zeile = 1
        spalte = 1


        Dim anzahlProjekte As Integer = AlleProjekte.Count

        With wsName
            ' jetzt werden alle Spalten auf Breite 25 gesetzt 
            tmpRange = CType(.Range(.Cells(zeile, spalte), .Cells(zeile, spalte).offset(0, 200)), Excel.Range)
            tmpRange.ColumnWidth = 25


            ' jetzt wird der Header geschrieben 
            With CType(.Cells(zeile, spalte), Excel.Range)
                .Value = "Projekte "
                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With

            With CType(.Cells(zeile, spalte + 1), Excel.Range)
                .Value = "Varianten"
                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With

            spalte = spalte + 1
        End With


        zeile = 2
        spalte = 1

        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

            Dim projName As String = kvp.Value.name
            Dim variantName As String = kvp.Value.variantName

            With wsName


                ' Name schreiben 
                CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.name

                ' Varianten-Name schreiben 
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = kvp.Value.variantName



            End With

            zeile = zeile + 1
            spalte = 1

        Next


        zeile = zeile + 1   ' eine Leerzeile
        spalte = 1
        With wsName
            With CType(.Cells(zeile, spalte), Excel.Range)
                .Value = "Szenarien"
                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With
        End With

        zeile = zeile + 1   ' eine Leerzeile
        spalte = 1

        ' alle möglichen Szenario-Namen eintragen
        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            Dim szenarioName As String = kvp.Value.constellationName

            With wsName


                ' SzenarioName schreiben 
                CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.constellationName

            End With

            zeile = zeile + 1
            spalte = 1

        Next

        Dim wsReportProfile As Excel.Worksheet
        appInstance.Worksheets.Add()
        wsReportProfile = CType(appInstance.ActiveSheet,
                                              Global.Microsoft.Office.Interop.Excel.Worksheet)
        wsReportProfile.Name = "ReportProfile"

        zeile = 1
        spalte = 1

        With wsReportProfile

            ' jetzt wird der Header geschrieben 
            With CType(.Cells(zeile, spalte), Excel.Range)
                .ColumnWidth = 40
                .Value = "ReportProfile"
                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With

        End With

        zeile = 2
        spalte = 1

        Dim dateiName As String = ""

        Try

            With wsReportProfile

                Dim i As Integer
                Dim dirname As String = My.Computer.FileSystem.CombinePath(awinPath, ReportProfileOrdner)

                ' ReportProfile vom Directory lesen
                Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)

                ' und in das Excel-File eintragen
                For i = 1 To listOfVorlagen.Count
                    Dim tmpstr() As String = Split(Dir(listOfVorlagen.Item(i - 1)), ".xml")
                    dateiName = tmpstr(0)
                    CType(.Cells(zeile, spalte), Excel.Range).Value = dateiName
                    zeile = zeile + 1

                Next i

            End With
        Catch ex As Exception

        End Try

        Dim wsInput As Excel.Worksheet
        appInstance.Worksheets.Add()
        wsInput = CType(appInstance.ActiveSheet,
                                              Global.Microsoft.Office.Interop.Excel.Worksheet)
        wsInput.Name = "Input"

        zeile = 1
        spalte = 1

        With wsInput
            ' jetzt werden alle Spalten auf Breite 40 gesetzt 
            tmpRange = CType(.Range(.Cells(zeile, spalte), .Cells(zeile, spalte).offset(0, 200)), Excel.Range)
            With tmpRange
                .RowHeight = 20
                .HorizontalAlignment = XlHAlign.xlHAlignCenter
                .VerticalAlignment = XlVAlign.xlVAlignCenter

                With .Font
                    .Name = "Arial"
                    .FontStyle = "Fett"
                    .Size = 11
                    .Strikethrough = False
                    .Superscript = False
                    .Subscript = False
                    .OutlineFont = False
                    .Shadow = False
                End With
            End With

            ' jetzt wird der Header geschrieben 

            With CType(.Cells(zeile, spalte), Excel.Range)
                .Value = "Name des Reports"
                .ColumnWidth = 40
            End With

            With CType(.Cells(zeile, spalte + 1), Excel.Range)
                .Value = "SpeicherModus"
                .ColumnWidth = 15
            End With

            With CType(.Cells(zeile, spalte + 2), Excel.Range)
                .Value = "Name des ReportProfils"
                .ColumnWidth = 45
            End With

            With CType(.Cells(zeile, spalte + 3), Excel.Range)
                .Value = "Names des Portfolios / Projekts"
                .ColumnWidth = 30
            End With

            With CType(.Cells(zeile, spalte + 4), Excel.Range)
                .Value = "VariantenName"
                .ColumnWidth = 30
            End With

            With CType(.Cells(zeile, spalte + 5), Excel.Range)
                .Value = "TimeStamp"
                .ColumnWidth = 30
            End With

            With CType(.Cells(zeile, spalte + 6), Excel.Range)
                .Value = " von"
                .ColumnWidth = 18
            End With

            With CType(.Cells(zeile, spalte + 7), Excel.Range)
                .Value = "bis"
                .ColumnWidth = 18
            End With

        End With



        'Dim expFName As String = awinPath & exportFilesOrdner & _
        '    "\Report_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"


        Dim expFName As String = exportOrdnerNames(PTImpExp.modulScen) &
            "\ReportGenTemplate_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Try
            appInstance.ActiveWorkbook.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True



    End Sub

    ''' <summary>
    ''' wird aus Formular NameSelection bzw. HrySelection aufgerufen
    ''' besetzt die Vorlagen Dropbox den entsprechenden Datei-NAmen
    ''' </summary>
    ''' <param name="menuOption"></param>
    ''' <param name="repVorlagenDropbox"></param>
    ''' <remarks></remarks>
    Public Sub frmHryNameReadPPTVorlagen(ByVal menuOption As Integer, ByRef repVorlagenDropbox As System.Windows.Forms.ComboBox, Optional ByVal mppreport As Boolean = False)


        Dim dirname As String
        Dim dateiName As String = ""


        If menuOption = PTmenue.multiprojektReport Or menuOption = PTmenue.einzelprojektReport Then

            If menuOption = PTmenue.multiprojektReport Then
                dirname = awinPath & RepPortfolioVorOrdner
            Else
                dirname = awinPath & RepProjectVorOrdner
            End If

            Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
            Try

                Dim i As Integer
                For i = 1 To listOfVorlagen.Count
                    dateiName = Dir(listOfVorlagen.Item(i - 1))
                    If dateiName.Contains("Typ II") Then
                        repVorlagenDropbox.Items.Add(dateiName)
                    End If

                Next i
            Catch ex As Exception

            End Try
        ElseIf menuOption = PTmenue.reportBHTC Or
            menuOption = PTmenue.reportMultiprojektTafel Then

            If mppreport Then
                dirname = awinPath & RepPortfolioVorOrdner
            Else
                dirname = awinPath & RepProjectVorOrdner
            End If


            Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
            Try

                Dim i As Integer
                For i = 1 To listOfVorlagen.Count

                    dateiName = Dir(listOfVorlagen.Item(i - 1))
                    repVorlagenDropbox.Items.Add(dateiName)

                Next i
            Catch ex As Exception

            End Try

        End If

    End Sub

    ''' <summary>
    ''' erstellt das Excel Export bzw. Vorlagen  File für die angegebenen Phasen, Meilensteine, Rollen und Kosten
    ''' vorläufig nur für Phasen und Meilensteine realisiert
    ''' </summary>
    ''' <param name="filterName">gibt den Namen des Filters an, der die Collections enthält </param>
    ''' <remarks></remarks>
    Friend Sub createDateiFromSelection(ByVal filterName As String, ByVal menueOption As Integer)

        Dim earliestDate As Date, latestDate As Date
        Dim phaseList As New SortedList(Of String, String)
        Dim milestonelist As New SortedList(Of String, String)

        Dim selphases As New Collection
        Dim selMilestones As New Collection
        Dim selRoles As New Collection
        Dim selCosts As New Collection
        Dim selBUs As New Collection
        Dim selTyps As New Collection


        Call retrieveSelections(filterName, menueOption, selBUs, selTyps,
                                 selphases, selMilestones, selRoles, selCosts)

        ' initialisieren 
        earliestDate = StartofCalendar.AddMonths(-12)
        latestDate = StartofCalendar.AddMonths(1200)

        Dim anteil As Double = 0.0
        Dim anzahlProjekte As Integer = ShowProjekte.Count
        Dim currentIX As Integer
        Dim hproj As clsProjekt
        Dim pName As String, msName As String

        Dim anzPlanobjekte As Integer = selphases.Count + selMilestones.Count
        Dim bestproj As String = ""
        Dim startFaktor As Double = 1.0
        Dim durationFaktor As Double = 0.000001
        Dim correctFaktor As Double = 0.00000001
        Dim korrFaktor As Double
        Dim refLaenge As Integer
        Dim fullName As String = ""
        Dim breadcrumb As String = ""
        Dim listName As String = ""

        ' die selphases und selMilestones enthalten jetzt 

        currentIX = 1
        Do While currentIX <= anzahlProjekte

            hproj = ShowProjekte.getProject(currentIX)

            If currentIX = 1 Then
                korrFaktor = 1.0
                refLaenge = hproj.dauerInDays
            Else
                Try
                    korrFaktor = hproj.dauerInDays / refLaenge
                Catch ex As Exception
                    korrFaktor = 1.0
                End Try

            End If

            ' es wird einfach der Reihenfolge nach eingetragen
            ' eine vorherige Überprüfung, welche Meilensteine grundsätzlich vorne stehen, wird nicht mehr gemacht 

            For Each pObject As Object In selphases

                pName = ""
                breadcrumb = ""
                fullName = CStr(pObject)
                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitHryFullnameTo2(fullName, pName, breadcrumb, type, pvName)

                ' jetzt muss eine Schleife gemacht werden über alle Vorkommen dieses Namens
                Dim anzahlElements As Integer = hproj.hierarchy.getPhaseIndices(pName, breadcrumb).Length

                For ce As Integer = 1 To anzahlElements

                    listName = fullName & "#" & ce.ToString("00#")

                    If phaseList.ContainsKey(listName) Then
                        ' nichts tun, dann ist sie schon eingeordnet 
                    Else

                        ' schlüssel kann gar nicht mehrfach vorkommen) 
                        phaseList.Add(listName, listName)


                    End If
                Next
            Next


            For Each pObject As Object In selMilestones

                msName = ""
                breadcrumb = ""
                fullName = CStr(pObject)
                Dim type As Integer = -1
                Dim pvName As String = ""
                Call splitHryFullnameTo2(fullName, msName, breadcrumb, type, pvName)

                ' jetzt muss eine Schleife gemacht werden über alle Vorkommen dieses Namens
                Dim anzahlElements As Integer = CInt(hproj.hierarchy.getMilestoneIndices(msName, breadcrumb).Length / 2)


                For ce As Integer = 1 To anzahlElements

                    listName = fullName & "#" & ce.ToString("00#")

                    If milestonelist.ContainsKey(listName) Then
                        ' nichts tun, dann ist sie schon eingeordnet 
                    Else

                        ' schlüssel kann gar nicht mehrfach vorkommen) 
                        milestonelist.Add(listName, listName)

                    End If

                Next

                ' alt 

            Next


            currentIX = currentIX + 1

        Loop

        ' jetzt sind die Elemente in der richtigen Reihenfolge eingeordnet 
        ' jetzt werden sie rausgeschrieben 
        Try
            If menueOption = PTmenue.excelExport Then
                Call exportSelectionToExcel(phaseList, milestonelist)
            ElseIf menueOption = PTmenue.vorlageErstellen Then
                Call createVorlageFromSelection(phaseList, milestonelist)
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try



    End Sub

    ''' <summary>
    ''' verallgemeinerte Import Routine, ähnlich wie BMWimport
    ''' wenn treatAsPhases = true, werden die einzelnen Pläne als Sammelvorgänge innerhalb ein und desselben Projektes aufgefasst  
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="isVorlage"></param>
    ''' <remarks></remarks>
    Public Sub planExcelImport(ByRef myCollection As Collection, ByVal isVorlage As Boolean, ByVal dateiname As String)

        Dim phaseHierarhy(9) As String
        Dim currentHierarchy As Integer = 0
        Dim zeile As Integer, spalte As Integer
        Dim pName As String = " "
        Dim phaseName As String = " "
        Dim currentDateiName As String
        Dim isMilestone As Boolean

        Dim lastRow As Integer

        Dim hproj As clsProjekt
        Dim vproj As clsProjektvorlage
        Dim vglName As String = ""
        Dim vglProj As New clsProjekt
        Dim geleseneProjekte As Integer
        Dim projektFarbe As Object
        Dim anfang As Integer, ende As Integer
        Dim cphase As clsPhase
        Dim cmilestone As clsMeilenstein
        Dim cbewertung As clsBewertung
        Dim ix As Integer
        Dim tmpStr(20) As String
        Dim completeName As String
        Dim nameSopTyp As String = " "
        Dim nameProduktlinie As String = ""
        Dim defaultBU As String = ""


        Dim startDate As Date, endDate As Date
        Dim startoffset As Long, duration As Long
        Dim vorlagenName As String = ""
        Dim variantenName As String = ""

        Dim itemName As String = ""
        Dim zufall As New Random(10)
        Dim itemDauer As Integer
        Dim colProtocol As Integer

        Dim schriftGroesse As Integer
        Dim schriftfarbe As Long

        ' Kennungen für die BMW Projekte
        Dim typKennung As String = ""
        Dim anlaufKennung As String = ""
        Dim anzProcessedElements As Integer = 0
        Dim anzSubstituted As Integer = 0
        Dim anzIgnored As Integer = 0
        Dim anzCorrect As Integer = 0

        ' 
        Dim logMessage As String = ""

        ' ur: 1.12.2015: wird nun Public awinSettings.fullProtokoll As Boolean = True  
        ' und damit global definiert, da auch in RXFImport benötigt.
        ' Dim fullProtocol As Boolean = True


        Dim milestoneIX As Integer = MilestoneDefinitions.Count + 1
        Dim phaseIX As Integer = PhaseDefinitions.Count + 1
        ' wird benötigt, um bei Phasen, die als doppelt erkannt wurden alle darunter liegenden Elemente auch zu ignorieren 
        Dim lastDuplicateIndent As Integer = 1000000

        ' bestimmen, des eventuell benötigten VariantenName. Dieser wird aus dem Dateinamen erstellt
        Dim tmpStrNew() As String
        tmpStrNew = Split(dateiname, "\", -1)
        variantenName = tmpStrNew(tmpStrNew.Length - 1)


        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        Dim colName As Integer
        Dim colAnfang As Integer
        Dim colEnde As Integer
        Dim colDauer As Integer = -1
        Dim colProduktlinie As Integer = -1
        Dim colAbbrev As Integer = -1
        Dim colVorgangsKlasse As Integer = -1
        Dim colDescription As Integer = -1
        Dim colVerantwortlich As Integer = -1
        Dim colPercentDone As Integer = -1
        Dim colTrafficLight As Integer = -1
        Dim colTLExplanation As Integer = -1
        Dim colDocUrl As Integer = -1
        Dim colDeliv As Integer = -1

        Dim pDescription As String = ""
        Dim firstZeile As Excel.Range
        Dim protocolRange As Excel.Range


        Dim suchstr(14) As String
        suchstr(ptPlanNamen.Name) = "Name"
        suchstr(ptPlanNamen.Anfang) = "Start"
        suchstr(ptPlanNamen.Ende) = "End"
        suchstr(ptPlanNamen.Beschreibung) = "Description"
        suchstr(ptPlanNamen.Vorgangsklasse) = "Appearance"
        suchstr(ptPlanNamen.BusinessUnit) = "Business Unit"
        suchstr(ptPlanNamen.Protocol) = "Übernommen als"
        suchstr(ptPlanNamen.Dauer) = "Duration"
        suchstr(ptPlanNamen.Abkuerzung) = "Abbreviation"
        suchstr(ptPlanNamen.Verantwortlich) = "Responsible"
        suchstr(ptPlanNamen.percentDone) = "%-Done"
        suchstr(ptPlanNamen.TrafficLight) = "traffic light"
        suchstr(ptPlanNamen.TLExplanation) = "Explanation"
        suchstr(ptPlanNamen.DocUrl) = "Document-Link"
        suchstr(ptPlanNamen.Deliv) = "Deliverables"

        zeile = 2
        spalte = 5
        geleseneProjekte = 0

        ' wie lautet der aktuelle Dateiname ? 
        currentDateiName = CType(appInstance.ActiveWorkbook, Excel.Workbook).Name

        ' wie lautet ggf der Default Produktlinien Name ? 
        Dim i As Integer
        Dim found As Boolean = False
        Dim tmpName As String
        i = 1
        While i <= businessUnitDefinitions.Count And Not found

            tmpName = businessUnitDefinitions.ElementAt(i - 1).Value.name
            If currentDateiName.Contains(tmpName) Then
                defaultBU = tmpName
                found = True
            Else
                i = i + 1
            End If

        End While



        Dim aktivesSheet As Excel.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

        With aktivesSheet
            firstZeile = CType(.Rows(1), Excel.Range)
        End With



        ' diese Daten müssen vorhanden sein - andernfalls Abbruch 
        Try
            colName = firstZeile.Find(What:=suchstr(ptPlanNamen.Name), LookAt:=XlLookAt.xlWhole).Column
            colAnfang = firstZeile.Find(What:=suchstr(ptPlanNamen.Anfang), LookAt:=XlLookAt.xlWhole).Column
            colEnde = firstZeile.Find(What:=suchstr(ptPlanNamen.Ende), LookAt:=XlLookAt.xlWhole).Column

        Catch ex As Exception
            Throw New ArgumentException("Fehler im Datei Aufbau ..." & vbLf & ex.Message)
        End Try

        Try
            colDauer = firstZeile.Find(What:=suchstr(ptPlanNamen.Dauer), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception
            colDauer = -1
        End Try


        Try
            colProduktlinie = firstZeile.Find(What:=suchstr(ptPlanNamen.BusinessUnit), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception
            colProduktlinie = -1
        End Try


        Try
            colAbbrev = firstZeile.Find(What:=suchstr(ptPlanNamen.Abkuerzung), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception
            colAbbrev = -1
        End Try

        Try
            colDescription = firstZeile.Find(What:=suchstr(ptPlanNamen.Beschreibung), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception
            colAbbrev = -1
        End Try

        Try
            colVorgangsKlasse = firstZeile.Find(What:=suchstr(ptPlanNamen.Vorgangsklasse), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try

        Try
            colVerantwortlich = firstZeile.Find(What:=suchstr(ptPlanNamen.Verantwortlich), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try

        Try
            colPercentDone = firstZeile.Find(What:=suchstr(ptPlanNamen.percentDone), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try

        Try
            colTrafficLight = firstZeile.Find(What:=suchstr(ptPlanNamen.TrafficLight), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try

        Try
            colTLExplanation = firstZeile.Find(What:=suchstr(ptPlanNamen.TLExplanation), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try

        Try
            colDocUrl = firstZeile.Find(What:=suchstr(ptPlanNamen.DocUrl), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try

        Try
            colDeliv = firstZeile.Find(What:=suchstr(ptPlanNamen.Deliv), LookAt:=XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try

        With aktivesSheet

            lastRow = System.Math.Max(CType(.Cells(40000, colName), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row,
                                          CType(.Cells(40000, colAnfang), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row)
        End With




        ' Hier wird die Stelle und die Informationen für das Visbo Protocoll ermittelt und gesetzt 
        Dim protocolCellName As String = "VISBO_Protocol"
        Dim pCell As Excel.Range

        With aktivesSheet
            Try
                colProtocol = .Range(protocolCellName).Column
                protocolRange = CType(.Range(.Cells(1, colProtocol - 3), .Cells(lastRow + 10, colProtocol + 200)), Excel.Range)
                protocolRange.Clear()
                protocolRange.Interior.Color = RGB(255, 255, 255)
                protocolRange.ClearFormats()

            Catch ex As Exception
                Try
                    colProtocol = CType(.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column + 4
                Catch ex1 As Exception
                    colProtocol = 20
                End Try
                pCell = .Range(.Cells(1, colProtocol), .Cells(1, colProtocol))
                appInstance.ActiveWorkbook.Names.Add(Name:=protocolCellName, RefersToR1C1:=pCell)

                ' dann müssen auch die Spaltenbreiten gesetzt werden 
                Dim tmpRange As Excel.Range
                With aktivesSheet

                    For i = -3 To 9
                        tmpRange = CType(aktivesSheet.Columns(colProtocol + i), Excel.Range)
                        tmpRange.ColumnWidth = 40
                    Next


                End With

            End Try


        End With


        ' Die Überschriften für das Protokoll werden alle wieder gesetzt 
        With aktivesSheet


            If awinSettings.fullProtocol Then

                CType(.Cells(1, colProtocol), Excel.Range).Value = "Projekt"
                CType(.Cells(1, colProtocol + 1), Excel.Range).Value = "Hierarchie"
                CType(.Cells(1, colProtocol + 2), Excel.Range).Value = "Plan-Element"
                CType(.Cells(1, colProtocol + 3), Excel.Range).Value = "Klasse"
                CType(.Cells(1, colProtocol + 4), Excel.Range).Value = "Abkürzung"
                CType(.Cells(1, colProtocol + 5), Excel.Range).Value = "Quelle"
                CType(.Cells(1, colProtocol + 8), Excel.Range).Value = "PT Hierarchie"
                CType(.Cells(1, colProtocol + 9), Excel.Range).Value = "PT Klasse"
                CType(.Cells(1, colProtocol + 10), Excel.Range).Value = "Verantwortlich"
                CType(.Cells(1, colProtocol + 11), Excel.Range).Value = "%-Done"
                CType(.Cells(1, colProtocol + 12), Excel.Range).Value = "Ampel"
                CType(.Cells(1, colProtocol + 13), Excel.Range).Value = "Explanation"
                CType(.Cells(1, colProtocol + 14), Excel.Range).Value = "Document-Link"
            End If

            ' wird immer geschrieben 
            CType(.Cells(1, colProtocol + 6), Excel.Range).Value = suchstr(ptPlanNamen.Protocol)
            CType(.Cells(1, colProtocol + 7), Excel.Range).Value = "Grund"

        End With

        Try

            With aktivesSheet

                Try
                    projektFarbe = CType(aktivesSheet.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color
                    ' das Folgende wird nur für die Projekt-Vorlagen benötigt (isVorlage = true) 
                    schriftfarbe = CLng(CType(aktivesSheet.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Font.Color)
                    schriftGroesse = CInt(CType(aktivesSheet.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Font.Size)

                Catch ex As Exception
                    projektFarbe = CType(aktivesSheet.Cells(zeile, 1), Excel.Range).Interior.ColorIndex
                End Try

                ' jetzt kommt der Check, ob Blanks als Indent verwendet werden oder echte Excel Indents
                Dim stdIndent As Boolean = True
                Dim stdIndentedRows As Integer = 0
                Dim blankIndentedRows As Integer = 0
                For ik As Integer = 1 To lastRow
                    If CType(.Cells(ik, colName), Excel.Range).IndentLevel > 0 Then
                        stdIndentedRows = stdIndentedRows + 1
                    End If
                    Dim tstString As String = CStr(CType(.Cells(ik, colName), Excel.Range).Value)
                    If tstString.StartsWith(" ") Then
                        blankIndentedRows = blankIndentedRows + 1
                    End If
                Next

                If stdIndentedRows > blankIndentedRows Then
                    stdIndent = True
                Else
                    stdIndent = False
                End If

                ' zeile ist an der Stelle 2
                While zeile <= lastRow

                    ' wenn es mit einem neuen Projekt beginnt, muss der lastDuplicateIndent zurückgesetzt sein 
                    lastDuplicateIndent = 1000000

                    ix = zeile + 1

                    Dim zellenFarbe As Long = CLng(CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color)
                    Do While zellenFarbe <> CLng(projektFarbe) And (ix <= lastRow)
                        ix = ix + 1
                        zellenFarbe = CLng(CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color)
                    Loop

                    anfang = zeile + 1
                    ende = ix - 1

                    ' hier wird Name, Typ, SOP, Business Unit, vname, Start-Datum, Dauer der Phase(1) ausgelesen  

                    ' ur: 24.06.2016:testweise auskomentiert
                    ' '' ''endDate = CDate(.Cells(RowIndex:=zeile, ColumnIndex:=colEnde).value)
                    ' '' ''startDate = CDate(.Cells(RowIndex:=zeile, ColumnIndex:=colAnfang).value)

                    ' '' ''completeName = CStr(.Cells(RowIndex:=zeile, ColumnIndex:=colName).value)

                    startDate = CDate(CType(.Cells(zeile, colAnfang), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    endDate = CDate(CType(.Cells(zeile, colEnde), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    completeName = CStr(CType(.Cells(zeile, colName), Global.Microsoft.Office.Interop.Excel.Range).Value)

                    ' andere Informationen auslesen ... 
                    pDescription = ""
                    If colDescription > 0 Then
                        pDescription = CStr(CType(.Cells(zeile, colDescription), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    End If

                    defaultBU = ""
                    If colProduktlinie > 0 Then

                        Try
                            Dim tmpBU As String
                            If colProduktlinie > 0 Then
                                tmpBU = CStr(CType(.Cells(zeile, colProduktlinie), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            Else
                                tmpBU = ""
                            End If


                            ' gibt es die Business Unit ? 
                            found = False
                            Dim bix As Integer = 1

                            If tmpBU.Length > 0 Then
                                While bix <= businessUnitDefinitions.Count And Not found
                                    If businessUnitDefinitions.ElementAt(bix - 1).Value.name = tmpBU Then

                                        found = True
                                        defaultBU = tmpBU

                                    Else
                                        bix = bix + 1
                                    End If
                                End While
                            End If


                            If Not found Then

                                CType(aktivesSheet.Cells(zeile, colProduktlinie), Excel.Range).Interior.Color = awinSettings.AmpelRot

                            End If

                        Catch ex1 As Exception

                        End Try

                    End If

                    Dim tmpvalue As String
                    Dim tmp2Str() As String

                    If colDauer > 0 Then

                        Try
                            tmpvalue = CStr(CType(aktivesSheet.Cells(zeile, colDauer), Excel.Range).Value).Trim
                            tmp2Str = tmpvalue.Trim.Split(New Char() {CChar(" ")}, 5)
                            itemDauer = CInt(tmp2Str(0))
                        Catch ex As Exception
                            itemDauer = -1
                        End Try
                    End If


                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    If duration < 0 Then
                        startDate = endDate
                        duration = -1 * duration
                        endDate = startDate.AddDays(duration)
                    End If

                    tmpStr = completeName.Trim.Split(New Char() {CChar("["), CChar("]")}, 5)

                    ' PT-71 Änderung 22.1.15 (tk) Der Projekt-Name soll der RPLAN Name sein 
                    'pName = tmpStr(0).Trim
                    ' damit alt: 
                    ' jetzt doch wieder hereingenommen, weil sich von einem Monat auf den anderen ein und dasselbe Projekte im SOP ändert .... 
                    Dim doADD As Boolean = False


                    pName = tmpStr(0)



                    ' prüfen, ob das Projekt überhaupt vollständig im Kalender liegt 
                    ' wenn nein, dann nicht importieren 
                    If DateDiff(DateInterval.Day, StartofCalendar, startDate) < 0 And Not isVorlage Then

                        Dim errMsg As String
                        If awinSettings.englishLanguage Then
                            errMsg = "project start is earlier than start of calendar in Visual Board ... No Import ... "
                        Else
                            errMsg = "Projekt liegt vor dem Kalender-Anfang und wird deshalb nicht importiert"
                        End If

                        Throw New ArgumentException(errMsg)

                    Else
                        '
                        ' jetzt wird das Projekt angelegt 
                        '
                        hproj = New clsProjekt


                        Try

                            hproj.name = pName
                            hproj.startDate = startDate

                            If DateDiff(DateInterval.Month, startDate, Date.Now) <= 0 Then
                                hproj.earliestStartDate = hproj.startDate.AddMonths(hproj.earliestStart)
                                hproj.latestStartDate = hproj.startDate.AddMonths(hproj.latestStart)
                            Else
                                hproj.earliestStartDate = startDate
                                hproj.latestStartDate = startDate
                            End If

                            hproj.StrategicFit = 5
                            hproj.Risiko = 5
                            hproj.businessUnit = defaultBU
                            hproj.description = pDescription

                            hproj.Erloes = 0.0


                        Catch ex As Exception
                            Throw New Exception("in erstelle Import Excel Projekte: " & vbLf & ex.Message)
                        End Try

                        ' jetzt wird die Import Hierarchie angelegt 
                        Dim pHierarchy As New clsImportFileHierarchy
                        Dim origHierarchy As New clsImportFileHierarchy

                        ' jetzt wird die Projekt-Hierarchie neu angelegt 
                        ' die erste Phase, die sogenannte Root Phase hat immer diesen Namen: 

                        ' jetzt werden all die Phasen angelegt , beginnend mit der ersten 
                        cphase = New clsPhase(parent:=hproj)
                        cphase.nameID = rootPhaseName
                        startoffset = 0
                        duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                        cphase.changeStartandDauer(startoffset, duration)

                        ' jetzt wird noch percentDone und Doument Link eingefügt 
                        If colPercentDone > 0 Then
                            Try
                                Dim tmpPD As Double = CDbl(CType(.Cells(zeile, colPercentDone), Excel.Range).Value)
                                cphase.percentDone = tmpPD
                            Catch ex As Exception

                            End Try
                        End If

                        If colDocUrl > 0 Then
                            Try
                                Dim tmpDU As String = CStr(CType(.Cells(zeile, colDocUrl), Excel.Range).Value)
                                cphase.DocURL = tmpDU
                            Catch ex As Exception

                            End Try
                        End If


                        hproj.AddPhase(cphase)

                        Try
                            pHierarchy.add(cphase, rootPhaseName, 0)
                            origHierarchy.add(cphase, rootPhaseName, 0)
                        Catch ex As Exception

                        End Try

                        Dim ampel As Integer = 0
                        Dim ampelExplanation As String = ""

                        ' wenn eine Ampel Bewertung für das Projekt abgegeben wurde 
                        If colTrafficLight > 0 Then
                            Try
                                ampel = CInt(CType(.Cells(zeile, colTrafficLight), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                hproj.ampelStatus = ampel
                            Catch ex As Exception

                            End Try
                        End If

                        If colTLExplanation > 0 Then
                            Try
                                ampelExplanation = CInt(CType(.Cells(zeile, colTLExplanation), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                hproj.ampelErlaeuterung = ampelExplanation
                            Catch ex As Exception

                            End Try
                        End If

                        ' jetzt die Projekt-Ampel ggf setzen 


                        Dim itemStartDate As Date
                        Dim itemEndDate As Date
                        Dim ok As Boolean = True

                        Dim curZeile As Integer
                        Dim txtVorgangsKlasse As String
                        Dim origVorgangsKlasse As String
                        Dim txtAbbrev As String
                        Dim verantwortlich As String = ""
                        Dim percentDone As Double = 0.0
                        Dim docURL As String = ""
                        ' ist notwendig um anhand der führenden Blanks die Hierarchie Stufe zu bestimmen 
                        Dim origItem As String = ""

                        ' 
                        ' Schleife, um alle Elemente des Projektes auszulesen
                        ' hier werden jetzt die einzelnen Zeilen = Phasen oder Meilensteine ausgelesen 
                        For curZeile = anfang To ende

                            origVorgangsKlasse = ""
                            txtVorgangsKlasse = ""
                            txtAbbrev = ""
                            logMessage = ""
                            verantwortlich = ""
                            percentDone = 0.0
                            ampel = 0
                            ampelExplanation = ""

                            Dim indentLevel As Integer

                            Try

                                Dim tmpName2 As String = CStr(CType(.Cells(curZeile, colName), Excel.Range).Value)

                                tmpStr = tmpName2.Split(New Char() {CChar("["), CChar("]")}, 5)
                                origItem = tmpStr(0)

                                If origItem.Trim.Length = 0 Then

                                    'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                    '            "leerer String wird ignoriert .."
                                    logMessage = "leerer String wird ignoriert .."
                                    ok = False

                                Else

                                    If stdIndent Then
                                        indentLevel = CType(.Cells(curZeile, colName), Excel.Range).IndentLevel
                                    Else
                                        indentLevel = pHierarchy.getLevel(origItem)
                                    End If

                                    ' hier checken, ob indentlevel > lastduplicateIndent; 
                                    ' wenn ja, dann protokollieren, Next for und lastduplicateIndent wieder auf hohen Wert setzen

                                    If indentLevel > lastDuplicateIndent Then
                                        ' Skip , weil es sich dann um Elemente handelt, deren Parent Phase als Duplikat ignoriert wurde 
                                        ' Protokollieren ...

                                        'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                        '            "ist Kind eines doppelten/nicht zugelassenen Elements und wird ignoriert"

                                        logMessage = "ist Kind eines doppelten/nicht zugelassenen Elements und wird ignoriert"
                                        ok = False

                                    Else
                                        lastDuplicateIndent = 1000000

                                        itemName = origItem.Trim

                                        anzProcessedElements = anzProcessedElements + 1


                                        If awinSettings.fullProtocol Then

                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 2), Excel.Range).Value = origItem.Trim
                                            CType(aktivesSheet.Cells(curZeile, colProtocol), Excel.Range).Value = completeName
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 5), Excel.Range).Value = currentDateiName
                                        End If



                                        logMessage = "ungültiges Startdatum ..."
                                        itemEndDate = CDate(CType(.Cells(curZeile, colEnde), Excel.Range).Value)
                                        logMessage = ""


                                        If IsNothing(CType(.Cells(curZeile, colAnfang), Excel.Range).Value) Then
                                            isMilestone = True
                                            itemStartDate = itemEndDate
                                        ElseIf CStr(CType(.Cells(curZeile, colAnfang), Excel.Range).Value).Trim = "" Then
                                            isMilestone = True
                                            itemStartDate = itemEndDate
                                        Else
                                            ' jetzt das Startdatum lesen 
                                            logMessage = "ungültiges Startdatum ..."
                                            itemStartDate = CDate(CType(.Cells(curZeile, colAnfang), Excel.Range).Value)
                                            logMessage = ""

                                            If DateDiff(DateInterval.Minute, itemStartDate, itemEndDate) = 0 Then
                                                isMilestone = True
                                            Else
                                                isMilestone = False
                                            End If
                                        End If



                                        ' jetzt prüfen, ob es sich um ein grundsätzlich zu ignorierendes Element handelt .. 
                                        If isMilestone Then
                                            If MilestoneDefinitions.Contains(itemName) Then
                                                ok = True
                                            ElseIf milestoneMappings.tobeIgnored(itemName) Then
                                                'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                                '                "nicht zugelassen (lt. Wörterbuch ignorieren)"

                                                logMessage = "nicht zugelassen (lt. Wörterbuch ignorieren)"
                                                ok = False
                                                lastDuplicateIndent = indentLevel
                                            Else
                                                ok = True
                                            End If


                                        Else

                                            If PhaseDefinitions.Contains(itemName) Then
                                                ok = True
                                            ElseIf phaseMappings.tobeIgnored(itemName) Then
                                                'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                                '                "nicht zugelassen (lt. Wörterbuch ignorieren)"
                                                logMessage = "nicht zugelassen (lt. Wörterbuch ignorieren)"
                                                lastDuplicateIndent = indentLevel
                                                ok = False
                                            Else
                                                ok = True

                                            End If

                                        End If

                                    End If

                                End If

                            Catch ex As Exception
                                itemName = ""
                                ok = False
                            End Try


                            If ok Then


                                startoffset = DateDiff(DateInterval.Day, hproj.startDate, itemStartDate)
                                duration = DateDiff(DateInterval.Day, itemStartDate, itemEndDate) + 1


                                ' jetzt werden vorgangsklasse und Abkürzung rausgelesen 
                                If colVorgangsKlasse > 0 Then
                                    Try

                                        origVorgangsKlasse = CStr((CType(.Cells(curZeile, colVorgangsKlasse), Excel.Range).Value)).Trim
                                        If duration > 1 Then
                                            txtVorgangsKlasse = mapToAppearance(origVorgangsKlasse, False)
                                            'CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                            '        "auf folgende Phasen Darstellungsklasse abgebildet: " & txtVorgangsKlasse.Trim
                                        Else
                                            txtVorgangsKlasse = mapToAppearance(origVorgangsKlasse, True)
                                            'CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                            '        "auf folgende Meilenstein Darstellungsklasse abgebildet: " & txtVorgangsKlasse.Trim
                                        End If




                                    Catch ex As Exception

                                        'CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                        '            "Fehler bei Abbildung auf Darstellungsklasse ... " & txtVorgangsKlasse.Trim

                                    End Try
                                End If


                                ' jetzt wird die Abkürzung rausgelesen 
                                If colAbbrev > 0 Then
                                    Try

                                        txtAbbrev = CStr((CType(.Cells(curZeile, colAbbrev), Excel.Range).Value)).Trim

                                    Catch ex As Exception
                                        txtAbbrev = ""
                                    End Try
                                End If

                                If colVerantwortlich > 0 Then
                                    Try
                                        verantwortlich = CStr(CType(.Cells(curZeile, colVerantwortlich), Excel.Range).Value)
                                    Catch ex As Exception
                                        verantwortlich = ""
                                    End Try
                                End If

                                ' jetzt %-Done auslesen 
                                If colPercentDone > 0 Then
                                    Try
                                        percentDone = CDbl(CType(.Cells(curZeile, colPercentDone), Excel.Range).Value)
                                    Catch ex As Exception
                                        percentDone = 0.0
                                    End Try
                                End If

                                ' jetzt Ampel-Farbe  auslesen 
                                If colTrafficLight > 0 Then
                                    Try
                                        ampel = CInt(CType(.Cells(curZeile, colTrafficLight), Excel.Range).Value)
                                    Catch ex As Exception
                                        ampel = 0
                                    End Try
                                End If

                                ' jetzt Ampel-Erläuterung  auslesen 
                                If colTLExplanation > 0 Then
                                    Try
                                        ampelExplanation = CStr(CType(.Cells(curZeile, colTLExplanation), Excel.Range).Value)
                                    Catch ex As Exception
                                        ampelExplanation = ""
                                    End Try
                                End If

                                ' jetzt Document Link auslesen 
                                If colDocUrl > 0 Then
                                    Try
                                        docURL = CStr(CType(.Cells(zeile, colDocUrl), Excel.Range).Value)
                                    Catch ex As Exception

                                    End Try
                                End If

                                '
                                ' jetzt muss protokolliert werden 
                                Dim oLevel As Integer
                                oLevel = origHierarchy.getLevel(origItem)
                                Dim oBreadCrumb As String = origHierarchy.getFootPrint(oLevel)


                                If awinSettings.fullProtocol Then

                                    ' Original Footprint
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 1), Excel.Range).Value = oBreadCrumb
                                    ' Textvorgangsklasse
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 3), Excel.Range).Value = origVorgangsKlasse
                                    ' Abkürzung
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 4), Excel.Range).Value = txtAbbrev
                                    ' %-Done
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 11), Excel.Range).Value = percentDone.ToString
                                    ' Ampel 
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 12), Excel.Range).Value = ampel.ToString
                                    ' Erläuterung 
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 13), Excel.Range).Value = ampelExplanation
                                    ' Dokumenten Link 
                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 14), Excel.Range).Value = docURL
                                End If


                                ' jetzt muss ggf die Phase in die Orig Hierarchie aufgenommen werden 
                                If Not isMilestone Then

                                    Dim ophase As clsPhase
                                    ophase = New clsPhase(parent:=hproj)
                                    ophase.nameID = calcHryElemKey(origItem.Trim, False)
                                    'ophase.changeStartandDauer(startoffset, duration)

                                    Try
                                        origHierarchy.add(ophase, "dummy", oLevel)
                                    Catch ex As Exception

                                    End Try


                                End If

                                Dim stdName As String
                                Dim parentElemName As String
                                Dim parentNodeID As String
                                Dim elemID As String

                                ' If duration > 1 Or itemDauer > 0 Then
                                If duration > 1 Then
                                    ' es handelt sich um eine Phase 


                                    parentElemName = pHierarchy.getPhaseBeforeLevel(indentLevel).name
                                    ' das folgende wurde am 31.3. ergänzt, um die Hierarchie aufbauen zu können
                                    parentNodeID = pHierarchy.getIDBeforeLevel(indentLevel)

                                    ' Plausibilitäts-Check: die beiden müssen identisch sein !!
                                    ' tk Debug: 27.11.15
                                    'If elemNameOfElemID(parentNodeID) <> parentElemName Then
                                    '    Call MsgBox("nicht konsistent in bmwImportProjekteITO15, zeile 663")
                                    'End If


                                    ' jetzt den tatsächlichen Namen bestimmen , ggf wird dazu der Parent Phase Name benötigt 
                                    Try

                                        If Not PhaseDefinitions.Contains(itemName) Then
                                            stdName = phaseMappings.mapToStdName(parentElemName, itemName)
                                        Else
                                            stdName = itemName
                                        End If

                                    Catch ex As Exception
                                        stdName = itemName
                                    End Try


                                    Dim ok1 As Boolean = True


                                    'Dim breadcrumb As String = pHierarchy.getFootPrint(indentLevel, "#")
                                    Dim parentPhase As clsPhase = pHierarchy.getPhaseBeforeLevel(indentLevel)
                                    Dim parentphaseName As String = ""

                                    If Not IsNothing(parentPhase) Then
                                        parentphaseName = parentPhase.name
                                    End If


                                    ' sollen Duplikate eliminiert werden ?
                                    If awinSettings.eliminateDuplicates And hproj.hierarchy.containsKey(calcHryElemKey(stdName, False)) Then
                                        ' nur dann kann es Duplikate geben 
                                        If hproj.isCloneToParent(stdName, parentPhase.nameID, itemStartDate, itemEndDate, 0.97) Then
                                            ok1 = False
                                            logMessage = stdName & " ist Duplikat zu Parent " & parentPhase.name & " und wird ignoriert "

                                        Else
                                            Dim duplicateSiblingID As String = hproj.getDuplicatePhaseSiblingID(stdName, parentPhase.nameID,
                                                                                                                 itemStartDate, itemEndDate, 0.97)

                                            If duplicateSiblingID = "" Then
                                                ok1 = True
                                            Else
                                                ok1 = False
                                                logMessage = stdName & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) &
                                                             " und wird ignoriert "
                                            End If
                                        End If



                                    End If



                                    ' jetzt muss geprüft werden, ob das Element in Std Definitions aufgenommen werden muss 
                                    Dim ok2 As Boolean = True
                                    If Not PhaseDefinitions.Contains(stdName) And ok1 Then

                                        Dim hphaseDef As clsPhasenDefinition
                                        hphaseDef = New clsPhasenDefinition

                                        hphaseDef.darstellungsKlasse = txtVorgangsKlasse
                                        hphaseDef.shortName = txtAbbrev
                                        hphaseDef.name = stdName
                                        hphaseDef.UID = phaseIX
                                        phaseIX = phaseIX + 1


                                        If isVorlage And awinSettings.alwaysAcceptTemplateNames Then
                                            ' in die Phase-Definitions aufnehmen 
                                            Try
                                                PhaseDefinitions.Add(hphaseDef)
                                            Catch ex As Exception
                                            End Try
                                        Else
                                            ' in Abhängigkeit vom Setting die Elemente aufnehmen oder nicht 
                                            Try
                                                If awinSettings.importUnknownNames Then
                                                    ok2 = True
                                                Else
                                                    ok2 = False
                                                    logMessage = "ist nicht in der Liste der zugelassenen Elemente enthalten"
                                                End If
                                                missingPhaseDefinitions.Add(hphaseDef)
                                            Catch ex As Exception
                                            End Try


                                        End If

                                    End If

                                    ' hier muss noch der letzte Check rein 

                                    If ok1 And ok2 Then

                                        ' hier muss jetzt überprüft werden, ob es Geschwister mit gleichen Namen gibt
                                        ' wenn ja , wird an den stdName solange eine ldfNR Ergänzung rangemacht, bis der NAme innerhalb der 
                                        ' Geschwistergruppe eindeutig ist

                                        ' Bestimmung des eindeutigen Namens innerhalb der Geschwister, unterschieden nach Meilensten  und Phase 
                                        If awinSettings.createUniqueSiblingNames Then
                                            stdName = hproj.hierarchy.findUniqueGeschwisterName(parentNodeID, stdName, False)
                                        End If

                                        elemID = hproj.hierarchy.findUniqueElemKey(stdName, False)

                                        ' das muss auf alle Fälle gemacht werden 
                                        cphase = New clsPhase(parent:=hproj)

                                        ' Änderung tk: jetzt muss die elemID in den Phasen Namen 
                                        cphase.nameID = elemID
                                        cphase.changeStartandDauer(startoffset, duration)

                                        ' tk 26.11.17, den Wert für verantwortlich mitaufnehmen ...
                                        cphase.verantwortlich = verantwortlich

                                        ' Vorgangslasse eintragen 
                                        cphase.appearance = txtVorgangsKlasse

                                        ' percentDone eintragen 
                                        cphase.percentDone = percentDone

                                        ' ampel eintragen 
                                        If ampel > 0 And ampel <= 3 Then
                                            cphase.ampelStatus = ampel
                                        End If

                                        ' ampel Erläuterung eintragen 
                                        If ampelExplanation <> "" Then
                                            cphase.ampelErlaeuterung = ampelExplanation
                                        End If

                                        If docURL <> "" Then
                                            cphase.DocURL = docURL
                                        End If


                                        ' der Aufbau der Hierarchie erfolgt in addphase
                                        hproj.AddPhase(cphase, origName:=origItem.Trim,
                                                       parentID:=pHierarchy.getIDBeforeLevel(indentLevel))

                                        ' wird übernommen als 
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Value = stdName

                                        Dim PTBreadCrumb As String = hproj.hierarchy.getBreadCrumb(elemID)


                                        If awinSettings.fullProtocol Then

                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 8), Excel.Range).Value = PTBreadCrumb
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 9), Excel.Range).Value = txtVorgangsKlasse
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 10), Excel.Range).Value = verantwortlich

                                        End If

                                        ' neuer Breadcrumb 
                                        'Dim PTBreadCrumb As String = pHierarchy.getFootPrint(indentLevel)

                                        If stdName.Trim <> origItem.Trim Then
                                            ' es hat eine Ersetzung stattgefunden 
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                            anzSubstituted = anzSubstituted + 1
                                        ElseIf PhaseDefinitions.Contains(stdName.Trim) Then
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGruen
                                            anzCorrect = anzCorrect + 1
                                        Else
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                                        End If

                                        ' nur wenn es aufgenommen ist, sollte es in die Hierarchie aufgenommen werden 
                                        Try
                                            pHierarchy.add(cphase, elemID, indentLevel)
                                        Catch ex As Exception

                                        End Try

                                    Else

                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = logMessage
                                        lastDuplicateIndent = indentLevel

                                        anzIgnored = anzIgnored + 1

                                    End If


                                ElseIf duration = 1 Then
                                    ' hier kommt die Behandlung eines Meilensteins


                                    Try



                                        ' hole die Parentphase
                                        cphase = pHierarchy.getPhaseBeforeLevel(indentLevel)
                                        cmilestone = New clsMeilenstein(parent:=cphase)
                                        cbewertung = New clsBewertung


                                        ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                        With cbewertung
                                            '.bewerterName = resultVerantwortlich
                                            .colorIndex = ampel
                                            .datum = Date.Now
                                            .description = ampelExplanation
                                        End With


                                        parentElemName = cphase.name
                                        ' jetzt den tatsächlichen Namen bestimmen , ggf wird dazu der Parent Phase Name benötigt 

                                        Try
                                            If Not MilestoneDefinitions.Contains(itemName) Then
                                                stdName = milestoneMappings.mapToStdName(parentElemName, itemName)
                                            Else
                                                stdName = itemName
                                            End If

                                        Catch ex As Exception
                                            stdName = itemName
                                        End Try

                                        Dim ok1 As Boolean = True

                                        If awinSettings.eliminateDuplicates And hproj.hierarchy.containsKey(calcHryElemKey(stdName, True)) Then
                                            ' nur dann kann es Duplikate geben 
                                            Dim duplicateSiblingID As String = hproj.getDuplicateMsSiblingID(stdName, cphase.nameID,
                                                                                                                 itemStartDate, 0)

                                            If duplicateSiblingID = "" Then
                                                ok1 = True
                                            Else
                                                ok1 = False
                                                logMessage = stdName & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) &
                                                             " und wird ignoriert "
                                            End If

                                        End If


                                        ' jetzt muss geprüft werden, ob stdName bereits aufgenommen ist
                                        Dim ok2 As Boolean = True
                                        If Not MilestoneDefinitions.Contains(stdName) And ok1 Then

                                            Dim hMilestoneDef As New clsMeilensteinDefinition

                                            With hMilestoneDef
                                                .name = stdName
                                                .belongsTo = parentElemName
                                                .shortName = txtAbbrev
                                                .darstellungsKlasse = txtVorgangsKlasse
                                                .UID = milestoneIX
                                            End With

                                            milestoneIX = milestoneIX + 1

                                            If isVorlage And awinSettings.alwaysAcceptTemplateNames Then
                                                ' in die Milestone-Definitions aufnehmen 
                                                Try
                                                    MilestoneDefinitions.Add(hMilestoneDef)
                                                Catch ex As Exception
                                                End Try

                                            Else

                                                logMessage = "ist nicht in der Liste der zugelassenen Elemente enthalten"

                                                ' in die Missing Milestone-Definitions aufnehmen 
                                                Try
                                                    ' das Element aufnehmen, in Abhängigkeit vom Setting 
                                                    If awinSettings.importUnknownNames Then
                                                        ok2 = True
                                                    Else
                                                        ok2 = False
                                                    End If

                                                    missingMilestoneDefinitions.Add(hMilestoneDef)
                                                Catch ex As Exception
                                                End Try
                                            End If


                                        End If

                                        If ok1 And ok2 Then


                                            ' Bestimmung des eindeutigen Namens innerhalb der Geschwister, unterschieden nach Meilenstein und Phase 
                                            If awinSettings.createUniqueSiblingNames Then
                                                stdName = hproj.hierarchy.findUniqueGeschwisterName(cphase.nameID, stdName, True)
                                            End If

                                            elemID = hproj.hierarchy.findUniqueElemKey(stdName, True)


                                            With cmilestone
                                                .nameID = elemID
                                                .setDate = itemEndDate
                                                ' tk 26.11.17 
                                                .verantwortlich = verantwortlich
                                                .appearance = txtVorgangsKlasse
                                                .percentDone = percentDone
                                                .DocURL = docURL

                                                If Not cbewertung Is Nothing Then
                                                    .addBewertung(cbewertung)
                                                End If
                                            End With

                                            If IsNothing(cphase.getMilestone(cmilestone.nameID)) Then

                                                With cphase
                                                    .addMilestone(cmilestone, origName:=origItem.Trim)
                                                End With

                                                ' Protokollieren
                                                CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Value = stdName.Trim

                                                ' neuer Breadcrumb 
                                                'Dim PTBreadCrumb As String = pHierarchy.getFootPrint(indentLevel)
                                                Dim PTBreadCrumb As String = hproj.hierarchy.getBreadCrumb(elemID)


                                                If awinSettings.fullProtocol Then

                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 8), Excel.Range).Value = PTBreadCrumb
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 9), Excel.Range).Value = txtVorgangsKlasse
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 10), Excel.Range).Value = verantwortlich
                                                End If

                                                If stdName.Trim <> origItem.Trim Then
                                                    ' es hat eine Ersetzung stattgefunden 
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                                    anzSubstituted = anzSubstituted + 1
                                                ElseIf MilestoneDefinitions.Contains(stdName.Trim) Then
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGruen
                                                    anzCorrect = anzCorrect + 1
                                                Else
                                                    CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelGruen

                                                End If


                                            Else

                                                ' Meilenstein existiert in dieser Phase bereits .... 
                                                CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value =
                                                        stdName.Trim & " existiert bereits: Datum 1: " & cphase.getMilestone(stdName).getDate.ToShortDateString &
                                                        "   , Datum 2: " & cmilestone.getDate.ToShortDateString

                                            End If
                                        Else

                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = logMessage
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                            anzIgnored = anzIgnored + 1

                                        End If


                                    Catch ex As Exception
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value =
                                                            ex.Message & ": " & vbLf & "Fehler in Zeile " & curZeile & ", Item-Name: " & itemName
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                    End Try


                                End If

                            Else
                                CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = logMessage
                                CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                anzIgnored = anzIgnored + 1
                            End If

                        Next


                        If Not isVorlage Then

                            ' das ist BMW spezifisch und wird jetzt de-aktiviert .... 
                            'Try
                            '    Dim sopDate As Date = hproj.getMilestone("SOP").getDate

                            '    If DateDiff(DateInterval.Month, StartofCalendar, sopDate) > 0 Then
                            '        Dim sopMonth As Integer = sopDate.Month
                            '        If sopMonth >= 3 And sopMonth <= 6 Then
                            '            anlaufKennung = "03"
                            '        ElseIf sopMonth >= 7 And sopMonth <= 10 Then
                            '            anlaufKennung = "07"
                            '        Else
                            '            anlaufKennung = "11"
                            '        End If
                            '    Else
                            '        anlaufKennung = "?"
                            '    End If

                            'Catch ex As Exception
                            '    anlaufKennung = "?"
                            'End Try

                            ' jetzt wird die Vorlagen Kennung bestimmt 
                            'Dim tstphase As clsPhase = Nothing
                            'Dim relNr As String
                            'tstphase = hproj.getPhase("Systemgestaltung")

                            'If IsNothing(tstphase) Then
                            '    tstphase = hproj.getPhase("I500")
                            '    If IsNothing(tstphase) Then
                            '        tstphase = hproj.getPhase("I300")
                            '        If IsNothing(tstphase) Then
                            '            relNr = "rel 4 "
                            '        Else
                            '            relNr = "rel 5 "
                            '        End If
                            '    Else
                            '        relNr = "rel 5 "
                            '    End If
                            'Else
                            '    relNr = "rel 5 "
                            'End If

                            'vorlagenName = relNr & typKennung & "-" & anlaufKennung
                            'Try
                            '    vorlagenName = vorlagenName.Trim
                            'Catch ex As Exception
                            '    vorlagenName = "unknown"
                            'End Try

                            'If Projektvorlagen.Contains(vorlagenName) Then
                            '    hproj.VorlagenName = vorlagenName
                            'Else
                            '    hproj.VorlagenName = vorlagenName & "*"
                            'End If

                            vorlagenName = ""
                            'If Projektvorlagen.Count >= 1 Then
                            '    vorlagenName = Projektvorlagen.getProject(0).VorlagenName
                            '    hproj.VorlagenName = vorlagenName
                            'End If

                        End If

                        Try

                            If isVorlage Then
                                hproj.farbe = projektFarbe
                                hproj.Schrift = schriftGroesse
                                hproj.Schriftfarbe = schriftfarbe
                            Else

                                If Projektvorlagen.Contains(vorlagenName) Then
                                    vproj = Projektvorlagen.getProject(vorlagenName)

                                    hproj.farbe = vproj.farbe
                                    hproj.Schrift = vproj.Schrift
                                    hproj.Schriftfarbe = vproj.Schriftfarbe
                                    hproj.earliestStart = vproj.earliestStart
                                    hproj.latestStart = vproj.latestStart

                                    'ElseIf Projektvorlagen.Contains("unknown") Then
                                    '    vproj = Projektvorlagen.getProject("unknown")
                                Else
                                    'Throw New Exception("es gibt weder die Vorlage 'unknown' noch die Vorlage " & vorlagenName)
                                    'hproj.farbe = awinSettings.AmpelNichtBewertet
                                    Try
                                        hproj.farbe = CInt(iProjektFarbe)
                                    Catch ex As Exception
                                        hproj.farbe = awinSettings.AmpelNichtBewertet
                                    End Try
                                    hproj.Schrift = Projektvorlagen.getProject(0).Schrift
                                    hproj.Schriftfarbe = RGB(10, 10, 10)
                                    hproj.earliestStart = 0
                                    hproj.latestStart = 0

                                End If




                            End If

                        Catch ex As Exception
                            Throw New Exception(ex.Message)
                        End Try


                        If Not isVorlage And awinSettings.fullProtocol Then

                            ' jetzt werden Projekt-Name, Business Unit und Vorlagen-Kennung weggeschreiben 
                            CType(aktivesSheet.Cells(anfang - 1, colProtocol - 3), Excel.Range).Value = hproj.name
                            CType(aktivesSheet.Cells(anfang - 1, colProtocol - 2), Excel.Range).Value = hproj.VorlagenName
                            CType(aktivesSheet.Cells(anfang - 1, colProtocol - 1), Excel.Range).Value = hproj.businessUnit
                        End If

                        ' jetzt muss das Projekt eingetragen werden 
                        ' ####################################################################
                        ' prüfen ob das Projekt bereits in Session oder Datenbank existiert 

                        vglName = calcProjektKey(hproj.name, hproj.variantName)
                        ' 
                        If ImportProjekte.Containskey(vglName) Then

                            ' dann existiert es bereits in der Session

                            vglProj = ImportProjekte.getProject(vglName)
                            If IsNothing(vglProj) Then
                                ' dieser Fall kann eigentlich gar nicht auftreten ... ? 
                                Call MsgBox("Fehler mit " & vglName)

                            Else
                                ' prüfen, ob es unterschiedlich ist; 
                                ' wenn ja , dann wird es unter dem Varianten Namen Datei-Name angelegt
                                ' wenn der auch schon existiert, dann Fehler udn nichts anlegen ...
                                Dim unterschiede As Collection = hproj.listOfDifferences(vglProj, True, 0)

                                If unterschiede.Count > 0 Then
                                    '' '' '' es gibt Unterschiede, also muss eine Variante angelegt werden 
                                    If hproj.variantName <> variantenName Then
                                        hproj.variantName = variantenName
                                        vglName = calcProjektKey(hproj.name, hproj.variantName)

                                        ' wenn die Variante bereits in der Session existiert ..
                                        ' wird die bisherige gelöscht , die neue über ImportProjekte neu aufgenommen  
                                        If AlleProjekte.Containskey(vglName) Then
                                            AlleProjekte.Remove(vglName)
                                        End If

                                    Else
                                        ' in diesem Fall wird die Variante über hproj neu angelegt 
                                        AlleProjekte.Remove(vglName)
                                    End If

                                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, hproj.tfZeile)

                                    Try
                                        myCollection.Add(vglName, vglName)
                                    Catch ex As Exception

                                    End Try

                                Else
                                    ' Projekt in der Form existiert bereits , keine Neu-Anlage
                                    ' es muss sichergestellt sein, dass es angezeigt wird und die Portfolio Definition entsprechend angepasst wird 
                                    ok = False
                                    hproj = vglProj

                                    Call replaceProjectVariant(hproj.name, hproj.variantName, False, True, hproj.tfZeile)

                                    Try
                                        myCollection.Add(vglName, vglName)
                                    Catch ex As Exception

                                    End Try
                                End If
                            End If


                        End If



                        If Not ImportProjekte.Containskey(calcProjektKey(hproj)) Then
                            ImportProjekte.Add(hproj, False)
                            myCollection.Add(calcProjektKey(hproj))

                        End If

                        zeile = ende + 1

                    End If

                End While

                ' jetzt wird die Statistik geschreiben ....
                'CType(activeWSListe.Cells(1, colProtocol + 10), Excel.Range).Value = "Anzahl Insgesamt"
                'CType(activeWSListe.Cells(2, colProtocol + 10), Excel.Range).Value = anzProcessedElements

                'CType(activeWSListe.Cells(1, colProtocol + 11), Excel.Range).Value = "Original Namen"
                'CType(activeWSListe.Cells(2, colProtocol + 11), Excel.Range).Value = anzCorrect

                'CType(activeWSListe.Cells(1, colProtocol + 12), Excel.Range).Value = "Korrigierte Namen"
                'CType(activeWSListe.Cells(2, colProtocol + 12), Excel.Range).Value = anzSubstituted

                'CType(activeWSListe.Cells(1, colProtocol + 13), Excel.Range).Value = "Ignorierte Namen"
                'CType(activeWSListe.Cells(2, colProtocol + 13), Excel.Range).Value = anzIgnored

                '
                ' jetzt werden die Missing Phase- und Milestone Definitions noch weggeschrieben 
                '

                ' aber nur, wenn awinSettings.fullProtokoll = true 


                If awinSettings.fullProtocol Then


                    Dim tmpzeile As Integer
                    tmpzeile = 1

                    Dim wsName As String = "unbekannte Phasen"
                    Dim txtrange As Excel.Range
                    Dim tmpWS As Excel.Worksheet

                    If missingPhaseDefinitions.Count > 0 Then
                        Try
                            tmpWS = CType(appInstance.ActiveWorkbook.Worksheets(wsName), Excel.Worksheet)
                            With tmpWS
                                txtrange = .Range(.Cells(1, 1), .Cells(5000, 8))
                            End With
                            txtrange.Clear()
                        Catch ex As Exception
                            tmpWS = CType(appInstance.ActiveWorkbook.Worksheets.Add(After:=aktivesSheet), Excel.Worksheet)
                            tmpWS.Name = wsName
                        End Try


                        CType(tmpWS.Cells(tmpzeile, 1), Excel.Range).Value = "Phasen-Name"
                        CType(tmpWS.Cells(tmpzeile, 6), Excel.Range).Value = "Abkürzung"
                        CType(tmpWS.Cells(tmpzeile, 7), Excel.Range).Value = "Darstellungsklasse"


                        Dim phDef As clsPhasenDefinition
                        For i = 1 To missingPhaseDefinitions.Count

                            phDef = missingPhaseDefinitions.getPhaseDef(i)
                            CType(tmpWS.Cells(tmpzeile + i, 1), Excel.Range).Value = phDef.name
                            CType(tmpWS.Cells(tmpzeile + i, 6), Excel.Range).Value = phDef.shortName
                            CType(tmpWS.Cells(tmpzeile + i, 7), Excel.Range).Value = phDef.darstellungsKlasse

                        Next
                    End If



                    '
                    ' jetzt werden die Missing Milestone Definitions noch weggeschrieben 
                    '
                    If missingMilestoneDefinitions.Count > 0 Then

                        tmpzeile = 1

                        wsName = "unbekannte Meilensteine"

                        Try
                            tmpWS = CType(appInstance.ActiveWorkbook.Worksheets(wsName), Excel.Worksheet)
                            With tmpWS
                                txtrange = .Range(.Cells(1, 1), .Cells(5000, 8))
                            End With
                            txtrange.Clear()
                        Catch ex As Exception
                            tmpWS = CType(appInstance.ActiveWorkbook.Worksheets.Add(After:=aktivesSheet), Excel.Worksheet)
                            tmpWS.Name = wsName
                        End Try


                        CType(tmpWS.Cells(tmpzeile, 1), Excel.Range).Value = "Meilenstein-Name"
                        CType(tmpWS.Cells(tmpzeile, 5), Excel.Range).Value = "Bezug"
                        CType(tmpWS.Cells(tmpzeile, 6), Excel.Range).Value = "Abkürzung"
                        CType(tmpWS.Cells(tmpzeile, 7), Excel.Range).Value = "Darstellungsklasse"


                        Dim msDef As clsMeilensteinDefinition
                        For i = 1 To missingMilestoneDefinitions.Count

                            msDef = missingMilestoneDefinitions.getMilestoneDef(i)
                            If Not IsNothing(msDef) Then
                                CType(tmpWS.Cells(tmpzeile + i, 1), Excel.Range).Value = msDef.name
                                CType(tmpWS.Cells(tmpzeile + i, 5), Excel.Range).Value = msDef.belongsTo
                                CType(tmpWS.Cells(tmpzeile + i, 6), Excel.Range).Value = msDef.shortName
                                CType(tmpWS.Cells(tmpzeile + i, 7), Excel.Range).Value = msDef.darstellungsKlasse
                            End If


                        Next


                    End If

                End If

                If appInstance.ActiveSheet.name <> aktivesSheet.Name Then
                    aktivesSheet.Activate()
                End If

            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei " & vbLf & ex.Message & vbLf &
                                 currentDateiName & vbLf)
        End Try


    End Sub

    ''' <summary>
    ''' exportiert das angegebene Projekt in die bereits geöffnete Datei 
    ''' Das Schreiben beginnt ab "zeile"
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="zeile"></param>
    ''' <remarks></remarks>
    Public Sub planExportProject(ByVal hproj As clsProjekt, ByRef zeile As Integer)

        Dim ip As Integer, im As Integer
        Dim startdate As Date, endDate As Date
        Dim curName As String
        Dim color As Long
        Dim ws As Excel.Worksheet
        Dim spalte As Integer = 1
        Dim cphase As clsPhase
        Dim cmilestone As clsMeilenstein
        Dim indentlevel As Integer = 0
        Dim indentDelta As Integer = 3

        ' diese Datei muss offen sein und das aktive Workbook
        ' wenn nein, dann aktivieren ! 
        Try
            If appInstance.ActiveWorkbook.Name <> excelExportVorlage Then
                appInstance.Workbooks(excelExportVorlage).Activate()
            End If
        Catch ex As Exception
            Throw New ArgumentException("Export Vorlage ist nicht die aktive Excel Datei")
        End Try

        ' bestimme die Farbe - sie steht im Excel Ausgabe File in der Zeile 2, Spalte 1 
        ws = CType(appInstance.ActiveWorkbook.Worksheets("Export VISBO Projekttafel"), Excel.Worksheet)

        Dim suchstr(13) As String
        suchstr(ptPlanNamen.Name) = "Name"
        suchstr(ptPlanNamen.Anfang) = "Start"
        suchstr(ptPlanNamen.Ende) = "End"
        suchstr(ptPlanNamen.Beschreibung) = "Description"
        suchstr(ptPlanNamen.Vorgangsklasse) = "Appearance"
        suchstr(ptPlanNamen.BusinessUnit) = "Business Unit"
        suchstr(ptPlanNamen.Protocol) = "Übernommen als"
        suchstr(ptPlanNamen.Dauer) = "Duration"
        suchstr(ptPlanNamen.Abkuerzung) = "Abbreviation"
        suchstr(ptPlanNamen.Verantwortlich) = "Responsible"
        suchstr(ptPlanNamen.percentDone) = "%-Done"
        suchstr(ptPlanNamen.TrafficLight) = "traffic light"
        suchstr(ptPlanNamen.TLExplanation) = "Explanation"
        suchstr(ptPlanNamen.DocUrl) = "Document-Link"

        ' jetzt werden die Spaltenüberschriften geschrieben 
        Dim üColor As Long = CLng(CType(ws.Cells(1, 1), Excel.Range).Interior.Color)

        CType(ws.Rows(1), Excel.Range).Interior.Color = üColor
        With CType(ws.Cells(1, 1), Excel.Range)
            .Copy()
        End With
        With CType(ws.Rows(1), Excel.Range)
            .PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, False, False)
        End With
        'With CType(ws.Rows(1), Excel.Range).Borders(XlBordersIndex.xlEdgeBottom)
        '    .LineStyle = XlLineStyle.xlContinuous
        '    .ColorIndex = 1
        '    .TintAndShade = 0
        '    .Weight = XlBorderWeight.xlThick
        'End With

        Dim colName As Integer = 1
        CType(ws.Cells(1, colName), Excel.Range).Value = suchstr(ptPlanNamen.Name)
        Dim colStart As Integer = 2
        CType(ws.Cells(1, colStart), Excel.Range).Value = suchstr(ptPlanNamen.Anfang)
        Dim colEnde As Integer = 3
        CType(ws.Cells(1, colEnde), Excel.Range).Value = suchstr(ptPlanNamen.Ende)
        Dim colBU As Integer = 4
        CType(ws.Cells(1, colBU), Excel.Range).Value = suchstr(ptPlanNamen.BusinessUnit)
        Dim colDescription As Integer = 5
        CType(ws.Cells(1, colDescription), Excel.Range).Value = suchstr(ptPlanNamen.Beschreibung)
        Dim colAppearance As Integer = 6
        CType(ws.Cells(1, colAppearance), Excel.Range).Value = suchstr(ptPlanNamen.Vorgangsklasse)
        Dim colAbbrev As Integer = 7
        CType(ws.Cells(1, colAbbrev), Excel.Range).Value = suchstr(ptPlanNamen.Abkuerzung)
        Dim colRespons As Integer = 8
        CType(ws.Cells(1, colRespons), Excel.Range).Value = suchstr(ptPlanNamen.Verantwortlich)
        Dim colPercent As Integer = 9
        CType(ws.Cells(1, colPercent), Excel.Range).Value = suchstr(ptPlanNamen.percentDone)
        Dim colAmpel As Integer = 10
        CType(ws.Cells(1, colAmpel), Excel.Range).Value = suchstr(ptPlanNamen.TrafficLight)
        Dim colExplan As Integer = 11
        CType(ws.Cells(1, colExplan), Excel.Range).Value = suchstr(ptPlanNamen.TLExplanation)
        Dim colDocUrl As Integer = 12
        CType(ws.Cells(1, colDocUrl), Excel.Range).Value = suchstr(ptPlanNamen.DocUrl)




        color = CLng(CType(ws.Cells(2, 1), Excel.Range).Interior.Color)

        ' jetzt wird das Projekt geschrieben 
        CType(ws.Cells(zeile, colName), Excel.Range).Value = hproj.getShapeText
        CType(ws.Cells(zeile, colStart), Excel.Range).Value = hproj.startDate.ToShortDateString
        CType(ws.Cells(zeile, colEnde), Excel.Range).Value = hproj.endeDate.ToShortDateString
        CType(ws.Cells(zeile, colBU), Excel.Range).Value = hproj.businessUnit
        CType(ws.Cells(zeile, colDescription), Excel.Range).Value = hproj.description
        CType(ws.Rows(zeile), Excel.Range).Interior.Color = color


        Dim indentPhase As String = "   "
        'Dim indentMS As String = "      "

        ' die erste Phase kann auch Meilensteine haben !
        cphase = hproj.getPhase(1)

        ' percentDone eintragen 
        Try
            Dim tmpPercentDone As String = (cphase.percentDone * 100).ToString & " %"
            CType(ws.Cells(zeile, colPercent), Excel.Range).Value = tmpPercentDone
        Catch ex As Exception

        End Try

        ' Document Link eintragen
        Try
            Dim docUrl As String = cphase.DocURL
            CType(ws.Cells(zeile, colDocUrl), Excel.Range).Value = docUrl
        Catch ex As Exception

        End Try


        indentlevel = hproj.hierarchy.getIndentLevel(cphase.nameID)

        For im = 1 To cphase.countMilestones
            zeile = zeile + 1
            cmilestone = cphase.getMilestone(im)
            startdate = cmilestone.getDate

            curName = cmilestone.name

            indentlevel = hproj.hierarchy.getIndentLevel(cmilestone.nameID)
            CType(ws.Cells(zeile, colName), Excel.Range).Value = erzeugeIndent(indentlevel) & curName

            If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                CType(ws.Cells(zeile, colStart), Excel.Range).Value = ""
                CType(ws.Cells(zeile, colEnde), Excel.Range).Value = startdate.ToShortDateString
            Else
                CType(ws.Cells(zeile, colStart), Excel.Range).Value = "Fehler !"
                CType(ws.Cells(zeile, colEnde), Excel.Range).Value = "Fehler !"
            End If

            ' jetzt Vorgangsklasse und Abbrev schreiben, falls vorhanden 
            Dim tmpAbbrev As String = MilestoneDefinitions.getAbbrev(curName)
            Dim tmpAppearance As String = MilestoneDefinitions.getAppearance(curName)

            CType(ws.Cells(zeile, colAbbrev), Excel.Range).Value = tmpAbbrev
            CType(ws.Cells(zeile, colAppearance), Excel.Range).Value = tmpAppearance

            ' jetzt Responsible, percentDone, TrafficLight und Explanation schreiben, falls vorhanden
            ' ur 18.12.17, den Wert für verantwortlich mitaufnehmen ...
            Dim tmpVerantwortlich As String = cmilestone.verantwortlich
            CType(ws.Cells(zeile, colRespons), Excel.Range).Value = tmpVerantwortlich

            ' percentDone eintragen 
            Dim tmpPercentDone As String = (cmilestone.percentDone * 100).ToString & " %"
            CType(ws.Cells(zeile, colPercent), Excel.Range).Value = tmpPercentDone

            ' TrafficLight eintragen 
            Dim tmpAmpel As Integer = cmilestone.ampelStatus
            CType(ws.Cells(zeile, colAmpel), Excel.Range).Value = tmpAmpel

            ' Explanation eintragen 
            Dim tmpExplan As String = cmilestone.ampelErlaeuterung
            CType(ws.Cells(zeile, colExplan), Excel.Range).Value = tmpExplan

            ' Document Link eintragen
            Try
                CType(ws.Cells(zeile, colDocUrl), Excel.Range).Value = cmilestone.DocURL
            Catch ex As Exception

            End Try


        Next



        For ip = 2 To hproj.AllPhases.Count
            zeile = zeile + 1
            cphase = hproj.getPhase(ip)
            startdate = cphase.getStartDate
            endDate = cphase.getEndDate
            curName = cphase.name

            indentlevel = hproj.hierarchy.getIndentLevel(cphase.nameID)
            CType(ws.Cells(zeile, spalte), Excel.Range).Value = erzeugeIndent(indentlevel) & curName

            If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                CType(ws.Cells(zeile, colStart), Excel.Range).Value = startdate.ToShortDateString
            Else
                CType(ws.Cells(zeile, colStart), Excel.Range).Value = "Fehler !"
            End If

            If DateDiff(DateInterval.Day, StartofCalendar, endDate) > 0 Then
                CType(ws.Cells(zeile, colEnde), Excel.Range).Value = endDate.ToShortDateString
            Else
                CType(ws.Cells(zeile, colEnde), Excel.Range).Value = "Fehler !"
            End If

            ' jetzt Vorgangsklasse und Abbrev schreiben, falls vorhanden 
            Dim tmpAbbrev As String = PhaseDefinitions.getAbbrev(curName)
            Dim tmpAppearance As String = PhaseDefinitions.getAppearance(curName)

            CType(ws.Cells(zeile, colAbbrev), Excel.Range).Value = tmpAbbrev
            CType(ws.Cells(zeile, colAppearance), Excel.Range).Value = tmpAppearance

            ' ur 18.12.17, den Wert für verantwortlich mitaufnehmen ...
            Dim tmpVerantwortlich As String = cphase.verantwortlich
            CType(ws.Cells(zeile, colRespons), Excel.Range).Value = tmpVerantwortlich

            ' percentDone eintragen 
            Dim tmpPercentDone As String = (cphase.percentDone * 100).ToString & " %"
            CType(ws.Cells(zeile, colPercent), Excel.Range).Value = tmpPercentDone

            ' ampel eintragen 
            Dim tmpAmpel As Integer = cphase.ampelStatus
            CType(ws.Cells(zeile, colAmpel), Excel.Range).Value = tmpAmpel

            ' ampel Erläuterung eintragen 
            Dim tmpExplan As String = cphase.ampelErlaeuterung
            CType(ws.Cells(zeile, colExplan), Excel.Range).Value = tmpExplan

            ' Document Link eintragen
            Try
                CType(ws.Cells(zeile, colDocUrl), Excel.Range).Value = cphase.DocURL
            Catch ex As Exception

            End Try

            For im = 1 To cphase.countMilestones
                zeile = zeile + 1
                cmilestone = cphase.getMilestone(im)
                startdate = cmilestone.getDate


                curName = cmilestone.name
                indentlevel = hproj.hierarchy.getIndentLevel(cmilestone.nameID)
                CType(ws.Cells(zeile, spalte), Excel.Range).Value = erzeugeIndent(indentlevel) & curName

                If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                    CType(ws.Cells(zeile, colStart), Excel.Range).Value = ""
                    CType(ws.Cells(zeile, colEnde), Excel.Range).Value = startdate.ToShortDateString
                Else
                    CType(ws.Cells(zeile, colStart), Excel.Range).Value = "Fehler !"
                    CType(ws.Cells(zeile, colEnde), Excel.Range).Value = "Fehler !"
                End If

                ' jetzt Vorgangsklasse und Abbrev schreiben, falls vorhanden 
                tmpAbbrev = MilestoneDefinitions.getAbbrev(curName)
                tmpAppearance = MilestoneDefinitions.getAppearance(curName)

                CType(ws.Cells(zeile, colAbbrev), Excel.Range).Value = tmpAbbrev
                CType(ws.Cells(zeile, colAppearance), Excel.Range).Value = tmpAppearance

                ' jetzt Responsible, percentDone, TrafficLight und Explanation schreiben, falls vorhanden
                ' ur 18.12.17, den Wert für verantwortlich mitaufnehmen ...
                tmpVerantwortlich = cmilestone.verantwortlich
                CType(ws.Cells(zeile, colRespons), Excel.Range).Value = tmpVerantwortlich

                ' percentDone eintragen 
                tmpPercentDone = (cmilestone.percentDone * 100).ToString & " %"
                CType(ws.Cells(zeile, colPercent), Excel.Range).Value = tmpPercentDone

                ' TrafficLight eintragen 
                tmpAmpel = cmilestone.ampelStatus
                CType(ws.Cells(zeile, colAmpel), Excel.Range).Value = tmpAmpel

                ' Explanation eintragen 
                tmpExplan = cmilestone.ampelErlaeuterung
                CType(ws.Cells(zeile, colExplan), Excel.Range).Value = tmpExplan

                ' Document Link eintragen
                Try
                    CType(ws.Cells(zeile, colDocUrl), Excel.Range).Value = cmilestone.DocURL
                Catch ex As Exception

                End Try

            Next

        Next

        ' jetzt muss um eine Zeile weitergeschaltet werden, damit immer auf eine freie Zeile geschrieben wird
        zeile = zeile + 1

    End Sub


    ''' <summary>
    ''' Einlesen eines RXF-Files (XML-Ausleitung von RPLAN) und dazu ein Protokoll in Tabellenblatt 'xmlfilename'protokoll in Datei Logfile
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="xmlfilename"></param>Name des RXF-Files
    ''' <param name="isVorlage"></param>Ist Vorlage, oder nicht
    ''' <remarks></remarks>
    Sub RXFImport(ByRef myCollection As Collection, ByVal xmlfilename As String,
                  ByVal isVorlage As Boolean, ByRef protokollliste As SortedList(Of Integer, clsProtokoll))
        ' akt. Name zum Zweck des Fehlersuchens
        Dim aktuellerName As String = ""

        'Variablen-Definitionen für Projectboard 

        Dim hproj As clsProjekt

        Dim vproj As clsProjektvorlage
        Dim vorlagenName As String = ""

        Dim ProjektdauerinDays As Integer = 0
        Dim cphase As clsPhase = Nothing

        Dim parentphase As clsPhase = Nothing
        Dim lastphase As clsPhase = Nothing

        Dim parentelemID As String = ""
        Dim lastelemID As String = ""

        Dim cBewertung As clsBewertung = Nothing

        Dim milestoneName As String = ""

        ' Ersetzen eines bestimmten Strings in der kompletten Datei 'xmlfilename'
        ' Zurückgeben des Namens der neuen Datei 'newXMLfilename'

        Dim newXMLfilename As String = replaceStringInFile(xmlfilename, "xsi:type=""subscribedtask""", "")



        ' XML-Datei Öffnen
        ' A FileStream is needed to read the XML document.
        Dim fs As New FileStream(newXMLfilename, FileMode.Open)

        ' Declare an object variable of the type to be deserialized.
        Dim Rplan As New rxf            ' Class rxf wird in clsRplanRXF.vb definiert

        Try

            ' Create an instance of the XmlSerializer class;
            ' specify the type of object to be deserialized.
            Dim deserializer As New XmlSerializer(GetType(rxf), "http://www.actano.de/2007/rxf")


            ' If the XML document has been altered with unknown
            ' nodes or attributes, handle them with the
            ' UnknownNode and UnknownAttribute events.

            ' Änderung tk: die beiden deserializer Kommandos müssen wieder aktiviert werden !
            'Call MsgBox("hier wurde RXF Import massgeblich verändert !!" & vbLf & _
            '             " lief bei Windows 10/Excel 2016 nicht")
            AddHandler deserializer.UnknownNode, AddressOf deserializer_UnknownNode
            AddHandler deserializer.UnknownAttribute, AddressOf deserializer_UnknownAttribute


            ' Einlesen des kompletten XML-Dokument im die Klasse rxf
            ' Use the Deserialize method to restore the object's state with
            ' data from the XML document. 
            Rplan = CType(deserializer.Deserialize(fs), rxf)

            ' Tabellenblatt "xmlfilename" im logfile.xlsx erzeugen fürs Protokoll (xmlfilename ohne ".rxf" Extension)

            Dim tstr As String() = Split(xmlfilename, "\", -1)
            Dim hstr As String = tstr(tstr.Length - 1)
            Dim quelle As String = hstr
            tstr = Split(hstr, ".", 2)

            Dim tabblattname As String = tstr(0)
            Dim wslogbuch As Excel.Worksheet = Nothing


            Dim protokollLine As New clsProtokoll("", quelle)
            Dim zeile As Integer = 3


            ' Projekt suchen; VISBO Projekt suchen unter der RPLANTasks mit gegebenen MainProject
            For i = 0 To Rplan.task.Length - 1

                If Not IsNothing(Rplan.task(i).mainProject) Then
                    ' akt. Task ist Projekt 

                    aktuellerName = Rplan.task(i).name

                    Dim aktTask_i As rxfTask = Rplan.task(i)
                    hproj = New clsProjekt

                    hproj.name = aktTask_i.name
                    hproj.VorlagenName = ""
                    hproj.leadPerson = aktTask_i.owner
                    hproj.startDate = aktTask_i.actualDate.start.Value
                    ProjektdauerinDays = calcDauerIndays(aktTask_i.actualDate.start.Value, aktTask_i.actualDate.finish.Value)

                    ' Protokollzeile bestücken
                    protokollLine = New clsProtokoll(hproj.name, quelle)


                    ' ProjektPhase erzeugen
                    cphase = New clsPhase(parent:=hproj)

                    With cphase
                        .nameID = rootPhaseName

                        Dim Duration As Long = calcDauerIndays(aktTask_i.actualDate.start.Value, aktTask_i.actualDate.finish.Value)
                        Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate, aktTask_i.actualDate.start.Value)

                        ' für die rootPhase muss gelten: offset = startoffset = 0 und duration = ProjektdauerIndays

                        Dim startOffset As Integer = 0
                        .changeStartandDauer(startOffset, Duration)
                        Dim phaseStartdate As Date = .getStartDate
                        Dim phaseEnddate As Date = .getEndDate

                    End With

                    ' ProjektPhase wird hinzugefügt
                    Dim hrchynode As New clsHierarchyNode
                    hrchynode.elemName = cphase.name
                    hrchynode.parentNodeKey = ""
                    hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)
                    parentphase = cphase
                    parentelemID = cphase.nameID
                    lastphase = cphase
                    lastelemID = cphase.nameID

                    ' Alle Tasks zu diesem Projekt mit deren Kinder und KindesKinder in hproj eintragen
                    Try
                        Call findAllTasksandInsert(aktTask_i, parentelemID, hproj, Rplan, protokollLine, zeile, protokollliste)
                    Catch ex As Exception
                        Dim a As Integer = 0
                    End Try


                    '
                    '' '' Bestimmung der BMW-Vorlage des jeweiligen Projektes
                    '' '' muss noch genauer herausgefunden werden, welche Vorlage für das jeweilige Projekt verwendet werden muss
                    '
                    vorlagenName = findBMWVorlagenName(hproj)
                    '
                    ''''    Ende Bestimmung der BMW-Vorlage zu diesem Projekt
                    '


                    If Projektvorlagen.Contains(vorlagenName) Then
                        vproj = Projektvorlagen.getProject(vorlagenName)

                        hproj.VorlagenName = vorlagenName
                        hproj.farbe = vproj.farbe
                        hproj.Schrift = vproj.Schrift
                        hproj.Schriftfarbe = vproj.Schriftfarbe
                        hproj.earliestStart = vproj.earliestStart
                        hproj.latestStart = vproj.latestStart

                        'ElseIf Projektvorlagen.Contains("unknown") Then
                        '    vproj = Projektvorlagen.getProject("unknown")
                    Else
                        'Throw New Exception("es gibt weder die Vorlage 'unknown' noch die Vorlage " & vorlagenName)
                        hproj.VorlagenName = ""
                        hproj.farbe = awinSettings.AmpelNichtBewertet
                        hproj.Schrift = Projektvorlagen.getProject(0).Schrift
                        hproj.Schriftfarbe = RGB(10, 10, 10)
                        hproj.earliestStart = 0
                        hproj.latestStart = 0

                    End If

                    Dim msPHdefcount As Integer = missingPhaseDefinitions.Count
                    Dim msMSdefcount As Integer = missingMilestoneDefinitions.Count

                    ' jetzt muss das Projekt eingetragen werden in die Listen Importierte Projekte und myCollection
                    ' Änderung tk: falls es das Projekt unter diesem Namen bereits gibt, wird eine Variante angelegt ... 
                    Dim lfdNr As Integer = 2
                    Do While ImportProjekte.Containskey(calcProjektKey(hproj))
                        hproj.variantName = lfdNr.ToString
                        lfdNr = lfdNr + 1
                    Loop

                    Dim hlptxt As String = ""
                    If lfdNr - 2 > 0 Then
                        If lfdNr - 2 = 1 Then
                            hlptxt = "es wurde eine Variante angelegt"
                        Else
                            hlptxt = "es wurden " & lfdNr - 2 & " Varianten angelegt."
                        End If
                        Call MsgBox("Projekt " & hproj.name & " kommt mehrmals vor! " & vbLf & hlptxt)
                    End If


                    ' jetzt ist sichergestellt, dass calcProjektKey nicht mehr vorkommt 
                    ImportProjekte.Add(hproj, False)
                    myCollection.Add(calcProjektKey(hproj))


                Else
                    ' aktuelle Task ist kein Projekt
                End If
            Next i

            '' '' Protokolldatei sichern
            ' ''Call writeProtokoll(protokollliste, tabblattname)


            ' RXF-Datei (entspricht XML-Datei) Schliessen
            fs.Close()

        Catch ex As Exception
            Call logfileSchreiben(ex.Message & vbLf & "Fehler bei Name " & CStr(aktuellerName), aktuellerName, anzFehler)
            Throw New ArgumentException("Fehler bei Name " & CStr(aktuellerName))

            ' RXF-Datei (entspricht XML-Datei) Schliessen
            fs.Close()
        End Try


    End Sub


    ''' <summary>
    ''' sucht zu der rxfTask 'task' alle Kinder und KindesKinder  und trägt diese in das Projekt 'hproj' ein 
    ''' dazu wird diese Routine rekursiv aufgerufen
    ''' </summary>
    ''' <param name="task"></param>rxfTask 'task', die Parent aller gesuchten Tasks ist
    ''' <param name="parentelemID"></param>Parent dieser rxfTask 'task'
    ''' <param name="hproj"></param>aktuelles aufzubauendes Projekt
    ''' <param name="RPLAN"></param>Komplette eingelesene rxf-Struktur 
    ''' <remarks></remarks>
    Private Sub findAllTasksandInsert(ByVal task As rxfTask, ByVal parentelemID As String, ByRef hproj As clsProjekt, ByVal RPLAN As rxf, ByRef prtLine As clsProtokoll, ByRef zeile As Integer, ByRef prtliste As SortedList(Of Integer, clsProtokoll))


        Dim cphase As clsPhase = Nothing
        Dim cmilestone As clsMeilenstein
        Dim parentphase As clsPhase = hproj.getPhaseByID(parentelemID)
        Dim lastphase As clsPhase = Nothing
        Dim lastelemID As String = ""

        Dim phaseNameID As String = ""
        Dim cBewertung As clsBewertung = Nothing

        Dim origMSname As String = ""
        Dim milestonedate As Date
        Dim isNotDuplikate As Boolean = True
        Dim isUnkownName As Boolean = False


        ' weitere Tasks finden, die zu diesem Projekt (mit ID=aktTask.id) gehören, d.h. ID muss als Parent auftreten
        For j = 0 To RPLAN.task.Length - 1

            isUnkownName = False            ' hier ist noch unklar, ob kown oder unkown Task

            If RPLAN.task(j).parent = task.id Then

                ' Änderung tk am 10.2.16 - wenn ein Fehler bei einem einzelnen Element auftritt, soll das nicht dazu führen, 
                ' dass der Import aller anderen abgebrochen wird ... eine entsprechende Fehlermeldung soll ins Protokoll kommen 
                ' alle anderen Elemente sollen importiert werden 
                Try

                    Dim aktTask_j As rxfTask = RPLAN.task(j)
                    Dim isMilestone As Boolean

                    Dim isKnownMsName As Boolean = MilestoneDefinitions.Contains(aktTask_j.name) Or
                                                missingMilestoneDefinitions.Contains(aktTask_j.name)

                    Dim isKnownPhName As Boolean = PhaseDefinitions.Contains(aktTask_j.name) Or
                                                missingPhaseDefinitions.Contains(aktTask_j.name)

                    Dim taskdauerinDays As Long = calcDauerIndays(aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value)
                    ' Herausfinden, ob aktTask_j Phase oder Meilenstein ist 

                    If taskdauerinDays > 1 Then
                        isMilestone = False

                        If aktTask_j.taskType.type = "MILESTONE" Then
                            Call logfileSchreiben("Korrektur, RXFImport: Phasen-Element mit verschiedenen Start- und Ende-Daten war als Meilenstein deklariert:",
                                                        aktTask_j.name & ": " & aktTask_j.actualDate.start.Value.ToShortDateString & " versus " &
                                                        aktTask_j.actualDate.finish.Value.ToShortDateString & vbLf &
                                                        "Projekt: " & hproj.name,
                                                        anzFehler)
                        End If

                    ElseIf aktTask_j.taskType.type = "MILESTONE" Then
                        isMilestone = True

                    ElseIf isKnownMsName And Not isKnownPhName Then
                        isMilestone = True
                        If aktTask_j.taskType.type <> "MILESTONE" Then
                            Call logfileSchreiben("Korrektur, RXFImport: bekanntes Meilenstein-Element  mit falscher Typ-Zuordnung:",
                                                        aktTask_j.name & " mit Typ " & aktTask_j.taskType.type & vbLf &
                                                        "Projekt: " & hproj.name,
                                                        anzFehler)
                        End If

                    ElseIf isKnownPhName And Not isKnownMsName Then
                        isMilestone = False


                    Else
                        isMilestone = True
                    End If

                    If Not isMilestone Then

                        ''''''  ist PHASE

                        If aktTask_j.name = "Projektphasen" Then

                            For i = 0 To aktTask_j.customvalue.Length - 1
                                If aktTask_j.customvalue(i).name = "UsA_SERVICE_SPALTE_A" Then
                                    hproj.businessUnit = aktTask_j.customvalue(i).Value
                                End If

                                If aktTask_j.customvalue(i).name = "UsA_SERVICE_SPALTE_B" Then
                                    hproj.VorlagenName = aktTask_j.customvalue(i).Value
                                End If
                            Next i

                        End If

                        ' überprüfen, ob die Phase evt. ignoriert werden soll (wird im  CustomizationFile in Tabelle Phase-Mappings definiert)
                        If Not phaseMappings.tobeIgnored(aktTask_j.name) Then
                            Dim mappedPhasename As String = ""

                            prtLine.planelement = aktTask_j.name
                            prtLine.hgColor = awinSettings.AmpelNichtBewertet

                            If PhaseDefinitions.Contains(aktTask_j.name) Then

                                mappedPhasename = aktTask_j.name
                                prtLine.hgColor = awinSettings.AmpelGruen

                            Else
                                ' aktTask_j.name existiert nicht in den PhaseDefinitions

                                'wenn der PhasenName gemappt werden kann und dieser dann in phasedefinitions enthalten ist, so wird phasename ersetzt
                                mappedPhasename = phaseMappings.mapToStdName(elemNameOfElemID(parentelemID), aktTask_j.name)

                                If PhaseDefinitions.Contains(mappedPhasename) Then
                                    ' neuer aktueller Name der Task
                                    prtLine.hgColor = awinSettings.AmpelGelb

                                Else
                                    ' PhasenName ist nicht bekannt
                                    isUnkownName = True


                                    Dim newPhaseDef As New clsPhasenDefinition


                                    ' Änderung tk 6.12.15: das muss auf den mappedPhasename gesetzt werdne, da sonst Eltern-Ersetzungen, die noch nicht 
                                    ' in der phasedefinitions sind , nicht in der Liste der unbekannten aufgenommen werden ... 
                                    'newPhaseDef.name = aktTask_j.name
                                    'mappedPhasename = aktTask_j.name

                                    newPhaseDef.name = mappedPhasename
                                    newPhaseDef.shortName = aktTask_j.remark

                                    newPhaseDef.darstellungsKlasse = mapToAppearance(aktTask_j.taskType.Value, False)
                                    newPhaseDef.UID = PhaseDefinitions.Count + 1
                                    ' muss in missingPhaseDefinitions noch eingetragen werden
                                    ' in add wird abgefragt, ob der Name schon existiert, wenn ja, wird nix gemacht 
                                    missingPhaseDefinitions.Add(newPhaseDef)

                                    ' Änderung tk: wird auskommentiert, das steht ja im Protokoll
                                    'Call logfileSchreiben(("Achtung, RXFImport: Phase '" & aktTask_j.name & "' existiert im CustomizationFile nicht!"), hproj.name, anzFehler)

                                End If
                            End If

                            ' Phase nur aufnehmen in das aktuelle Projekt, wenn 
                            ' awinSettings.importUnkownNames=true ist oder auch isUnkownName = false

                            If Not isUnkownName Or awinSettings.importUnknownNames Then

                                Dim phaseStartdate As Date
                                Dim phaseEnddate As Date
                                cphase = New clsPhase(hproj)

                                With cphase

                                    Dim Duration As Integer = calcDauerIndays(aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value)
                                    Dim offset As Integer = DateDiff(DateInterval.Day, hproj.startDate, aktTask_j.actualDate.start.Value)

                                    .changeStartandDauer(offset, Duration)
                                    phaseStartdate = .getStartDate
                                    phaseEnddate = .getEndDate

                                    isNotDuplikate = True
                                    ' sollen Duplikate eliminiert werden ?
                                    If awinSettings.eliminateDuplicates And hproj.hierarchy.containsKey(calcHryElemKey(mappedPhasename, False)) Then
                                        ' nur dann kann es Duplikate geben 
                                        If hproj.isCloneToParent(mappedPhasename, parentphase.nameID, phaseStartdate, phaseEnddate, 0.97) Then
                                            isNotDuplikate = False
                                            prtLine.planelement = aktTask_j.name
                                            prtLine.hgColor = awinSettings.AmpelRot
                                            prtLine.grund = "Phase wurde eliminiert: Duplikat zur Parent-Phase"
                                            'Call logfileSchreiben("Fehler in RXFImport: " & mappedPhasename & " ist Duplikat zu Parent " & parentphase.name & " und wird ignoriert ", hproj.name, anzFehler)

                                        Else
                                            Dim duplicateSiblingID As String = hproj.getDuplicatePhaseSiblingID(mappedPhasename, parentphase.nameID,
                                                                                                                phaseStartdate, phaseEnddate, 0.97)

                                            If duplicateSiblingID = "" Then
                                                isNotDuplikate = True
                                            Else
                                                isNotDuplikate = False
                                                prtLine.planelement = aktTask_j.name
                                                prtLine.hgColor = awinSettings.AmpelRot
                                                prtLine.grund = "Phase wurde eliminiert: Duplikat zur Geschwister-Phase"
                                                'Call logfileSchreiben(" Fehler in RXFImport: " & mappedPhasename & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) & _
                                                '" und wird ignoriert ", hproj.name, anzFehler)
                                            End If
                                        End If

                                    End If

                                End With

                                If isNotDuplikate Then

                                    ' hier muss für gleiche PhasenNamen als Geschwister noch eine lfdNummer angehängt werden
                                    ' es muss überprüft werden, ob es Geschwister mit gleichem Namen gibt:
                                    ' wenn ja, wird an den mappedPhaseName eine LFdNr. ergänzt,bis der Name innerhalb der Geschwistergruppe eindeutig ist.

                                    If awinSettings.createUniqueSiblingNames Then
                                        mappedPhasename = hproj.hierarchy.findUniqueGeschwisterName(parentelemID, mappedPhasename, False)
                                    End If

                                    cphase.nameID = hproj.hierarchy.findUniqueElemKey(mappedPhasename, False)

                                    ' Phase wird ins Projekt mitaufgenommen

                                    Dim phrchynode As New clsHierarchyNode
                                    phrchynode.elemName = cphase.name
                                    phrchynode.parentNodeKey = parentelemID

                                    hproj.AddPhase(cphase, origName:=aktTask_j.name, parentID:=phrchynode.parentNodeKey)
                                    phrchynode.indexOfElem = hproj.AllPhases.Count

                                    ' merken von letzem Element (Knoten,Phase,Meilenstein)
                                    'lasthrchynode = phrchynode
                                    lastelemID = cphase.nameID
                                    lastphase = cphase

                                    prtLine.hierarchie = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                                    prtLine.PThierarchie = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                                    prtLine.planelement = aktTask_j.name
                                    prtLine.abkürzung = PhaseDefinitions.getAbbrev(cphase.name)
                                    prtLine.planeleÜbern = cphase.name

                                    prtLine.klasse = aktTask_j.taskType.Value
                                    prtLine.PTklasse = mapToAppearance(aktTask_j.taskType.Value, False)

                                    prtliste.Add(zeile, prtLine)
                                    zeile = zeile + 1
                                    'prtLine.writeLog(zeile)

                                    Dim quelle As String = prtLine.quelle

                                    prtLine = New clsProtokoll(hproj.name, quelle)
                                    prtLine.actDate = ""


                                    Call findAllTasksandInsert(aktTask_j, lastelemID, hproj, RPLAN, prtLine, zeile, prtliste)

                                Else
                                    prtliste.Add(zeile, prtLine)
                                    zeile = zeile + 1

                                    Dim quelle As String = prtLine.quelle
                                    prtLine = New clsProtokoll(hproj.name, quelle)
                                    prtLine.actDate = ""
                                End If

                            Else
                                prtLine.planelement = aktTask_j.name
                                prtLine.hgColor = awinSettings.AmpelRot
                                prtLine.grund = "Phase wurde ignoriert: unbekannter Bezeichner"

                                prtliste.Add(zeile, prtLine)
                                zeile = zeile + 1

                                Dim quelle As String = prtLine.quelle
                                prtLine = New clsProtokoll(hproj.name, quelle)
                                prtLine.actDate = ""
                            End If

                        Else
                            prtLine.planelement = aktTask_j.name
                            prtLine.hgColor = awinSettings.AmpelRot
                            prtLine.grund = "Phase wurde ignoriert: gemäß Eintrag TOBEIGNORED im Wörterbuch"

                            prtliste.Add(zeile, prtLine) ' Protokollzeile in Liste eintragen
                            zeile = zeile + 1

                            Dim quelle As String = prtLine.quelle
                            prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                            prtLine.actDate = ""

                        End If       'Ende of tobeignored phase


                    Else
                        ' ist MEILENSTEIN

                        Dim mappedMSname As String = ""

                        If Not milestoneMappings.tobeIgnored(aktTask_j.name) Then

                            If MilestoneDefinitions.Contains(aktTask_j.name) Then

                                mappedMSname = aktTask_j.name
                                prtLine.hgColor = awinSettings.AmpelGruen

                            Else
                                'wenn der MeilensteinName gemappt werden kann und dieser dann in milestonedefinitions enthalten ist, so wird Meilensteinname ersetzt
                                mappedMSname = milestoneMappings.mapToStdName(elemNameOfElemID(parentelemID), aktTask_j.name)
                                If MilestoneDefinitions.Contains(mappedMSname) Then

                                    prtLine.hgColor = awinSettings.AmpelGelb
                                Else

                                    isUnkownName = True

                                    Dim msDef As New clsMeilensteinDefinition


                                    ' Änderung tk 6.12.15: das muss auf den mappedMSNamen gesetzt werdne, da sonst Eltern-Ersetzungen, die noch nicht 
                                    ' in der milestonedefinitions sind , nicht in der Liste der unbekannten aufgenommen werden ... 
                                    'msDef.name = aktTask_j.name
                                    'mappedMSname = aktTask_j.name

                                    msDef.name = mappedMSname
                                    msDef.schwellWert = 0
                                    msDef.belongsTo = parentphase.name
                                    msDef.shortName = aktTask_j.remark

                                    msDef.darstellungsKlasse = mapToAppearance(aktTask_j.taskType.Value, True)
                                    msDef.UID = MilestoneDefinitions.Count + 1

                                    Try
                                        missingMilestoneDefinitions.Add(msDef)

                                        'Call logfileSchreiben(("Achtung, RXFImport: Meilenstein '" & aktTask_j.name & "' existiert im CustomizationFile nicht!"), hproj.name, anzFehler)

                                    Catch ex As Exception
                                    End Try

                                End If
                            End If

                            ' Meilenstein wird nur in das aktuelle Projekt aufgenommen, wenn awinSettings.importUnkownNames = true 
                            ' und der Name bekannt ist (isUnkownName = false)

                            If Not isUnkownName Or awinSettings.importUnknownNames Then

                                cmilestone = New clsMeilenstein(parent:=parentphase)
                                cBewertung = New clsBewertung

                                origMSname = aktTask_j.name

                                If DateDiff(DateInterval.Day, aktTask_j.actualDate.start.Value, aktTask_j.actualDate.finish.Value) = 0 Then
                                    milestonedate = aktTask_j.actualDate.start.Value
                                Else
                                    Throw New Exception("Fehler, RXFImport: Der Meilenstein hat verschiedene Start- und End-Daten:" & vbLf &
                                                        aktTask_j.actualDate.start.Value.ToShortDateString & " versus " &
                                                        aktTask_j.actualDate.finish.Value.ToShortDateString & vbLf &
                                                        "Projekt: " & hproj.name)
                                End If



                                ' wenn der freefloat nicht zugelassen ist und der Meilenstein ausserhalb der Phasen-Grenzen liegt 
                                ' muss abgebrochen werden 

                                If Not awinSettings.milestoneFreeFloat And
                                    (DateDiff(DateInterval.Day, parentphase.getStartDate, milestonedate) < 0 Or
                                     DateDiff(DateInterval.Day, parentphase.getEndDate, milestonedate) > 0) Then

                                    'Call logfileSchreiben(("Fehler, RXFImport: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf & _
                                    '                    origMSname & " nicht innerhalb " & parentphase.name & vbLf & _
                                    '                         "Korrigieren Sie bitte diese Inkonsistenz in der Datei '"), hproj.name, anzFehler)
                                    Throw New Exception("Fehler, RXFImport: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf &
                                                        origMSname & " nicht innerhalb " & parentphase.name & vbLf &
                                                             "Korrigieren Sie bitte diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
                                End If

                                Dim resultVerantwortlich As String = aktTask_j.owner
                                Dim bewertungsAmpel As Integer = 0
                                Dim explanation As String = aktTask_j.note

                                ' Ergänzung tk 2.11 deliverables ergänzt 
                                Dim deliverables As String = ""

                                If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                                    ' es gibt keine Bewertung
                                    bewertungsAmpel = 0
                                End If
                                ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                With cBewertung
                                    '.bewerterName = resultVerantwortlich
                                    .colorIndex = bewertungsAmpel
                                    .datum = Date.Now
                                    .description = explanation
                                    ' deliverables sind jetzt Bestandteil von clsMeilenstein (List (of String))  
                                    '.deliverables = deliverables
                                End With

                                isNotDuplikate = True
                                If awinSettings.eliminateDuplicates And hproj.hierarchy.containsKey(calcHryElemKey(mappedMSname, True)) Then
                                    ' nur dann kann es Duplikate geben 
                                    Dim duplicateSiblingID As String = hproj.getDuplicateMsSiblingID(mappedMSname, parentphase.nameID,
                                                                                                         milestonedate, 0)

                                    If duplicateSiblingID = "" Then
                                        isNotDuplikate = True
                                    Else
                                        isNotDuplikate = False
                                        prtLine.planelement = aktTask_j.name
                                        prtLine.hgColor = awinSettings.AmpelRot
                                        prtLine.grund = "Meilenstein wurde eliminiert: Duplikat zur Geschwister-Phase"
                                        'Call logfileSchreiben("Fehler, RXFImport:" & mappedMSname & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) & _
                                        '" und wird ignoriert ", hproj.name, anzFehler)
                                    End If

                                End If

                                If isNotDuplikate Then

                                    With cmilestone
                                        .setDate = milestonedate
                                        '.verantwortlich = resultVerantwortlich

                                        ' hier muss für gleiche PhasenNamen als Geschwister noch eine lfdNummer angehängt werden
                                        ' es muss überprüft werden, ob es Geschwister mit gleichem Namen gibt:
                                        ' wenn ja, wird an den mappedPhaseName eine LFdNr. ergänzt,bis der Name innerhalb der Geschwistergruppe eindeutig ist.

                                        If awinSettings.createUniqueSiblingNames Then
                                            mappedMSname = hproj.hierarchy.findUniqueGeschwisterName(parentelemID, mappedMSname, True)
                                        End If

                                        .nameID = hproj.hierarchy.findUniqueElemKey(mappedMSname, True)
                                        If Not cBewertung Is Nothing Then
                                            .addBewertung(cBewertung)
                                        End If
                                    End With

                                    With parentphase
                                        .addMilestone(cmilestone, origName:=origMSname)
                                    End With

                                    prtLine.hierarchie = hproj.hierarchy.getBreadCrumb(cmilestone.nameID)
                                    prtLine.PThierarchie = hproj.hierarchy.getBreadCrumb(cmilestone.nameID)
                                    prtLine.planelement = aktTask_j.name
                                    prtLine.abkürzung = MilestoneDefinitions.getAbbrev(cmilestone.name)
                                    prtLine.planeleÜbern = cmilestone.name

                                    prtLine.klasse = aktTask_j.taskType.Value
                                    prtLine.PTklasse = mapToAppearance(aktTask_j.taskType.Value, True)

                                    prtliste.Add(zeile, prtLine)
                                    zeile = zeile + 1

                                    Dim quelle As String = prtLine.quelle
                                    prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                                    prtLine.actDate = ""

                                Else
                                    prtliste.Add(zeile, prtLine) ' Protokollzeile in Liste eintragen
                                    zeile = zeile + 1

                                    Dim quelle As String = prtLine.quelle
                                    prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                                    prtLine.actDate = ""

                                End If

                            Else
                                prtLine.planelement = aktTask_j.name
                                prtLine.hgColor = awinSettings.AmpelRot
                                prtLine.grund = "Meilenstein wurde ignoriert: unbekannter Bezeichner"

                                prtliste.Add(zeile, prtLine) ' Protokollzeile in Liste eintragen
                                zeile = zeile + 1

                                Dim quelle As String = prtLine.quelle
                                prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                                prtLine.actDate = ""
                            End If

                        Else
                            prtLine.planelement = aktTask_j.name
                            prtLine.hgColor = awinSettings.AmpelRot
                            prtLine.grund = "Meilenstein wurde ignoriert gemäß Eintrag im Wörterbuch"

                            prtliste.Add(zeile, prtLine) ' Protokollzeile in Liste eintragen
                            zeile = zeile + 1

                            Dim quelle As String = prtLine.quelle
                            prtLine = New clsProtokoll(hproj.name, quelle) ' neue Protokollzeile
                            prtLine.actDate = ""

                        End If     ' Ende: Meilenstein soll ignoriert werden



                    End If      '  Ende: ist MEILENSTEIN

                Catch ex As Exception

                    Call logfileSchreiben(ex.Message, hproj.name, anzFehler)


                End Try


            End If

        Next j    ' Ende Schleife über alle Tasks
    End Sub

    ''' <summary>
    ''' nach BMW-Vorgaben:
    ''' bestimmt aus dem übergebenen VorlagenNamen ( =  der CustomValue "UsA_SERVICE_SPALTE_B" aus Phase "Projektphasen" ) 
    ''' den tatsächlichen VorlagenNamen des Projekts 
    '''     ''' </summary>
    ''' <param name="hproj"></param>aktuelles zu lesendes Projekt
    ''' <returns></returns>fertig zusammengesetzter VorlagenName des Projekts (gemäß BMW vorschriften
    ''' <remarks></remarks>
    Private Function findBMWVorlagenName(ByVal hproj As clsProjekt) As String


        Dim vorNam1 = "rel 4"
        Dim typkennung As String = hproj.VorlagenName      ' hier ist aber nur enthalten, eA, wA, E  usw.
        Dim anlaufkennung As String = "03"
        Dim firstMS As Integer = hproj.hierarchy.getIndexOf1stMilestone


        Dim hrchyhproj As clsHierarchy = hproj.hierarchy
        For phi = 1 To firstMS - 1
            Dim phID As String = hrchyhproj.getIDAtIndex(phi)
            Dim phName As String = elemNameOfElemID(phID)
            If phName.Contains("I-Stufen") Then
                Dim pharray() As String = Split(phName, " ", 5)
                vorNam1 = "rel 5"
            End If
        Next phi
        For msi = firstMS To hrchyhproj.count - 1
            Dim msID As String = hrchyhproj.getIDAtIndex(msi)
            Dim msName As String = elemNameOfElemID(msID)
            If msName.Contains("SOP") Then
                Dim msarray() As String = Split(msName, " ", 5)
                Try
                    Dim sopdate As Date = hproj.getMilestoneDate(msID)

                    If DateDiff(DateInterval.Month, StartofCalendar, sopdate) > 0 Then
                        Dim sopMonth As Integer = sopdate.Month
                        If sopMonth >= 3 And sopMonth <= 6 Then
                            anlaufkennung = "03"
                        ElseIf sopMonth >= 7 And sopMonth <= 10 Then
                            anlaufkennung = "07"
                        Else
                            anlaufkennung = "11"
                        End If
                    Else
                        anlaufkennung = sopdate.Month.ToString("D2")       ' Monat mindestens zweistellig angeben
                    End If

                Catch ex As Exception
                    anlaufkennung = "?"
                End Try

            End If
        Next
        Try
            If Not IsNothing(typkennung) Then

                If typkennung.Contains("SB") Then
                    typkennung = "SBWE"
                ElseIf typkennung.Contains("eA") Then
                    typkennung = "eA"
                ElseIf typkennung.Contains("wA") Then
                    typkennung = "wA"
                ElseIf typkennung.Contains("E") Then
                    typkennung = "E"
                Else
                    typkennung = "?"
                End If
            Else
                typkennung = "?"
            End If
        Catch ex As Exception
            typkennung = "?"
        End Try

        findBMWVorlagenName = vorNam1 & " " & typkennung & "-" & anlaufkennung

    End Function


    ''' <summary>
    ''' Behandelt den Fehler UnkonwnNode beim Einlesen eines XML-Files (oder RXF-Files)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub deserializer_UnknownNode(sender As Object, e As XmlNodeEventArgs)
        Call MsgBox(("XMLImport: Unknown Node:" & e.Name & ControlChars.Tab & e.Text))
    End Sub 'serializer_UnknownNode


    ''' <summary>
    ''' Behandelt den Fehler UnkonwnAttribute beim Einlesen eines XML-Files (oder RXF-Files)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub deserializer_UnknownAttribute(sender As Object, e As XmlAttributeEventArgs)
        Dim attr As System.Xml.XmlAttribute = e.Attr
        Call MsgBox(("XMLImport: Unknown attribute " & attr.Name & "='" & attr.Value & "'"))
    End Sub 'serializer_UnknownAttribute

    ''' <summary>
    ''' in der ganzen Datei sfilename wird der String searchstr durch replacestr ersetzt
    ''' </summary>
    ''' <param name="sfilename"></param>Name der Datei, in der die Ersetzung erfolgen soll
    ''' <param name="searchstr"></param>zu ersetzender String
    ''' <param name="replacestr"></param>neuer String
    ''' <returns></returns>Name der neuen Datei
    ''' <remarks></remarks>
    Private Function replaceStringInFile(ByVal sfilename As String, ByVal searchstr As String, ByVal replacestr As String) As String

        'Declare ALL of your variables :)
        Const ForReading = 1    '
        Dim fileToRead As String = sfilename  ' the path of the file to read
        Dim tstr() As String = Split(sfilename, ".", 2)
        Dim fileToWrite As String = tstr(0) & ".new"     ' the path of a new file
        Dim FSO As Object
        Dim readFile As Object  'the file you will READ
        Dim writeFile As Object 'the file you will CREATE
        Dim repLine As Object   'the array of lines you will WRITE
        Dim ln As Object
        Dim l As Long

        FSO = CreateObject("Scripting.FileSystemObject")
        readFile = FSO.OpenTextFile(fileToRead, ForReading, False)
        writeFile = FSO.CreateTextFile(fileToWrite, True, False)

        '# Read entire file into an array & close it
        repLine = Split(readFile.ReadAll, vbNewLine)
        readFile.Close()

        '# iterate the array and do the replacement line by line

        For Each ln In repLine
            ln = Replace(ln, searchstr, replacestr)
            repLine(l) = ln
            l = l + 1
        Next

        '# Write to the array items to the file
        writeFile.Write(Join(repLine, vbNewLine))
        writeFile.Close()

        '# clean up
        readFile = Nothing
        writeFile = Nothing
        FSO = Nothing
        replaceStringInFile = fileToWrite

    End Function

    ''' <summary>
    ''' importiert ein MS Project File 
    ''' </summary>
    ''' <param name="modus"></param>
    ''' <param name="filename"></param>
    ''' <param name="hproj"></param>
    ''' <param name="mapProj"></param>
    ''' <param name="importdate"></param>
    Sub awinImportMSProject(ByVal modus As String, ByVal filename As String, ByRef hproj As clsProjekt, ByRef mapProj As clsProjekt, ByRef importdate As Date)

        Dim mapStruktur As String = awinSettings.mappingVorlage

        Dim prj As MSProject.Application
        Dim msproj As MSProject.Project
        Dim i As Integer = 1
        Dim lastphase As clsPhase
        Dim lasthrchyNode As clsHierarchyNode
        Dim lastelemID As String = ""
        Dim lastlevel As Integer = 0
        Dim Xwerte() As Double
        Dim active_proj As String = ""      ' Name des aktuell aktiven Projektes

        ' hier wird eingetragen, welches vordefinierte Flag das customized Field VISBO usw. repräsentiert
        Dim visboflag As MSProject.PjField = Nothing
        Dim visbo_taskclass As MSProject.PjField = Nothing
        Dim visbo_abbrev As MSProject.PjField = Nothing
        Dim visbo_ampel As MSProject.PjField = Nothing
        Dim visbo_ampeltext As MSProject.PjField = Nothing
        Dim visbo_deliverables As MSProject.PjField = Nothing
        Dim visbo_responsible As MSProject.PjField = Nothing
        Dim visbo_percentDone As MSProject.PjField = Nothing
        Dim visbo_mapping As MSProject.PjField = Nothing

        ' Liste, die aufgebaut wird beim Einlesen der Tasks. Hier wird vermerkt, welche Task das Visbo-Flag mit YES und welche mit NO
        ' gesetzt hat d.h. berücksichtigt werden soll
        ' Diese Liste enthält keine Elemente, wenn das VISBO-Flag nicht definiert ist
        Dim visboFlagListe As New SortedList(Of String, Boolean)


        Dim outputCollection As New Collection
        Dim outputline As String = ""


        Try

            'On Error Resume Next
            Try
                prj = CType(GetObject(, "msproject.application"), MSProject.Application)
            Catch ex As Exception
                prj = CType(CreateObject("msproject.application"), MSProject.Application)

                If IsNothing(prj) Then
                    Call MsgBox("MSproject ist nicht installiert")
                    Exit Sub
                End If
            End Try

            If modus <> "BHTC" Then

                ' ''prj.FileOpen(Name:="\\KOYTEK-NAS\backup\Ute\VISBO\MS Project Beispiele\ute.mpp", _
                ' ''             ReadOnly:=True, FormatID:="MSProject.MPP")

                prj.FileOpen(Name:=filename,
                            ReadOnly:=True, FormatID:="MSProject.MPP")


            End If


            Dim anzProj As Integer = prj.Projects.Count

            If anzProj > 0 Then


                ' VISBO-Flag dient dazu, Tasks, die nicht benötigt werden in der MultiprojektPlanung nicht mit einzulesen
                ' in die Projekt-Tafel

                ' Ist dieses VISBO-Flag definiert?
                Dim pjFlag As String = ""

                Try
                    visboflag = CType(prj.FieldNameToFieldConstant("VISBO", MSProject.PjFieldType.pjTask), MSProject.PjField)
                    pjFlag = prj.FieldConstantToFieldName(visboflag)

                Catch ex As Exception
                    visboflag = 0
                End Try

                Try
                    visbo_taskclass = CType(prj.FieldNameToFieldConstant(awinSettings.visboTaskClass, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_taskclass = 0
                End Try
                Try
                    visbo_abbrev = CType(prj.FieldNameToFieldConstant(awinSettings.visboAbbreviation, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_abbrev = 0
                End Try
                Try
                    visbo_ampel = CType(prj.FieldNameToFieldConstant(awinSettings.visboAmpel, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_ampel = 0
                End Try
                Try
                    visbo_ampeltext = CType(prj.FieldNameToFieldConstant(awinSettings.visboAmpelText, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_ampeltext = 0
                End Try
                Try
                    visbo_deliverables = CType(prj.FieldNameToFieldConstant(awinSettings.visbodeliverables, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_deliverables = 0
                End Try
                Try
                    visbo_responsible = CType(prj.FieldNameToFieldConstant(awinSettings.visboresponsible, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_responsible = 0
                End Try
                Try
                    visbo_percentDone = CType(prj.FieldNameToFieldConstant(awinSettings.visbopercentDone, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_percentDone = 0
                End Try
                Try
                    visbo_mapping = CType(prj.FieldNameToFieldConstant(awinSettings.visboMapping, MSProject.PjFieldType.pjTask), MSProject.PjField)
                Catch ex As Exception
                    visbo_mapping = 0
                End Try

                If modus = "BHTC" Then
                    ' In Missing..Definitions sind noch die Definitionen des vorausgegangenen Projekts definiert.
                    ' Diese sollen nicht mehr aktiv sein.
                    missingPhaseDefinitions.Clear()
                    missingMilestoneDefinitions.Clear()
                    '' Einlesen des aktiven Projekts
                    msproj = prj.ActiveProject
                Else
                    '' Einlesen des zuletzt gelesenen Projekts
                    msproj = prj.Projects.Item(anzProj)

                End If

                ' '' '' Einlesen der diversen Projekte, die geladen wurden (gilt nur für BHTC), sonst immer nur das zuletzt geladene
                '' ''For proj_i = beginnProjekt To endeProjekt

                ' Herausfinden, welches Startdatum des Projektes das früheste ist, da sonst die RootPhase zu spät anfängt
                ' und manche Phasen dann einen negative startoffset bekommen
                Dim ProjectStartDate As Date

                ProjectStartDate = CDate(msproj.ProjectStart)

                If CDate(msproj.Start) < ProjectStartDate Then
                    ProjectStartDate = CDate(msproj.Start)
                End If

                If CDate(msproj.EarlyStart) < ProjectStartDate Then
                    ProjectStartDate = CDate(msproj.EarlyStart)
                End If


                hproj = New clsProjekt(CDate(ProjectStartDate).Date, CDate(ProjectStartDate).Date, CDate(ProjectStartDate).Date)


                hproj.Erloes = 0


                Dim ProjektdauerIndays As Integer = calcDauerIndays(hproj.startDate, CDate(msproj.Finish).Date)
                Dim startOffset As Long = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(0))

                ' Projektname ohne "."
                Dim hhstr() As String
                hhstr = Split(msproj.Name, ".", -1)
                ' alle evtl auftretenden #, (, ) werden ersetzt durch unkritische Zeichen ... 
                hproj.name = makeValidProjectName(hhstr(0))
                'hproj.idauer = DateDiff(DateInterval.Month, CType(msproj.DefaultFinishTime, Date), CType(msproj.DefaultStartTime, Date))

                '' '' merken für BHTC, da hier der Report für das aktive Projekt gemacht werden soll 
                ' ''If prj.ActiveProject.Name = msproj.Name Then
                ' ''    active_proj = hproj.name
                ' ''End If

                Dim anzSubprojects As Integer = msproj.Subprojects.Count

                hproj.description = msproj.ProjectNotes
                hproj.UID = msproj.UniqueID
                Dim hrsPerDay As Double = msproj.HoursPerDay

                Dim projUID As Object = msproj.DatabaseProjectUniqueID

                ' ------------------------------------------------------------------------------------------------------
                ' Erzeugen und eintragen der Projekt-Phase (= erste Phase mit Dauer des Projekts)
                ' ------------------------------------------------------------------------------------------------------
                Try
                    Dim cphase As New clsPhase(hproj)

                    ' ProjektPhase wird erzeugt
                    cphase = New clsPhase(parent:=hproj)
                    cphase.nameID = rootPhaseName

                    ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                    With cphase
                        .nameID = rootPhaseName
                        Dim cphaseStartOffset As Integer = 0
                        .changeStartandDauer(cphaseStartOffset, ProjektdauerIndays)
                    End With
                    ' rootPhaseName - Phase wird hinzugefügt
                    hproj.AddPhase(cphase)

                Catch ex1 As Exception
                    Throw New ArgumentException("Fehler in awinImportMSProject, Erzeugen ProjektPhase")
                End Try




                Dim anzTasks As Integer = msproj.Tasks.Count
                anzTasks = msproj.NumberOfTasks
                Dim projSumTask As MSProject.Task = msproj.ProjectSummaryTask


                Dim resPool As MSProject.Resources = msproj.Resources

                Dim res(resPool.Count) As Object
                For i = 1 To resPool.Count
                    res(i) = resPool.Item(i)
                Next


                For i = 1 To anzTasks

                    Dim msTask As MSProject.Task

                    Dim cphase As New clsPhase(parent:=hproj)


                    msTask = msproj.Tasks.Item(i)


                    ' hier: evt. Prüfung ob eine VISBO Projekt-Tafel relevante Task
                    ' oder: ob eine Task auf dem kritischen Pfad liegt

                    ' Wenn sumTask = true, dann ist die aktuelle Task eine Summary-Task
                    Dim sumTask As Boolean = CType(msTask.Summary, Boolean)

                    ' Herausfinden der Hierarchiestufe
                    Dim hstr() As String = Split(msTask.WBS, ".", -1)
                    Dim tasklevel As Integer = hstr.Count


                    ' hier muss der Uniquename(ID) erzeugt werden evt. aus PhaseDefinitions

                    If Not CType(msTask.Milestone, Boolean) _
                        Or
                        (CType(msTask.Milestone, Boolean) And CType(msTask.Summary, Boolean)) Then

                        ' Ergänzung tk für Demo BHTC 
                        ' falls Synonyme definiert sind, ersetzen durch Std-Name, sonst bleibt Name unverändert 
                        Dim origPhName As String = msTask.Name
                        msTask.Name = phaseMappings.mapToStdName("", msTask.Name)

                        '' '' Budgets Kosten und Work der SammelTasks aufsummieren
                        '' '' ur: 25.08.2017: Testweise
                        ' ''Dim co As Double = 0
                        ' ''Dim wo As Double = 0
                        ' ''If CType(msTask.Summary, Boolean) Then

                        ' ''    Dim hstrco() As String = Split(msTask.BudgetCost, msproj.CurrencySymbol)
                        ' ''    If hstrco.Length > 1 Then
                        ' ''        'Dim co As Double = Val(msTask.BudgetCost)
                        ' ''        co = Val(hstrco(1))
                        ' ''    End If

                        ' ''    Dim hstrwo() As String = Split(msTask.BudgetWork, msproj.CurrencySymbol)
                        ' ''    If hstrwo.Length > 1 Then
                        ' ''        wo = Val(hstrwo(1))
                        ' ''        'Dim wo As Double = Val(msTask.BudgetWork)
                        ' ''    End If

                        ' ''    hproj.Erloes = hproj.Erloes + co + wo

                        ' ''End If

                        ' nachsehen, ob msTask.Name in PhaseDefinitions definiert ist
                        If Not PhaseDefinitions.Contains(msTask.Name) Then
                            Dim newPhaseDef As New clsPhasenDefinition
                            newPhaseDef.name = msTask.Name
                            ' Abbreviation, falls Customfield visbo_abbrev definiert ist
                            If visbo_abbrev <> 0 Then          ' VISBO-Abbrev ist definiert
                                newPhaseDef.shortName = msTask.GetField(visbo_abbrev)
                            Else
                                newPhaseDef.shortName = msTask.Name
                            End If
                            ' Task Class, falls Customfield visbo_taskclass definiert ist
                            If visbo_taskclass <> 0 Then          ' VISBO-TaskClass ist definiert
                                newPhaseDef.darstellungsKlasse = msTask.GetField(visbo_taskclass)
                            Else
                                newPhaseDef.darstellungsKlasse = ""
                            End If
                            cphase.appearance = newPhaseDef.darstellungsKlasse

                            newPhaseDef.UID = PhaseDefinitions.Count + 1
                            'PhaseDefinitions.Add(newPhaseDef)
                            missingPhaseDefinitions.Add(newPhaseDef)
                        Else
                            cphase.appearance = PhaseDefinitions.getAppearance(msTask.Name)
                        End If

                        With cphase

                            Dim phBewertung As New clsBewertung
                            If Not istElemID(msTask.Name) Then

                                .nameID = hproj.hierarchy.findUniqueElemKey(msTask.Name, False)
                            End If

                            If visboflag <> 0 Then          ' VISBO-Flag ist definiert

                                Dim hflag As Boolean = readCustomflag(msTask, visboflag)

                                ' Liste, ob Task in Projekt für die Projekt-Tafel aufgenommen werden soll, oder nicht
                                'visboFlagListe.Add(.nameID, msTask.GetField(visboflag) = pbYes)
                                visboFlagListe.Add(.nameID, hflag)

                            End If

                            'percentDone, falls Customfiels visbo_percentDone definiert ist
                            If visbo_percentDone <> 0 Then
                                Dim strPercentDone As String = msTask.GetField(visbo_percentDone)
                                Dim hpercent() As String = Split(strPercentDone, "%", , )
                                Dim vPercentDone As Double
                                Try
                                    vPercentDone = Convert.ToDouble(hpercent(0))

                                Catch e As FormatException
                                    vPercentDone = 0.0
                                Catch e As OverflowException
                                    Call MsgBox(hpercent(1) & " is outside the range of a Double.")
                                End Try
                                ' Änderung tk: percentDone sollte immer Werte zwischen 0..1 haben 
                                cphase.percentDone = vPercentDone / 100

                            End If


                            ' Deliverables, falls Customfield visbo_delivaerables definiert ist
                            Dim count As Integer = 0
                            Dim hvDel() As String
                            If visbo_deliverables <> 0 Then          ' VISBO Deliverables ist definiert
                                Dim vDeliverable As String = ""
                                If visbo_deliverables = MSProject.PjField.pjTaskIndicators Then
                                    vDeliverable = msTask.Notes
                                    hvDel = Split(vDeliverable, vbCr, , )
                                    count = hvDel.Length
                                Else
                                    vDeliverable = msTask.GetField(visbo_deliverables)
                                    hvDel = Split(vDeliverable, ";", , )
                                    count = hvDel.Length
                                End If
                                For iDel As Integer = 0 To count - 1
                                    If Not cphase.containsDeliverable(hvDel(iDel)) Then

                                        Try
                                            cphase.addDeliverable(hvDel(iDel).Trim)
                                        Catch ex As Exception

                                        End Try

                                    End If
                                Next iDel

                            End If

                            ' Responsible, falls Customfield visbo_responsible definiert ist
                            If visbo_responsible <> 0 Then          ' VISBO-Responsible ist definiert
                                Dim vResponsible As String = msTask.GetField(visbo_responsible)
                                cphase.verantwortlich = vResponsible
                            End If


                            ' Ampel-Erläuterung, falls Customfield visbo_ampeltext definiert ist
                            If visbo_ampeltext <> 0 Then
                                Dim vAmpelText As String = ""
                                If visbo_ampeltext = MSProject.PjField.pjTaskIndicators Then
                                    vAmpelText = msTask.Notes
                                Else
                                    vAmpelText = msTask.GetField(visbo_ampeltext)
                                End If
                                phBewertung.description = vAmpelText
                            End If

                            If visbo_ampel <> 0 Then

                                Dim visboAmpel As String = msTask.GetField(visbo_ampel)

                                Select Case visboAmpel

                                    Case "none"
                                        phBewertung.colorIndex = PTfarbe.none
                                    Case "red"
                                        phBewertung.colorIndex = PTfarbe.red
                                    Case "green"
                                        phBewertung.colorIndex = PTfarbe.green
                                    Case "yellow"
                                        phBewertung.colorIndex = PTfarbe.yellow
                                    Case Else
                                        phBewertung.colorIndex = PTfarbe.none

                                End Select

                            Else
                                phBewertung.colorIndex = PTfarbe.none
                            End If
                            cphase.addBewertung(phBewertung)

                            ' Änderung 28.11.13: jetzt wird die Phasen Länge exakt bestimmt , über startoffset in Tagen und dauerinDays als Länge
                            Dim cphaseStartOffset As Long
                            Dim dauerIndays As Long

                            cphaseStartOffset = DateDiff(DateInterval.Day, hproj.startDate, CDate(msTask.Start).Date)
                            dauerIndays = calcDauerIndays(CDate(msTask.Start).Date, CDate(msTask.Finish).Date)

                            .changeStartandDauer(cphaseStartOffset, dauerIndays)
                            .offset = 0

                            ' hier muss eine Routine aufgerufen werden, die die Dauer in Tagen berechnet !!!!!!
                            Dim phaseStartdate As Date = .getStartDate
                            Dim phaseEnddate As Date = .getEndDate

                            ' Verification Check
                            If DateDiff(DateInterval.Day, CDate(msTask.Start).Date, phaseStartdate.Date) <> 0 Then
                                outputline = "Task (Phase) : " & msTask.Name & "beginnt: " & CDate(msTask.Start).Date.ToShortDateString & "(MSProject) - " & phaseStartdate.ToShortDateString & "(VISBO)"
                                outputCollection.Add(outputline)
                                outputline = "Task (Phase) : " & msTask.Name & "endet: " & CDate(msTask.Finish).Date.ToShortDateString & "(MSProject) - " & phaseEnddate.ToShortDateString & "(VISBO)"
                                outputCollection.Add(outputline)
                            End If


                            Dim anzRessources As Integer = msTask.Resources.Count

                            ' Resourcen je MSTask durchgehen
                            Dim j As Integer = 0
                            Dim ccost As clsKostenart = Nothing
                            Dim crole As clsRolle = Nothing



                            Dim ass As MSProject.Assignment

                            If msproj.CurrencyCode = "EUR" And msTask.Assignments.Count > 0 Then


                                For Each ass In msTask.Assignments


                                    Dim msRess As MSProject.Resource = ass.Resource

                                    Select Case ass.Resource.Type
                                        Case MSProject.PjResourceTypes.pjResourceTypeMaterial To _
                                           MSProject.PjResourceTypes.pjResourceTypeCost
                                            Try

                                                Dim k As Integer = 0

                                                If CostDefinitions.containsName(ass.ResourceName) Then
                                                    k = CInt(CostDefinitions.getCostdef(ass.ResourceName).UID)
                                                Else
                                                    ' Kostenart existiert noch nicht
                                                    ' wird hier neu aufgenommen
                                                    Dim newCostDef As New clsKostenartDefinition
                                                    newCostDef.name = ass.ResourceName
                                                    newCostDef.farbe = RGB(120, 120, 120)   ' Farbe: grau
                                                    newCostDef.UID = CostDefinitions.Count + 1
                                                    If Not missingCostDefinitions.containsName(newCostDef.name) Then
                                                        missingCostDefinitions.Add(newCostDef)
                                                    End If

                                                    CostDefinitions.Add(newCostDef)

                                                    ' Änderung tk: muss auf costdefinitions gesetzt werden 
                                                    ' k = CInt(missingCostDefinitions.getCostdef(ass.ResourceName).UID)
                                                    k = CInt(CostDefinitions.getCostdef(ass.ResourceName).UID)
                                                End If

                                                Dim work As Double = CType(ass.Work, Double)
                                                Dim cost As Double = CType(ass.Cost, Double)

                                                Dim startdate As Date = CDate(msTask.Start)
                                                Dim endedate As Date = CDate(msTask.Finish)

                                                Dim anzmonth As Integer = CInt(DateDiff(DateInterval.Month, startdate, endedate))
                                                Dim anzdays As Integer = CInt(DateDiff(DateInterval.Day, startdate, endedate))
                                                Dim anzhours As Integer = CInt(DateDiff(DateInterval.Hour, startdate, endedate))

                                                If anzhours > 0 And anzdays = 0 And anzmonth = 0 Then
                                                    anzdays = 1
                                                    anzmonth = 1
                                                End If
                                                If anzdays > 0 And anzmonth = 0 Then
                                                    anzmonth = 1
                                                End If


                                                ReDim Xwerte(anzmonth - 1)

                                                Dim m As Integer
                                                For m = 1 To anzmonth

                                                    Try
                                                        Xwerte(m - 1) = CType(cost / anzmonth, Double)
                                                    Catch ex As Exception
                                                        Xwerte(m - 1) = 0.0
                                                    End Try

                                                Next m

                                                ccost = New clsKostenart(anzmonth - 1)

                                                With ccost
                                                    .KostenTyp = k
                                                    .Xwerte = Xwerte
                                                End With


                                                With cphase
                                                    .AddCost(ccost)
                                                End With
                                            Catch ex As Exception
                                                '
                                                ' handelt es sich um die Kostenart Definition?
                                                '
                                            End Try
                                            'Call MsgBox("Kosten = " & ass.ResourceName)

                                        Case MSProject.PjResourceTypes.pjResourceTypeWork

                                            Try
                                                Dim r As Integer = 0


                                                If RoleDefinitions.containsName(ass.ResourceName) Then
                                                    r = CInt(RoleDefinitions.getRoledef(ass.ResourceName).UID)
                                                Else
                                                    ' Rolle existiert noch nicht
                                                    ' wird hier neu aufgenommen

                                                    Dim newRoleDef As New clsRollenDefinition
                                                    newRoleDef.name = ass.ResourceName
                                                    newRoleDef.farbe = RGB(120, 120, 120)
                                                    newRoleDef.defaultKapa = 200000

                                                    ' OvertimeRate in Tagessatz umrechnen
                                                    Dim hoverstr() As String = Split(CStr(ass.Resource.OvertimeRate), "/", -1)
                                                    hoverstr = Split(hoverstr(0), "€", -1)
                                                    'newRoleDef.tagessatzExtern = CType(hoverstr(0), Double) * msproj.HoursPerDay

                                                    ' StandardRate in Tagessatz umrechnen
                                                    Dim hstdstr() As String = Split(CStr(ass.Resource.StandardRate), "/", -1)
                                                    hstdstr = Split(hstdstr(0), "€", -1)
                                                    newRoleDef.tagessatzIntern = CType(hstdstr(0), Double) * msproj.HoursPerDay

                                                    newRoleDef.UID = RoleDefinitions.Count + 1
                                                    If Not missingRoleDefinitions.containsName(newRoleDef.name) Then
                                                        missingRoleDefinitions.Add(newRoleDef)
                                                    End If

                                                    RoleDefinitions.Add(newRoleDef)


                                                    ' Änderung tk: das muss von roledefinitions geholt werden ...
                                                    ' r = CInt(missingRoleDefinitions.getRoledef(ass.ResourceName).UID)
                                                    r = CInt(RoleDefinitions.getRoledef(ass.ResourceName).UID)

                                                End If



                                                Dim work As Double = CType(ass.Work, Double)
                                                'Dim duration As Double = CType(ass.Duration, Double)
                                                Dim unit As Double = CType(ass.Units, Double)
                                                Dim budgetWork As Double = CType(ass.BudgetWork, Double)

                                                Dim startdate As Date = CDate(msTask.Start).Date
                                                Dim endedate As Date = CDate(msTask.Finish).Date

                                                ' Änderung tk: wurde ersetzt durch tk Anpassung: keine Gleichverteilung auf die Monate, sondern 
                                                ' entsprechend der Lage der Monate ; es muss auch beachtet werden, dass anzmonth von 3.5 - 1.6 2 Monate sind; 
                                                ' die Berechnung Datediff ergibt aber nur 1 Monat '
                                                'Dim anzmonth As Integer = CInt(DateDiff(DateInterval.Month, startdate, endedate))
                                                'Dim anzdays As Integer = CInt(DateDiff(DateInterval.Day, startdate, endedate))
                                                'Dim anzhours As Integer = CInt(DateDiff(DateInterval.Hour, startdate, endedate))

                                                'If anzhours > 0 And anzdays = 0 And anzmonth = 0 Then
                                                '    anzdays = 1
                                                '    anzmonth = 1
                                                'End If
                                                'If anzdays > 0 And anzmonth = 0 Then
                                                '    anzmonth = 1
                                                'End If


                                                'ReDim Xwerte(anzmonth - 1)
                                                ' Ende Auskommentierung tk  

                                                ' tk Anpassung ...
                                                Dim oldWerte(0) As Double
                                                Dim anzmonth As Integer = getColumnOfDate(endedate) - getColumnOfDate(startdate) + 1
                                                oldWerte(0) = work
                                                ReDim Xwerte(anzmonth - 1)
                                                Call cphase.berechneBedarfe(startdate, endedate, oldWerte, 1.0, Xwerte)


                                                For m As Integer = 1 To anzmonth
                                                    Xwerte(m - 1) = Xwerte(m - 1) / 60 / 8
                                                Next

                                                ' Ende tk Anpassung


                                                ' Änderung tk: wieder auskommentieren - alter Code: hier wurde gleichverteilt  
                                                'For m As Integer = 1 To anzmonth

                                                '    Try
                                                '        ' Xwerte in Anzahl Tage; in MSProject alle Werte in anz. Minuten
                                                '        Xwerte(m - 1) = CType(work / anzmonth / 60 / 8, Double)

                                                '    Catch ex As Exception
                                                '        Xwerte(m - 1) = 0.0
                                                '    End Try

                                                'Next m

                                                ' Check , um Unterschiede in der Summe herausfinden zu können
                                                ' die waren immer 0 ... 
                                                'Dim aChck As Double = Xwerte1.Sum - Xwerte.Sum

                                                crole = New clsRolle(anzmonth - 1)
                                                With crole
                                                    .uid = r
                                                    .Xwerte = Xwerte
                                                End With

                                                With cphase
                                                    .addRole(crole)
                                                End With
                                            Catch ex As Exception

                                            End Try

                                            'Call MsgBox("Work = " & ass.ResourceName & " mit " & CStr(ass.Work) & "Arbeit")
                                    End Select
                                Next ass


                            End If

                            ' Hierarchie-Aufbau
                            Dim cphaseParent As Object = msTask.Parent

                            Dim hrchynode As New clsHierarchyNode
                            hrchynode.elemName = cphase.name

                            If tasklevel = 0 Then
                                hrchynode.parentNodeKey = ""

                            ElseIf tasklevel = 1 Then
                                hrchynode.parentNodeKey = rootPhaseName

                            ElseIf tasklevel - lastlevel = 1 Then
                                hrchynode.parentNodeKey = lastelemID

                            ElseIf tasklevel - lastlevel = 0 Then
                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(lastelemID)

                            ElseIf lastlevel - tasklevel >= 1 Then
                                Dim hilfselemID As String = lastelemID
                                For l As Integer = 1 To lastlevel - tasklevel
                                    hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                Next l
                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(hilfselemID)
                            Else
                                Throw New ArgumentException("Fehler beim Import! Hierarchie kann nicht richtig aufgebaut werden")
                            End If

                            ' Bestimmung des eindeutigen Namens innerhalb der Geschwister, unterschieden nach Meilensten  und Phase
                            Dim newStdName As String = ""
                            If awinSettings.createUniqueSiblingNames Then
                                newStdName = hproj.hierarchy.findUniqueGeschwisterName(hrchynode.parentNodeKey, msTask.Name, False)
                            Else
                                newStdName = msTask.Name
                            End If

                            cphase.nameID = hproj.hierarchy.findUniqueElemKey(newStdName, False)

                            hproj.AddPhase(cphase, origName:=origPhName, parentID:=hrchynode.parentNodeKey)

                            ' '' ''hproj.hierarchy.addNode(hrchynode, cphase.nameID)
                            hrchynode.indexOfElem = hproj.AllPhases.Count
                            ' merken von letzem Element (Knoten,Phase,Meilenstein)
                            lasthrchyNode = hrchynode
                            lastelemID = cphase.nameID
                            lastphase = cphase
                            lastlevel = tasklevel
                        End With


                        Dim oBreadCrumb As String = hproj.hierarchy.getBreadCrumb(lastelemID)

                    Else
                        ' mstask ist ein Meilenstein und kein Summary-Meilenstein


                        ' Ergänzung tk für Demo BHTC 
                        ' falls Synonyme definiert sind, ersetzen durch Std-Name, sonst bleibt Name unverändert 
                        Dim origMsName As String = msTask.Name
                        msTask.Name = milestoneMappings.mapToStdName("", msTask.Name)
                        '


                        Dim hierarchy As String = msTask.WBS
                        'Dim oBreadCrumb As String = hproj.hierarchy.getBreadCrumb(lastelemID)
                        Dim msPhase As clsPhase = Nothing
                        Dim parentID As String = rootPhaseName

                        lastlevel = hproj.hierarchy.getIndentLevel(lastelemID)

                        If lastlevel = -1 Then          ' lastelemID existiert in der hierarchy nicht, also wird Meilenstein der Rootphase zugeordnet
                            parentID = rootPhaseName

                        ElseIf tasklevel = lastlevel Then
                            parentID = hproj.hierarchy.getParentIDOfID(lastelemID)

                        ElseIf tasklevel > lastlevel Then
                            parentID = lastelemID

                        ElseIf tasklevel = 1 And tasklevel < lastlevel Then
                            parentID = rootPhaseName

                        ElseIf lastlevel - tasklevel >= 1 Then
                            Dim hilfselemID As String = lastelemID
                            For l As Integer = 1 To lastlevel - tasklevel
                                hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                            Next l
                            parentID = hproj.hierarchy.getParentIDOfID(hilfselemID)

                        End If

                        msPhase = hproj.getPhaseByID(parentID)

                        Dim cmilestone As New clsMeilenstein(msPhase)


                        ' prüfen, ob MeilensteinDefinition bereits vorhanden
                        If Not MilestoneDefinitions.Contains(msTask.Name) Then
                            Dim msDef As New clsMeilensteinDefinition
                            msDef.belongsTo = msPhase.name
                            msDef.name = msTask.Name
                            ' Abbreviation, falls Customfield visbo_abbrev definiert ist
                            If visbo_abbrev <> 0 Then          ' VISBO-Abbrev ist definiert
                                msDef.shortName = msTask.GetField(visbo_abbrev)
                            Else
                                msDef.shortName = ""
                            End If
                            ' Task Class, falls Customfield visbo_taskclass definiert ist
                            If visbo_taskclass <> 0 Then          ' VISBO-TaskClass ist definiert
                                msDef.darstellungsKlasse = msTask.GetField(visbo_taskclass)
                            Else
                                msDef.darstellungsKlasse = ""
                            End If
                            cmilestone.appearance = msDef.darstellungsKlasse

                            msDef.schwellWert = 0
                            msDef.UID = MilestoneDefinitions.Count + 1
                            'MilestoneDefinitions.Add(msDef)
                            Try
                                missingMilestoneDefinitions.Add(msDef)
                            Catch ex As Exception
                            End Try
                        Else
                            cmilestone.appearance = MilestoneDefinitions.getAppearance(msTask.Name)
                        End If

                        ' MeilensteinDefinition vorhanden?
                        If MilestoneDefinitions.Contains(msTask.Name) _
                            Or missingMilestoneDefinitions.Contains(msTask.Name) Then

                            Dim msBewertung As New clsBewertung

                            cmilestone.setDate = CType(msTask.Start, Date)

                            ' Bestimmung des eindeutigen Namens innerhalb der Geschwister, unterschieden nach Meilensten  und Phase
                            Dim newStdName As String = ""
                            If awinSettings.createUniqueSiblingNames Then
                                newStdName = hproj.hierarchy.findUniqueGeschwisterName(msPhase.nameID, msTask.Name, True)
                            Else
                                newStdName = msTask.Name
                            End If

                            cmilestone.nameID = hproj.hierarchy.findUniqueElemKey(newStdName, True)
                            Dim testDate As Date = cmilestone.getDate

                            ' Check der Daten: wenn nicht identisch, dann Output bringen
                            If DateDiff(DateInterval.Day, CDate(msTask.Start).Date, cmilestone.getDate) <> 0 Then
                                outputline = "Task(Milestone): " & msTask.Name & "beginnt: " & CDate(msTask.Start).Date.ToShortDateString & "(MSProject) - " & cmilestone.getDate.ToShortDateString & "(VISBO)"
                                outputCollection.Add(outputline)
                            End If

                            'percentDone, falls Customfiels visbo_percentDone definiert ist
                            If visbo_percentDone <> 0 Then
                                Dim strPercentDone As String = msTask.GetField(visbo_percentDone)
                                Dim hpercent() As String = Split(strPercentDone, "%", , )
                                Dim vPercentDone As Double
                                Try
                                    vPercentDone = Convert.ToDouble(hpercent(0))

                                Catch e As FormatException
                                    vPercentDone = 0.0
                                Catch e As OverflowException
                                    Call MsgBox(hpercent(1) & " is outside the range of a Double.")
                                End Try
                                ' Änderung tk: percentDone sollte immer Werte zwischen 0..1 haben 
                                cmilestone.percentDone = vPercentDone / 100

                            End If

                            ' Deliverables, falls Customfield visbo_delivaerables definiert ist
                            Dim count As Integer = 0
                            Dim hvDel() As String
                            If visbo_deliverables <> 0 Then          ' VISBO Deliverables ist definiert
                                Dim vDeliverable As String = ""
                                If visbo_deliverables = MSProject.PjField.pjTaskIndicators Then
                                    vDeliverable = msTask.Notes
                                    hvDel = Split(vDeliverable, vbCr, , )
                                    count = hvDel.Length
                                Else
                                    vDeliverable = msTask.GetField(visbo_deliverables)
                                    hvDel = Split(vDeliverable, ";", , )
                                    count = hvDel.Length
                                End If
                                For iDel As Integer = 0 To count - 1
                                    If Not cmilestone.containsDeliverable(hvDel(iDel)) Then

                                        Try
                                            cmilestone.addDeliverable(hvDel(iDel).Trim)
                                        Catch ex As Exception

                                        End Try

                                    End If
                                Next iDel

                            End If

                            ' Responsible, falls Customfield visbo_responsible definiert ist
                            If visbo_responsible <> 0 Then          ' VISBO-Responsible ist definiert
                                Dim vResponsible As String = msTask.GetField(visbo_responsible)
                                cmilestone.verantwortlich = vResponsible
                            End If

                            ' Ampel-Erläuterung, falls Customfield visbo_ampeltext definiert ist
                            If visbo_ampeltext <> 0 Then
                                Dim vAmpelText As String = ""
                                If visbo_ampeltext = MSProject.PjField.pjTaskIndicators Then
                                    vAmpelText = msTask.Notes
                                Else
                                    vAmpelText = msTask.GetField(visbo_ampeltext)
                                End If
                                msBewertung.description = vAmpelText
                            End If

                            If visbo_ampel <> 0 Then

                                Dim visboAmpel As String = msTask.GetField(visbo_ampel)

                                Select Case visboAmpel

                                    Case "none"
                                        msBewertung.colorIndex = PTfarbe.none
                                    Case "red"
                                        msBewertung.colorIndex = PTfarbe.red
                                    Case "green"
                                        msBewertung.colorIndex = PTfarbe.green
                                    Case "yellow"
                                        msBewertung.colorIndex = PTfarbe.yellow
                                    Case Else
                                        msBewertung.colorIndex = PTfarbe.none

                                End Select

                            Else
                                msBewertung.colorIndex = PTfarbe.none
                            End If

                            cmilestone.addBewertung(msBewertung)


                            If visboflag <> 0 Then        ' Ist VISBO-flag definiert?

                                Dim hflag As Boolean = readCustomflag(msTask, visboflag)
                                ' Liste, ob Meilenstein in Projekt für die Projekt-Tafel aufgenommen werden soll, oder nicht
                                visboFlagListe.Add(cmilestone.nameID, hflag)
                            End If

                            Try
                                With msPhase
                                    .addMilestone(cmilestone, origName:=origMsName)
                                End With
                            Catch ex1 As Exception
                                Throw New Exception(ex1.Message)
                            End Try
                        Else
                            Throw New ArgumentException("Fehler: Meilenstein konnte nicht gefunden werden")
                        End If
                    End If

                    '' Testweise hier eingetragen

                    Dim anzVorgaenger As Integer = msTask.PredecessorTasks.Count
                    Dim anzNachfolger As Integer = msTask.SuccessorTasks.Count
                    Dim dependencies As MSProject.TaskDependencies = msTask.TaskDependencies

                    Dim startTask As Date = CType(msTask.Start, Date)
                    Dim endeTask As Date = CType(msTask.Finish, Date)




                Next i          ' Ende Schleife über alle Tasks/Phasen eines Projektes

                ' Ausgabe der Checks-Fehler
                If outputCollection.Count > 0 Then
                    Call showOutPut(outputCollection, "Import " & hproj.name & " Standard", "folgende Ungereimtheiten in den Daten wurden festgestellt")
                End If


                Dim ele_i As Integer = 0
                Dim msStart As Integer = hproj.hierarchy.getIndexOf1stMilestone

                ' Liste der Phasen/Meilensteine durchgehen und die Phasen/Meilensteine die den visbo-Flag nicht gesetzt haben aus der Hierarchie löschen
                For ele_i = 0 To visboFlagListe.Count - 1

                    Dim elemID As String = visboFlagListe.ElementAt(ele_i).Key
                    If hproj.hierarchy.containsKey(elemID) Then

                        If Not visboFlagListe.ElementAt(ele_i).Value Then

                            If elemIDIstMeilenstein(elemID) Then

                                ' Meilenstein muss entfernt werden

                                Dim hrchynode As clsHierarchyNode = hproj.hierarchy.nodeItem(elemID)
                                If hrchynode.childCount > 0 Then
                                    Call MsgBox("Knoten " & elemNameOfElemID(elemID) & " kann nicht aus der Hierarchie entfernt werden")
                                Else
                                    hproj.removeMeilenstein(elemID)
                                End If

                            Else        ' Element elemID ist eine Phase


                                If isRemovable(elemID, hproj, visboFlagListe) Then

                                    ' es wird die Phase elemID mit allen seinen Kindern gelöscht
                                    hproj.removePhase(elemID, True)

                                    ' ''Call MsgBox("isRemovable = true" & vbLf & _
                                    ' ''            elemID & " kann entfernt werden")
                                Else

                                    '' ''Call MsgBox("isRemovable = false" & vbLf & _
                                    '' ''            elemID & " kann nicht entfernt werden ")

                                End If
                            End If

                        Else
                            ' Phase/Meilenstein bleibt erhalten
                        End If

                    Else
                        ' Element elemID wurde bereits entfernt '
                        ' Call MsgBox("das Element elemID= " & elemID & " wurde bereits entfernt")
                    End If

                Next  ' Schleife über alle Phasen/Meilensteine zum entfernern derer, die VISBO-Flag nicht gesetzt haben




                Dim key As String = calcProjektKey(hproj.name, hproj.variantName)

                If modus = "BHTC" Then

                    ' prüfen, ob AlleProjekte das Projekt bereits enthält 
                    ' danach ist sichergestellt, daß AlleProjekte das Projekt bereits enthält 
                    If AlleProjekte.Containskey(key) Then
                        AlleProjekte.Remove(key)
                    End If

                    AlleProjekte.Add(hproj)

                Else
                    If ImportProjekte.Containskey(key) Then
                        ImportProjekte.Remove(key)
                    End If

                    ImportProjekte.Add(hproj)

                End If

                If modus = "BHTC" Then
                    ' Alle Projekte in ShowProjekte löschen
                    ShowProjekte.Clear()
                End If

                If Not ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Add(hproj)
                Else
                    ShowProjekte.Remove(hproj.name)
                    ShowProjekte.Add(hproj)
                    'Call MsgBox("Projekt " & hproj.name & " ist bereits in der Projekt-Liste enthalten")
                End If

                ' Fehlermeldung: Falsche Währung vordefiniert.
                If msproj.CurrencyCode <> "EUR" Then
                    Call MsgBox("Vorsicht: Es wurden keine Ressourcen eingelesen, da die definierte Währung nicht EUR sondern " & msproj.CurrencyCode & " ist.")
                End If


                If modus = "BHTC" Or visbo_mapping <> 0 Then

                    ' ---------------------
                    ' Mapping gewünscht
                    ' ---------------------

                    If visbo_mapping <> 0 Then

                        mapProj = mappingProject(msproj, mapStruktur, hproj, visbo_mapping)

                        If IsNothing(mapProj) Then
                            Call MsgBox("Kein Mapping erfolgt")
                        End If

                    End If

                    ' --------------------
                    ' Mapping hier beendet
                    ' --------------------

                    ' ----------------------------------------
                    ' Eintrag in ShowProjekte und AlleProjekte 
                    ' ----------------------------------------
                    '
                    ' ist erforderlich für die Erstellung des Reports

                    If Not IsNothing(mapProj) Then

                        key = calcProjektKey(mapProj.name, mapProj.variantName)

                        If modus = "BHTC" Then

                            ' prüfen, ob AlleProjekte das Projekt bereits enthält 
                            ' danach ist sichergestellt, daß AlleProjekte das Projekt bereits enthält 
                            If AlleProjekte.Containskey(key) Then
                                AlleProjekte.Remove(key)
                            End If

                            AlleProjekte.Add(mapProj)

                        Else
                            If ImportProjekte.Containskey(key) Then
                                ImportProjekte.Remove(key)
                            End If

                            ImportProjekte.Add(mapProj)

                        End If

                        If modus = "BHTC" Then
                            ' Alle Projekte entfernen
                            ShowProjekte.Clear()
                        End If


                        If Not ShowProjekte.contains(mapProj.name) Then
                            ShowProjekte.Add(mapProj)

                        End If

                        ' Fehlermeldung: Falsche Währung vordefiniert.
                        If msproj.CurrencyCode <> "EUR" Then
                            Call MsgBox("Vorsicht: Es wurden keine Ressourcen eingelesen, da die definierte Währung nicht EUR sondern " & msproj.CurrencyCode & " ist.")
                        End If


                    End If

                End If

                ' Wenn Aufruf aus VisualBoard, so muss MS Project wieder geschlossen werden
                If modus <> "BHTC" Then
                    prj.FileExit(MSProject.PjSaveType.pjDoNotSave)
                End If

            Else

                Call MsgBox("Bitte zunächst ein Projekt öffnen !")

            End If
        Catch ex As Exception
            Call MsgBox(ex)
        End Try

        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' liest einen ProjektSteckbrief mit Hierarchie ein 
    ''' Außerdem gibt es die Spalte Summe, in der die Summe der Kosten enthalten sein kann.
    ''' 
    ''' </summary>
    ''' <param name="hprojekt"></param>
    ''' <param name="hprojTemp"></param>
    ''' <param name="isTemplate"></param>
    ''' <param name="importDatum"></param>
    ''' <remarks></remarks>
    Public Sub awinImportProjectmitHrchy(ByRef hprojekt As clsProjekt, ByRef hprojTemp As clsProjektvorlage, ByVal isTemplate As Boolean, ByVal importDatum As Date)

        Dim zeile As Integer, spalte As Integer
        Dim hproj As New clsProjekt
        Dim ProjektdauerIndays As Integer = 0
        Dim endedateProjekt As Date

        Dim projektAmpelFarbe As Integer
        Dim projektAmpelText As String

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet

        Try

            zeile = 1
            spalte = 1
            ' ------------------------------------------------------------------------------------------------------
            ' Einlesen der Stammdaten
            ' ------------------------------------------------------------------------------------------------------

            Try
                Dim wsGeneralInformation As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Stammdaten"),
                    Global.Microsoft.Office.Interop.Excel.Worksheet)
                With wsGeneralInformation

                    .Unprotect(Password:="x")       ' Blattschutz aufheben

                    ' Projekt-Name auslesen
                    hproj.name = makeValidProjectName(CType(.Range("Projekt_Name").Value, String))
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
                    Dim startOffset As Long = DateDiff(DateInterval.Day, hproj.startDate, hproj.startDate.AddMonths(0))

                    ' Budget
                    Try
                        hproj.Erloes = CType(.Range("Budget").Value, Double)
                    Catch ex1 As Exception

                    End Try


                    ' Ampel-Farbe
                    projektAmpelFarbe = CType(.Range("Bewertung").Value, Integer)
                    If projektAmpelFarbe >= 0 And projektAmpelFarbe <= 3 Then
                        ' zulässiger Wert
                    Else
                        projektAmpelFarbe = 0
                    End If


                    ' Ampel-Bewertung 
                    projektAmpelText = CType(.Range("BewertgErläuterung").Value, String)
                    ' das kann jetzt noch gar nicht zugewiesen werden, weil es noch keine Phasen gibt
                    ' Ampel-Beschreibung und Farbe ist jetzt Attribut der Phase(1), der Projekt-Phase
                    'hproj.ampelErlaeuterung = ampelText


                End With
            Catch ex As Exception
                Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Stammdaten", hproj.name, anzFehler)
                Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Stammdaten")
            End Try

            ' ------------------------------------------------------------------------------------------------------
            ' Einlesen der Attribute
            ' ------------------------------------------------------------------------------------------------------

            Try
                Dim wsAttribute As Excel.Worksheet
                Try
                    wsAttribute = CType(appInstance.ActiveWorkbook.Worksheets("Attribute"),
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
                                hproj.variantName = ""
                            End If
                        Catch ex1 As Exception
                            hproj.variantName = ""
                        End Try

                        '   Varianten-Beschreibung
                        Try
                            hproj.variantDescription = ""
                            Dim tmprng As Excel.Range = CType(.Range("Variant_Description"), Excel.Range)
                            If Not IsNothing(tmprng) Then
                                If Not IsNothing(tmprng.Value) Then
                                    hproj.variantDescription = CType(.Range("Variant_Description").Value, String)
                                End If
                            End If

                            'If Not IsNothing(hproj.variantDescription) Then
                            '    hproj.variantDescription = hproj.variantDescription.Trim
                            'Else
                            '    hproj.variantDescription = ""
                            'End If

                        Catch ex1 As Exception
                            hproj.variantDescription = ""
                        End Try

                        ' Business Unit - kein Problem wenn nicht da   
                        Try
                            hproj.businessUnit = CType(.Range("Business_Unit").Value, String)
                        Catch ex As Exception

                        End Try

                        ' Risiko
                        hproj.Risiko = CDbl(.Range("Risiko").Value)


                        ' Strategic Fit
                        hproj.StrategicFit = CDbl(.Range("Strategischer_Fit").Value)


                        ' Ergänzung tk 19.5 es können hier auch sogenannte Custom Fields eingelesen werden ...
                        Try
                            Dim cfRange As Excel.Range = CType(.Range("IndivName2"), Excel.Range)
                            Dim startzeile As Integer = cfRange.Row
                            Dim cfValueColumn As Integer = cfRange.Column
                            Dim lastZeile As Integer = CInt(CType(.Cells(10000, 2), Excel.Range).End(XlDirection.xlUp).Row)

                            ' jetzt die Custom-Fields einlesen 
                            For i As Integer = startzeile To lastZeile

                                Try

                                    Dim cfName As String = CStr(CType(.Cells(i, cfValueColumn - 1), Excel.Range).Value).Trim
                                    Dim cfUid As Integer = customFieldDefinitions.getUid(cfName)

                                    If cfUid > -1 Then ' dann existiert diese Custom Field Definition 
                                        Dim cfType As Integer = customFieldDefinitions.getTyp(cfUid)

                                        If Not IsNothing(cfType) Then
                                            Select Case cfType
                                                Case ptCustomFields.Str
                                                    Dim cfvalue As String = CStr(CType(.Cells(i, cfValueColumn), Excel.Range).Value)
                                                    hproj.addSetCustomSField(cfUid, cfvalue)
                                                Case ptCustomFields.Dbl
                                                    Dim cfvalue As Double = CDbl(CType(.Cells(i, cfValueColumn), Excel.Range).Value)
                                                    hproj.addSetCustomDField(cfUid, cfvalue)
                                                Case ptCustomFields.bool
                                                    Dim cfvalue As Boolean = CBool(CType(.Cells(i, cfValueColumn), Excel.Range).Value)
                                                    hproj.addSetCustomBField(cfUid, cfvalue)
                                                Case Else
                                                    ' Custom Field Type nicht bekannt ...
                                                    Call logfileSchreiben("unbekanntes Custom-Field, wird ignoriert: ", hproj.name & " " & cfName & "," & cfType, anzFehler)
                                            End Select
                                        Else
                                            ' Custom Field UID nicht existent ...
                                            Call logfileSchreiben("uid von Custom-Field existiert nicht ...", hproj.name & " " & cfName & "," & cfUid, anzFehler)
                                        End If
                                    Else
                                        ' Custom Field Definition nicht bekannt ...
                                        Call logfileSchreiben("unbekanntes Custom-Field, wird ignoriert: ", hproj.name & " " & cfName, anzFehler)
                                    End If

                                Catch ex As Exception

                                End Try

                            Next


                        Catch ex As Exception

                        End Try



                    End With
                End If
            Catch ex As Exception
                Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Attribute", hproj.name, anzFehler)
                Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Attribute")
            End Try


            ' ------------------------------------------------------------------------------------------------------
            ' Einlesen der Termine ur: 06.10.2015: nun vor dem Einlesen der Phasen
            ' ------------------------------------------------------------------------------------------------------


            Try
                Dim wsTermine As Excel.Worksheet
                Try
                    wsTermine = CType(appInstance.ActiveWorkbook.Worksheets("Termine"),
                                                                 Global.Microsoft.Office.Interop.Excel.Worksheet)
                Catch ex As Exception
                    wsTermine = Nothing
                End Try

                If Not IsNothing(wsTermine) Then
                    Try
                        With wsTermine
                            Dim lastrow As Integer
                            Dim phaseNameID As String
                            Dim milestoneName As String
                            Dim milestoneDate As Date
                            Dim bewertungsAmpel As Integer
                            Dim explanation As String
                            Dim deliverables As String
                            Dim responsible As String = ""
                            Dim percentDone As Double = 0.0
                            Dim bewertungsdatum As Date = importDatum
                            Dim tbl As Excel.Range
                            Dim rowOffset As Integer
                            Dim columnOffset As Integer


                            .Unprotect(Password:="x")       ' Blattschutz aufheben

                            tbl = .Range("ErgebnTabelle")
                            rowOffset = tbl.Row
                            columnOffset = tbl.Column

                            lastrow = CInt(CType(.Cells(40000, columnOffset), Excel.Range).End(XlDirection.xlUp).Row)

                            ' ur: 12.05.2015: hier wurde die Sortierung der ErgebnTabelle entfernt

                            Dim cphase As New clsPhase(parent:=hproj)
                            Dim lastPhase As New clsPhase(parent:=hproj)
                            Dim breadCrumb As String = ""
                            Dim lastLevel As Integer = 0
                            Dim lasthrchynode As New clsHierarchyNode
                            Dim lastelemID As String = ""
                            Dim hilfselemID As String = ""


                            ' die beiden ersten Spalten verbinden, sofern nicht schon gemacht und abspeichern
                            Dim verbRange As Excel.Range
                            Dim iv As Integer

                            For iv = 0 To lastrow - rowOffset + 1
                                verbRange = .Range(.Cells(rowOffset + iv, columnOffset), .Cells(rowOffset + iv, columnOffset + 1))
                                verbRange.Merge()
                            Next


                            For zeile = rowOffset To lastrow


                                Dim cMilestone As clsMeilenstein
                                Dim cBewertung As New clsBewertung

                                Dim objectName As String
                                Dim startDate As Date, endeDate As Date
                                ' 
                                Dim errMessage As String = ""
                                Dim aktLevel As Integer = 0

                                Dim isPhase As Boolean = False
                                Dim isMeilenstein As Boolean = False
                                Dim cphaseExisted As Boolean = True

                                Dim duration As Long
                                Dim offset As Long

                                ' 10.5.18 document URL String ergänzt 
                                Dim docURL As String = ""


                                Try
                                    ' String aus erster Spalte der Tabelle lesen

                                    objectName = CStr(CType(.Cells(zeile, columnOffset), Excel.Range).Value).Trim

                                    ' Level abfragen

                                    Dim x As Integer = CInt(CType(.Cells(zeile, columnOffset), Excel.Range).IndentLevel)
                                    If x Mod einrückTiefe <> 0 Then
                                        Call logfileSchreiben("Fehler, Lesen Termine: die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl", hproj.name, anzFehler)
                                        Throw New ArgumentException("Fehler, Lesen Termine: die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl")
                                    End If
                                    aktLevel = CInt(x / einrückTiefe)

                                Catch ex As Exception
                                    objectName = Nothing
                                    Call logfileSchreiben("Fehler, Lesen Termine: In Tabelle 'Termine' ist der PhasenName nicht angegeben ", hproj.name, anzFehler)
                                    Throw New Exception("Fehler, Lesen Termine: In Tabelle 'Termine' ist der PhasenName nicht angegeben ")
                                    Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                End Try

                                ' erste Zeile gelesen; muss RootPhasename sein
                                If zeile = rowOffset Then

                                    If (aktLevel <> 0 And objectName <> elemNameOfElemID(rootPhaseName)) Then
                                        Call logfileSchreiben("Fehler, Lesen Termine: In Tabelle 'Termine' fehlt die ProjektPhase '.' !", hproj.name, anzFehler)
                                        Throw New Exception("Fehler, Lesen Termine: In Tabelle 'Termine' fehlt die ProjektPhase '.' !")
                                        Exit For ' Ende der For-Schleife, wenn keine laufende Nummer mehr existiert
                                    Else
                                        ' erzeuge ProjektPhase rootPhaseName
                                        isPhase = True
                                        isMeilenstein = False
                                        Try
                                            startDate = CDate(CType(.Cells(zeile, columnOffset + 2), Excel.Range).Value)
                                        Catch ex As Exception
                                            startDate = Date.MinValue
                                        End Try
                                        Try
                                            endeDate = CDate(CType(.Cells(zeile, columnOffset + 3), Excel.Range).Value)
                                        Catch ex As Exception
                                            endeDate = Date.MinValue
                                        End Try

                                        ' das Feld %Done wird hier ausgelesen ...
                                        Try
                                            ' Ergänzung ur: 09.11.2017 %Done  ergänzt 
                                            percentDone = CType(CType(.Cells(zeile, columnOffset + 8), Excel.Range).Value, Double)
                                            If IsNothing(percentDone) Then
                                                percentDone = 0.0
                                            End If
                                        Catch ex As Exception
                                            percentDone = 0.0
                                        End Try

                                        ' das Feld document-Url wird hier ausgelesen ...
                                        Try
                                            ' Ergänzung tk: 10.05.2018 document-URL  ergänzt 
                                            docURL = CType(CType(.Cells(zeile, columnOffset + 9), Excel.Range).Value, String)
                                            If IsNothing(docURL) Then
                                                docURL = ""
                                            End If
                                        Catch ex As Exception
                                            docURL = ""
                                        End Try



                                        ' ProjektPhase wird erzeugt
                                        cphase = New clsPhase(parent:=hproj)


                                        ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                                        With cphase
                                            .nameID = rootPhaseName
                                            .percentDone = percentDone
                                            .DocURL = docURL

                                            duration = calcDauerIndays(startDate, endeDate)
                                            offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)

                                            If duration < 1 Or offset < 0 Then
                                                If startDate = Date.MinValue And endeDate = Date.MinValue Then
                                                    Call logfileSchreiben("Fehler, Lesen Termine:  zu '" & objectName & "' wurde kein Datum eingetragen!", hproj.name, anzFehler)
                                                    Throw New Exception("Fehler, Lesen Termine:  zu '" & objectName & "' wurde kein Datum eingetragen!")
                                                Else
                                                    Dim exMsg As String = "Fehler, Lesen Termine: unzulässige Angaben für Offset (>=0) und Dauer (>=1): " &
                                                                        "Offset= " & offset.ToString &
                                                                        ", Duration= " & duration.ToString & " " & objectName & "; "

                                                    Call logfileSchreiben(exMsg, hproj.name, anzFehler)
                                                    Throw New Exception(exMsg)
                                                End If
                                            End If

                                            ' für die rootPhase muss gelten: offset = startoffset = 0 und duration = ProjektdauerIndays
                                            If duration <> ProjektdauerIndays Or offset <> 0 Then

                                                Dim exMsg As String = "Fehler, Lesen Termine: unzulässige Angaben für Offset und Dauer: der ProjektPhase " &
                                                                        "Offset= " & offset.ToString &
                                                                        ", Duration=" & duration.ToString & " " & objectName & "; " &
                                                                        ", ProjektDauer=" & ProjektdauerIndays.ToString
                                                Call logfileSchreiben(exMsg, hproj.name, anzFehler)
                                                Throw New Exception(exMsg)
                                            Else
                                                Dim startOffset As Integer = 0
                                                .changeStartandDauer(startOffset, ProjektdauerIndays)
                                                Dim phaseStartdate As Date = .getStartDate
                                                Dim phaseEnddate As Date = .getEndDate

                                            End If

                                        End With


                                        ' ProjektPhase wird hinzugefügt
                                        Dim hrchynode As New clsHierarchyNode
                                        hrchynode.elemName = cphase.name
                                        hrchynode.parentNodeKey = ""
                                        hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)
                                        lastPhase = cphase
                                        lastelemID = cphase.nameID
                                    End If

                                Else
                                    ' alle weiteren Phasen oder Meilensteine
                                    Try
                                        startDate = CDate(CType(.Cells(zeile, columnOffset + 2), Excel.Range).Value)
                                    Catch ex As Exception
                                        startDate = Date.MinValue
                                    End Try
                                    Try
                                        endeDate = CDate(CType(.Cells(zeile, columnOffset + 3), Excel.Range).Value)
                                    Catch ex As Exception
                                        endeDate = Date.MinValue
                                    End Try

                                    If startDate = Date.MinValue And endeDate <> Date.MinValue Then
                                        isPhase = False
                                        isMeilenstein = True
                                    ElseIf startDate <> Date.MinValue And endeDate <> Date.MinValue Then

                                        duration = calcDauerIndays(startDate, endeDate)
                                        offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)

                                        If duration < 1 Or offset < 0 Then
                                            If startDate = Date.MinValue And endeDate = Date.MinValue Then
                                                Call logfileSchreiben(("Fehler, Lesen Termine:  zu '" & objectName & "' wurde kein Datum eingetragen!"), hproj.name, anzFehler)
                                                Throw New Exception("Fehler, Lesen Termine:  zu '" & objectName & "' wurde kein Datum eingetragen!")
                                            Else
                                                Dim exmsg As String = "Fehler, Lesen Termine: unzulässige Angaben für Offset und Dauer: " &
                                                                    offset.ToString & ", " & duration.ToString & ": " & objectName

                                                Call logfileSchreiben(exmsg, hproj.name, anzFehler)
                                                Throw New Exception(exmsg)
                                            End If
                                        End If

                                        isPhase = True
                                        isMeilenstein = False

                                    End If

                                    ' eingelesener String objectname ist eine Phase

                                    If isPhase Then

                                        cphase = New clsPhase(parent:=hproj)

                                        If PhaseDefinitions.Contains(objectName) Or isMissingDefinitionOK(objectName, isTemplate, False) Then

                                            With cphase
                                                .nameID = hproj.hierarchy.findUniqueElemKey(objectName, False)

                                                duration = calcDauerIndays(startDate, endeDate)
                                                offset = DateDiff(DateInterval.Day, hproj.startDate, startDate)

                                                .changeStartandDauer(offset, duration)
                                                Dim phaseStartdate As Date = .getStartDate
                                                Dim phaseEnddate As Date = .getEndDate

                                            End With


                                            Try
                                                bewertungsAmpel = CType(CType(.Cells(zeile, columnOffset + 4), Excel.Range).Value, Integer)
                                                If IsNothing(bewertungsAmpel) Then
                                                    bewertungsAmpel = 0
                                                End If
                                            Catch ex As Exception
                                                bewertungsAmpel = 0
                                            End Try

                                            Try
                                                explanation = CType(CType(.Cells(zeile, columnOffset + 5), Excel.Range).Value, String)
                                                If IsNothing(explanation) Then
                                                    explanation = ""
                                                End If
                                            Catch ex As Exception
                                                explanation = ""
                                            End Try

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

                                            ' das Feld Deliverables wird hier ausgelesen ...
                                            Try
                                                ' Ergänzung tk 2.11 deliverables ergänzt 
                                                deliverables = CType(CType(.Cells(zeile, columnOffset + 6), Excel.Range).Value, String)
                                                If IsNothing(deliverables) Then
                                                    deliverables = ""
                                                End If
                                            Catch ex As Exception
                                                deliverables = ""
                                            End Try

                                            ' das Feld Responsible wird hier ausgelesen ...
                                            Try
                                                ' Ergänzung tk 26.10.17 responsible ergänzt 
                                                responsible = CType(CType(.Cells(zeile, columnOffset + 7), Excel.Range).Value, String)
                                                If IsNothing(responsible) Then
                                                    responsible = ""
                                                End If
                                            Catch ex As Exception
                                                responsible = ""
                                            End Try

                                            ' das Feld %Done wird hier ausgelesen ...
                                            Try
                                                ' Ergänzung ur: 09.11.2017 %Done  ergänzt 
                                                percentDone = CType(CType(.Cells(zeile, columnOffset + 8), Excel.Range).Value, Double)
                                                If IsNothing(percentDone) Then
                                                    percentDone = 0.0
                                                End If
                                            Catch ex As Exception
                                                percentDone = 0.0
                                            End Try

                                            ' das Feld document-Url wird hier ausgelesen ...
                                            Try
                                                ' Ergänzung tk: 10.05.2018 document-URL  ergänzt 
                                                docURL = CType(CType(.Cells(zeile, columnOffset + 9), Excel.Range).Value, String)
                                                If IsNothing(docURL) Then
                                                    docURL = ""
                                                End If
                                            Catch ex As Exception
                                                docURL = ""
                                            End Try


                                            With cphase
                                                .percentDone = percentDone
                                                .verantwortlich = responsible
                                                .DocURL = docURL
                                                If Not IsNothing(cBewertung) Then
                                                    .addBewertung(cBewertung)
                                                End If

                                                ' ur: 09.11.2017
                                                ' hier müssen die Deliverables jetzt auseinander dividiert werden in die einzelnen Items
                                                Try
                                                    If deliverables.Trim.Length > 0 Then
                                                        Dim splitStr() As String = deliverables.Split(New Char() {CChar(vbLf), CChar(vbCr)}, 100)

                                                        ' tk 29.5.16 Deliverables jetzt als einzelnen Items 
                                                        For ix As Integer = 1 To splitStr.Length
                                                            .addDeliverable(splitStr(ix - 1))
                                                        Next
                                                    End If
                                                Catch ex As Exception

                                                End Try
                                            End With


                                            Dim hrchynode As New clsHierarchyNode
                                            hrchynode.elemName = cphase.name

                                            If aktLevel = 0 Then
                                                hrchynode.parentNodeKey = ""

                                            ElseIf aktLevel = 1 Then
                                                hrchynode.parentNodeKey = rootPhaseName

                                            ElseIf aktLevel - lastLevel = 1 Then
                                                hrchynode.parentNodeKey = lastelemID

                                            ElseIf aktLevel - lastLevel = 0 Then
                                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(lastelemID)

                                            ElseIf lastLevel - aktLevel >= 1 Then
                                                hilfselemID = lastelemID
                                                For l As Integer = 1 To lastLevel - aktLevel
                                                    hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                                Next l
                                                hrchynode.parentNodeKey = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                            Else
                                                Call logfileSchreiben(("Fehler, Lesen Termine: Hierarchie kann nicht richtig aufgebaut werden:" & cphase.nameID), hproj.name, anzFehler)
                                                Throw New ArgumentException("Fehler, Lesen Termine:  Hierarchie kann nicht richtig aufgebaut werden" & cphase.nameID)
                                            End If

                                            hproj.AddPhase(cphase, parentID:=hrchynode.parentNodeKey)
                                            '' ''hproj.hierarchy.addNode(hrchynode, cphase.nameID)
                                            hrchynode.indexOfElem = hproj.AllPhases.Count
                                            ' merken von letzem Element (Knoten,Phase,Meilenstein)
                                            lasthrchynode = hrchynode
                                            lastelemID = cphase.nameID
                                            lastPhase = cphase
                                            lastLevel = aktLevel

                                        Else
                                            ' objectname existiert nicht in den PhaseDefinitions
                                            ' muss in missingPhaseDefinitions noch eingetragen werden
                                            Call logfileSchreiben(("Fehler, Lesen Termine: Phase '" & objectName & "' existiert im CustomizationFile nicht!"), hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler, Lesen Termine:Phase '" & objectName & "' existiert im CustomizationFile nicht!")
                                        End If

                                    ElseIf isMeilenstein Then

                                        If MilestoneDefinitions.Contains(objectName) Or isMissingDefinitionOK(objectName, isTemplate, True) Then

                                            Dim hrchynode As New clsHierarchyNode
                                            hrchynode.elemName = cphase.name

                                            If aktLevel = 0 Then
                                                ' Fehler, denn Meilenstein kann nicht parallel zu Rootphase sein??
                                                Call logfileSchreiben(("Fehler, Lesen Termine: Hierarchie kann nicht richtig aufgebaut werden:" & vbLf & "Level des Meilensteins ist nicht akzeptabel" & objectName), hproj.name, anzFehler)
                                                Throw New ArgumentException("Fehler, Lesen Termine: Hierarchie kann nicht richtig aufgebaut werden:" & vbLf & "Level des Meilensteins ist nicht akzeptabel" & objectName)

                                            ElseIf aktLevel = 1 Then
                                                phaseNameID = rootPhaseName

                                            ElseIf aktLevel - lastLevel = 1 Then
                                                phaseNameID = lastelemID

                                            ElseIf aktLevel - lastLevel = 0 Then
                                                phaseNameID = hproj.hierarchy.getParentIDOfID(lastelemID)

                                            ElseIf lastLevel - aktLevel >= 1 Then
                                                hilfselemID = lastelemID
                                                For l As Integer = 1 To lastLevel - aktLevel
                                                    hilfselemID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                                Next l
                                                phaseNameID = hproj.hierarchy.getParentIDOfID(hilfselemID)
                                            Else
                                                Call logfileSchreiben(("Fehler, Lesen Termine: Hierarchie kann nicht richtig aufgebaut werden: Meilenstein " & objectName), hproj.name, anzFehler)
                                                Throw New ArgumentException("Fehler, Lesen Termine:  Hierarchie kann nicht richtig aufgebaut werden: Meilenstein " & objectName)
                                            End If


                                            Dim hilfsPhase As clsPhase = hproj.getPhaseByID(phaseNameID)
                                            cMilestone = New clsMeilenstein(parent:=hproj.getPhaseByID(phaseNameID))
                                            cBewertung = New clsBewertung

                                            milestoneName = objectName.Trim
                                            milestoneDate = endeDate

                                            ' wenn der freefloat nicht zugelassen ist und der Meilenstein ausserhalb der Phasen-Grenzen liegt 
                                            ' muss abgebrochen werden 

                                            If Not awinSettings.milestoneFreeFloat And
                                                (DateDiff(DateInterval.Day, hilfsPhase.getStartDate, milestoneDate) < 0 Or
                                                 DateDiff(DateInterval.Day, hilfsPhase.getEndDate, milestoneDate) > 0) Then

                                                Call logfileSchreiben(("Fehler, Lesen Termine: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf &
                                                                    milestoneName & " nicht innerhalb " & hilfsPhase.name & vbLf &
                                                                         "Korrigieren Sie bitte diese Inkonsistenz in der Datei '"), hproj.name, anzFehler)
                                                Throw New Exception("Fehler, Lesen Termine: Der Meilenstein liegt ausserhalb seiner Phase" & vbLf &
                                                                    milestoneName & " nicht innerhalb " & hilfsPhase.name & vbLf &
                                                                         "Korrigieren Sie bitte diese Inkonsistenz in der Datei '" & vbLf & hproj.name & ".xlsx'")
                                            End If


                                            ' wenn kein Datum angegeben wurde, soll das Ende der Phase als Datum angenommen werden 
                                            If DateDiff(DateInterval.Month, hproj.startDate, milestoneDate) < -1 Then
                                                milestoneDate = hproj.startDate.AddDays(hilfsPhase.startOffsetinDays + hilfsPhase.dauerInDays)
                                            Else
                                                If DateDiff(DateInterval.Day, endedateProjekt, endeDate) > 0 Then
                                                    Call logfileSchreiben(("Fehler, Lesen Termine: der Meilenstein '" & milestoneName & "' liegt später als das Ende des gesamten Projekts" & vbLf &
                                                                "Bitte korrigieren Sie dies im Tabellenblatt Ressourcen der Datei '"), hproj.name & ".xlsx", anzFehler)
                                                End If

                                            End If

                                            Try
                                                bewertungsAmpel = CType(CType(.Cells(zeile, columnOffset + 4), Excel.Range).Value, Integer)
                                                If IsNothing(bewertungsAmpel) Then
                                                    bewertungsAmpel = 0
                                                End If
                                            Catch ex As Exception
                                                bewertungsAmpel = 0
                                            End Try

                                            If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                                                ' es gibt keine Bewertung
                                                bewertungsAmpel = 0
                                            End If

                                            Try
                                                explanation = CType(CType(.Cells(zeile, columnOffset + 5), Excel.Range).Value, String)
                                                If IsNothing(explanation) Then
                                                    explanation = ""
                                                End If
                                            Catch ex As Exception
                                                explanation = ""
                                            End Try


                                            ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                            With cBewertung
                                                '.bewerterName = resultVerantwortlich
                                                .colorIndex = bewertungsAmpel
                                                .datum = importDatum
                                                .description = explanation
                                            End With


                                            Try
                                                ' Ergänzung tk 2.11 deliverables ergänzt 
                                                deliverables = CType(CType(.Cells(zeile, columnOffset + 6), Excel.Range).Value, String)
                                                If IsNothing(deliverables) Then
                                                    deliverables = ""
                                                End If
                                            Catch ex As Exception
                                                deliverables = ""
                                            End Try


                                            Try
                                                ' Ergänzung tk 26.10.17 responsible ergänzt 
                                                responsible = CType(CType(.Cells(zeile, columnOffset + 7), Excel.Range).Value, String)
                                                If IsNothing(responsible) Then
                                                    responsible = ""
                                                End If
                                            Catch ex As Exception
                                                responsible = ""
                                            End Try

                                            ' das Feld %Done wird hier ausgelesen ...
                                            Try
                                                ' Ergänzung ur: 09.11.2017 %Done  ergänzt 
                                                percentDone = CType(CType(.Cells(zeile, columnOffset + 8), Excel.Range).Value, Double)
                                                If IsNothing(percentDone) Then
                                                    percentDone = 0.0
                                                End If
                                            Catch ex As Exception
                                                percentDone = 0.0
                                            End Try

                                            ' das Feld document-Url wird hier ausgelesen ...
                                            Try
                                                ' Ergänzung tk: 10.05.2018 document-URL  ergänzt 
                                                docURL = CType(CType(.Cells(zeile, columnOffset + 9), Excel.Range).Value, String)
                                                If IsNothing(docURL) Then
                                                    docURL = ""
                                                End If
                                            Catch ex As Exception
                                                docURL = ""
                                            End Try

                                            With cMilestone
                                                .setDate = milestoneDate
                                                .verantwortlich = responsible
                                                .nameID = hproj.hierarchy.findUniqueElemKey(milestoneName, True)
                                                .percentDone = percentDone
                                                .DocURL = docURL
                                                If Not cBewertung Is Nothing Then
                                                    .addBewertung(cBewertung)
                                                End If
                                            End With

                                            ' tk 29.5.16
                                            ' hier müssen die Deliverables jetzt auseinander dividiert werden in die einzelnen Items
                                            Try
                                                If deliverables.Trim.Length > 0 Then
                                                    Dim splitStr() As String = deliverables.Split(New Char() {CChar(vbLf), CChar(vbCr)}, 100)

                                                    ' tk 29.5.16 Deliverables jetzt als einzelnen Items 
                                                    For ix As Integer = 1 To splitStr.Length
                                                        cMilestone.addDeliverable(splitStr(ix - 1))
                                                    Next
                                                End If
                                            Catch ex As Exception

                                            End Try


                                            Try
                                                With hproj.getPhaseByID(phaseNameID)
                                                    .addMilestone(cMilestone)
                                                End With
                                            Catch ex1 As Exception
                                                Throw New Exception(ex1.Message)
                                            End Try

                                        Else
                                            ' objectname existiert nicht in den PhaseDefinitions
                                            ' muss in missingPhaseDefinitions noch eingetragen werden
                                            Call logfileSchreiben(("Fehler, Lesen Termine: Meilenstein '" & objectName & "' existiert im CustomizationFile nicht!"), hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler, Lesen Termine:Meilenstein '" & objectName & "' existiert im CustomizationFile nicht!")
                                        End If

                                    End If



                                End If

                            Next zeile
                        End With

                    Catch ex As Exception
                        Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Termine: '" & ex.Message, hproj.name, anzFehler)
                        'Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Termine von '" & hproj.name & "' " & vbLf & ex.Message)
                        Throw New ArgumentException(ex.Message)

                    End Try


                Else

                    Call MsgBox("keine Termine definiert")
                    Throw New ArgumentException("Es wurden keine Termine definiert! Projekt " & hproj.name & " kann nicht eingelesen werden")
                End If
            Catch ex As Exception
                Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Termine: '" & ex.Message, hproj.name, anzFehler)
                Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Termine von '" & hproj.name & "' " & vbLf & ex.Message)

            End Try


            ' ------------------------------------------------------------------------------------------------------
            ' Einlesen der Ressourcen
            ' ------------------------------------------------------------------------------------------------------
            Dim wsRessourcen As Excel.Worksheet
            Try
                wsRessourcen = CType(appInstance.ActiveWorkbook.Worksheets("Ressourcen"),
                                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                wsRessourcen = Nothing
                ' '' '' '' ------------------------------------------------------------------------------------------------------
                ' '' '' '' Erzeugen und eintragen der Projekt-Phase (= erste Phase mit Dauer des Projekts)
                ' '' '' '' ------------------------------------------------------------------------------------------------------
                '' '' ''Try
                '' '' ''    Dim cphase As New clsPhase(hproj)

                '' '' ''    ' ProjektPhase wird erzeugt
                '' '' ''    cphase = New clsPhase(parent:=hproj)
                '' '' ''    cphase.nameID = rootPhaseName

                '' '' ''    ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                '' '' ''    With cphase
                '' '' ''        .nameID = rootPhaseName
                '' '' ''        Dim startOffset As Integer = 0
                '' '' ''        .changeStartandDauer(startOffset, ProjektdauerIndays)
                '' '' ''    End With
                '' '' ''    ' ProjektPhase wird hinzugefügt
                '' '' ''    hproj.AddPhase(cphase)

                '' '' ''Catch ex1 As Exception
                '' '' ''    Throw New ArgumentException("Fehler in awinImportProject, Erzeugen ProjektPhase")
                '' '' ''End Try

            End Try

            If Not IsNothing(wsRessourcen) Then

                Try
                    With wsRessourcen
                        Dim rng As Excel.Range
                        Dim zelle As Excel.Range
                        Dim ressSumOffset As Integer = 1
                        Dim ressOff As Integer = 2
                        Dim chkPhase As Boolean = True
                        Dim chkRolle As Boolean = True
                        Dim firsttime As Boolean = False
                        Dim fertig As Boolean = True
                        Dim summe As Double = -1        ' summe = -1: bedeutet, Summe wird nicht verwendet, oder hat einen unsinnigen Wert
                        Dim Xwerte As Double() = Nothing
                        Dim oldXwerte As Double()
                        Dim crole As clsRolle
                        Dim cphase As clsPhase = Nothing
                        Dim lastphase As clsPhase = Nothing
                        Dim lastelemID As String = ""
                        Dim ccost As clsKostenart
                        Dim phaseName As String = ""
                        Dim aktLevel As Integer = 0   'speichert den Level direkt nach dem Lesen der Phase
                        Dim cphaseLevel As Integer = 0 'speichert den Level der momentan in cphase gespeicherten Phase
                        Dim lastlevel As Integer = 0  'speichert den Level des vorausgehenden elements
                        Dim breadcrumb As String = ""
                        Dim anfang As Integer, ende As Integer  ', projDauer As Integer

                        Dim farbeAktuell As Object
                        Dim r As Integer, k As Integer


                        .Unprotect(Password:="x")       ' Blattschutz aufheben


                        'Dim tmpws As Excel.Range = CType(wsRessourcen.Range("Phasen_des_Projekts"), Excel.Range)
                        Dim oldrng = .Range("Phasen_des_Projekts")
                        ' Änderung tk: es muss die Spalte der Rollen betrachtet werden , wenn die Spalte der Phasen betrachtet wird, werden bei der letzten Phase die Rollen nicht mitgenommen 
                        Dim columnOffset As Integer = oldrng.Column

                        ' es muss das Maximum aus den beiden Spalten Pahse und Ressourcen gesucht werden
                        Dim lastrow1 As Integer = CInt(CType(.Cells(40000, columnOffset), Excel.Range).End(XlDirection.xlUp).Row)
                        Dim lastRow2 As Integer = CInt(CType(.Cells(40000, columnOffset + 2), Excel.Range).End(XlDirection.xlUp).Row)
                        Dim lastRow As Integer = System.Math.Max(lastrow1, lastRow2)
                        ' ´Verlängerung des Range "Phasen_des_Projekts" bis zur lastrow
                        rng = wsRessourcen.Range(.Cells(oldrng.Row, oldrng.Column), .Cells(lastRow, oldrng.Column))
                        'rng.Name = "Phasen_des_Projekts"

                        Dim testrange As Excel.Range = CType(.Cells(10, 2000), Excel.Range)
                        Dim gefundenRange As Excel.Range = testrange.Find(What:="Summe")
                        If IsNothing(gefundenRange) Then
                            ' alte Version des Steckbriefes 
                            ressOff = 1
                            ressSumOffset = -1              ' keine Summe vorhanden
                            Call logfileSchreiben("alte Version des ProjektSteckbriefes: ohne 'Summe'", hproj.name, anzFehler)
                        Else

                            ' die beiden ersten Spalten verbinden, sofern nicht schon gemacht und abspeichern
                            Dim verbRange As Excel.Range
                            Dim iv As Integer

                            For iv = 0 To rng.Rows.Count - 1
                                verbRange = .Range(.Cells(rng.Row + iv, rng.Column), .Cells(rng.Row + iv, rng.Column + 1))
                                verbRange.Merge()
                            Next

                            ressOff = gefundenRange.Column - rng.Column - 1
                            ressSumOffset = gefundenRange.Column - rng.Column - 2
                            'Call logfileSchreiben("neue Version des ProjektSteckbriefes: mit 'Summe'", hproj.name, anzFehler)


                            '' die beiden ersten Spalten verbinden, sofern nicht schon gemacht und abspeichern
                            'Dim verbRange As Excel.Range
                            'Dim iv As Integer

                            'For iv = 0 To rng.Rows.Count - 1
                            '    verbRange = .Range(.Cells(rng.Row + iv, rng.Column), .Cells(rng.Row + iv, rng.Column + 1))
                            '    verbRange.Merge()
                            'Next
                        End If

                        Dim hstr As String = CStr(CType(rng.Cells(1), Excel.Range).Value)
                        hstr = elemNameOfElemID(rootPhaseName)

                        If CStr(CType(rng.Cells(1), Excel.Range).Value) <> elemNameOfElemID(rootPhaseName) Then


                            ' ProjektPhase wird hinzugefügt, sofern sie nich
                            cphase = New clsPhase(parent:=hproj)
                            fertig = False


                            ' Phasen Dauer wird gleich der Dauer des Projekts gesetzt
                            With cphase
                                .nameID = rootPhaseName
                                Dim startOffset As Integer = 0
                                .changeStartandDauer(startOffset, ProjektdauerIndays)
                                Dim phaseStartdate As Date = .getStartDate
                                Dim phaseEnddate As Date = .getEndDate
                                firsttime = True
                            End With
                            'Call MsgBox("Projektnamen/Phasen Konflikt in awinImportProjekt" & vbLf & "Problem wurde behoben")

                        End If


                        zeile = 0

                        For Each zelle In rng

                            zeile = zeile + 1



                            ' nachsehen, ob Phase angegeben oder Rolle/Kosten
                            hstr = CStr(zelle.Value)
                            Dim x As Integer = CInt(zelle.IndentLevel)
                            If x Mod einrückTiefe <> 0 Then
                                Call logfileSchreiben("Fehler beim Lesen Ressourcen: die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl", hproj.name, anzFehler)
                                Throw New ArgumentException("Fehler beim Lesen Ressourcen: die Einrückung ist keine durch '" & CStr(einrückTiefe) & "' teilbare Zahl")
                            End If
                            aktLevel = CInt(x / einrückTiefe)

                            If Len(CType(zelle.Value, String)) > 0 Then
                                phaseName = CType(zelle.Value, String).Trim
                            Else
                                phaseName = ""
                            End If

                            ' hier wird die Rollen bzw Kosten Information ausgelesen
                            Dim hname As String = ""
                            Try

                                If Not IsNothing(CType(zelle.Offset(0, 1), Excel.Range).Value) Then
                                    hname = CType(CType(zelle.Offset(0, 1), Excel.Range).Value, String).Trim
                                End If

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

                                    If Not fertig Then

                                        lastelemID = cphase.nameID
                                        lastphase = cphase
                                        lastlevel = cphaseLevel
                                    End If

                                    ' in cphase wird die Phase mit Namen phaseName, bereits über Termine in der Hierarchie des Projekts eingetragen
                                    ' gespeichert
                                    ' das muss später überprüft werden können, um ggf gleichnamige Phasen in einer Breadcrumb Stufe richtig zuordnen zu können

                                    ' wenn in einer und derselben Hierarchy-Stufe mehrere gleichnamige Phasen vorkommen, so muss später anhand der Liste der 
                                    ' Phase-Nummern geprüft werden, welche denn die richtige Phase ist 
                                    Dim phaseIndex() As Integer

                                    If phaseName = hproj.name Or phaseName = elemNameOfElemID(rootPhaseName) Then

                                        cphase = hproj.getPhaseByID(rootPhaseName)
                                        ReDim phaseIndex(0)
                                        phaseIndex(0) = 1
                                        'das ist derselbe Effekt wie der untenstehende Befehl, nur schneller; und das Ergebnis muss ja gleich sein 
                                        ' phaseIndex = hproj.hierarchy.getPhaseIndices(cphase.name, "")

                                    Else

                                        ' erzeugen des breadcrumb, um nachsehen zu können, ob diese Phase in der gleichen Hierarchiestufe
                                        ' bereits über Termine eingelesen wurde
                                        If aktLevel > lastlevel Then

                                            If breadcrumb = "" Then
                                                breadcrumb = "."
                                            Else
                                                breadcrumb = breadcrumb & "#" & lastphase.name
                                            End If

                                        ElseIf aktLevel = lastlevel Then
                                            ' aktlevel = lastlevel: also nicht tun
                                        Else

                                            While aktLevel < lastlevel
                                                Dim hhstr As String = ""
                                                Dim type As Integer = -1
                                                Dim pvName As String = ""
                                                Call splitHryFullnameTo2(breadcrumb, hhstr, breadcrumb, type, pvName)
                                                lastlevel = lastlevel - 1
                                            End While

                                        End If

                                        ' Prüfung, ob die Phase phaseName in der bereits aus Termine bestehenden Hierarchie mit dem gleiche breadcrumb existiert, sonst Fehler


                                        If Not hproj.hierarchy.containsPhase(phaseName, breadcrumb) Then

                                            Dim xxx As Boolean = hproj.hierarchy.containsPhase(phaseName, breadcrumb)
                                            ReDim phaseIndex(0)
                                            Call logfileSchreiben("Fehler beim Lesen Ressourcen: bei Phase '" & phaseName & "#" & breadcrumb & "'", hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler beim Lesen Ressourcen: bei Phase '" & phaseName & "#" & breadcrumb & "'")
                                        Else

                                            phaseIndex = hproj.hierarchy.getPhaseIndices(phaseName, breadcrumb)

                                            cphase = hproj.getPhase(phaseIndex(0))
                                            cphaseLevel = hproj.hierarchy.getIndentLevel(cphase.nameID)

                                        End If

                                    End If

                                    fertig = False

                                    ' ur: 12.10.2015: neu:  Bedarfe nur als Summe angegeben

                                    ' Auslesen der Phasen Dauer und anschließend vergleichen, ob die in Termine mit der in Ressource übereinstimmt
                                    ' d.h. rel.Anfang und rel.Ende müssen übereinstimmen, wenn relStart und relEnde nicht übereinstimmen, so werden Sie einfach so gesetzt.

                                    Dim maxcol As Integer = hproj.anzahlRasterElemente
                                    Dim col As Integer


                                    col = 1
                                    While CInt(zelle.Offset(0, ressOff + col).Interior.ColorIndex) = -4142 And
                                             Not (CType(zelle.Offset(0, ressOff + col).Value, String) = "x") And
                                             col <= maxcol

                                        col = col + 1

                                    End While


                                    If col >= maxcol Then

                                        ' Phase und deren Länge wird nicht dargestellt in Tabellenblatt Ressourcen
                                        anfang = cphase.relStart
                                        ende = cphase.relEnde

                                    Else
                                        ' Phasenlänge wird dargestellt in Tabellenblatt Ressourcen, also überprüfen

                                        anfang = col

                                        Try
                                            ende = anfang + 1

                                            If CInt(zelle.Offset(0, ressOff + anfang).Interior.ColorIndex) = -4142 Then
                                                While CType(zelle.Offset(0, ressOff + ende).Value, String) = "x"
                                                    ende = ende + 1
                                                End While
                                                ende = ende - 1
                                            Else
                                                farbeAktuell = zelle.Offset(0, ressOff + anfang).Interior.Color
                                                While CInt(zelle.Offset(0, ressOff + ende).Interior.Color) = CInt(farbeAktuell)

                                                    ende = ende + 1
                                                End While
                                                ende = ende - 1
                                            End If

                                        Catch ex As Exception
                                            Call logfileSchreiben("Fehler beim Lesen Ressourcen: Es wurden keine oder falsche Angaben zur Phasendauer der Phase '" & phaseName & "' gemacht." & vbLf &
                                                                       "Bitte überprüfen Sie dies.", hproj.name, anzFehler)
                                            Throw New ArgumentException("Fehler beim Lesen Ressourcen: Es wurden keine oder falsche Angaben zur Phasendauer der Phase '" & phaseName & "' gemacht." & vbLf &
                                                                       "Bitte überprüfen Sie dies.")
                                        End Try

                                    End If


                                    ' Prüfung, ob die Phase cphase in Termine und Ressourcen übereinstimmt in relStart und relEnde
                                    Dim rightOneFound As Boolean = (anfang = cphase.relStart And ende = cphase.relEnde)
                                    Dim tmpIX As Integer = 1

                                    If phaseIndex.Length > 1 Then
                                        While Not rightOneFound And tmpIX <= phaseIndex.Length - 1
                                            cphase = hproj.getPhase(phaseIndex(tmpIX))
                                            rightOneFound = (anfang = cphase.relStart And ende = cphase.relEnde)
                                            tmpIX = tmpIX + 1
                                        End While
                                    End If


                                    If Not rightOneFound Then

                                        'Call MsgBox("Fehler beim Lesen der Ressourcen: die Dauer der Phase " & cphase.name & "' ist fehlerhaft")
                                        Throw New ArgumentException("Fehler beim Lesen der Ressourcen: die Dauer der Phase '" & cphase.name & "' ist fehlerhaft")

                                    End If


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
                                            If RoleDefinitions.containsName(hname) Then
                                                Try
                                                    r = CInt(RoleDefinitions.getRoledef(hname).UID)


                                                    ''ur:12.10.2015: Eingabe einer Summe in Ressourcen nun möglich, 
                                                    Try
                                                        summe = CDbl(zelle.Offset(0, 1 + ressSumOffset).Value)
                                                    Catch ex As Exception
                                                        summe = -1
                                                    End Try

                                                    If summe > 0.0 Then    ' Verteilung der Summe auf die Monate über Dauer der Phase

                                                        ReDim oldXwerte(0)
                                                        oldXwerte(0) = summe

                                                        With cphase

                                                            anfang = .relStart
                                                            ende = .relEnde
                                                            ReDim Xwerte(ende - anfang)

                                                            .berechneBedarfe(.getStartDate, .getEndDate, oldXwerte, 1, Xwerte)
                                                        End With

                                                        ''ur:12.10.2015:  eingefügt

                                                    Else

                                                        '  Anfang Check , ob richtige Kästchen Werte enthalten
                                                        Dim msgstr As String = " Fehler bei der Verteilung benötigter Kapazitäten" & vbCrLf & "für Rolle " & hname & " in Spalte "
                                                        Dim checkok As Boolean = True

                                                        Dim i As Integer
                                                        For i = 1 To hproj.anzahlRasterElemente

                                                            Dim wertvorhanden As Boolean = (CDbl(zelle.Offset(0, i + ressOff).Value) <> 0.0)
                                                            If (i < anfang Or i > ende) And wertvorhanden Then
                                                                msgstr = msgstr & " ," & i
                                                                checkok = False
                                                            End If

                                                        Next
                                                        If Not checkok Then
                                                            Call logfileSchreiben(msgstr, hproj.name, anzFehler)
                                                            'Call MsgBox(msgstr)
                                                            'Throw New ArgumentException(msgstr)
                                                        End If
                                                        ' Ende Check

                                                        ReDim Xwerte(ende - anfang)

                                                        Dim m As Integer
                                                        For m = anfang To ende

                                                            Try
                                                                Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + ressOff).Value)
                                                            Catch ex As Exception
                                                                Xwerte(m - anfang) = 0.0
                                                            End Try

                                                        Next m


                                                    End If

                                                    ' das muss doch eigentlich heissen: ende - anfang !? 
                                                    'crole = New clsRolle(ende - anfang + 1)
                                                    crole = New clsRolle(ende - anfang)
                                                    With crole
                                                        .uid = r
                                                        .Xwerte = Xwerte
                                                    End With

                                                    With cphase
                                                        .addRole(crole)
                                                    End With

                                                Catch ex As Exception
                                                    Throw New Exception(ex.Message)
                                                End Try

                                            ElseIf CostDefinitions.containsName(hname) Then

                                                Try

                                                    k = CInt(CostDefinitions.getCostdef(hname).UID)

                                                    ''ur:12.10.2015: Eingabe einer Summe in Ressourcen nun möglich, 
                                                    Try
                                                        summe = CDbl(zelle.Offset(0, 1 + ressSumOffset).Value)
                                                    Catch ex As Exception
                                                        summe = -1
                                                    End Try

                                                    If summe > 0.0 Then        'Summe wird verteilt auf Dauer der Phase

                                                        ReDim oldXwerte(0)
                                                        oldXwerte(0) = summe

                                                        With cphase

                                                            anfang = .relStart
                                                            ende = .relEnde
                                                            ReDim Xwerte(ende - anfang)

                                                            .berechneBedarfe(.getStartDate, .getEndDate, oldXwerte, 1, Xwerte)
                                                        End With

                                                    Else


                                                        ''ur:12.10.2015: 
                                                        '  Anfang Check , ob richtige Kästchen Werte enthalten
                                                        Dim msgstr As String = " Fehler bei der Verteilung benötigter Kapazitäten:" & vbCrLf & "für Kostenart " & hname & " in Spalte "
                                                        Dim checkok As Boolean = True

                                                        Dim i As Integer
                                                        For i = 1 To hproj.anzahlRasterElemente

                                                            Dim wertvorhanden As Boolean = (CDbl(zelle.Offset(0, i + ressOff).Value) <> 0.0)
                                                            If (i < anfang Or i > ende) And wertvorhanden Then
                                                                msgstr = msgstr & " ," & i
                                                                checkok = False
                                                            End If

                                                        Next
                                                        If Not checkok Then
                                                            Call logfileSchreiben(msgstr, hproj.name, anzFehler)
                                                            'Call MsgBox(msgstr)
                                                            'Throw New ArgumentException(msgstr)
                                                        End If
                                                        ' Ende Check

                                                        ReDim Xwerte(ende - anfang)
                                                        Dim m As Integer
                                                        For m = anfang To ende
                                                            Try
                                                                Xwerte(m - anfang) = CDbl(zelle.Offset(0, m + ressOff).Value)
                                                            Catch ex As Exception
                                                                Xwerte(m - anfang) = 0.0
                                                            End Try

                                                        Next m

                                                    End If

                                                    'ccost = New clsKostenart(ende - anfang + 1)
                                                    ccost = New clsKostenart(ende - anfang)
                                                    With ccost
                                                        .KostenTyp = k
                                                        .Xwerte = Xwerte
                                                    End With


                                                    With cphase
                                                        .AddCost(ccost)
                                                    End With

                                                Catch ex As Exception
                                                    Throw New Exception(ex.Message)
                                                End Try

                                            End If

                                        Case False  ' es wurde weder Phase noch Rolle angegeben. 
                                            If firsttime Then
                                                firsttime = False
                                            Else 'beim 2. mal:  ENDE von For-Schleife for each Zelle

                                                Exit For
                                            End If

                                    End Select

                            End Select

                        Next zelle


                    End With
                Catch ex As Exception
                    Call logfileSchreiben("Fehler in awinImportProjectmitHrchy, Lesen Ressourcen: " & ex.Message, hproj.name, anzFehler)
                    Throw New ArgumentException("Fehler in awinImportProjectmitHrchy, Lesen Ressourcen von '" & hproj.name & "' " & vbLf & ex.Message)
                End Try

            End If

            ' ------------------------------------------------------------------
            '   Ende Einlesen Ressourcen
            ' -------------------------------------------------------------------

        Catch ex As Exception
            Call logfileSchreiben("Fehler in awinImportProjectmitHrchy " & ex.Message, hproj.name, anzFehler)
            Throw New ArgumentException("Fehler in awinImportProjectmitHrchy '" & hproj.name & "' " & vbLf & ex.Message)
        End Try

        ' da Ampelfarbe , Beschreibung jetzt in Phase ist, muss das hier , nach Einlesen der Phasen
        hproj.ampelStatus = projektAmpelFarbe
        hproj.ampelErlaeuterung = projektAmpelText


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
            projVorlage.hierarchy = hproj.hierarchy
            hprojTemp = projVorlage

        Else
            hprojekt = hproj
        End If

    End Sub


    ''' <summary>
    ''' prüft, ob es sich beim Namen/der Email um eine bekannte Email Adresse handelt. 
    ''' Bei Erfolg wird das Werte-Paar in mappingNameID eingetragen
    ''' </summary>
    ''' <param name="userName"></param>
    ''' <param name="roleType"></param>
    ''' <param name="specifics"></param>
    ''' <returns></returns>
    Public Function isValidCustomUserRole(ByVal userName As String,
                                          ByVal roleType As ptCustomUserRoles,
                                          ByRef specifics As String) As Boolean

        Dim stillOk As Boolean = True
        Dim specificsWithIDs As String = ""
        Dim tmpNameUID As String = ""

        Try
            If userName.Length > 0 And userName.Contains("@") And userName.Contains(".") Then
                If roleType = ptCustomUserRoles.RessourceManager Then
                    If RoleDefinitions.containsName(specifics) Then
                        ' alles ok
                        stillOk = True
                        Dim teamID As Integer = -1
                        specificsWithIDs = CStr(RoleDefinitions.getRoledef(specifics).UID)
                    Else
                        stillOk = False
                    End If
                ElseIf roleType = ptCustomUserRoles.PortfolioManager Then
                    Dim tmpStr() As String = specifics.Split(New Char() {CChar(";")})

                    For Each tmpName As String In tmpStr

                        stillOk = stillOk And RoleDefinitions.containsName(tmpName.Trim)

                        If RoleDefinitions.containsName(tmpName.Trim) Then
                            tmpNameUID = CStr(RoleDefinitions.getRoledef(tmpName.Trim).UID)
                            If specificsWithIDs = "" Then
                                specificsWithIDs = tmpNameUID
                            Else
                                specificsWithIDs = specificsWithIDs & ";" & tmpNameUID
                            End If

                        Else
                            Call MsgBox("unbekannte Orga-Einheit: " & tmpName.Trim)
                        End If
                    Next
                End If
            Else
                stillOk = False
            End If

        Catch ex As Exception
            stillOk = False
        End Try

        specifics = specificsWithIDs
        isValidCustomUserRole = stillOk
    End Function


    ''' <summary>
    ''' erzeugt die Projekte, die in der Batch-Datei angegeben sind
    ''' stellt sie in ImportProjekte 
    ''' erstellt ein Szenario mit Namen der Batch-Datei; die Sortierung erfolgt über die Reihenfolge in der Batch-Datei 
    ''' das wird sichergestellt über Eintrag der tfzeile in hproj ... 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinImportProjektInventur()
        Dim zeile As Integer, spalte As Integer

        Dim tfZeile As Integer = 2
        Dim listOfpNames As New SortedList(Of String, String)
        Dim pName As String = ""
        Dim variantName As String = ""
        Dim vorlageName As String = ""
        Dim start As Date, inputStart As Date
        Dim startElem As String = ""
        Dim endElem As String = ""
        Dim ende As Date, inputEnde As Date
        Dim budget As Double
        Dim budgetInput As String = ""
        Dim dauer As Integer = 0
        Dim sfit As Double, risk As Double
        Dim capacityNeeded As String = ""
        Dim externCostInput As String = ""

        Dim description As String = ""
        Dim businessUnit As String = ""
        Dim createdProjects As Integer = 0
        Dim responsiblePerson As String = ""
        Dim custFields As New Collection
        ' wieviele Spalten müssen mindesten drin sein ... also was ist der standard 
        Dim nrOfStdColumns As Integer = 15

        Dim lastRow As Integer
        Dim lastColumn As Integer
        'Dim startSpalte As Integer
        Dim vglName As String = ""
        Dim hproj As clsProjekt
        Dim vproj As clsProjektvorlage
        Dim geleseneProjekte As Integer
        Dim ProjektdauerIndays As Integer = 0
        Dim ok As Boolean = False
        Dim refDauer As Double
        Dim vorgabeDauer As Double
        Dim abstandAnfang As Double
        Dim abstandEnde As Double
        Dim lastSpaltenValue As Integer

        Dim dauerFaktor As Double = 1.0
        Dim refProj As New clsProjekt

        Dim firstZeile As Excel.Range
        ' Änderung tk 5.6.16 wird jetzt an der Aufruf Schnittstelle gemacht 
        ''Dim scenarioName As String = appInstance.ActiveWorkbook.Name
        ''Dim tmpName As String = ""

        ' ''Dim namesForConstellation As New Collection
        ' '' bestimme den Namen des Szenarios - das ist gleich der Name der Excel Datei 
        ''Dim positionIX As Integer = scenarioName.IndexOf(".xls") - 1
        ''tmpName = ""
        ''For ih As Integer = 0 To positionIX
        ''    tmpName = tmpName & scenarioName.Chars(ih)
        ''Next
        ''scenarioName = tmpName.Trim

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1
        geleseneProjekte = 0

        ' später, um mal das Einlesen einigermaßen intelligent zu machen .... 
        'Dim suchstr(1) As String
        'suchstr(ptInventurSpalten.Name) = "Name"
        'suchstr(ptInventurSpalten.Vorlage) = "Vorlage"
        'suchstr(ptInventurSpalten.Start) = "Start-Datum"
        'suchstr(ptInventurSpalten.Ende) = "Ende-Datum"
        'suchstr(ptInventurSpalten.startElement) = "Bezug Start"
        'suchstr(ptInventurSpalten.endElement) = "Bezug Ende"
        'suchstr(ptInventurSpalten.Dauer) = "Dauer [Tage]"
        'suchstr(ptInventurSpalten.Budget) = "Budget [T€]"
        'suchstr(ptInventurSpalten.Risiko) = "Risiko"
        'suchstr(ptInventurSpalten.Strategie) = "Strategie"
        'suchstr(ptInventurSpalten.Kapazitaet) = "benötigte Kapazität"
        'suchstr(ptInventurSpalten.Businessunit) = "Business Unit"
        'suchstr(ptInventurSpalten.Beschreibung) = "Beschreibung"


        'Dim inputColumns(11) As Integer



        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Liste"),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)

                ' für später ... siehe oben, intelligent ...
                '' jetzt werden die Spalten bestimmt 
                'Try
                '    For i As Integer = 0 To 13
                '        inputColumns(i) = firstZeile.Find(What:=suchstr(i)).Column
                '    Next
                'Catch ex As Exception

                'End Try

                'lastColumn = firstZeile.End(XlDirection.xlToLeft).Column
                lastColumn = firstZeile.Columns.Count
                lastColumn = CType(firstZeile, Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastColumn = CType(.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                While zeile <= lastRow
                    ok = False
                    Dim sMilestone As clsMeilenstein = Nothing
                    Dim eMilestone As clsMeilenstein = Nothing
                    ' Kommentare zurücksetzen ...
                    Try
                        CType(.Range(.Cells(zeile, 1), .Cells(zeile, lastColumn)), Global.Microsoft.Office.Interop.Excel.Range).ClearComments()
                    Catch ex As Exception

                    End Try

                    ' hier muss jetzt alles zurückgesetzt werden 
                    ' ansonsten könnten alte Werte übernommen werden aus der Projekt-Information von vorher ..
                    pName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    If IsNothing(pName) Then
                        CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name fehlt ..")
                    ElseIf pName.Trim.Length < 2 Then

                        Try
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name muss mindestens 2 Buchstaben haben und eindeutig sein ..")
                        Catch ex As Exception

                        End Try

                    ElseIf Not isValidProjectName(pName) Then
                        Try
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Name darf keine #, (, ), Zeilenumbrüche enthalten ..")
                        Catch ex As Exception

                        End Try
                    Else
                        variantName = ""
                        custFields.Clear()
                        capacityNeeded = ""

                        ' falls ein Varianten-Name mit angegeben wurde: pname#variantNAme 
                        Try
                            Dim tmpStr() As String = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value).Split(New Char() {CChar("#")}, 2)
                            If tmpStr.Length > 1 Then
                                pName = makeValidProjectName(tmpStr(0))
                                variantName = tmpStr(1).Trim
                            End If
                        Catch ex As Exception
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            variantName = ""
                        End Try

                        vorlageName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        lastSpaltenValue = spalte + 1

                        If IsNothing(vorlageName) Then
                            CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        ElseIf vorlageName.Trim.Length = 0 Then
                            CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        Else
                            If Projektvorlagen.Liste.ContainsKey(vorlageName) Then

                                vproj = Projektvorlagen.getProject(vorlageName)
                                refProj = New clsProjekt
                                vproj.copyTo(refProj)
                                refProj.startDate = Date.Now

                                Try

                                    lastSpaltenValue = spalte + 2
                                    responsiblePerson = CStr(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    lastSpaltenValue = spalte + 3
                                    start = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    ' eines der beiden Daten Start bzw Ende darf ohne Angabe bleiben ...
                                    'If start < StartofCalendar Then
                                    '    Throw New ArgumentException("Datum vor Kalender-Start")
                                    'End If

                                    lastSpaltenValue = spalte + 4
                                    ende = CDate(CType(.Cells(zeile, spalte + 4), Global.Microsoft.Office.Interop.Excel.Range).Value)


                                    If start < StartofCalendar And ende < StartofCalendar Then
                                        Throw New ArgumentException("sowohl Start wie Ende-Datum liegen vor dem Kalender-Start")
                                    End If

                                    lastSpaltenValue = spalte + 5
                                    startElem = CStr(CType(.Cells(zeile, spalte + 5), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    lastSpaltenValue = spalte + 6
                                    endElem = CStr(CType(.Cells(zeile, spalte + 6), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    lastSpaltenValue = spalte + 7
                                    dauer = CInt(CType(.Cells(zeile, spalte + 7), Global.Microsoft.Office.Interop.Excel.Range).Value)

                                    ' Konsistenzprüfung 
                                    If start > StartofCalendar And ende > StartofCalendar And dauer > 0 Then
                                        Throw New ArgumentException("Überbestimmt: es kann nicht Start, Ende und Dauer angegeben werden .. ")
                                    End If

                                    lastSpaltenValue = spalte + 8
                                    budgetInput = CStr(CType(.Cells(zeile, spalte + 8), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If budgetInput <> "calcNeeded" And IsNumeric(budgetInput) Then
                                        budget = CDbl(CType(.Cells(zeile, spalte + 8), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                        If budget < 0 Then
                                            Throw New ArgumentException("negative Werte nicht zugelassen!")
                                        End If
                                    ElseIf budgetInput = "calcNeeded" Then
                                        ' das bedeutet, dass das Budget errechnet werden soll ... 
                                        budget = -999
                                    ElseIf budgetInput = "" Then
                                        budget = 0
                                    Else
                                        Throw New ArgumentException("mit dieser Angabe konnte nichts angefangen werden ...")
                                    End If


                                    lastSpaltenValue = spalte + 9
                                    capacityNeeded = CStr(CType(.Cells(zeile, spalte + 9), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If Not isValidRoleCostInput(capacityNeeded, True) Then
                                        Throw New ArgumentException("ungültige Kapa-Angabe")
                                    End If

                                    lastSpaltenValue = spalte + 10
                                    externCostInput = CStr(CType(.Cells(zeile, spalte + 10), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If Not isValidRoleCostInput(externCostInput, False) Then
                                        Throw New ArgumentException("ungültige Kosten-Angabe")
                                    End If

                                    ' Konsistenzprüfung ...
                                    ' es darf nicht sein, dass Budget und externe Kosten berechnet werden sollen ...
                                    If budget = -999 And externCostInput = "filltobudget" Then
                                        Throw New ArgumentException("unterbestimmt: es können nicht sowohl Budget als auch externe Kosten berechnet werden")
                                    End If

                                    lastSpaltenValue = spalte + 11
                                    risk = CDbl(CType(.Cells(zeile, spalte + 11), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If risk < 0 Or risk > 10.0 Then
                                        Throw New ArgumentException("Kennzahl muss zwischen [0 und 10] liegen")
                                    End If

                                    lastSpaltenValue = spalte + 12
                                    sfit = CDbl(CType(.Cells(zeile, spalte + 12), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If sfit < 0 Or risk > 10.0 Then
                                        Throw New ArgumentException("Kennzahl muss zwischen [0 und 10] liegen")
                                    End If


                                    lastSpaltenValue = spalte + 13
                                    businessUnit = CStr(CType(.Cells(zeile, spalte + 13), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If Not IsNothing(businessUnit) Then
                                        Dim bi As Integer = 0
                                        Dim found As Boolean = False
                                        While bi <= businessUnitDefinitions.Count - 1 And Not found
                                            If businessUnitDefinitions.ElementAt(bi).Value.name = businessUnit Then
                                                found = True
                                            Else
                                                bi = bi + 1
                                            End If
                                        End While

                                        If Not found Then
                                            Throw New ArgumentException("Business Unit unbekannt ..")
                                        End If
                                    End If


                                    lastSpaltenValue = spalte + 14
                                    description = CStr(CType(.Cells(zeile, spalte + 14), Global.Microsoft.Office.Interop.Excel.Range).Value)


                                    If lastColumn > nrOfStdColumns Then
                                        ' es gibt evtl Custom fields 
                                        Dim arrayOfSpalten() As Integer
                                        ReDim arrayOfSpalten(lastColumn - 1 - nrOfStdColumns)

                                        For i As Integer = nrOfStdColumns To lastColumn - 1
                                            arrayOfSpalten(i - nrOfStdColumns) = i + spalte ' spalte = immer 1
                                        Next

                                        Call readCustomFieldsFromExcel(arrayOfSpalten, 1, zeile, activeWSListe)

                                    End If

                                    vglName = calcProjektKey(pName.Trim, variantName)
                                    inputStart = start
                                    inputEnde = ende

                                    If DateDiff(DateInterval.Day, StartofCalendar, start) >= 0 Then

                                        If DateDiff(DateInterval.Day, start, ende) > 0 Then
                                            ' nichts tun , Ende-Datum ist ein gültiges Datum
                                            ok = True
                                        ElseIf DateDiff(DateInterval.Day, StartofCalendar, ende) >= 0 Then
                                            ' auch Ende ist ein gültiges Datum , liegt nur vor Start
                                            ' also vertauschen der beiden 
                                            Dim tmpDate As Date = ende
                                            ende = start
                                            start = tmpDate
                                            ok = True
                                        Else
                                            ' Ende Datum wird anhand der Laufzeit der Vorlage oder der Dauer berechnet
                                            If dauer > 0 Then
                                                ProjektdauerIndays = dauer
                                            Else
                                                ProjektdauerIndays = vproj.dauerInDays
                                            End If
                                            ende = calcDatum(start, ProjektdauerIndays)
                                            ok = True
                                        End If

                                    ElseIf DateDiff(DateInterval.Day, StartofCalendar, ende) >= 0 Then
                                        ' hier ist Start kein gültiges Datum innerhalb der Projekt-Tafel 
                                        ' Start Datum wird anhand der Laufzeit der Vorlage berechnet
                                        If dauer > 0 Then
                                            ProjektdauerIndays = -1 * dauer
                                        Else
                                            ProjektdauerIndays = -1 * vproj.dauerInDays
                                        End If

                                        start = calcDatum(ende, ProjektdauerIndays)

                                        If DateDiff(DateInterval.Day, StartofCalendar, start) >= 0 Then
                                            ' Start ist ein korrektes Datum 
                                            ok = True
                                        Else
                                            CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "Start liegt vor Kalender-Start "
                                            ok = False
                                        End If

                                    Else
                                        CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = "ungültiges Start- und Ende-Datum"
                                        ok = False
                                    End If

                                Catch ex As Exception

                                    ok = False
                                    'Call MsgBox(ex.Message)
                                    CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                    CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:=ex.Message)
                                End Try

                                ' jetzt die Daten richtig berechnen, falls Bezug Start , Bezug Ende angegeben ist 

                                vorgabeDauer = calcDauerIndays(start, ende)
                                Try

                                    If Not IsNothing(startElem) Then
                                        If startElem.Trim.Length > 0 Then
                                            sMilestone = refProj.getMilestone(startElem)
                                        End If
                                    End If

                                    If Not IsNothing(endElem) Then
                                        If endElem.Trim.Length > 0 Then
                                            eMilestone = refProj.getMilestone(endElem)
                                        End If
                                    End If

                                    ' jetzt werden Start und Ende ggf neu bestimmt, so dass die Bezugs-Elemente genau so liegen 
                                    If Not IsNothing(sMilestone) Then
                                        abstandAnfang = DateDiff(DateInterval.Day, refProj.startDate, sMilestone.getDate) * -1
                                        If Not IsNothing(eMilestone) Then
                                            abstandEnde = DateDiff(DateInterval.Day, eMilestone.getDate, refProj.endeDate)
                                            refDauer = calcDauerIndays(sMilestone.getDate, eMilestone.getDate)
                                        Else
                                            refDauer = calcDauerIndays(sMilestone.getDate, refProj.endeDate)
                                        End If
                                    Else
                                        If Not IsNothing(eMilestone) Then
                                            abstandEnde = DateDiff(DateInterval.Day, eMilestone.getDate, refProj.endeDate)
                                            refDauer = calcDauerIndays(refProj.startDate, eMilestone.getDate)
                                        Else
                                            refDauer = vorgabeDauer
                                        End If
                                    End If

                                    If refDauer < 0 Then
                                        refDauer = -1 * refDauer
                                    ElseIf refDauer = 0 Then
                                        refDauer = vorgabeDauer
                                    End If

                                    dauerFaktor = vorgabeDauer / refDauer

                                    ' rechne den neuen Start aus 
                                    If Not IsNothing(sMilestone) Then
                                        start = start.AddDays(CInt(dauerFaktor * abstandAnfang))
                                        ende = start.AddDays(CInt(dauerFaktor * vproj.dauerInDays - 1))
                                    ElseIf Not IsNothing(eMilestone) Then
                                        ende = start.AddDays(CInt(dauerFaktor * vproj.dauerInDays - 1))
                                    End If

                                Catch ex As Exception
                                    ' nichts tn 
                                End Try


                            Else
                                'CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value = ".?."
                                CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                ok = False
                            End If

                            ' jetzt die Aktion durchführen, wenn alles ok 
                            If ok Then

                                'Projekt anlegen ,Verschiebung um 
                                hproj = New clsProjekt(start, start.AddMonths(-1), start.AddMonths(1))

                                ' #####################################################################
                                ' Erstellen des Projekts nach den Angaben aus der Batch-Datei 
                                '
                                hproj = erstelleInventurProjekt(pName, vorlageName, variantName,
                                                             start, ende, budget, zeile, sfit, risk,
                                                             capacityNeeded, externCostInput, businessUnit, description, custFields,
                                                             responsiblePerson, 0.0)


                                If Not IsNothing(hproj) Then

                                    ' ein neu angelegtes Projekt bekommt immer den Status geplant ... 

                                    'prüfen ob Rundungsfehler bei Setzen Meilenstein passiert sind ... 
                                    If Not IsNothing(sMilestone) Then
                                        If DateDiff(DateInterval.Day, hproj.getMilestone(startElem).getDate, inputStart) <> 0 Then
                                            'Call MsgBox("Differenz Start:" & DateDiff(DateInterval.Day, hproj.getMilestone(startElem).getDate, inputStart))
                                            hproj.getMilestone(startElem).setDate = inputStart
                                        End If
                                    End If

                                    If Not IsNothing(eMilestone) Then
                                        If DateDiff(DateInterval.Day, hproj.getMilestone(endElem).getDate, inputEnde) <> 0 Then
                                            'Call MsgBox("Differenz Ende:" & DateDiff(DateInterval.Day, hproj.getMilestone(endElem).getDate, inputEnde))
                                            hproj.getMilestone(endElem).setDate = inputEnde
                                        End If
                                    End If

                                Else
                                    ok = False
                                    CType(.Range(.Cells(zeile, 1), .Cells(zeile, 15)), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                    CType(.Cells(zeile, lastSpaltenValue), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt konnte nicht erzeugt werden ...")
                                End If


                                If ok Then ' wenn es nicht explizit auf false gesetzt wurde, ist es an dieser Stelle immer noch true 
                                    Dim pkey As String = ""
                                    If Not IsNothing(hproj) Then
                                        Try
                                            pkey = calcProjektKey(hproj)

                                            If ImportProjekte.Containskey(pkey) Then
                                                CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                                CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Name existiert bereits")
                                            Else

                                                createdProjects = createdProjects + 1
                                                ' jetzt in die Liste der 
                                                If Not listOfpNames.ContainsValue(hproj.name) Then
                                                    hproj.tfZeile = tfZeile
                                                    Dim tmpKey As String = calcSortKeyCustomTF(tfZeile)
                                                    listOfpNames.Add(tmpKey, hproj.name)
                                                    tfZeile = tfZeile + 1
                                                Else
                                                    hproj.tfZeile = CInt(listOfpNames.ElementAt(listOfpNames.IndexOfValue(hproj.name)).Key)
                                                End If

                                                ImportProjekte.Add(hproj, False)
                                            End If


                                            'myCollection.Add(calcProjektKey(hproj))
                                        Catch ex As Exception
                                            CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                            CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:=ex.Message)
                                        End Try

                                    End If

                                End If

                            End If
                        End If


                    End If


                    geleseneProjekte = geleseneProjekte + 1
                    zeile = zeile + 1

                End While


            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Portfolio-Datei" & ex.Message)

        End Try


        Call MsgBox("gelesen: " & geleseneProjekte & vbLf &
                    "erzeugt: " & createdProjects & vbLf &
                    "importiert: " & ImportProjekte.Count)

    End Sub


    ''' <summary>
    ''' liest aus einer Excel-Tabelle die ggf vorhandenen CustomFields, der arrayOfSpalten gibt an, welche Spalten ausgelesen werden müssen 
    ''' </summary>
    ''' <param name="arrayOfSpalten">Indices der Spalten, müssen nicht zusammenhängend sein</param>
    ''' <param name="Headerzeile"><include file='welcher Zeile steht die Überschrift ' path='[@name=""]'/></param>
    ''' <param name="curZeile">welche Zeile der Tabelle soll gerade ausgelesen werden </param>
    ''' <param name="currentWS">das Excel.Worksheet, das die Tabelle enthält</param>
    ''' <returns></returns>
    Private Function readCustomFieldsFromExcel(ByVal arrayOfSpalten() As Integer,
                                               ByVal Headerzeile As Integer, ByVal curzeile As Integer,
                                               ByVal currentWS As Excel.Worksheet) As Collection

        Dim custFields As New Collection

        If Not IsNothing(arrayOfSpalten) Then

            For i As Integer = 0 To arrayOfSpalten.Length - 1

                Dim spalte As Integer = arrayOfSpalten(i)

                With currentWS
                    Try
                        Dim cfName As String = CStr(CType(.Cells(Headerzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        Dim uniqueID As Integer = customFieldDefinitions.getUid(cfName)

                        If uniqueID > 0 Then
                            ' es ist eine Custom Field

                            Dim cfType As Integer = customFieldDefinitions.getTyp(uniqueID)
                            Dim cfValue As Object = Nothing
                            Dim tstStr As String

                            Select Case cfType
                                Case ptCustomFields.Str

                                    cfValue = CStr(CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                Case ptCustomFields.Dbl

                                    cfValue = CDbl(CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                Case ptCustomFields.bool

                                    cfValue = CBool(CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            End Select

                            Dim cfObj As New clsCustomField
                            With cfObj
                                .uid = uniqueID
                                .wert = cfValue
                                tstStr = CStr(.wert)
                            End With
                            custFields.Add(cfObj)
                        End If
                    Catch ex As Exception
                        CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(curzeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:=ex.Message)
                    End Try
                End With


            Next

        End If

        readCustomFieldsFromExcel = custFields

    End Function


    ''' <summary>
    ''' liest alle in der Massen-Edit referenzierten Projekte ein und ersetzt die Werte dafür  
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub importiereMassenEdit()

        Dim err As New clsErrorCodeMsg

        Dim projectName As String = ""
        Dim variantName As String = ""

        Dim phaseName As String = ""
        Dim phaseNameID As String = ""
        Dim rcName As String = ""

        Dim isRole As Boolean = False
        Dim isCost As Boolean = False

        Dim zeile As Integer = 2
        Dim spalte As Integer = 1
        Dim lastRow As Integer

        Dim startColumnData As Integer
        Dim endColumnData As Integer

        Dim tmpValues() As Double = Nothing

        Dim von As Integer, bis As Integer
        Dim vonDate As Date, bisDate As Date
        Dim ok As Boolean = False
        Dim hproj As clsProjekt = Nothing
        Dim vproj As clsProjekt = Nothing

        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("VISBO"),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

            With activeWSListe

                startColumnData = CType(.Range("StartData"), Excel.Range).Column
                endColumnData = CType(.Range("EndData"), Excel.Range).Column

                vonDate = CType(CType(.Range("StartData"), Excel.Range).Value, Date)
                bisDate = CType(CType(.Range("EndData"), Excel.Range).Value, Date)

                von = getColumnOfDate(vonDate)
                bis = getColumnOfDate(bisDate)

                ' jetzt die TimeZone markieren , ohne die sonstigen Konsequenzen .. 
                ' überlegen, ob hier nicht awinchangeTimeSpan aufgerufen werden sollte ...

                Call awinShowtimezone(von, bis, True)
                showRangeLeft = von
                showRangeRight = bis

                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                ' jetzt Zeile für Zeile auslesen 
                While zeile <= lastRow

                    Dim valuesDidChange As Boolean = False

                    Try

                    Catch ex As Exception
                        ' dann ist irgendwo was schief gegangen ... 

                    End Try

                    ' die Farben in der Zeile zurücksetzen , aber nicht in den Datenbereichen, weil sonst die Info zu den Phasen weg ist 
                    CType(.Range(.Cells(zeile, 1), .Cells(zeile, startColumnData - 1)), Excel.Range).Interior.ColorIndex = XlColorIndex.xlColorIndexNone
                    Dim namesOK As Boolean = True

                    Try
                        projectName = CStr(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                    Catch ex As Exception
                        projectName = ""
                        namesOK = False
                    End Try

                    Try
                        variantName = CStr(CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                    Catch ex As Exception
                        variantName = ""
                    End Try

                    Try
                        phaseName = CStr(CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                    Catch ex As Exception
                        phaseName = ""
                        namesOK = False
                    End Try


                    Try
                        Dim cellComment As Excel.Comment = CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).Comment
                        If Not IsNothing(cellComment) Then
                            phaseNameID = cellComment.Text
                        Else
                            phaseNameID = calcHryElemKey(phaseName, False)
                        End If
                    Catch ex As Exception
                        phaseNameID = calcHryElemKey(phaseName, False)
                    End Try

                    Try
                        rcName = CStr(CType(.Cells(zeile, 5), Global.Microsoft.Office.Interop.Excel.Range).Value)
                    Catch ex As Exception
                        rcName = ""
                        namesOK = False
                    End Try

                    If namesOK Then

                        ok = False

                        Dim pKey As String = calcProjektKey(projectName, variantName)
                        If AlleProjekte.Containskey(pKey) Then
                            hproj = AlleProjekte.getProject(pKey)
                            ok = True
                        Else
                            ' in der Datenbank nachsehen und laden ... 
                            If Not noDB Then

                                '
                                ' prüfen, ob es in der Datenbank existiert ... wenn ja,  laden und anzeigen

                                If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                                    If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(projectName, variantName, Date.Now, err) Then

                                        ' Projekt ist noch nicht im Hauptspeicher geladen, es muss aus der Datenbank geholt werden.
                                        hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(projectName, variantName, Date.Now, err)
                                        ' jetzt in AlleProjekte eintragen ... 
                                        If Not IsNothing(hproj) Then
                                            AlleProjekte.Add(hproj)
                                            ok = True
                                        End If

                                    Else
                                        ' nicht in Session, nicht in Datenbank: nicht ok !
                                        ok = False
                                    End If
                                Else
                                    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Massen-Edit ..")
                                End If


                            Else
                                ' nicht in Session, keine Datenbank aktiv: nicht ok !
                                ok = False

                            End If


                        End If

                        If ok Then

                            If Not ImportProjekte.Containskey(pKey) Then
                                ImportProjekte.Add(hproj, False)
                            End If

                            ' hier kommt die eigentliche Behandlung , andernfalls Zeile rot einfärben ... 
                            ' hier ist das hproj gelesen 
                            ' jetzt prüfen, ob es die Phase gibt 
                            Dim cphase As clsPhase = hproj.getPhaseByID(phaseNameID)
                            If Not IsNothing(cphase) Then
                                ' es gibt die Phase

                                If RoleDefinitions.containsName(rcName) Then
                                    isRole = True
                                    isCost = False

                                ElseIf CostDefinitions.containsName(rcName) Then
                                    isCost = True
                                    isRole = False
                                Else
                                    isCost = False
                                    isRole = False
                                End If

                                ' jetzt werden die Werte ausgelesen ... 
                                ' die müssen an der Stelle ausgelesen werden, weil eine fehlende Rolle/kostenart nur angemeckert werden soll, 
                                ' wenn auch tmpValues.sum > 0 
                                ReDim tmpValues(bis - von)
                                Dim i As Integer

                                For i = 0 To bis - von

                                    Try
                                        tmpValues(i) = CDbl(CType(.Cells(zeile, startColumnData + 2 * i), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                        If tmpValues(i) < 0 Then
                                            tmpValues(i) = 0
                                            CType(.Cells(zeile, startColumnData + 2 * i), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                                        End If
                                    Catch ex As Exception
                                        CType(.Cells(zeile, startColumnData + 2 * i), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                                    End Try

                                Next


                                ' nur weitermachen, wenn es entweder eine gültige Rolle oder gültige Kostenart ist 
                                If isRole Or isCost Then


                                    If tmpValues.Sum > 0 Then

                                        Dim ixZeitraum As Integer, ix As Integer, anzLoops As Integer
                                        Call awinIntersectZeitraum(getColumnOfDate(cphase.getStartDate), getColumnOfDate(cphase.getEndDate),
                                                                   ixZeitraum, ix, anzLoops)

                                        If anzLoops > 0 Then
                                            ' es gibt eine Überdeckung
                                            If isRole Then
                                                Dim tmpRole As clsRolle = cphase.getRole(rcName)
                                                ' wenn die Rolle in diesem Projekt noch nicht da war, dann wird eine neue Instanz angelegt 
                                                Dim didntExist As Boolean = False

                                                If IsNothing(tmpRole) Then
                                                    didntExist = True
                                                    Dim dimension As Integer = cphase.relEnde - cphase.relStart
                                                    tmpRole = New clsRolle(dimension)

                                                    With tmpRole
                                                        .uid = RoleDefinitions.getRoledef(rcName).UID
                                                    End With
                                                End If

                                                Dim xWerte() As Double = tmpRole.Xwerte

                                                ' jetzt werden die Werte überschrieben ...
                                                For al As Integer = 1 To anzLoops
                                                    If xWerte(ix + al - 1) <> tmpValues(ixZeitraum + al - 1) Then
                                                        valuesDidChange = True
                                                    End If
                                                    xWerte(ix + al - 1) = tmpValues(ixZeitraum + al - 1)
                                                Next

                                                If didntExist Then
                                                    cphase.addRole(tmpRole)
                                                End If

                                            ElseIf isCost Then
                                                Dim tmpCost As clsKostenart = cphase.getCost(rcName)
                                                ' wenn die Kostenart in diesem Projekt noch nicht da war, dann wird eine neue Instanz angelegt 
                                                Dim didntExist As Boolean = False

                                                If IsNothing(tmpCost) Then
                                                    didntExist = True
                                                    Dim dimension As Integer = cphase.relEnde - cphase.relStart
                                                    tmpCost = New clsKostenart(dimension)

                                                    With tmpCost
                                                        .KostenTyp = CostDefinitions.getCostdef(rcName).UID
                                                    End With
                                                End If

                                                Dim xWerte() As Double = tmpCost.Xwerte

                                                ' jetzt werden die Werte überschrieben ...
                                                For al As Integer = 1 To anzLoops
                                                    If xWerte(ix + al - 1) <> tmpValues(ixZeitraum + al - 1) Then
                                                        valuesDidChange = True
                                                    End If
                                                    xWerte(ix + al - 1) = tmpValues(ixZeitraum + al - 1)
                                                Next

                                                If didntExist Then
                                                    cphase.AddCost(tmpCost)
                                                End If

                                            End If

                                        End If
                                    Else
                                        ' Löschen der Rolle bzw. Kostenart aus dieser Phase
                                        valuesDidChange = True
                                        If isRole Then
                                            Call cphase.removeRoleByName(rcName)
                                        ElseIf isCost Then
                                            Call cphase.removeCostByName(rcName)
                                        End If
                                    End If


                                Else
                                    ' es gibt die Rolle / Kostenart nicht 
                                    If tmpValues.Sum > 0 Then
                                        CType(.Cells(zeile, 5), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                                    Else
                                        ' keine Aktion notwendig 
                                    End If

                                End If

                            Else
                                ' es gibt die Phase nicht 
                                CType(.Cells(zeile, 4), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                            End If
                        Else
                            ' Projekt- Variante existiert nicht !
                            CType(.Range(.Cells(zeile, 2), .Cells(zeile, 3)), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelRot
                        End If

                    End If



                    If valuesDidChange Then
                        hproj.marker = True
                    End If

                    zeile = zeile + 1

                End While


            End With


        Catch ex As Exception
            Call MsgBox("Fehler beim Import der Massen-Edit Datei" & vbLf & ex.Message)
        End Try



    End Sub
    ''' <summary>
    ''' erzeugt eine Szenario Definition
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Function importScenarioDefinition(ByVal scenarioName As String) As clsConstellation

        Dim err As New clsErrorCodeMsg

        Dim zeile As Integer, spalte As Integer


        Dim tfZeile As Integer = 2
        Dim listOfpNames As New SortedList(Of String, String)
        Dim pName As String = ""
        Dim variantName As String = ""

        Dim lastRow As Integer
        Dim lastColumn As Integer
        'Dim startSpalte As Integer

        Dim geleseneProjekte As Integer


        Dim firstZeile As Excel.Range

        Dim newC As New clsConstellation
        newC.constellationName = scenarioName
        newC.sortCriteria = ptSortCriteria.customTF


        zeile = 2
        spalte = 1
        geleseneProjekte = 0




        Try
            Dim activeWSListe As Excel.Worksheet
            Try
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets("VISBO"),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End Try

            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)
                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                While zeile <= lastRow

                    ' Kommentare zurücksetzen ...
                    Try
                        CType(.Range(.Cells(zeile, 1), .Cells(zeile, lastColumn)), Global.Microsoft.Office.Interop.Excel.Range).ClearComments()
                    Catch ex As Exception

                    End Try

                    ' hier muss jetzt alles zurückgesetzt werden 
                    ' ansonsten könnten alte Werte übernommen werden aus der Projekt-Information von vorher ..
                    pName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)

                    If IsNothing(pName) Then
                        CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name fehlt ..")
                    ElseIf pName.Trim.Length < 2 Then

                        Try
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name muss mindestens 2 Buchstaben haben und eindeutig sein ..")
                        Catch ex As Exception

                        End Try


                    Else
                        variantName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        If IsNothing(variantName) Then
                            variantName = ""
                        End If


                        If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, variantName, Date.Now, err) Then
                            ' als Constellation Item aufnehmen 
                            Dim cItem As New clsConstellationItem

                            With cItem
                                .projectName = pName
                                .variantName = variantName
                                .show = True
                                .projectTyp = ptPRPFType.project.ToString
                                .zeile = zeile
                            End With

                            newC.add(cItem)

                        End If

                    End If

                    geleseneProjekte = geleseneProjekte + 1
                    zeile = zeile + 1

                End While


            End With
        Catch ex As Exception

            Throw New Exception("Fehler In Portfolio-Datei" & ex.Message)
        End Try



        Call MsgBox("gelesen: " & geleseneProjekte & vbLf &
                    "Portfolio erzeugt: " & scenarioName)

        importScenarioDefinition = newC

    End Function

    ''' <summary>
    ''' setzt die für den Allianz 1 Import Type notwendigen Felder
    ''' </summary>
    ''' <param name="PTroleNamesToConsider"></param>
    ''' <param name="PTcolRoleNamesToConsider"></param>
    ''' <param name="TEroleNamesToConsider"></param>
    ''' <param name="TEcolRoleNamesToConsider"></param>
    ''' <param name="currentWS"></param>
    ''' <param name="importTyp"></param>
    Private Sub setAllianzImportArrays(ByRef PTroleNamesToConsider() As String,
                                       ByRef PTcolRoleNamesToConsider() As Integer,
                                       ByRef TEroleNamesToConsider() As String,
                                       ByRef TEcolRoleNamesToConsider() As Integer,
                                       ByVal currentWS As Excel.Worksheet,
                                       ByVal importTyp As ptVisboImportTypen)

        Dim tmpRoleNames() As String
        Dim tmpColBz() As String
        Dim tmpCols() As Integer

        Dim tmpTEroleNames() As String
        Dim tmpTEcolBZ() As String
        Dim tmpTECols() As Integer

        Dim errRoles As String = ""
        Dim ok As Boolean = True
        Dim zeile As Integer = 2


        If importTyp = ptVisboImportTypen.allianzMassImport1 Then

            zeile = 2
            ' am besten hier aus awinsettings einlesen ...
            ' sowohl die PTRoleNames als auch die T€RoleNames 
            tmpRoleNames = {"D-BOSV-KB0", "D-BOSV-KB1", "D-BOSV-KB2", "D-BOSV-KB3", "D-BOSV-SBF1", "D-BOSV-SBF2", "DRUCK", "D-BOSV-SBP1", "D-BOSV-SBP2", "D-BOSV-SBP3", "AMIS",
                        "IT-BVG", "IT-KuV", "IT-PSQ", "A-IT04", "AZ Technology", "IT-SFK", "Op-DFS", "KaiserX IT"}

            tmpColBz = {"DB1", "DC1", "DD1", "DE1", "DG1", "DH1", "DI1", "DK1", "DL1", "DM1", "DN1", "DP1", "DQ1", "DR1", "DS1", "DT1", "DU1", "DV1", "DW1"}

            ReDim tmpCols(tmpRoleNames.Length - 1)

            tmpTEroleNames = Nothing
            tmpTEcolBZ = Nothing
            ReDim tmpTECols(0)

        Else
            zeile = 3
            tmpRoleNames = {"D-BITSV-KB0", "D-BITSV-KB1", "D-BITSV-KB2", "D-BITSV-KB3", "D-BITSV-SBF1", "D-BITSV-SBF2", "D-BITSV-SBF-DRUCK", "D-BITSV-SBP1", "D-BITSV-SBP2", "D-BITSV-SBP3", "AMIS"}
            tmpColBz = {"CP1", "CQ1", "CR1", "CS1", "CU1", "CV1", "CW1", "CY1", "CZ1", "DA1", "DB1"}

            ReDim tmpCols(tmpRoleNames.Length - 1)

            tmpTEroleNames = {"D-BITKuV", "D-BITLuA", "D-BITKIS", "D-BITEPM", "D-BIT-FMV", "D-IT-BVG", "D-BITKVI", "D-IT-PSQ", "A-IT04", "D-IT-AS", "AMOS", "KX BIT", "KX IT", "D-IT-ISM"}
            tmpTEcolBZ = {"AN1", "AP1", "AQ1", "AR1", "AS1", "AT1", "AU1", "AV1", "AW1", "AX1", "AY1", "AZ1", "BA1", "BB1"}

            ReDim tmpTECols(tmpTEroleNames.Length - 1)
        End If


        If (tmpRoleNames.Length <> tmpColBz.Length) Or (tmpTEroleNames.Length <> tmpTEcolBZ.Length) Then
            Throw New ArgumentException("ungleiche Anzahl Namen und Spalten-Ids")
        Else
            Dim tmpAnzahl As Integer = tmpRoleNames.Length

            ' Plausibilitätsprüfung: nur weitermachen, wenn auch alle Rollen in der RollenDefinition drin sind 

            For Each tmpRoleName As String In tmpRoleNames
                If RoleDefinitions.containsName(tmpRoleName) Then
                    ' ok 
                Else
                    errRoles = errRoles & tmpRoleName & "; "
                End If
            Next

            For Each tmpRoleName As String In tmpTEroleNames
                If RoleDefinitions.containsName(tmpRoleName) Then
                    ' ok 
                Else
                    errRoles = errRoles & tmpRoleName & "; "
                End If
            Next

            If errRoles.Length = 0 Then
                ' jetzt weiter machen .. die col
                With currentWS
                    For i As Integer = 1 To tmpAnzahl
                        tmpCols(i - 1) = CType(.Range(tmpColBz(i - 1)), Excel.Range).Column

                        ' test tk 9.6.18
                        If tmpRoleNames(i - 1).StartsWith("D-") Then

                            Dim tmpValue As String = CStr(CType(.Cells(zeile, tmpCols(i - 1)), Excel.Range).Value).Trim
                            Dim chkTxt As String = tmpRoleNames(i - 1).Trim.Substring(2)

                            ok = tmpValue.StartsWith(chkTxt) Or tmpRoleNames(i - 1) = "DRUCK"
                        Else
                            ok = ok And (CStr(CType(.Cells(zeile, tmpCols(i - 1)), Excel.Range).Value).StartsWith(tmpRoleNames(i - 1)) Or
                            tmpRoleNames(i - 1) = "DRUCK")
                        End If


                        If Not ok Then
                            Call MsgBox("Fehler in Spalte mit Angaben zu (?) " & tmpRoleNames(i - 1))
                            ok = True
                        End If
                    Next

                    tmpAnzahl = tmpTEroleNames.Length

                    For i As Integer = 1 To tmpAnzahl

                        tmpTECols(i - 1) = CType(.Range(tmpTEcolBZ(i - 1)), Excel.Range).Column

                        Dim tmpValue As String = CStr(CType(.Cells(zeile, tmpTECols(i - 1)), Excel.Range).Value).Trim

                        ok = tmpValue.StartsWith(tmpTEroleNames(i - 1))

                        If Not ok Then
                            Call MsgBox("Fehler in Spalte mit Angaben zu (?) " & tmpRoleNames(i - 1))
                            ok = True
                        End If
                    Next

                End With

            Else
                Throw New ArgumentException("nicht bekannte Rolle(n: " & errRoles)
            End If
        End If

        PTroleNamesToConsider = tmpRoleNames
        PTcolRoleNamesToConsider = tmpCols

        TEroleNamesToConsider = tmpTEroleNames
        TEcolRoleNamesToConsider = tmpTECols


    End Sub

    ''' <summary>
    ''' bestimmt den Import-Typ und das Worksheet, das eingelesen werden soll ..
    ''' </summary>
    ''' <param name="importType"></param>
    ''' <returns></returns>
    Private Function bestimmeWsAndImporttype(ByRef importType As ptVisboImportTypen) As Excel.Worksheet
        Dim resultWS As Excel.Worksheet = Nothing
        Dim wb As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim tmpImportType As Integer = -1

        Try
            tmpImportType = CInt(wb.Range(visboImportKennung).Value)

            If [Enum].IsDefined(GetType(ptVisboImportTypen), tmpImportType) Then
                resultWS = CType(CType(wb.Range(visboImportKennung), Excel.Range).Parent, Excel.Worksheet)
                importType = tmpImportType
            End If


        Catch ex As Exception
            resultWS = Nothing
        End Try

        bestimmeWsAndImporttype = resultWS

    End Function

    ''' <summary>
    ''' importiert alle Custom User Roles 
    ''' </summary>
    ''' <param name="outputCollection"></param>
    ''' <returns></returns>
    Public Function ImportCustomUserRoles(ByRef outputCollection As Collection) As clsCustomUserRoles

        Dim importedUserRoles As New clsCustomUserRoles
        Dim UserRoleSheet As Excel.Worksheet = CType(appInstance.ActiveSheet, Global.Microsoft.Office.Interop.Excel.Worksheet)

        Try
            Dim errMsg As String = ""
            Dim myRange As Excel.Range = UserRoleSheet.UsedRange
            Dim maxZeile As Integer = myRange.Rows.Count
            Dim curType As ptCustomUserRoles = ptCustomUserRoles.OrgaAdmin
            Dim emailAdresse As String = ""
            Dim userRole As String = ""
            Dim roleSpecifics As String = ""
            Dim saveSpecificsForErrMsg As String = ""

            For zeile As Integer = 2 To maxZeile

                Try

                    If Not IsNothing(CType(UserRoleSheet.Cells(zeile, 1), Excel.Range).Value) And Not IsNothing(CType(UserRoleSheet.Cells(zeile, 2), Excel.Range).Value) Then

                        emailAdresse = CStr(CType(UserRoleSheet.Cells(zeile, 1), Excel.Range).Value).Trim
                        userRole = CStr(CType(UserRoleSheet.Cells(zeile, 2), Excel.Range).Value).Trim
                        roleSpecifics = CStr(CType(UserRoleSheet.Cells(zeile, 3), Excel.Range).Value)

                        If Not IsNothing(roleSpecifics) Then
                            roleSpecifics = roleSpecifics.Trim
                        Else
                            roleSpecifics = ""
                        End If

                        saveSpecificsForErrMsg = roleSpecifics

                        Dim tmpstr() As String = userRole.Split(New Char() {CChar("-")})
                        curType = CType(tmpstr(0), ptCustomUserRoles)

                        If isValidCustomUserRole(emailAdresse, curType, roleSpecifics) Then
                            importedUserRoles.addCustomUserRole(emailAdresse, "", curType, roleSpecifics)
                        Else
                            errMsg = "Zeile " & zeile & "- Error: no valid Custom User Role: " & emailAdresse & "; " & userRole & "; " & saveSpecificsForErrMsg
                            outputCollection.Add(errMsg)
                            CType(UserRoleSheet.Cells(zeile, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed
                        End If
                    End If

                Catch ex As Exception
                    errMsg = "Zeile " & zeile & "- Error: no valid Custom User Role: " & emailAdresse & "; " & userRole & "; " & saveSpecificsForErrMsg
                    outputCollection.Add(errMsg)
                    CType(UserRoleSheet.Cells(zeile, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed
                End Try


            Next
        Catch ex As Exception

        End Try

        ImportCustomUserRoles = importedUserRoles
    End Function

    ''' <summary>
    ''' Voraussetzungen: das File ist geöffnet 
    ''' </summary>
    ''' <returns></returns>
    Public Function ImportOrganisation(ByRef outputCollection As Collection) As clsOrganisation

        Dim importedOrga As New clsOrganisation
        Dim orgaSheet As Excel.Worksheet = CType(appInstance.ActiveSheet, Global.Microsoft.Office.Interop.Excel.Worksheet)

        ' auslesen der Gültigkeit
        Dim validFrom As Date = Date.Now
        Try
            validFrom = CDate(CType(orgaSheet.Cells(1, 2), Excel.Range).Value)
        Catch ex As Exception

        End Try

        Dim oldOrga As clsOrganisation = Nothing

        If Not IsNothing(validOrganisations) Then
            If validOrganisations.count > 0 Then
                oldOrga = validOrganisations.getOrganisationValidAt(validFrom)
            End If
        End If


        ' Auslesen der Rollen Definitionen 
        Dim newRoleDefinitions As New clsRollen
        Call readRoleDefinitions(orgaSheet, newRoleDefinitions, outputCollection)

        If awinSettings.visboDebug Then
            Call MsgBox("readRoleDefinitions")
        End If

        ' Auslesen der Kosten Definitionen 
        Dim newCostDefinitions As New clsKostenarten
        Call readCostDefinitions(orgaSheet, newCostDefinitions, outputCollection)

        If awinSettings.visboDebug Then
            Call MsgBox("readCostDefinitions")
        End If

        ' und jetzt werden noch die Gruppen-Definitionen ausgelesen 
        Call readRoleDefinitions(orgaSheet, newRoleDefinitions, outputCollection, readingGroups:=True)

        ' jetzt kommen die Validierungen .. wenn etwas davon schief geht 
        If newRoleDefinitions.Count > 0 Then
            ' jetzt sind die Rollen alle aufgebaut und auch die Teams definiert 
            ' jetzt kommt der Validation-Check 

            Dim TeamsAreNotOK As Boolean = checkTeamDefinitions(newRoleDefinitions, outputCollection)
            Dim existingOverloads As Boolean = checkTeamMemberOverloads(newRoleDefinitions, outputCollection)

            If outputCollection.Count > 0 Then
                ' wird an der aurufenden Stelle ausgegeben ... 
            ElseIf TeamsAreNotOK Or existingOverloads Then
                ' darf eigentlich nicht vorkommen, weil man dann im oberen Zweig landen müsste ...
            Else
                'bis hier ist alles in Ordnung 
                With importedOrga
                    .allRoles = newRoleDefinitions
                    .allCosts = newCostDefinitions
                    .validFrom = validFrom
                End With

                If Not importedOrga.validityCheckWith(oldOrga, outputCollection) = True Then
                    ' wieder zurück setzen ..
                    importedOrga = New clsOrganisation
                Else

                End If
            End If

        End If

        ImportOrganisation = importedOrga
    End Function

    ''' <summary>
    ''' erzeugt die Projekte, die in der Batch-Datei angegeben sind
    ''' stellt sie in ImportProjekte 
    ''' erstellt ein Szenario mit Namen der Batch-Datei; die Sortierung erfolgt über die Reihenfolge in der Batch-Datei 
    ''' das wird sichergestellt über Eintrag der tfzeile in hproj ... 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub importAllianzType1(ByVal startdate As Date, ByVal endDate As Date)

        Dim zeile As Integer, spalte As Integer

        Dim importType As ptVisboImportTypen

        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""
        Dim custFields As New Collection
        Dim description As String = ""
        Dim responsiblePerson As String = ""
        Dim sFit As Double = 5.0
        Dim risk As Double = 5.0
        Dim budget As Double = 0.0
        Dim businessUnit As String = ""
        Dim allianzProjektNummer As String = ""
        Dim allianzStatus As String = ""
        Dim ampelText As String
        Dim projVorhabensBudget As Double = 0.0

        Dim logmsg() As String

        Dim programName As String = ""
        Dim current1program As clsConstellation = Nothing
        Dim last1Budget As Double = 0.0
        Dim lfdNr1program As Integer = 2

        ' nimmt die vollen Namen der 
        Dim fullNameListe1 As New SortedList(Of String, String)

        Dim createdProjects As Integer = 0
        Dim createdPrograms As Integer = 0
        Dim emptyPrograms As Integer = 0


        Dim vorlageName As String = "Rel"
        Dim lastRow As Integer
        Dim lastColumn As Integer
        Dim geleseneProjekte As Integer
        Dim ok As Boolean = False

        ' für den Output 
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""

        ' Standard-Definition
        Dim anzReleases As Integer = 5

        Try
            anzReleases = Projektvorlagen.getProject("Rel").CountPhases - 1
        Catch ex As Exception

        End Try


        ' enthält die prozentualen Anteile in den Releases 
        Dim relPrz() As Double
        ReDim relPrz(anzReleases - 1)

        ' Projekt-Eintrag/Zeile, die in der Excel Datei ignoriert werden soll 
        Dim nameTobeIgnored As String = "xxx"

        ' enthält die Phasen Namen
        Dim phNames() As String
        ReDim phNames(anzReleases - 1)

        ' enthält die Spalten-Nummer, ab der die Release Phasen Anteile stehen 
        Dim colRelPrzStart As Integer

        ' enthält die Info, welche Rollen-Namen berücksichtigt werden sollen 
        Dim roleNamesToConsider() As String = Nothing

        ' enthält die Spalten-Nummern, wo die einzelnen Rollen-Namen zu finden sind
        Dim colRoleNamesToConsider() As Integer = Nothing

        ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
        Dim roleNeeds() As Double = Nothing

        ' enthält die Info, welche Rollen-Namen berücksichtigt werden sollen 
        Dim TEroleNamesToConsider() As String = Nothing

        ' enthält die Spalten-Nummern, wo die einzelnen Rollen-Namen zu finden sind
        Dim colTEroleNamesToConsider() As Integer = Nothing

        ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
        Dim costNeeds() As Double = Nothing

        ' enthält die Spalten, wo die einzelnen Felder stehen , korreliert mit der Enum allianzSpalten
        Dim colFields() As Integer

        Dim firstZeile As Excel.Range

        Dim enumAllianzCount As Integer = [Enum].GetNames(GetType(allianzSpalten)).Length
        ReDim colFields(enumAllianzCount)



        spalte = 1
        geleseneProjekte = 0

        ' jetzt werden die Phase-Names besetzt
        Try
            For i = 1 To anzReleases
                phNames(i - 1) = Projektvorlagen.getProject(vorlageName).getPhase(i + 1).name
            Next
        Catch ex As Exception
            Call MsgBox("Probleme mit Vorlage " & vorlageName)
            Exit Sub
        End Try


        Try


            Dim currentWS As Excel.Worksheet = bestimmeWsAndImporttype(importType)

            If IsNothing(currentWS) Then
                Call MsgBox("Import File nicht erkannt - bitte " & visboImportKennung & "-Feld in Excel-Datei eintragen!")
            ElseIf (importType <> ptVisboImportTypen.allianzMassImport1 And importType <> ptVisboImportTypen.allianzMassImport2) Then
                Call MsgBox("keine Allianz-Projektliste: " & visboImportKennung & "muss Wert 5 oder 6 haben!")
                Exit Sub
            End If

            Dim isOldAllianzImport As Boolean = (importType = ptVisboImportTypen.allianzMassImport1)


            With currentWS



                ' jetzt wird festgelegt, ab wo die relativen Verteilungs-Werte für die Releases stehen 
                If isOldAllianzImport Then
                    colRelPrzStart = .Range("AI1").Column
                    firstZeile = CType(.Rows(2), Excel.Range)
                    zeile = 3
                Else
                    colRelPrzStart = .Range("AB1").Column
                    firstZeile = CType(.Rows(3), Excel.Range)
                    zeile = 5
                End If


                ' damit werden die Arrays besetzt, welche Rollen gesucht sind und in welchen Spalten die Angaben dazu zu finden sind ... 
                Call setAllianzImportArrays(roleNamesToConsider, colRoleNamesToConsider,
                                            TEroleNamesToConsider, colTEroleNamesToConsider,
                                            currentWS, importType)


                'lastColumn = firstZeile.End(XlDirection.xlToLeft).Column

                lastColumn = CType(.Cells(1, 3000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column

                If isOldAllianzImport Then
                    lastRow = CType(.Cells(5000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                Else
                    lastRow = CType(.Cells(5000, "A"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                End If



                ' um die CustomFields lesen zu können ... 
                Dim colCustomFields(3) As Integer

                If isOldAllianzImport Then
                    ' T-BWLA
                    colCustomFields(0) = CInt(CType(.Range("A1"), Excel.Range).Column)
                    ' PGML
                    colCustomFields(1) = CInt(CType(.Range("B1"), Excel.Range).Column)
                    ' POB
                    colCustomFields(2) = CInt(CType(.Range("C1"), Excel.Range).Column)
                    ' Key Cluster
                    colCustomFields(3) = CInt(CType(.Range("D1"), Excel.Range).Column)

                    ' jetzt die Spalten bestimmen, wo die Werte stehen
                    Try
                        colFields(allianzSpalten.Name) = CType(.Range("H1"), Excel.Range).Column
                        colFields(allianzSpalten.itemType) = CType(.Range("G1"), Excel.Range).Column
                        colFields(allianzSpalten.AmpelText) = CType(.Range("U1"), Excel.Range).Column
                        colFields(allianzSpalten.BusinessUnit) = CType(.Range("AB1"), Excel.Range).Column
                        colFields(allianzSpalten.Responsible) = CType(.Range("AC1"), Excel.Range).Column
                        colFields(allianzSpalten.Projektnummer) = CType(.Range("AD1"), Excel.Range).Column
                        colFields(allianzSpalten.Status) = CType(.Range("AF1"), Excel.Range).Column
                        colFields(allianzSpalten.Budget) = CType(.Range("M1"), Excel.Range).Column
                        colFields(allianzSpalten.pvBudget) = CType(.Range("N1"), Excel.Range).Column
                    Catch ex As Exception
                        Dim errmsg As String = "fehlerhafte Range Definition ..."
                        Throw New ArgumentException(errmsg)
                    End Try
                Else

                    ' T-BWLA
                    colCustomFields(0) = CInt(CType(.Range("D1"), Excel.Range).Column)
                    ' PGML
                    colCustomFields(1) = CInt(CType(.Range("E1"), Excel.Range).Column)
                    ' POB
                    colCustomFields(2) = CInt(CType(.Range("F1"), Excel.Range).Column)
                    ' Key Cluster
                    colCustomFields(3) = CInt(CType(.Range("H1"), Excel.Range).Column)

                    ' jetzt die Spalten bestimmen, wo die Werte stehen
                    Try
                        colFields(allianzSpalten.Name) = CType(.Range("J1"), Excel.Range).Column
                        colFields(allianzSpalten.itemType) = CType(.Range("A1"), Excel.Range).Column
                        colFields(allianzSpalten.AmpelText) = CType(.Range("L1"), Excel.Range).Column
                        colFields(allianzSpalten.BusinessUnit) = CType(.Range("Y1"), Excel.Range).Column
                        colFields(allianzSpalten.Responsible) = CType(.Range("Z1"), Excel.Range).Column
                        colFields(allianzSpalten.Projektnummer) = CType(.Range("K1"), Excel.Range).Column
                        colFields(allianzSpalten.Status) = CType(.Range("EA1"), Excel.Range).Column
                        colFields(allianzSpalten.Budget) = CType(.Range("M1"), Excel.Range).Column
                        colFields(allianzSpalten.pvBudget) = CType(.Range("O1"), Excel.Range).Column
                    Catch ex As Exception
                        Dim errmsg As String = "fehlerhafte Range Definition ..."
                        Throw New ArgumentException(errmsg)
                    End Try

                End If

                Dim realRoleNamesToConsider() As String = Nothing
                If isOldAllianzImport Then
                    realRoleNamesToConsider = roleNamesToConsider
                Else
                    ReDim realRoleNamesToConsider(roleNamesToConsider.Length + TEroleNamesToConsider.Length - 1)
                    For i As Integer = 0 To roleNamesToConsider.Length - 1
                        realRoleNamesToConsider(i) = roleNamesToConsider(i)
                    Next
                    Dim i_offset As Integer = roleNamesToConsider.Length

                    For i As Integer = 0 To TEroleNamesToConsider.Length - 1
                        realRoleNamesToConsider(i + i_offset) = TEroleNamesToConsider(i)
                    Next
                End If

                ' tk Test Logfile schreiben ...
                Try
                    ReDim logmsg(realRoleNamesToConsider.Count)
                    logmsg(0) = ""
                    For ix As Integer = 1 To realRoleNamesToConsider.Count
                        logmsg(ix) = realRoleNamesToConsider(ix - 1)
                    Next
                    Call logfileSchreiben(logmsg)
                Catch ex As Exception

                End Try

                ' jetzt die zugelassenen Werte für 
                Dim pgmlinie() As Integer
                Dim projektvorhaben() As Integer

                If isOldAllianzImport Then
                    ReDim pgmlinie(0)
                    ReDim projektvorhaben(0)
                    pgmlinie(0) = 1
                    projektvorhaben(0) = 4
                Else
                    ReDim pgmlinie(0)
                    ReDim projektvorhaben(2)

                    pgmlinie(0) = 2
                    projektvorhaben(0) = 5
                    projektvorhaben(1) = 6
                    projektvorhaben(2) = 7
                End If

                ' jetzt müssen die Dimensionen gesetzt werden 
                Dim tmpLen As Integer = roleNamesToConsider.Length

                If Not IsNothing(roleNamesToConsider) Then

                    If importType = ptVisboImportTypen.allianzMassImport2 Then

                        If Not IsNothing(TEroleNamesToConsider) Then
                            tmpLen = tmpLen + TEroleNamesToConsider.Length
                        End If

                    End If

                    ReDim roleNeeds(tmpLen - 1)

                ElseIf Not IsNothing(TEroleNamesToConsider) Then

                    tmpLen = TEroleNamesToConsider.Length
                    ReDim roleNeeds(tmpLen - 1)

                End If


                While zeile <= lastRow

                    ' Werte zurücksetzen ..
                    ReDim roleNeeds(tmpLen - 1)

                    ok = False

                    ' Kommentare zurücksetzen ...
                    Try
                        CType(.Range(.Cells(zeile, 1), .Cells(zeile, lastColumn)), Global.Microsoft.Office.Interop.Excel.Range).ClearComments()
                    Catch ex As Exception

                    End Try


                    ' lese den Projekt-Namen
                    Try
                        pName = CStr(CType(.Cells(zeile, colFields(allianzSpalten.Name)), Excel.Range).Value)
                        ok = True
                    Catch ex As Exception
                        pName = "?"
                    End Try


                    If IsNothing(pName) Then
                        CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name fehlt ..")

                    ElseIf pName.Trim = nameTobeIgnored Then
                        CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                        CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="wird ignoriert ..")

                    ElseIf pName.Trim.Length < 2 Then

                        Try
                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Projekt-Name muss mindestens 2 Buchstaben haben und eindeutig sein ..")
                        Catch ex As Exception

                        End Try

                    Else

                        Dim itemType As Integer

                        custFields.Clear()
                        description = pName

                        If Not isValidProjectName(pName) Then
                            pName = makeValidProjectName(pName)
                        End If

                        Try
                            ' weitere Informationen auslesen 

                            Try
                                itemType = CInt(CType(.Cells(zeile, colFields(allianzSpalten.itemType)), Excel.Range).Value)
                            Catch ex As Exception
                                itemType = 0
                            End Try

                            If projektvorhaben.Contains(itemType) Then
                                ' ok weitermachen
                                ok = True

                                Try
                                    projVorhabensBudget = CDbl(CType(.Cells(zeile, colFields(allianzSpalten.pvBudget)), Excel.Range).Value)
                                Catch ex As Exception
                                    projVorhabensBudget = 0.0
                                End Try

                            Else
                                ok = False
                                ' jetzt muss geschaut werden, ob es sich um eine 1-er Konstellation handelt, dann soll 
                                ' ein neues Portfolio aufgemacht werden .. 
                                If pgmlinie.Contains(itemType) Then
                                    ' die bisherige Constellation wegschreiben ...

                                    If Not IsNothing(current1program) Then
                                        ' ggf hier wieder rausnehmen ...

                                        If current1program.count > 0 Then
                                            If projectConstellations.Contains(current1program.constellationName) Then
                                                projectConstellations.Remove(current1program.constellationName)
                                            End If

                                            createdPrograms = createdPrograms + 1
                                            projectConstellations.Add(current1program)

                                            ' jetzt das union-Projekt erstellen ; 
                                            Dim unionProj As clsProjekt = calcUnionProject(current1program, True, Date.Now.Date.AddHours(23).AddMinutes(59), budget:=last1Budget)

                                            Try
                                                ' Test, ob das Budget auch ausreicht
                                                ' wenn nein, einfach Warning ausgeben 
                                                Dim tmpGesamtCost As Double = unionProj.getGesamtKostenBedarf.Sum
                                                If unionProj.Erloes - tmpGesamtCost < 0 Then
                                                    outPutLine = "Warnung: Budget-Überschreitung bei Programmlinie" & unionProj.name & " (Budget=" & unionProj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                                    outputCollection.Add(outPutLine)

                                                    Dim logtxt(2) As String
                                                    logtxt(0) = "Budget-Überschreitung"
                                                    logtxt(1) = "Programmlinie"
                                                    logtxt(2) = unionProj.name
                                                    Dim values(2) As Double
                                                    values(0) = unionProj.Erloes
                                                    values(1) = tmpGesamtCost
                                                    If values(0) > 0 Then
                                                        values(2) = tmpGesamtCost / unionProj.Erloes
                                                    Else
                                                        values(2) = 9999999999
                                                    End If
                                                    Call logfileSchreiben(logtxt, values)
                                                End If

                                            Catch ex As Exception

                                            End Try

                                            ' Status gleich auf 1: beauftragt setzen 
                                            unionProj.Status = ProjektStatus(PTProjektStati.beauftragt)

                                            If ImportProjekte.Containskey(calcProjektKey(unionProj)) Then
                                                ImportProjekte.Remove(calcProjektKey(unionProj), updateCurrentConstellation:=False)
                                            End If

                                            ImportProjekte.Add(unionProj, updateCurrentConstellation:=False)
                                            ' test
                                            Dim everythingOK As Boolean = testUProjandSingleProjs(current1program)
                                            If Not everythingOK Then
                                                outPutLine = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben: " & current1program.constellationName
                                                outputCollection.Add(outPutLine)

                                                ReDim logmsg(1)
                                                logmsg(0) = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben:"
                                                logmsg(1) = current1program.constellationName
                                                Call logfileSchreiben(logmsg)
                                            End If
                                            ' ende test
                                        Else
                                            emptyPrograms = emptyPrograms + 1
                                        End If

                                    End If

                                    current1program = New clsConstellation(ptSortCriteria.customTF, itemType.ToString & " - " & pName)
                                    lfdNr1program = 2

                                    Try
                                        last1Budget = CDbl(CType(.Cells(zeile, colFields(allianzSpalten.Budget)), Excel.Range).Value)
                                    Catch ex As Exception
                                        last1Budget = 0.0
                                    End Try

                                    'With current1program
                                    '    .constellationName = itemType.ToString & " - " & pName
                                    'End With

                                End If
                            End If


                            If ok Then
                                Try

                                    If isOldAllianzImport Then
                                        custFields = readCustomFieldsFromExcel(colCustomFields, 2, zeile, currentWS)
                                    Else
                                        custFields = readCustomFieldsFromExcel(colCustomFields, 2, zeile, currentWS)
                                    End If


                                    ' lese , wieviel Prozent der Gesamtsumme jeweils auf die Release verteilt werden soll 
                                    For i As Integer = 0 To anzReleases - 1
                                        Try
                                            relPrz(i) = CDbl(CType(.Cells(zeile, colRelPrzStart + i), Excel.Range).Value)
                                            If IsNothing(relPrz(i)) Then
                                                relPrz(i) = 0.0
                                            End If
                                        Catch ex As Exception
                                            relPrz(i) = 0.0
                                        End Try
                                    Next

                                    ' Plausibilitäts-Check - wenn es sich nicht auf 100% summiert, dann lieber alles auf die rootPhase verteilen und nichts auf die Release Phasen
                                    Dim a As Double = relPrz.Sum
                                    If relPrz.Sum > 0 Then
                                        If relPrz.Sum < 0.99 Or relPrz.Sum > 1.01 Then
                                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color = awinSettings.AmpelGelb
                                            CType(.Cells(zeile, lastColumn), Global.Microsoft.Office.Interop.Excel.Range).AddComment(Text:="Prz-Sätze addieren nicht auf 100% ... alles in Projektphase ")
                                            If relPrz.Sum < 0.99 Then
                                                outPutLine = pName & " .Prozent-Sätze > 0 , aber < 1; Gesamt-Summe auf Gesamt-Projekt verteilt  .."
                                            Else
                                                outPutLine = pName & " .Prozent-Sätze > 1.0 , Gesamt-Summe auf Gesamt-Projekt verteilt  .."
                                            End If

                                            outputCollection.Add(outPutLine)

                                            ReDim relPrz(anzReleases - 1)
                                        End If
                                    End If


                                    ' was ist der Gesamtbedarf dieser Rolle in dem besagten Vorhaben ? 
                                    For i As Integer = 0 To colRoleNamesToConsider.Length - 1
                                        Try
                                            If IsNothing(CType(.Cells(zeile, colRoleNamesToConsider(i)), Excel.Range).Value) Then
                                                roleNeeds(i) = 0.0
                                            Else
                                                Dim tmpValue As Double = CDbl(CType(.Cells(zeile, colRoleNamesToConsider(i)), Excel.Range).Value) * nrOfDaysMonth
                                                If tmpValue >= 0 Then
                                                    roleNeeds(i) = tmpValue
                                                Else
                                                    roleNeeds(i) = 0.0
                                                End If
                                            End If
                                        Catch ex As Exception
                                            roleNeeds(i) = 0.0
                                        End Try

                                    Next

                                    If Not isOldAllianzImport Then
                                        Dim i_offset As Integer = colRoleNamesToConsider.Length

                                        For i As Integer = 0 To colTEroleNamesToConsider.Length - 1
                                            Try
                                                If IsNothing(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value) Then
                                                    roleNeeds(i + i_offset) = 0.0
                                                Else
                                                    roleNeeds(i + i_offset) = 0.0
                                                    Dim cellValue As Double = CDbl(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value)
                                                    If cellValue > 0 Then

                                                        Dim tmpRoleDef As clsRollenDefinition = RoleDefinitions.getRoledef(TEroleNamesToConsider(i))

                                                        If Not IsNothing(tmpRoleDef) Then
                                                            ' jetzt handelt es sich um T€ - Werte , das heisst die anzahl Manntage erreichnet sich aus value*1000/tagessatz
                                                            Dim tagessatz As Double
                                                            Dim tmpValue As Double
                                                            Try
                                                                tagessatz = RoleDefinitions.getRoledef(TEroleNamesToConsider(i)).tagessatzIntern
                                                                If tagessatz = 0 Then
                                                                    tagessatz = 800
                                                                    Call MsgBox("tagessatz = 0 ! Rolle " & TEroleNamesToConsider(i))
                                                                End If

                                                                tmpValue = CDbl(CType(.Cells(zeile, colTEroleNamesToConsider(i)), Excel.Range).Value) * 1000 / tagessatz

                                                                If tmpValue >= 0 Then
                                                                    roleNeeds(i + i_offset) = tmpValue
                                                                Else
                                                                    roleNeeds(i + i_offset) = 0.0
                                                                End If
                                                            Catch ex As Exception

                                                            End Try
                                                        End If
                                                    End If


                                                End If
                                            Catch ex As Exception
                                                roleNeeds(i + i_offset) = 0.0
                                            End Try

                                        Next
                                    End If


                                Catch ex As Exception
                                    ok = False
                                End Try

                                ' jetzt werden noch weitere Infos eingelesen ..
                                Try ' Ampelbeschreibung

                                    ampelText = CStr(CType(.Cells(zeile, colFields(allianzSpalten.AmpelText)), Excel.Range).Value)
                                Catch ex As Exception
                                    ampelText = ""
                                End Try

                                Try ' Business Unit

                                    businessUnit = CStr(CType(.Cells(zeile, colFields(allianzSpalten.BusinessUnit)), Excel.Range).Value)
                                Catch ex As Exception
                                    businessUnit = ""
                                End Try

                                Try ' Projektleiter
                                    responsiblePerson = CStr(CType(.Cells(zeile, colFields(allianzSpalten.Responsible)), Excel.Range).Value)
                                Catch ex As Exception
                                    responsiblePerson = ""
                                End Try

                                Try ' Budget
                                    budget = 0.0
                                    If projektvorhaben.Contains(itemType) Then
                                        budget = CStr(CType(.Cells(zeile, colFields(allianzSpalten.pvBudget)), Excel.Range).Value)

                                    ElseIf pgmlinie.Contains(itemType) Then
                                        budget = CStr(CType(.Cells(zeile, colFields(allianzSpalten.Budget)), Excel.Range).Value)
                                        ' wenn dieses Null ist, so soll die andere Spalte genommen werden 
                                        Try
                                            If budget = 0 Then
                                                budget = CStr(CType(.Cells(zeile, colFields(allianzSpalten.pvBudget)), Excel.Range).Value)
                                            End If
                                        Catch ex As Exception

                                        End Try

                                    End If

                                Catch ex As Exception
                                    budget = 0.0
                                End Try


                                Try ' Projekt-Nummer

                                    allianzProjektNummer = CStr(CType(.Cells(zeile, colFields(allianzSpalten.Projektnummer)), Excel.Range).Value)
                                Catch ex As Exception
                                    allianzProjektNummer = ""
                                End Try

                                Try ' Status
                                    If itemType = 6 Then
                                        allianzStatus = ProjektStatus(PTProjektStati.geplant)
                                    Else
                                        allianzStatus = ProjektStatus(PTProjektStati.beauftragt)
                                    End If

                                Catch ex As Exception
                                    allianzStatus = ProjektStatus(PTProjektStati.geplant)
                                End Try

                            End If



                        Catch ex As Exception
                            Call MsgBox("Fehler bei Informationen auslesen: Projekt " & pName)
                            ok = False
                        End Try



                        If ok Then

                            ' Varianten-Name wird hier nicht ausgelesen ..., deshalb Default Wert annehmen 
                            variantName = ""

                            'Projekt anlegen ,Verschiebung um 
                            Dim hproj As clsProjekt = Nothing

                            ' #####################################################################
                            ' Erstellen des Projekts nach den Angaben aus der Batch-Datei 
                            '

                            ' lege ein Allianz IT - Projekt an
                            hproj = erstelleProjektausParametern(pName, variantName, vorlageName, startdate, endDate, budget, sFit, risk, allianzProjektNummer,
                                                                 description, custFields, businessUnit, responsiblePerson, allianzStatus,
                                                                 zeile, realRoleNamesToConsider, roleNeeds, Nothing, Nothing, phNames, relPrz, False)

                            Try
                                ' Test, ob das Budget auch ausreicht
                                ' wenn nein, einfach Warning ausgeben 
                                Dim tmpGesamtCost As Double = hproj.getGesamtKostenBedarf.Sum
                                If hproj.Erloes - tmpGesamtCost < 0 Then
                                    outPutLine = "Warnung: Budget-Überschreitung bei " & pName & " (Budget=" & hproj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                    outputCollection.Add(outPutLine)

                                    Dim logtxt(2) As String
                                    logtxt(0) = "Budget-Überschreitung"
                                    logtxt(1) = "Projekt"
                                    logtxt(2) = pName
                                    Dim values(2) As Double
                                    values(0) = hproj.Erloes
                                    values(1) = tmpGesamtCost
                                    If values(0) > 0 Then
                                        values(2) = tmpGesamtCost / hproj.Erloes
                                    Else
                                        values(2) = 9999999999
                                    End If
                                    Call logfileSchreiben(logtxt, values)
                                End If

                            Catch ex As Exception

                            End Try

                            ' Test tk 
                            Try
                                ReDim logmsg(0)
                                logmsg(0) = pName

                                For ix As Integer = 1 To realRoleNamesToConsider.Count

                                    Dim tmpRollenName As String = realRoleNamesToConsider(ix - 1)
                                    Dim sollBedarf As Double = roleNeeds(ix - 1)


                                    Dim tmpCollection As New Collection
                                    tmpCollection.Add(tmpRollenName)
                                    Dim istBedarf As Double = hproj.getRessourcenBedarf(tmpRollenName,
                                                                                        inclSubRoles:=True).Sum

                                    If Math.Abs(sollBedarf - istBedarf) > 0.001 Then
                                        outPutLine = "Differenz bei " & pName & ", " & tmpRollenName & ": " & Math.Abs(sollBedarf - istBedarf).ToString("#0.##")
                                        outputCollection.Add(outPutLine)
                                    End If

                                Next

                                Dim sollBedarfGesamt As Double = roleNeeds.Sum
                                Dim istBedarfGesamt As Double = hproj.getAlleRessourcen.Sum

                                If Math.Abs(sollBedarfGesamt - istBedarfGesamt) > 0.001 Then
                                    outPutLine = "Gesamt Differenz bei " & pName & ": " & Math.Abs(sollBedarfGesamt - istBedarfGesamt).ToString("#0.##")
                                    outputCollection.Add(outPutLine)
                                End If

                                Call logfileSchreiben(logmsg, roleNeeds)

                            Catch ex As Exception

                            End Try



                            ' Ende Test tk 

                            If Not IsNothing(hproj) Then


                                ' jetzt ist alles so weit ok 
                                Dim pkey As String = ""
                                If Not IsNothing(hproj) Then
                                    Try
                                        pkey = calcProjektKey(hproj)

                                        If ImportProjekte.Containskey(pkey) Then
                                            outPutLine = "Name existiert mehrfach: " & pName
                                            outputCollection.Add(outPutLine)
                                        Else
                                            createdProjects = createdProjects + 1
                                            ImportProjekte.Add(hproj, False)

                                            ' jetzt soll das in die Constellation 
                                            Dim cItem As New clsConstellationItem
                                            With cItem
                                                .projectName = hproj.name
                                                .variantName = hproj.variantName
                                                .show = True
                                                .projectTyp = CType(hproj.projectType, ptPRPFType).ToString
                                                .zeile = lfdNr1program
                                            End With

                                            current1program.add(cItem)
                                            lfdNr1program = lfdNr1program + 1
                                        End If

                                    Catch ex As Exception
                                        outPutLine = "Fehler bei " & pName & vbLf & "Error: " & ex.Message
                                        outputCollection.Add(outPutLine)
                                    End Try

                                End If


                            Else
                                ok = False
                                outPutLine = "Fehler beim Erzeugen des Projektes " & pName
                                outputCollection.Add(outPutLine)
                            End If

                        End If

                    End If


                    geleseneProjekte = geleseneProjekte + 1
                    zeile = zeile + 1

                End While

                ' jetzt die letzte ggf vorkommende Constellation aufnehmen 
                If Not IsNothing(current1program) Then

                    If current1program.count > 0 Then
                        ' ggf aus der Liste aller Constellations wieder rausnehmen 

                        If projectConstellations.Contains(current1program.constellationName) Then
                            projectConstellations.Remove(current1program.constellationName)
                        End If

                        createdPrograms = createdPrograms + 1
                        projectConstellations.Add(current1program)

                        ' jetzt das union-Projekt erstellen 
                        Dim unionProj As clsProjekt = calcUnionProject(current1program, True, Date.Now.Date.AddHours(23).AddMinutes(59), budget:=last1Budget)

                        Try
                            ' Test, ob das Budget auch ausreicht
                            ' wenn nein, einfach Warning ausgeben 
                            Dim tmpGesamtCost As Double = unionProj.getGesamtKostenBedarf.Sum
                            If unionProj.Erloes - tmpGesamtCost < 0 Then
                                outPutLine = "Warnung: Budget-Überschreitung bei Programmlinie " & unionProj.name & " (Budget=" & unionProj.Erloes.ToString("#0.##") & ", Gesamtkosten=" & tmpGesamtCost.ToString("#0.##")
                                outputCollection.Add(outPutLine)

                                Dim logtxt(2) As String
                                logtxt(0) = "Budget-Überschreitung"
                                logtxt(1) = "Programmlinie"
                                logtxt(2) = unionProj.name
                                Dim values(2) As Double
                                values(0) = unionProj.Erloes
                                values(1) = tmpGesamtCost
                                If values(0) > 0 Then
                                    values(2) = tmpGesamtCost / unionProj.Erloes
                                Else
                                    values(2) = 9999999999
                                End If
                                Call logfileSchreiben(logtxt, values)
                            End If

                        Catch ex As Exception

                        End Try

                        ' Status wird gleich auf 1: beauftragt gesetzt
                        unionProj.Status = ProjektStatus(PTProjektStati.beauftragt)

                        If ImportProjekte.Containskey(calcProjektKey(unionProj)) Then
                            ImportProjekte.Remove(calcProjektKey(unionProj), updateCurrentConstellation:=False)
                        End If

                        ImportProjekte.Add(unionProj, updateCurrentConstellation:=False)

                        ' test
                        Dim everythingOK As Boolean = testUProjandSingleProjs(current1program)
                        If Not everythingOK Then

                            outPutLine = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben: " & current1program.constellationName
                            outputCollection.Add(outPutLine)

                            ReDim logmsg(1)
                            logmsg(0) = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben:"
                            logmsg(1) = current1program.constellationName
                            Call logfileSchreiben(logmsg)

                        End If
                        ' ende test
                    Else
                        emptyPrograms = emptyPrograms + 1
                    End If


                End If

            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Import-Datei: " & ex.Message)

        End Try

        If outputCollection.Count > 0 Then
            Call showOutPut(outputCollection, "Import Type 1", "")
        End If

        If emptyPrograms = 0 Then
            Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Projekte erzeugt: " & createdProjects & vbLf &
                    "Programme erzeugt: " & createdPrograms & vbLf &
                    "insgesamt importiert: " & ImportProjekte.Count)
        Else
            Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Projekte erzeugt: " & createdProjects & vbLf &
                    "Programme erzeugt: " & createdPrograms & vbLf &
                    "Programme nicht erzeugt, weil leer: " & emptyPrograms & vbLf &
                    "insgesamt importiert: " & ImportProjekte.Count)
        End If


    End Sub

    ''' <summary>
    ''' gibt den Wert einer Excel Zelle als String zurück
    ''' wenn Nothing in der Excel Zelle ist, wir dder leere String zurückgegeben
    ''' </summary>
    ''' <param name="excelCell">ist die Zelle, um die es geht</param>
    ''' <returns></returns>
    Public Function getStringFromExcelCell(ByVal excelCell As Excel.Range) As String

        Dim tmpResult As String = ""
        Try
            If excelCell.Cells.Count = 1 Then
                tmpResult = CStr(excelCell.Value).Trim
                If IsNothing(tmpResult) Then
                    tmpResult = ""
                End If
            End If

        Catch ex As Exception

        End Try


        getStringFromExcelCell = tmpResult
    End Function

    ''' <summary>
    ''' gibt den Wert einer Excel-Zelle als Double-Wert zurück; wenn Nothing oder keine Zahl wird 0.0 zurückgegeben
    ''' </summary>
    ''' <param name="excelCell"></param>
    ''' <returns></returns>
    Public Function getDoubleFromExcelCell(ByVal excelCell As Excel.Range) As Double

        Dim tmpResult As Double = 0.0
        Try
            If Not IsNothing(excelCell) Then
                If CStr(excelCell.Value) <> "" Then
                    If IsNumeric(excelCell.Value) Then
                        tmpResult = CDbl(excelCell.Value)
                    End If
                End If

            End If
        Catch ex As Exception
            tmpResult = 0.0
        End Try

        getDoubleFromExcelCell = tmpResult
    End Function

    ''' <summary>
    ''' importiert die Offline Ressourcen Dateien 
    ''' </summary>
    ''' <param name="wb"></param>
    ''' <param name="outputCollection"></param>
    Public Sub ImportOfflineData(ByVal wb As String, ByRef outputCollection As Collection)

        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""


        Dim upDatedProjects As Integer = 0
        Dim errorProjects As Integer = 0

        Dim firstRow As Integer = 0
        Dim lastRow As Integer = 0
        Dim lastColumn As Integer = 0

        Dim ok As Boolean = False

        ' die Projekte

        Dim hproj As clsProjekt = Nothing
        Dim newProj As clsProjekt = Nothing
        Dim projektKundenNummer As String = ""

        Dim potentialParentList() As Integer = RoleDefinitions.getIDArray(awinSettings.allianzI2DelRoles)


        ' welche Rollen sollen gelöscht werden; die werden dann danach gesetzt, ob es sich um einen Ressource-Manager handelt, 
        ' der nur einen Teil importieren kann oder um den Portfolio Manager, der eine neue komplette Vorgabe macht 
        Dim deleteRoles As New Collection


        '' new approach 
        Dim currentWS As Excel.Worksheet = Nothing
        Try
            currentWS = CType(CType(appInstance.Workbooks.Item(wb), Excel.Workbook).Worksheets.Item("VISBO"), Excel.Worksheet)
        Catch ex As Exception
            logmessage = "Sheet VISBO nicht gefunden ... Abbruch"
            outputCollection.Add(logmessage)
            Exit Sub
        End Try

        Dim dataRange As Excel.Range = currentWS.UsedRange


        ' bestimme, wo PName, VariantName, Kunden-Nummer, PhaseName, RoleName, Value(PT), Value(T€) steht
        Dim colPName As Integer = 1
        Dim colVName As Integer = 2
        Dim colKdNr As Integer = 3
        Dim colVerantwortlich As Integer = 4
        Dim colPhaseName As Integer = 5
        Dim colRoleName As Integer = 6
        Dim colSumPT As Double = 7
        Dim colSumTe As Double = 8

        ' für logfile 
        Dim tmpanz As Long = 0

        ' bestimme die maximale Anzahl Zeilen
        'Dim lastRow As Integer = 0


        firstRow = CType(dataRange.Rows.Item(2), Excel.Range).Row
        'lastRow = CType(dataRange, Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
        lastRow = CType(currentWS.Cells(dataRange.Rows.Count, "A"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
        lastColumn = CType(dataRange.Columns.Item(dataRange.Columns.Count), Excel.Range).Column



        ' überprüfe bzw. stelle sicher, dass die Datei sortiert nach Pname, VariantName, PhaseName, RoleName ist
        ' in einer ersten Version wird dies manuell sichergestellt ... 
        Dim firstRowOfProject As Integer = firstRow
        Dim lastRowOFProject As Integer = 1
        Dim currentZeile As Integer = 2
        ' ------------------------------------------------
        '
        ' was passiert hier alles ? 
        ' 
        ' für jedes Projekt 

        '    lade das Projekt bzw. die Projekt-Variante
        '    erzeuge eine temporäre Variante 
        '    bestimme die Anzahl Phasen in dem Projekt  
        '    bestimme die rolePhaseValues 
        '    bestimme die deleteRoles aus den "Eltern" der in rolePhaseValues auftretenden Rollen
        '    lösche die deleteRoles+Kinder+Kindeskinder aus jeder Phase des Projektes 
        '    für jede Rolle: 
        '       trage sie In den auftretenden Phasen ein; wenn sie Istdaten enthält, verteile die Summe entsprechend bzw gib eine Fehlermeldung aus 
        '    wenn alles gut gegangen ist, dann mach daraus die ursprüngliche Variante 


        Do While currentZeile <= lastRow

            ' bestimme first- and lastRowOfProject 
            firstRowOfProject = lastRowOFProject + 1
            Dim currentPName As String = getStringFromExcelCell(currentWS.Cells(firstRowOfProject, colPName))
            Dim currentVName As String = getStringFromExcelCell(currentWS.Cells(firstRowOfProject, colVName))
            Dim currentKdNummer As String = getStringFromExcelCell(currentWS.Cells(firstRowOfProject, colKdNr))

            Dim tmpZeile As Integer = firstRowOfProject + 1
            Dim tmpPName As String = getStringFromExcelCell(currentWS.Cells(tmpZeile, colPName))
            Dim tmpVname As String = getStringFromExcelCell(currentWS.Cells(tmpZeile, colVName))

            Dim newPNameFound As Boolean = (tmpPName <> currentPName) Or (tmpVname <> currentVName)

            Do While tmpZeile < lastRow And Not newPNameFound
                ' bestimme die neue lastRowOfProject 
                tmpZeile = tmpZeile + 1
                tmpPName = getStringFromExcelCell(currentWS.Cells(tmpZeile, colPName))
                tmpVname = getStringFromExcelCell(currentWS.Cells(tmpZeile, colVName))
                newPNameFound = (tmpPName <> currentPName) Or (tmpVname <> currentVName)
            Loop

            If newPNameFound Then
                lastRowOFProject = tmpZeile - 1
            Else
                lastRowOFProject = tmpZeile
            End If

            ' jetzt muss der pName normiert werden ..
            If Not isValidProjectName(currentPName) Then
                currentPName = makeValidProjectName(currentPName)
            End If

            ' lade die Projekt-Variante 
            hproj = getProjektFromSessionOrDB(currentPName, currentVName, AlleProjekte, Date.Now, currentKdNummer)
            ' ins Protokoll 



            If IsNothing(hproj) Then
                Dim logtxt(1) As String
                logtxt(0) = "Projekt existiert nicht ... "

                If currentKdNummer = "" Then
                    logtxt(1) = currentPName
                Else
                    logtxt(1) = currentPName & "; " & currentKdNummer
                End If

                logmessage = logtxt(0) & logtxt(1)
                outputCollection.Add(logmessage)

                Call logfileSchreiben(logtxt)

                ' jetzt noch im Input File markieren 
                CType(currentWS.Cells(firstRowOfProject, colPName), Excel.Range).Interior.Color = XlRgbColor.rgbRed

            Else
                If hproj.hasActualValues Then
                    ' noch nicht erlaubt ...

                Else
                    ' Schreiben Protokoll, wenn name und Projektnummer nicht zueinander passen 
                    If hproj.name = currentPName Then
                        ' es wurde über den Namen gefunden 
                        If hproj.kundenNummer <> currentKdNummer And (hproj.kundenNummer <> "" And currentKdNummer <> "") Then

                            logmessage = "Projekt über Name gefunden, aber Projekt-Nummern passen nicht zueinander; Datei: " & currentKdNummer & "; DB: " & hproj.kundenNummer
                            outputCollection.Add(logmessage)

                            Dim logtxt(3) As String
                            logtxt(0) = "Projekt über Name gefunden, aber Projekt-Nummern passen nicht zueinander"
                            logtxt(1) = hproj.name
                            logtxt(2) = currentKdNummer
                            logtxt(3) = "DB: " & hproj.kundenNummer

                            Call logfileSchreiben(logtxt)


                        End If
                    ElseIf hproj.kundenNummer = currentKdNummer Then
                        If hproj.name <> currentPName Then
                            logmessage = "Projekt über Projekt-Nummer gefunden, aber Projekt-Namen passen nicht zueinander" & currentPName & "; DB: " & hproj.name
                            outputCollection.Add(logmessage)

                            Dim logtxt(3) As String
                            logtxt(0) = "Projekt über Projekt-Nummer gefunden, aber Projekt-Namen passen nicht zueinander"
                            logtxt(1) = currentKdNummer
                            logtxt(2) = currentPName
                            logtxt(3) = "DB: " & hproj.name

                            Call logfileSchreiben(logtxt)
                        End If

                    End If



                    Dim anzPhasen As Integer = hproj.CountPhases

                    ' enthält die eingeplanten PT für die einzelnen Releases  
                    Dim phValues() As Double
                    ' enthält die Phasen Namen
                    Dim phNameIDs() As String

                    ReDim phValues(anzPhasen - 1)
                    ReDim phNameIDs(anzPhasen - 1)

                    For ip As Integer = 1 To hproj.CountPhases
                        Dim cPhase As clsPhase = hproj.getPhase(ip)
                        phNameIDs(ip - 1) = cPhase.nameID
                    Next

                    ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
                    Dim rolePhaseValues As New SortedList(Of String, Double())

                    ' wenn wenigstens ein Fehler beim Projekt auftritt , dann wird es nicht eingetragen 
                    Dim atleastOneError As Boolean = False

                    For iz As Integer = firstRowOfProject To lastRowOFProject

                        Dim currentCell As Excel.Range = CType(currentWS.Cells(iz, colRoleName), Excel.Range)
                        Dim roleNameID As String = getRCNameIDfromExcelCell(currentCell)
                        Dim chckRoleName As String = ""

                        Try
                            chckRoleName = CStr(currentCell.Value)
                        Catch ex As Exception

                        End Try

                        currentCell = CType(currentWS.Cells(iz, colPhaseName), Excel.Range)
                        Dim phaseNameID As String = getPhaseNameIDfromExcelCell(currentCell)
                        If phaseNameID = "" Then
                            phaseNameID = rootPhaseName
                        End If


                        ' nur weitermachen, wenn valide Angaben 
                        If phNameIDs.Contains(phaseNameID) And RoleDefinitions.containsNameID(roleNameID) Then
                            Dim curDelRole As String = ""

                            curDelRole = RoleDefinitions.chooseParentFromList(roleNameID, potentialParentList)
                            If curDelRole.Length > 0 Then
                                If Not deleteRoles.Contains(curDelRole) Then
                                    deleteRoles.Add(curDelRole, curDelRole)
                                End If
                            End If


                            ' bestimme jetzt den Index 
                            Dim found As Boolean = False

                            Dim ix As Integer = 0
                            Do While ix <= phNameIDs.Length - 1 And Not found
                                If phNameIDs(ix) = phaseNameID Then
                                    found = True
                                Else
                                    ix = ix + 1
                                End If
                            Loop


                            Dim weiterMachen As Boolean = False

                            If found Then

                                Dim sumPT As Double = getDoubleFromExcelCell(currentWS.Cells(iz, colSumPT))
                                Dim sumTE As Double = getDoubleFromExcelCell(currentWS.Cells(iz, colSumTe))

                                If sumPT > 0 And sumTE = 0 Then
                                    ' Angabe in PT, der Wert passt schon 
                                    weiterMachen = True
                                ElseIf sumPT = 0 And sumTE > 0 Then
                                    ' Angabe in T€
                                    ' der Wert muss in PT umgerechnet werden 
                                    Dim teamID As Integer = -1
                                    Dim tagessatz As Double = RoleDefinitions.getRoleDefByIDKennung(roleNameID, teamID).tagessatzIntern
                                    If tagessatz <= 0 Then
                                        weiterMachen = False
                                    Else
                                        sumPT = sumTE * 1000 / tagessatz
                                        weiterMachen = True
                                    End If

                                Else
                                    ' nichts tun...
                                    weiterMachen = False
                                End If

                                If weiterMachen Then
                                    If rolePhaseValues.ContainsKey(roleNameID) Then
                                        phValues = rolePhaseValues.Item(roleNameID)
                                        phValues(ix) = phValues(ix) + sumPT
                                    Else
                                        ReDim phValues(anzPhasen - 1)
                                        phValues(ix) = sumPT
                                        rolePhaseValues.Add(roleNameID, phValues)
                                    End If

                                End If


                            Else
                                ' Fehler ! kann eigentlich nicht passieren, denn dann wäre er erst gar nicht in den then Zweig gekommen ..? 
                                atleastOneError = True
                                Dim errCol As Integer
                                Dim phaseName As String = CStr(CType(currentWS.Cells(iz, colPhaseName), Excel.Range).Value)
                                logmessage = "Phase-Name existiert nicht: " & phaseName
                                errCol = colPhaseName

                                outputCollection.Add(logmessage)

                                Dim logtxt(2) As String
                                logtxt(0) = "Phase-Name existiert nicht: "
                                logtxt(1) = hproj.name
                                logtxt(2) = phaseName

                                Call logfileSchreiben(logtxt)


                                ' jetzt noch im Input File markieren 
                                CType(currentWS.Cells(iz, errCol), Excel.Range).Interior.Color = XlRgbColor.rgbRed
                                ' tk 16.2.19 hier dürfen keine Kommentare geschrieben werden ! 
                                ' beim nächsten Mal auslesen versucht er das als PhaseID zu interpretieren ! 
                                'If Not IsNothing(CType(currentWS.Cells(iz, errCol), Excel.Range).Comment) Then
                                '    CType(currentWS.Cells(iz, errCol), Excel.Range).ClearComments()
                                'End If
                                'CType(currentWS.Cells(iz, errCol), Excel.Range).AddComment(logmessage)
                            End If



                        Else
                            ' nichts tun
                            Dim errCol As Integer
                            Dim logtxt(2) As String
                            If Not phNameIDs.Contains(phaseNameID) Then
                                atleastOneError = True
                                Dim phaseName As String = CStr(CType(currentWS.Cells(iz, colPhaseName), Excel.Range).Value)
                                logmessage = "Phase-Name existiert nicht: " & phaseName

                                logtxt(0) = "Phase-Name existiert nicht: "
                                logtxt(1) = currentPName
                                logtxt(2) = phaseName

                                errCol = colPhaseName
                            ElseIf Not RoleDefinitions.containsNameID(roleNameID) Then
                                atleastOneError = True
                                Dim roleName As String = CStr(CType(currentWS.Cells(iz, colRoleName), Excel.Range).Value)
                                logmessage = "Rollen-Name existiert nicht: " & roleName

                                logtxt(0) = "Rollen-Name existiert nicht: "
                                logtxt(1) = currentPName
                                logtxt(2) = roleName

                                errCol = colRoleName
                            End If

                            outputCollection.Add(logmessage)

                            Call logfileSchreiben(logtxt)

                            ' jetzt noch im Input File markieren 
                            CType(currentWS.Cells(iz, errCol), Excel.Range).Interior.Color = XlRgbColor.rgbRed
                            ' tk 16.2.19 hier dürfen keine Kommentare geschrieben werden ! 
                            ' beim nächsten Mal auslesen versucht er das als PhaseID zu interpretieren ! 
                            'If Not IsNothing(CType(currentWS.Cells(iz, errCol), Excel.Range).Comment) Then
                            '    CType(currentWS.Cells(iz, errCol), Excel.Range).ClearComments()
                            'End If
                            'CType(currentWS.Cells(iz, errCol), Excel.Range).AddComment(logmessage)

                        End If


                    Next

                    ' prüfen, ob auch kein Fehler beim Import aufgetreten ist ... 
                    If Not atleastOneError Then
                        ' jetzt muss das Projekt verändert werden 
                        ' 1. die deleteRoles alle löschen 

                        ' 2. die roleValues ergänzen

                        ' jetzt wird der Merge auf das Projekt gemacht 
                        ' dabei wird die updateSummaryRole und alle dazu gehörenden SubRoles gelöscht 
                        ' es müssen aber auch die Gruppe gelöscht werden ... 

                        ' test tk 
                        Dim formerLeft As Integer = showRangeLeft
                        Dim formerRight As Integer = showRangeRight
                        showRangeLeft = getColumnOfDate(CDate("1.1.2019"))
                        showRangeRight = getColumnOfDate(CDate("31.12.2019"))

                        Dim testprojekte As New clsProjekte
                        testprojekte.Add(hproj)

                        Dim gesamtVorher As Double = hproj.getAlleRessourcen().Sum
                        Dim gesamtVorher2 As Double = testprojekte.getRoleValuesInMonth("Orga", considerAllSubRoles:=True).Sum

                        ' tk test ...
                        If Math.Abs(gesamtVorher - gesamtVorher2) >= 0.001 Then
                            logmessage = hproj.name & " Einzelproj <> Portfolio" & gesamtVorher.ToString & " <> " & gesamtVorher2.ToString
                            outputCollection.Add(logmessage)
                        End If
                        ' tk test ...

                        ' jetzt alle Rollen und SubRoles von updateSummaryRole löschen 
                        newProj = hproj.deleteRolesAndCosts(deleteRoles, Nothing, True)
                        Dim gesamtNachher As Double = newProj.getAlleRessourcen().Sum

                        ' tk test ...
                        For Each tmpRoleName As String In deleteRoles
                            Dim roleSumNachher As Double = newProj.getRessourcenBedarf(tmpRoleName,
                                                                                       inclSubRoles:=True).Sum

                            If Not roleSumNachher = 0 Then
                                logmessage = "Rolle " & tmpRoleName & " wurde nicht gelöscht ... Fehler bei" & newProj.name
                                outputCollection.Add(logmessage)
                            End If
                        Next
                        ' tk test ...


                        ' jetzt alle Rollen / Phasen Werte hinzufügen 

                        newProj = newProj.merge(rolePhaseValues, phNameIDs, True)

                        ' tk test 
                        For Each kvp As KeyValuePair(Of String, Double()) In rolePhaseValues
                            Dim teilErgebnis As Double = newProj.getRessourcenBedarf(kvp.Key, inclSubRoles:=False).Sum
                            If Math.Abs(teilErgebnis - kvp.Value.Sum) >= 0.001 Then
                                logmessage = "TeilErgebnis ungleich Vorgabe: " & teilErgebnis.ToString("#0.##") & " <> " & kvp.Value.Sum.ToString("#0.##")
                                outputCollection.Add(logmessage)
                            End If

                        Next


                        ' jetzt in die Import-Projekte eintragen 
                        upDatedProjects = upDatedProjects + 1
                        ImportProjekte.Add(newProj, updateCurrentConstellation:=False)

                        ' wegen test 
                        showRangeLeft = formerLeft
                        showRangeRight = formerRight

                    End If ' if not atleastOneError ...

                End If ' if hproj.hasActualValues

            End If

            currentZeile = lastRowOFProject + 1

        Loop

        ' wird an der aufrufenden Stelle gezeigt .. 
        'If outputCollection.Count > 0 Then
        '    Call showOutPut(outputCollection, "Import Offline Ressourcen Zuordnungen", "")
        'End If

        Call MsgBox("Zeilen gelesen: " & (lastRow - firstRow + 1).ToString & vbLf &
                    "Projekte aktualisiert: " & upDatedProjects)


    End Sub

    ''' <summary>
    ''' aktualisiert Projekte mit den für BOSV-KB angegebenen Werten 
    ''' dabei werden die neuen Daten in das Projekt "gemerged"; d.h alle Werte zu anderen Rollen als BOSV-KB bleiben erhalten 
    ''' Ebenso alle Attribute ; es werden also nur die Rollen-Bedarfe zu BOSV-KB ausgetauscht ...  
    ''' </summary>
    Public Sub importAllianzType2()
        Dim zeile As Integer, spalte As Integer

        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""


        Dim upDatedProjects As Integer = 0
        Dim errorProjects As Integer = 0

        ' für den Output 
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""


        Dim vorlageName As String = "Rel"
        Dim lastRow As Integer
        Dim lastColumn As Integer
        Dim geleseneProjekte As Integer
        Dim ok As Boolean = False

        ' die Projekte

        Dim hproj As clsProjekt = Nothing
        Dim newProj As clsProjekt = Nothing
        Dim projektKundenNummer As String = ""

        ' welche Rollen sollen gelöscht werden
        Dim deleteRoles As New Collection

        ' jetzt werden die aufgebaut ...
        If awinSettings.allianzI2DelRoles = "" Then

            deleteRoles.Add("D-BOSV-KB0")
            deleteRoles.Add("D-BOSV-KB1")
            deleteRoles.Add("D-BOSV-KB2")
            deleteRoles.Add("D-BOSV-KB3")
            deleteRoles.Add("Grp-BOSV-KB")

        Else
            Dim tmpStr() As String = awinSettings.allianzI2DelRoles.Split(New Char() {CChar(";")})
            For Each tmpRCName As String In tmpStr
                If RoleDefinitions.containsName(tmpRCName.Trim) Then
                    deleteRoles.Add(tmpRCName.Trim)
                End If
            Next
        End If

        ' diese Rollen und Subroles sollen alle vorher gelöscht werden und dann mit den neuen Werten ersetzt werden 
        ' Amis soll nicht gelöscht werden, deshalb die explizite Aufführung


        ' Standard-Definition
        Dim anzPhasen As Integer = 5

        Try
            anzPhasen = Projektvorlagen.getProject(vorlageName).CountPhases
        Catch ex As Exception
            Call MsgBox("in ImportAllianzType2: " & vbLf & "es gibt keine Projektvorlage " & vorlageName & ".xlsx!" & vbLf & "-> Abbruch ...")
            Exit Sub
        End Try


        ' enthält die eingeplanten PT für die einzelnen Releases  
        Dim phValues() As Double
        ReDim phValues(anzPhasen - 1)

        ' nimmt die Farbe auf, die steuert, dass diese Zeile nicht eingelesen wird ... 
        Dim projectStartingColor As Integer

        Dim currentColor As Integer


        ' enthält die Phasen Namen
        Dim phNameIDs() As String
        ReDim phNameIDs(anzPhasen - 1)

        ' enthält die Spalten-Nummer, ab der die Release Phasen Mann-Tage stehen 
        Dim colRelValues As Integer

        ' enthät die Saplte, wo der ProjektName steht ...
        Dim colPname As Integer

        ' enthält die Spalten-Nummer, wo die einzelnen Rollen-Namen zu finden sind
        Dim colRoleName As Integer = -1

        ' jetzt werden die ImportProjekte zurückgesetzt ...
        ImportProjekte.Clear()

        Dim firstZeile As Excel.Range


        zeile = 2
        spalte = 1
        geleseneProjekte = 0

        ' jetzt werden die Phase-Names besetzt
        Try
            For i = 1 To anzPhasen
                phNameIDs(i - 1) = Projektvorlagen.getProject(vorlageName).getPhase(i).nameID
            Next
        Catch ex As Exception
            Call MsgBox("Probleme mit Vorlage " & vorlageName)
            Exit Sub
        End Try

        ' enthält, wieviel Manntage von dieser Rolle insgesamt benötigt werden 
        Dim rolePhaseValues As New SortedList(Of String, Double())


        Try

            Dim found As Boolean = False
            Dim wsi As Integer = 1
            Dim wsCount As Integer = appInstance.ActiveWorkbook.Worksheets.Count


            While Not found And wsi <= wsCount
                If CType(appInstance.ActiveWorkbook.Worksheets.Item(wsi),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet).Name.StartsWith("Projekte") Then
                    found = True
                Else
                    wsi = wsi + 1
                End If
            End While

            If Not found Then
                Call MsgBox("keine Projekte-Tabelle gefunden ...")
                Exit Sub
            End If

            Dim currentWS As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets.Item(wsi),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With currentWS


                firstZeile = CType(.Rows(2), Excel.Range)

                ' jetzt wird festgelegt, ab wo die absoluten PT-Werte für die Releases stehen 
                colRelValues = CType(.Range("M1"), Excel.Range).Column

                colPname = CType(.Range("B1"), Excel.Range).Column

                ' wo stehen die Team-Bezeichner
                colRoleName = .Range("D1").Column

                projectStartingColor = CInt(CType(.Cells(2, 2), Excel.Range).Interior.Color)


                'lastColumn = firstZeile.End(XlDirection.xlToLeft).Column

                lastColumn = CType(.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlToLeft).Column
                lastRow = CType(.Cells(20000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row



                While zeile < lastRow

                    Dim oldProj As clsProjekt = Nothing

                    currentColor = CInt(CType(.Cells(zeile, 2), Excel.Range).Interior.Color)
                    If currentColor = projectStartingColor Then
                        ' jetzt kommt die Behandlung ...   

                        Try

                            pName = CStr(CType(.Cells(zeile, colPname), Excel.Range).Value).Trim
                            geleseneProjekte = geleseneProjekte + 1
                            ok = isKnownProject(pName, projektKundenNummer, AlleProjekte)

                        Catch ex As Exception
                            ok = False
                        End Try

                        ' startzeile muss jetzt gemerkt werden ...
                        zeile = zeile + 1
                        currentColor = CInt(CType(.Cells(zeile, 2), Excel.Range).Interior.Color)
                        Dim startzeile As Integer = zeile
                        Dim endeZeile As Integer = startzeile

                        ' jetzt schon zeile auf das nächste Projekt positionieren ...
                        Do Until currentColor = projectStartingColor And Not zeile > lastRow
                            zeile = zeile + 1
                            currentColor = CInt(CType(.Cells(zeile, 2), Excel.Range).Interior.Color)
                        Loop
                        ' in zeile ist jetzt das nächste Projekt 

                        If Not ok Then

                            outPutLine = "Projekt nicht bekannt: " & pName
                            outputCollection.Add(outPutLine)

                        Else

                            Dim pvKey As String = calcProjektKey(pName, "")
                            oldProj = AlleProjekte.getProject(pvKey)

                            ' jetzt werden die Values für ein Projekt ausgelsen 
                            rolePhaseValues.Clear()
                            ' in zeile steht das nächste Projekt, in zeile-1 dann der letzte Eintrag des aktuellen Projekts
                            endeZeile = zeile - 1


                            ' jetzt kann rolePhaseValues dimensioniert werden 
                            For iz As Integer = startzeile To endeZeile
                                Dim phaseValues(anzPhasen - 1) As Double
                                Dim roleName As String = ""

                                If Not IsNothing(CType(.Cells(iz, colRoleName), Excel.Range).Value) Then
                                    roleName = CStr(CType(.Cells(iz, colRoleName), Excel.Range).Value).Trim


                                    If roleName <> "" Then

                                        If RoleDefinitions.containsName(roleName) Then

                                            ' jetzt muss die RCNameID bestimmt werden 
                                            Dim rcNameID As String = RoleDefinitions.getRoledef(roleName).UID.ToString
                                            For ip As Integer = 1 To anzPhasen - 1
                                                phaseValues(ip) = CDbl(CType(.Cells(iz, colRelValues + ip - 1), Excel.Range).Value)
                                            Next

                                            If phaseValues.Sum = 0 Then
                                                ' nichts tun
                                            Else
                                                If rolePhaseValues.ContainsKey(rcNameID) Then
                                                    ' addieren ...
                                                    For px As Integer = 1 To anzPhasen - 1
                                                        rolePhaseValues.Item(rcNameID)(px) = rolePhaseValues.Item(rcNameID)(px) + phaseValues(px)
                                                    Next
                                                Else
                                                    ' neu aufnehmen 
                                                    rolePhaseValues.Add(rcNameID, phaseValues)
                                                End If
                                            End If
                                        Else
                                            outPutLine = "Team / Rolle nicht bekannt: " & roleName
                                            outputCollection.Add(outPutLine)
                                        End If

                                    End If

                                End If

                            Next

                            ' jetzt wird der Merge auf das Projekt gemacht 
                            ' dabei wird die updateSummaryRole und alle dazu gehörenden SubRoles gelöscht 
                            ' es müssen aber auch die Gruppe gelöscht werden ... 

                            ' test tk 
                            Dim formerLeft As Integer = showRangeLeft
                            Dim formerRight As Integer = showRangeRight
                            showRangeLeft = getColumnOfDate(CDate("1.1.2018"))
                            showRangeRight = getColumnOfDate(CDate("31.12.2018"))

                            Dim testprojekte As New clsProjekte
                            testprojekte.Add(oldProj)

                            Dim gesamtVorher As Double = oldProj.getAlleRessourcen().Sum
                            Dim gesamtVorher2 As Double = testprojekte.getRoleValuesInMonth("Orga", considerAllSubRoles:=True).Sum
                            Dim bosvVorher As Double = oldProj.getRessourcenBedarf("D-BOSV-KB", inclSubRoles:=True).Sum

                            ' tk test ...
                            If Math.Abs(gesamtVorher - gesamtVorher2) >= 0.001 Then
                                Call MsgBox(oldProj.name & " Einzelproj <> Portfolio" & gesamtVorher.ToString & " <> " & gesamtVorher2.ToString)
                            End If
                            ' tk test ...

                            ' jetzt alle Rollen und SubRoles von updateSummaryRole löschen 
                            newProj = oldProj.deleteRolesAndCosts(deleteRoles, Nothing, True)
                            Dim gesamtNachher As Double = newProj.getAlleRessourcen().Sum

                            ' tk test ...
                            For Each tmpRoleName As String In deleteRoles
                                Dim bosvNachher As Double = newProj.getRessourcenBedarf(tmpRoleName, inclSubRoles:=True).Sum

                                If Not bosvNachher = 0 Then
                                    Call MsgBox(tmpRoleName & " wurde nicht gelöscht ... Fehler bei" & newProj.name)
                                End If
                            Next
                            ' tk test ...


                            ' jetzt alle Rollen / Phasen Werte hinzufügen 
                            Dim addValues As Double = 0.0
                            For Each kvp As KeyValuePair(Of String, Double()) In rolePhaseValues
                                addValues = addValues + kvp.Value.Sum
                            Next
                            newProj = newProj.merge(rolePhaseValues, phNameIDs, True)

                            Dim bosvErgebnis As Double = newProj.getRessourcenBedarf("Grp-BOSV-KB", inclSubRoles:=True).Sum

                            If Math.Abs(bosvErgebnis - addValues) >= 0.001 Then
                                outPutLine = "addValues ungleich ergebnis: " & addValues.ToString("#0.##") & " <> " & bosvErgebnis.ToString("#0.##")
                                outputCollection.Add(outPutLine)
                            End If

                            ' jetzt in die Import-Projekte eintragen 
                            upDatedProjects = upDatedProjects + 1
                            ImportProjekte.Add(newProj, updateCurrentConstellation:=False)

                            ' wegen test 
                            showRangeLeft = formerLeft
                            showRangeRight = formerRight
                        End If

                    End If

                End While


            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Import-Datei" & ex.Message)

        End Try

        If outputCollection.Count > 0 Then
            Call showOutPut(outputCollection, "Import Detail-Planungs Typ 2", "")
        End If

        Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Projekte aktualisiert: " & upDatedProjects)


    End Sub

    ''' <summary>
    ''' importiert die Ist-Datensätze zu allen Projekten, die identifiziert werden können  
    ''' </summary>
    ''' <param name="monat">gibt an, bis wohin einschließlich Ist-Werte gelesen werden </param>
    ''' <param name="readAll">gibt an, ob Vergangenheit und Zukunft gelesen werden soll</param>
    ''' <param name="createUnknown">gibt an, ob Unbekannte Projekte angelegt werden sollen</param>
    Public Sub ImportAllianzType3(ByVal monat As Integer, ByVal readAll As Boolean, ByVal createUnknown As Boolean,
                                  ByRef outputCollection As Collection)


        ' im Key steht der Projekt-Name, im Value steht eine sortierte Liste mit key=Rollen-Name, values die Ist-Werte
        Dim validProjectNames As New SortedList(Of String, SortedList(Of String, Double()))

        ' nimmt dann später pro Projekt die vorkommenden Rollen auf - setzt voraus, dass die Datei nach Projekt-Namen, dann nach Jahr, dann nach Monat sortiert ist ...  
        Dim projectRoleNames(,) As String = Nothing

        ' nimmt dann die Werte pro Projekt, Rolle und Monat auf  
        Dim projectRoleValues(,,) As Double = Nothing

        ' für die Meldungen
        Dim outPutLine As String = ""

        Dim lastRow As Integer = -1
        Dim updatedProjects As Integer = 0

        Dim logF_Fehler As Integer = 0
        ' nimmt die Texte für die LogFile Zeile auf
        ' Array kann beliebig lang werden 
        Dim logArray() As String
        Dim logDblArray() As Double

        ' nimmt auf, zu welcher Orga-Einheit die Ist-Daten erfasst werden ... 
        Dim referatsCollection As New Collection



        Dim lastValidMonth As Integer = monat
        If readAll Then
            lastValidMonth = 12
        End If

        ' jetzt muss als erstes auf das korrekte Worksheet positioniert werden 
        ' das aktive Sheet muss das richtige sein ... und die richtige Header Struktur haben 
        Try


            If monat < 1 Or monat > 12 Then
                logmessage = "ungültige Angabe des ActualDataUntil-Monats: " & monat
                outputCollection.Add(logmessage)
                Exit Sub
            End If

            ' jetzt kommt die eigentliche Import Behandlung 
            Dim currentWS As Excel.Worksheet = Nothing
            Try
                currentWS = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                           Global.Microsoft.Office.Interop.Excel.Worksheet)
                If Not currentWS.Name.Contains("Bericht") Then
                    currentWS = CType(appInstance.ActiveWorkbook.Worksheets("Bericht_RL_Kapa_Excel"),
                                                           Global.Microsoft.Office.Interop.Excel.Worksheet)
                End If
            Catch ex As Exception
                logmessage = "Keine Tabelle mit Namen 'Bericht_RL_Kapa_Excel> gefunden' ... Abbruch"
                outputCollection.Add(logmessage)
                Exit Sub
            End Try


            ' tk, 2.8.2018 Behandlung LookupTable 
            Dim lookUpTableWS As Excel.Worksheet = Nothing

            ' die lookupTable nimmt die Projekt-Nummer als KEy auf und den korrespondierenden NAmen aus der Rupi-Liste
            ' bei Aufbau der lookupTable werden die Rupi-Liste NAmen bereits in valide Namen gewandelt ... 
            Dim lookupTable As SortedList(Of String, String) = Nothing

            Try
                lookUpTableWS = CType(appInstance.ActiveWorkbook.Worksheets("lookupTable"),
                                                           Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                lookUpTableWS = Nothing
            End Try

            ' wenn jetzt eine Tabelle vorhanden ist, dann muss die LookupTable aufgebaut werden 
            If Not IsNothing(lookUpTableWS) Then

                With lookUpTableWS

                    Dim lupTLastZeile As Integer = CType(.Cells(20000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                    Dim lupTZeile As Integer = 2
                    If lupTLastZeile >= lupTZeile Then

                        lookupTable = New SortedList(Of String, String)

                        While lupTZeile <= lupTLastZeile
                            Try
                                Dim pNr As String = CStr(CType(.Cells(lupTZeile, 2), Excel.Range).Value).Trim
                                Dim rupiPName As String = CStr(CType(.Cells(lupTZeile, 3), Excel.Range).Value).Trim

                                If Not isValidProjectName(rupiPName) Then
                                    rupiPName = makeValidProjectName(rupiPName)
                                End If

                                If pNr <> "" Then
                                    If Not lookupTable.ContainsKey(pNr) Then
                                        lookupTable.Add(pNr, rupiPName)
                                    End If
                                End If

                            Catch ex As Exception

                            End Try

                            lupTZeile = lupTZeile + 1

                        End While

                    End If
                End With


            End If

            Dim lookupsExist As Boolean = False
            If Not IsNothing(lookupTable) Then
                lookupsExist = (lookupTable.Count > 0)
            End If


            ' hat die Datei die richtige Header-Struktur ? 
            Dim firstZeile As Excel.Range = currentWS.Rows(1)

            If Not isCorrectAllianzImportStructure(firstZeile, 3) Then
                logmessage = "Datei hat nicht den für den Istdaten-Import erforderlichen Spalten-Aufbau!"
                outputCollection.Add(logmessage)

                Exit Sub
            End If

            ' hier wird das Logfile jetzt geöffnet 
            Call logfileOpen()

            With currentWS

                Dim zeile As Integer = 2
                lastRow = CType(.Cells(20000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                ' welche Werte sollen ausgelesen werden, wo stehen die 
                Dim colISExtern As Integer = CType(.Range("F1"), Excel.Range).Column
                Dim colResource As Integer = CType(.Range("G1"), Excel.Range).Column
                Dim colProjectNr As Integer = CType(.Range("C1"), Excel.Range).Column
                Dim colPname As Integer = CType(.Range("D1"), Excel.Range).Column

                Dim colYear As Integer = CType(.Range("H1"), Excel.Range).Column
                Dim colMonth As Integer = CType(.Range("I1"), Excel.Range).Column
                Dim colEuroActuals As Integer = CType(.Range("L1"), Excel.Range).Column
                Dim colReferat As Integer = CType(.Range("A1"), Excel.Range).Column
                Dim colPTZuweisung As Integer = CType(.Range("J1"), Excel.Range).Column

                Dim cacheProjekte As New clsProjekteAlle

                ' im key steht der NAme aus der Datei , im Value steht der Name i CacheProjekte
                Dim handledNames As New SortedList(Of String, String)
                ' nimmt die unbekannten / nicht erkannten Role-Names auf 
                Dim unKnownRoleNames As New SortedList(Of String, Boolean)

                ' 1. Schleife: ermittle die Menge aller bekannten Projekte ... 
                While zeile <= lastRow

                    Try
                        ' zu welchem Referat gehört die Rolle ? das wird später benötigt, um die zugehörigen Werte zurücksetzen zu können 
                        Dim tmpReferat As String = CStr(CType(.Cells(zeile, colReferat), Excel.Range).Value).Trim
                        If Not IsNothing(tmpReferat) Then
                            If RoleDefinitions.containsName(tmpReferat) Then
                                ' wenn es sich um KB handelt , dann ist KB0 gemeint ...
                                If tmpReferat = "D-BOSV-KB" Then
                                    tmpReferat = "D-BOSV-KB0"
                                    ' jetzt auch AMIS in ReferatsCollection aufnehmen ..
                                    If Not referatsCollection.Contains("AMIS") Then
                                        referatsCollection.Add("AMIS", "AMIS")
                                    End If
                                End If
                                If Not referatsCollection.Contains(tmpReferat) Then
                                    referatsCollection.Add(tmpReferat, tmpReferat)
                                End If
                            End If
                        End If

                        Dim tmpPName As String = CStr(CType(.Cells(zeile, colPname), Excel.Range).Value).Trim
                        Dim tmpPNr As String = CStr(CType(.Cells(zeile, colProjectNr), Excel.Range).Value).Trim

                        Dim pName As String = getAllianzPNameFromPPN(tmpPName, tmpPNr)

                        Dim fullRoleName As String = CStr(CType(.Cells(zeile, colResource), Excel.Range).Value).Trim
                        Dim isExtern As Boolean = False

                        Try
                            Dim tstExtern As String = CStr(CType(.Cells(zeile, colISExtern), Excel.Range).Value)
                            isExtern = (tstExtern = "Extern")
                        Catch ex1 As Exception
                            isExtern = False
                        End Try

                        Dim curMonat As Integer = CInt(CType(.Cells(zeile, colMonth), Excel.Range).Value)
                        Dim curIstEuroValue As Double = CDbl(CType(.Cells(zeile, colEuroActuals), Excel.Range).Value)
                        Dim curZuwPTValue As Double = CDbl(CType(.Cells(zeile, colPTZuweisung), Excel.Range).Value)

                        Dim shallContinue As Boolean = False
                        Dim oldProj As clsProjekt = Nothing

                        If Not handledNames.ContainsKey(tmpPName) Then

                            ' wenn ok = true zurück kommt, dann ist in cacheProjekte das Projekt mit Varianten-Name "" drin .. 
                            If isKnownProject(pName, tmpPNr, cacheProjekte, lookupTable, createUnknown) Then
                                shallContinue = True
                                handledNames.Add(tmpPName, pName)

                            Else
                                outPutLine = "unbekanntes Projekt: " & pName & "; P-Nr: " & tmpPNr
                                outputCollection.Add(outPutLine)

                                ReDim logArray(4)
                                logArray(0) = "unbekannte PNr / Projekt "
                                logArray(1) = tmpPNr
                                logArray(2) = pName
                                logArray(3) = ""
                                logArray(4) = ""
                                Call logfileSchreiben(logArray)

                                shallContinue = False
                                handledNames.Add(tmpPName, "")

                            End If

                        Else
                            shallContinue = (handledNames.Item(tmpPName).Length > 0)
                        End If

                        If Not readAll Then
                            shallContinue = shallContinue And curMonat >= 1 And curMonat <= monat And curIstEuroValue >= 0
                        Else
                            ' readAll:
                            shallContinue = shallContinue And curMonat >= 1 And curMonat <= lastValidMonth And curZuwPTValue >= 0
                        End If

                        If shallContinue Then
                            'If shallContinue And curMonat >= 1 And curMonat <= monat And curIstEuroValue >= 0 Then
                            ' nur dann handelt es sich um ein zuordenbares Projekt ... 
                            ' nur dann geht es um gültige Werte

                            ' jetzt wird eine sortedlist of sortedlist aufgebaut
                            ' Projekte, dann Rollen mit Values 
                            ' 
                            pName = handledNames.Item(tmpPName)
                            Dim pvkey As String = calcProjektKey(pName, "")
                            oldProj = cacheProjekte.getProject(pvkey)

                            If Not IsNothing(oldProj) Then

                                Dim roleName As String = getAllianzRoleNameFromValue(fullRoleName, isExtern)

                                ' tk 7.8.18, um zu verhindern, dass Rollen mehrfach gezählt werden, wenn sie in der Config Datei / Ist-Daten  Datei zu unterschiedlichen Referaten gehören ...   
                                ' sicherstellen, dass roleName auch existiert ..
                                If roleName <> "" Then
                                    If Not RoleDefinitions.hasAnyChildParentRelationsship(roleName, referatsCollection) Then
                                        ' in diesem Fall muss die Eltern-Rolle der RoleName noch aufgenommen werden ..
                                        Dim tmpparentRole As String = ""
                                        Try
                                            Dim roleUID As Integer = RoleDefinitions.getRoledef(roleName).UID

                                            tmpparentRole = RoleDefinitions.getParentRoleOf(roleUID).name
                                            If Not referatsCollection.Contains(tmpparentRole) Then
                                                referatsCollection.Add(tmpparentRole, tmpparentRole)
                                            End If
                                        Catch ex As Exception

                                        End Try

                                        ReDim logArray(4)
                                        logArray(0) = "Rolle hat anderes Referat wie in Konfiguration"
                                        logArray(1) = ""
                                        logArray(2) = tmpReferat
                                        logArray(3) = fullRoleName
                                        logArray(4) = tmpparentRole
                                        Call logfileSchreiben(logArray)

                                    End If
                                End If



                                If roleName = "" And tmpReferat <> "" Then
                                    ' dann ersetzen durch 
                                    roleName = tmpReferat

                                    If Not unKnownRoleNames.ContainsKey(fullRoleName) Then
                                        unKnownRoleNames.Add(fullRoleName, True)
                                        'outPutLine = "unbekannt: " & fullRoleName
                                        'outputCollection.Add(outPutLine)
                                        logmessage = "unbekannte Rolle wird ersetzt durch Referat " & fullRoleName & " -> " & tmpReferat
                                        outputCollection.Add(logmessage)

                                        ReDim logArray(4)
                                        logArray(0) = "unbekannte Rolle wird ersetzt durch Referat"
                                        logArray(1) = ""
                                        logArray(2) = ""
                                        logArray(3) = fullRoleName
                                        logArray(4) = tmpReferat
                                        Call logfileSchreiben(logArray)
                                    End If
                                End If

                                If roleName = "" Then

                                    If Not unKnownRoleNames.ContainsKey(fullRoleName) Then
                                        unKnownRoleNames.Add(fullRoleName, True)
                                        'outPutLine = "unbekannt: " & fullRoleName
                                        'outputCollection.Add(outPutLine)
                                        logmessage = "unbekannte Rolle ohne Referat: " & fullRoleName
                                        outputCollection.Add(logmessage)

                                        ReDim logArray(4)
                                        logArray(0) = "unbekannte Rolle ohne Referat"
                                        logArray(1) = ""
                                        logArray(2) = ""
                                        logArray(3) = fullRoleName
                                        logArray(4) = "?"
                                        Call logfileSchreiben(logArray)
                                    End If


                                Else
                                    ' Aufbauen des Eintrags
                                    Dim roleValues As New SortedList(Of String, Double())
                                    Dim tmpValues() As Double

                                    'ReDim tmpValues(monat - 1)
                                    ' lastValidMonth ist entweder der monat oder aber 12, falls alles gelesen werden soll 
                                    ReDim tmpValues(lastValidMonth - 1)

                                    Dim hrole As clsRollenDefinition = RoleDefinitions.getRoledef(roleName)

                                    If Not IsNothing(hrole) Then
                                        Dim tagessatz As Double = hrole.tagessatzIntern
                                        If tagessatz <= 0 Then
                                            tagessatz = 800.0
                                        End If

                                        If Not validProjectNames.ContainsKey(pName) Then

                                            roleValues = New SortedList(Of String, Double())
                                            ' wird doch überhaupt nicht gebraucht
                                            'ReDim tmpValues(monat - 1)
                                            If readAll Then
                                                ' es muss unterschieden werden, ob es sich um Ist-Daten oder um Zuwesiung handelt ...  
                                                If curMonat <= monat Then
                                                    tmpValues(curMonat - 1) = curIstEuroValue / tagessatz
                                                Else
                                                    tmpValues(curMonat - 1) = curZuwPTValue
                                                End If
                                            Else
                                                ' es handelt sich um Ist-Euro, also muss umgerechnet werden 
                                                tmpValues(curMonat - 1) = curIstEuroValue / tagessatz
                                            End If


                                            roleValues.Add(roleName, tmpValues)
                                            validProjectNames.Add(pName, roleValues)

                                        Else
                                            roleValues = validProjectNames.Item(pName)
                                            If roleValues.ContainsKey(roleName) Then
                                                ' rolle ist bereits enthalten 
                                                ' also summieren 
                                                tmpValues = roleValues.Item(roleName)
                                                If readAll Then
                                                    ' es muss unterschieden werden, ob es sich um Ist-Daten oder um Zuwesiung handelt ...  
                                                    If curMonat <= monat Then
                                                        tmpValues(curMonat - 1) = tmpValues(curMonat - 1) + curIstEuroValue / tagessatz
                                                    Else
                                                        tmpValues(curMonat - 1) = tmpValues(curMonat - 1) + curZuwPTValue
                                                    End If
                                                Else
                                                    tmpValues(curMonat - 1) = tmpValues(curMonat - 1) + curIstEuroValue / tagessatz
                                                End If

                                            Else
                                                ' Rolle ist noch nicht enthalten 
                                                'ReDim tmpValues(monat - 1)

                                                If readAll Then
                                                    ' es muss unterschieden werden, ob es sich um Ist-Daten oder um Zuwesiung handelt ...  
                                                    If curMonat <= monat Then
                                                        tmpValues(curMonat - 1) = curIstEuroValue / tagessatz
                                                    Else
                                                        tmpValues(curMonat - 1) = curZuwPTValue
                                                    End If
                                                Else
                                                    ' es handelt sich um Ist-Euro, also muss umgerechnet werden 
                                                    tmpValues(curMonat - 1) = curIstEuroValue / tagessatz
                                                End If

                                                roleValues.Add(roleName, tmpValues)
                                            End If

                                        End If
                                    Else
                                        ' darf/kann eigentlich nicht sein ...
                                        logmessage = "unbekannte Rolle ohne Referat: " & roleName
                                        outputCollection.Add(logmessage)

                                        ReDim logArray(3)
                                        logArray(0) = "Rollendefinition nicht gefunden ... Fehler 100412: "
                                        logArray(1) = ""
                                        logArray(2) = ""
                                        logArray(3) = roleName
                                        Call logfileSchreiben(logArray)
                                    End If



                                End If



                            Else
                                ' darf/kann eigentlich nicht sein ...
                                logmessage = "Fehler 100411: Projekt mit Name nicht gefunden: " & pName
                                outputCollection.Add(logmessage)

                                ReDim logArray(1)
                                logArray(0) = "Fehler 100411: Projekt mit Name nicht gefunden: "
                                logArray(1) = pvkey
                                Call logfileSchreiben(logArray)
                            End If

                        End If


                    Catch ex As Exception
                        outPutLine = "Fehler 99232: " & ex.Message
                        outputCollection.Add(outPutLine)

                        ' darf/kann eigentlich nicht sein ...
                        ReDim logArray(1)
                        logArray(0) = "Projekt Fehler 100413 in Zeile: "
                        logArray(1) = zeile.ToString
                        Call logfileSchreiben(logArray)
                    End Try

                    zeile = zeile + 1

                End While

                ' jetzt kommt die zweite Bearbeitungs-Welle
                ' das Rausschreiben der Test Records 

                ' Protokoll schreiben ...
                For Each vPKvP As KeyValuePair(Of String, SortedList(Of String, Double())) In validProjectNames

                    Dim protocolLine As String = ""
                    For Each rVKvP As KeyValuePair(Of String, Double()) In vPKvP.Value

                        ' jetzt schreiben 
                        Dim hrole As clsRollenDefinition = RoleDefinitions.getRoledef(rVKvP.Key)
                        Dim curTagessatz As Double = hrole.tagessatzIntern

                        ReDim logArray(3)
                        logArray(0) = "Importiert wurde: "
                        logArray(1) = ""
                        logArray(2) = vPKvP.Key
                        logArray(3) = rVKvP.Key


                        ReDim logDblArray(rVKvP.Value.Length - 1)
                        For i As Integer = 0 To rVKvP.Value.Length - 1
                            ' umrechnen, damit es mit dem Input File wieder vergleichbar wird 
                            logDblArray(i) = rVKvP.Value(i) * curTagessatz
                        Next

                        Call logfileSchreiben(logArray, logDblArray)
                    Next

                Next
                ' Protokoll schreiben Ende ... 

                Dim gesamtIstValue As Double = 0.0

                For Each vPKvP As KeyValuePair(Of String, SortedList(Of String, Double())) In validProjectNames

                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(vPKvP.Key, "", cacheProjekte, Date.Now)
                    Dim oldPlanValue As Double = 0.0
                    Dim newIstValue As Double = 0.0

                    If Not IsNothing(hproj) Then
                        ' es wird pro Projekt eine Variante erzeugt 
                        Dim istDatenVName As String = ptVariantFixNames.acd.ToString
                        Dim newProj As clsProjekt = hproj.createVariant(istDatenVName, "temporär für Ist-Daten-Aufnahme")

                        ' es werden in jeder Phase, die einen der actual Monate enthält, die Werte gelöscht ... 
                        ' gleichzeitig werden die bisherigen Soll-Werte dieser Zeit in T€ gemerkt ...
                        ' True: die Werte werden auf Null gesetzt 
                        Dim gesamtvorher As Double = newProj.getGesamtKostenBedarf().Sum * 1000

                        'oldPlanValue = newProj.getSetRoleCostUntil(referatsCollection, monat, True)
                        oldPlanValue = newProj.getSetRoleCostUntil(referatsCollection, lastValidMonth, True)
                        'Dim checkOldPlanValue As Double = newProj.getSetRoleCostUntil(referatsCollection, monat, False)

                        newIstValue = calcIstValueOf(vPKvP.Value)

                        gesamtIstValue = gesamtIstValue + newIstValue

                        ' die Werte der neuen Rollen in PT werden in der RootPhase eingetragen 
                        Call newProj.mergeActualValues(rootPhaseName, vPKvP.Value)

                        Dim gesamtNachher As Double = newProj.getGesamtKostenBedarf().Sum * 1000
                        Dim checkNachher As Double = gesamtvorher - oldPlanValue + newIstValue
                        ' Test tk 
                        'Dim checkIstValue As Double = newProj.getSetRoleCostUntil(referatsCollection, monat, False)
                        Dim checkIstValue As Double = newProj.getSetRoleCostUntil(referatsCollection, lastValidMonth, False)

                        If gesamtNachher <> checkNachher Then
                            Dim abc As Integer = 0
                        End If

                        If checkIstValue <> newIstValue Then
                            Dim abc As Integer = 0
                        End If

                        ReDim logArray(3)
                        logArray(0) = "Import Istdaten old/new/diff/check1/check2"
                        logArray(1) = ""
                        logArray(2) = vPKvP.Key
                        logArray(3) = ""

                        ReDim logDblArray(4)
                        logDblArray(0) = oldPlanValue
                        logDblArray(1) = newIstValue
                        logDblArray(2) = oldPlanValue - newIstValue
                        logDblArray(3) = checkIstValue
                        logDblArray(4) = gesamtNachher - checkNachher

                        Call logfileSchreiben(logArray, logDblArray)

                        ' die Differenz aus Soll und Ist zwischen Beautragung / Actual sowie Last / Actual in T€ wird gemerkt und dem Projekt als Attribut mitgegeben  

                        With newProj
                            .actualDataUntil = newProj.startDate.AddMonths(monat - 1).AddDays(15)
                            .variantName = ""
                            .variantDescription = ""
                        End With


                        ' jetzt in die Import-Projekte eintragen 
                        updatedProjects = updatedProjects + 1
                        ImportProjekte.Add(newProj, updateCurrentConstellation:=False)

                    Else
                        ReDim logArray(4)
                        logArray(0) = " Projekt existiert nicht !!?? ... kann eigentlich nicht sein ..."
                        logArray(1) = ""
                        logArray(2) = vPKvP.Key
                        logArray(3) = ""
                        logArray(4) = ""

                        Call logfileSchreiben(logArray)
                    End If

                Next

                ' tk Test 
                ReDim logArray(3)
                logArray(0) = "Import von insgesamt " & updatedProjects & " Projekten (Gesamt-Euro): "
                logArray(1) = ""
                logArray(2) = ""
                logArray(3) = ""

                ReDim logDblArray(0)
                logDblArray(0) = gesamtIstValue
                Call logfileSchreiben(logArray, logDblArray)


            End With


        Catch ex As Exception
            ReDim logArray(1)
            logArray(0) = "Exception aufgetreten 100457: "
            logArray(1) = ex.Message
            Call logfileSchreiben(logArray)
            Throw New Exception("Fehler in Import-Datei Typ 3" & ex.Message)
        End Try


        logmessage = vbLf & "Zeilen gelesen: " & lastRow - 1 & vbLf &
                    "Projekte aktualisiert: " & updatedProjects
        outputCollection.Add(logmessage)

        logmessage = vbLf & "detailllierte Protokollierung LogFile ./requirements/logfile.xlsx"
        outputCollection.Add(logmessage)

        If outputCollection.Count > 0 Then
            Call showOutPut(outputCollection, "Import Ist-Daten", "")
        End If

    End Sub

    ''' <summary>
    ''' übernimmt für Projekte, die bislang noch keine Projekt-Nummern hatten, die Projekt-Nummer  
    ''' </summary>
    Public Sub importAllianzType4()
        Dim zeile As Integer, spalte As Integer

        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""


        Dim upDatedProjects As Integer = 0
        Dim errorProjects As Integer = 0

        ' für den Output 
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""


        Dim lastRow As Integer
        Dim geleseneProjekte As Integer
        Dim ok As Boolean = False

        ' die Projekte

        Dim hproj As clsProjekt = Nothing
        Dim projektKundenNummer As String = ""



        ' enthät die Saplte, wo der "alte" ProjektName steht ...
        Dim colPname As Integer = 4
        ' enthält die Spalte, wo die PRojekt-Nummer drin steht 

        ' enthält die Spalten-Nummer, wo die einzelnen Rollen-Namen zu finden sind
        Dim colPNr As Integer = 3


        ' jetzt werden die ImportProjekte zurückgesetzt ...
        ImportProjekte.Clear()

        Dim firstZeile As Excel.Range


        zeile = 2
        spalte = 1
        geleseneProjekte = 0


        Try

            Dim found As Boolean = False
            Dim wsi As Integer = 1
            Dim wsCount As Integer = appInstance.ActiveWorkbook.Worksheets.Count


            While Not found And wsi <= wsCount
                If CType(appInstance.ActiveWorkbook.Worksheets.Item(wsi),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet).Name.StartsWith("logBuch") Then
                    found = True
                Else
                    wsi = wsi + 1
                End If
            End While

            If Not found Then
                Call MsgBox("keine Projekte-Tabelle gefunden ...")
                Exit Sub
            End If

            Dim currentWS As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets.Item(wsi),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With currentWS


                firstZeile = CType(.Rows(2), Excel.Range)
                lastRow = CType(.Cells(20000, "D"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row



                While zeile < lastRow

                    Dim oldProj As clsProjekt = Nothing

                    Try

                        pName = CStr(CType(.Cells(zeile, colPname), Excel.Range).Value).Trim
                        projektKundenNummer = CStr(CType(.Cells(zeile, colPNr), Excel.Range).Value).Trim

                        If pName <> "" And projektKundenNummer <> "" Then
                            ' jetzt kann ggf der Update erfolgen ... 
                            geleseneProjekte = geleseneProjekte + 1
                            ok = isKnownProject(pName, projektKundenNummer, AlleProjekte)
                        End If


                    Catch ex As Exception
                        ok = False
                    End Try

                    If ok Then
                        hproj = getProjektFromSessionOrDB(pName, "", AlleProjekte, Date.Now)
                        If Not IsNothing(hproj) Then

                            If hproj.kundenNummer = "" Then
                                hproj.kundenNummer = projektKundenNummer
                                ' jetzt in die Import-Projekte eintragen 
                                upDatedProjects = upDatedProjects + 1
                                ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                            Else
                                outPutLine = "Projekt hat bereits eine Kunden-Nummer: " & pName & " old-Nr: " & hproj.kundenNummer & "; new-Nr: " & projektKundenNummer
                                outputCollection.Add(outPutLine)
                            End If



                        Else
                            outPutLine = "Projekt nicht in Datenbank gefunden: " & pName
                            outputCollection.Add(outPutLine)
                        End If
                    Else
                        outPutLine = "Projekt nicht bekannt: " & pName
                        outputCollection.Add(outPutLine)
                    End If

                    zeile = zeile + 1

                End While


            End With
        Catch ex As Exception

            Throw New Exception("Fehler in Import-Datei" & ex.Message)

        End Try

        If outputCollection.Count > 0 Then
            Call showOutPut(outputCollection, "Import Detail-Planungs Typ 4", "")
        End If

        Call MsgBox("Zeilen gelesen: " & geleseneProjekte & vbLf &
                    "Projekte aktualisiert: " & upDatedProjects)


    End Sub


    ''' <summary>
    ''' gibt true zurück, wenn die Spalten-Struktur dem erforderlichen Import Typ entspricht 
    ''' false sonst
    ''' </summary>
    ''' <param name="firstZeile"></param>
    ''' <param name="importTyp"></param>
    ''' <returns></returns>
    Private Function isCorrectAllianzImportStructure(ByVal firstZeile As Excel.Range, ByVal importTyp As Integer) As Boolean

        Dim tmpResult As Boolean = False

        Select Case importTyp
            Case 1
            Case 2
            Case 3
                Dim headerCheck() As String = {"Referat", "Projekttyp", "Projektnummer", "Projekt", "Vorgang/Aktivität", "Intern/Extern", "Ressource/Planungsebene", "Jahr", "Monat", "IST", "(PT)"}
                Dim colCheck() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 10}

                Try
                    tmpResult = True ' initiale Vorbesetzung 
                    Dim ix As Integer = 0
                    Do While tmpResult = True And ix <= headerCheck.Length - 1
                        tmpResult = tmpResult And CStr(CType(firstZeile.Cells(1, colCheck(ix)), Excel.Range).Value).Contains(headerCheck(ix))
                        ix = ix + 1
                    Loop

                Catch ex As Exception
                    tmpResult = False
                End Try


            Case Else
                ' nicht vorgesehen 
        End Select



        isCorrectAllianzImportStructure = tmpResult

    End Function


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ruleSet"></param>
    ''' <remarks></remarks>
    Public Sub awinReadAddOnRules(ByRef ruleSet As clsAddElements)

        Dim zeile As Integer, spalte As Integer
        Dim newName As String
        Dim duration As Integer = 0
        Dim isPhase As Boolean

        Dim referenceNameMS As String = ""
        Dim referenceNamePH As String = ""
        Dim refISStart As Boolean = True
        Dim abstandsRegel As String = ""
        Dim offset As Integer
        Dim deliverables As String = ""
        Dim newRule As clsAddElementRuleItem
        ' faktor = 1 bedeutet Tage; faktor = 7 bedeutet Wochen 
        Dim faktor As Integer = 1

        Dim lastRow As Integer

        Dim ok As Boolean = False

        Dim firstZeile As Excel.Range

        ' der Name des Rule-Sets wird später der Name der Phase, die ergänzt wird 
        Dim fileName As String = appInstance.ActiveWorkbook.Name
        Dim tmpName As String = ""

        ' bestimme den Namen des Szenarios - das ist gleich der Name der Excel Datei 
        Dim positionIX As Integer = fileName.IndexOf(".xls") - 1
        tmpName = ""
        For ih As Integer = 0 To positionIX
            tmpName = tmpName & fileName.Chars(ih)
        Next
        ruleSet.name = tmpName.Trim

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1

        Try
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Tabelle1"),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                firstZeile = CType(.Rows(1), Excel.Range)
                lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                While zeile <= lastRow
                    ok = False

                    Try
                        ' Name des neuen Elements lesen  
                        newName = CStr(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                        ' Dauer des neuen Elements lesen; bestimmt damit, ob es sich um eine Phase oder einen MEilenstein handelt
                        Try
                            duration = CInt(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                            If duration > 0 Then
                                isPhase = True
                            Else
                                isPhase = False
                            End If
                        Catch ex1 As Exception
                            duration = 0
                            isPhase = False
                        End Try

                        ' Ergebnisse des Meilensteins lesen 
                        deliverables = CStr(CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        If IsNothing(deliverables) Then
                            deliverables = ""
                        Else
                            If deliverables.Length > 0 Then
                                deliverables = deliverables.Trim
                            End If
                        End If

                        ' Rollenbedarfe der Phase lesen, spalte 4 

                        ' Kostenbedarfe der Phase lesen , spalte 5

                        ' Referenz-Name des Meilensteins lesen 
                        Try
                            referenceNameMS = CStr(CType(.Cells(zeile, 6), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                        Catch ex1 As Exception
                            referenceNameMS = ""
                        End Try


                        ' Referenz-Name der Phase  lesen 
                        Try
                            referenceNamePH = CStr(CType(.Cells(zeile, 7), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                        Catch ex1 As Exception
                            referenceNamePH = ""
                        End Try


                        ' Start oder Ende der Phase lesen 
                        Try
                            If CStr(CType(.Cells(zeile, 8), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim = "Ende" Then
                                refISStart = False
                            Else
                                refISStart = True
                            End If
                        Catch ex As Exception
                            refISStart = True
                        End Try

                        abstandsRegel = CStr(CType(.Cells(zeile, 9), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                        If abstandsRegel.EndsWith("w") Or tmpName.EndsWith("W") Then
                            faktor = 7
                        Else
                            faktor = 1
                        End If
                        Dim tmpstr() As String

                        tmpstr = abstandsRegel.Trim.Split(New Char() {CChar("w"), CChar("W"), CChar("d"), CChar("D")}, 5)
                        offset = CInt(tmpstr(0)) * faktor

                        ' wenn ein Meilenstein - Name angegeben wurde, wird jetzt die Regel für den Meilenstein angelegt
                        If referenceNameMS.Length > 0 Then
                            newRule = New clsAddElementRuleItem
                            With newRule
                                .newElemName = newName
                                .referenceName = referenceNameMS
                                .referenceIsPhase = False
                                .offset = offset
                            End With

                            If ruleSet.containsElement(newName, isPhase) Then
                                ruleSet.addRule(newRule, isPhase)
                            Else
                                Dim newElem As New clsAddElementRules(newName, isPhase, duration, deliverables)
                                ruleSet.addElem(newElem, isPhase)
                                ruleSet.addRule(newRule, isPhase)
                            End If

                        End If

                        '
                        If referenceNamePH.Length > 0 Then
                            newRule = New clsAddElementRuleItem
                            With newRule
                                .newElemName = newName
                                .referenceName = referenceNamePH
                                .referenceIsPhase = True
                                .referenceDateIsStart = refISStart
                                .offset = offset
                            End With

                            If ruleSet.containsElement(newName, isPhase) Then
                                ruleSet.addRule(newRule, isPhase)
                            Else
                                Dim newElem As New clsAddElementRules(newName, isPhase, duration, deliverables)
                                ruleSet.addElem(newElem, isPhase)
                                ruleSet.addRule(newRule, isPhase)
                            End If

                        End If


                    Catch ex As Exception

                    End Try

                    zeile = zeile + 1

                End While

            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei Module Import ...")
        End Try


    End Sub


    ''' <summary>
    ''' liest die im Diretory ../ressource manager liegenden detaillierten Kapa files zu den Rollen aus
    ''' und hinterlegt es an entsprechender Stelle im hrole.kapazitaet
    ''' wenn die Details als Rollen angelegt sind, dann werden diese Rollen gleich mitausgelesen 
    ''' </summary>
    ''' <param name="hrole"></param>
    ''' <remarks></remarks>
    Friend Sub readKapaOfRole(ByRef hrole As clsRollenDefinition)
        Dim kapaFileName As String
        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim summenZeile As Integer, extSummenZeile As Integer
        Dim spalte As Integer = 2
        Dim blattname As String = "Kapazität"
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim tmpDate As Date
        Dim tmpKapa As Double
        Dim extTmpKapa As Double
        Dim lastSpalte As Integer


        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        kapaFileName = awinPath & projektRessOrdner & "\" & hrole.name & " Kapazität.xlsx"


        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(kapaFileName) Then

            Try
                appInstance.Workbooks.Open(kapaFileName)
                ok = True

                Try

                    currentWS = CType(appInstance.Worksheets(blattname), Global.Microsoft.Office.Interop.Excel.Worksheet)
                    summenZeile = currentWS.Range("intern_sum").Row
                    lastSpalte = CType(currentWS.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column

                    ' bevor jetzt die eigentliche Kapa dieser Rolle aus intern_sum ausgelesen wird, wird geschaut, ob 
                    ' es eine zusammengesetzte Rolle ist
                    ' das wird dadurch entschieden, ob bis zur summenzeile bekannte Rollen auftauchen. Das sind dann die Sub-Roles 


                    Dim atleastOneSubRole As Boolean = False
                    Dim aktzeile As Integer = 2
                    Do While aktzeile < summenZeile

                        Dim subRoleName As String = CStr(CType(currentWS.Cells(aktzeile, spalte - 1), Excel.Range).Value)

                        If Not IsNothing(subRoleName) Then
                            subRoleName = subRoleName.Trim
                            If subRoleName.Length > 0 And RoleDefinitions.containsName(subRoleName) Then

                                Dim subRole As clsRollenDefinition = RoleDefinitions.getRoledef(subRoleName)

                                Try
                                    atleastOneSubRole = True
                                    ' es ist eine Sub-Rolle

                                    hrole.addSubRole(subRole.UID, subRoleName)

                                    spalte = 2
                                    tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)

                                    ' erstmal dahin positionieren, wo das Datum auch mit StartOfCalendar harmoniert 
                                    Do While DateDiff(DateInterval.Month, StartofCalendar, tmpDate) < 0 And spalte <= lastSpalte
                                        spalte = spalte + 1
                                        tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                                    Loop

                                    Do While spalte < 241 And spalte <= lastSpalte

                                        index = getColumnOfDate(tmpDate)
                                        If index >= 1 Then
                                            tmpKapa = CDbl(CType(currentWS.Cells(aktzeile, spalte), Excel.Range).Value)

                                            If index <= 240 And index > 0 And tmpKapa >= 0 Then
                                                subRole.kapazitaet(index) = tmpKapa
                                            End If
                                        End If

                                        spalte = spalte + 1
                                        tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                                    Loop

                                Catch ex As Exception

                                End Try


                            End If

                        End If

                        aktzeile = aktzeile + 1
                        ' jetzt spalte wieder auf 2 setzen 
                        spalte = 2
                    Loop

                    ' die internen Kapas einer Sammelrolle sind NULL 
                    If atleastOneSubRole Then
                        For i As Integer = 1 To 240
                            hrole.kapazitaet(i) = 0
                        Next
                    End If



                    Try
                        extSummenZeile = currentWS.Range("extern_sum").Row
                    Catch ex As Exception
                        extSummenZeile = 0
                    End Try

                    tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)

                    Do While DateDiff(DateInterval.Month, StartofCalendar, tmpDate) > 0 And
                            spalte < 241 And spalte <= lastSpalte
                        index = getColumnOfDate(tmpDate)
                        tmpKapa = CDbl(CType(currentWS.Cells(summenZeile, spalte), Excel.Range).Value)


                        If extSummenZeile > 0 Then
                            extTmpKapa = CDbl(CType(currentWS.Cells(extSummenZeile, spalte), Excel.Range).Value)
                        Else
                            extTmpKapa = 0.0
                        End If

                        If index <= 240 And index > 0 Then

                            If atleastOneSubRole Then
                                ' alles ist Null , wird erst später aufgrund der Sub-Rollen berechnet 
                            Else
                                If tmpKapa >= 0 Then
                                    hrole.kapazitaet(index) = tmpKapa
                                End If
                            End If

                            If extTmpKapa >= 0 Then
                                'hrole.externeKapazitaet(index) = extTmpKapa
                            End If


                        End If

                        spalte = spalte + 1
                        tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                    Loop

                Catch ex2 As Exception

                End Try

                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
            Catch ex As Exception

            End Try

        End If


        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' liest die externen Verträge gemäß Allianz Struktur 
    ''' </summary>
    ''' <param name="meldungen"></param>
    Public Sub readMonthlyExternKapasEV(ByRef meldungen As Collection)

        Dim kapaFolder As String


        Dim ok As Boolean = True

        Dim spalte As Integer = 2
        Dim blattname As String = "Werte in Euro"
        Dim currentWS As Excel.Worksheet = Nothing

        Dim errMsg As String = ""
        Dim anzFehler As Integer = 0

        Dim aktzeile As Integer = 1
        Dim saveNeeded As Boolean = False


        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        kapaFolder = awinPath & projektRessOrdner

        Try
            Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(kapaFolder)

            For i = 0 To listOfImportfiles.Count - 1

                Dim dateiName As String = My.Computer.FileSystem.CombinePath(kapaFolder, listOfImportfiles.Item(i))

                If Not IsNothing(dateiName) Then

                    If My.Computer.FileSystem.FileExists(dateiName) And dateiName.Contains("Extern") Then

                        errMsg = "Reading external Capacities " & dateiName
                        Call logfileSchreiben(errMsg, "", anzFehler)

                        Try
                            appInstance.Workbooks.Open(dateiName)
                            ok = True

                            Try

                                currentWS = CType(appInstance.Worksheets(blattname), Global.Microsoft.Office.Interop.Excel.Worksheet)

                                Dim colRessource As Integer = 8
                                Dim colBeginn As Integer = 9
                                Dim colEnde As Integer = 10
                                Dim colVV As Integer = 15

                                Dim lastRow As Integer = CType(currentWS.Cells(16000, "H"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                                ' jetzt wird Zeile für Zeile nachgesehen, ob das eine Basic Role ist und dann die Kapas besetzt 

                                aktzeile = 3
                                Do While aktzeile <= lastRow

                                    Dim subRoleName As String = CStr(CType(currentWS.Cells(aktzeile, colRessource), Excel.Range).Value)

                                    If Not IsNothing(subRoleName) Then
                                        subRoleName = subRoleName.Trim
                                        If subRoleName.Length > 0 And RoleDefinitions.containsName(subRoleName) Then

                                            Dim subRole As clsRollenDefinition = RoleDefinitions.getRoledef(subRoleName)

                                            ' nur weiter machen, wenn es keine SummenRolle ist ... und es ausserdem einen Tagessatz gibt ..
                                            ' weil andernfalls Dividion durch Null passieren würde 
                                            If Not subRole.isCombinedRole And subRole.tagessatzIntern > 0 Then

                                                ' lese das Vertragsvolumen
                                                Try
                                                    Dim vertragsVolumen As Double = 0.0
                                                    If Not IsNothing(CType(currentWS.Cells(aktzeile, colVV), Excel.Range).Value) Then
                                                        vertragsVolumen = CDbl(CType(currentWS.Cells(aktzeile, colVV), Excel.Range).Value)
                                                    End If

                                                    Dim startDate As Date = CDate(CType(currentWS.Cells(aktzeile, colBeginn), Excel.Range).Value)
                                                    Dim endeDate As Date = CDate(CType(currentWS.Cells(aktzeile, colEnde), Excel.Range).Value)

                                                    If vertragsVolumen >= 0 Then
                                                        Dim dimension As Integer = getColumnOfDate(endeDate) - getColumnOfDate(startDate)
                                                        Dim vorgabeArray(0) As Double
                                                        vorgabeArray(0) = vertragsVolumen / subRole.tagessatzIntern
                                                        Dim volumenArray() As Double = calcVerteilungAufMonate(startDate, endeDate, vorgabeArray, 1.0)

                                                        Dim startCol As Integer = getColumnOfDate(startDate)
                                                        For ix As Integer = 0 To volumenArray.Length - 1
                                                            If ix + startCol <= 240 And ix + startCol > 0 And volumenArray(ix) >= 0 Then
                                                                subRole.kapazitaet(ix + startCol) = volumenArray(ix)
                                                            End If
                                                        Next

                                                    End If
                                                Catch ex As Exception

                                                End Try


                                            Else
                                                If subRole.isCombinedRole Then
                                                    errMsg = "File " & dateiName & ": " & subRoleName & " is combinedRole; combinedRoles are calculated automatically"
                                                    meldungen.Add(errMsg)
                                                ElseIf subRole.tagessatzIntern <= 0 Then
                                                    errMsg = "File " & dateiName & ": " & subRoleName & " no dayrate / tagessatz available "
                                                    meldungen.Add(errMsg)
                                                End If

                                                Call logfileSchreiben(errMsg, "", anzFehler)
                                            End If
                                        Else
                                            If subRoleName.Length > 0 Then
                                                errMsg = "File " & dateiName & ": " & subRoleName & " does not exist ..."
                                                meldungen.Add(errMsg)
                                                Call logfileSchreiben(errMsg, "", anzFehler)
                                            End If
                                        End If

                                    End If

                                    aktzeile = aktzeile + 1
                                    ' jetzt spalte wieder auf 2 setzen 
                                    spalte = 2
                                Loop

                            Catch ex2 As Exception
                                errMsg = "File " & dateiName & ": Fehler / Error  ... " & vbLf & ex2.Message
                                meldungen.Add(errMsg)
                                Call logfileSchreiben(errMsg, "", anzFehler)

                                If Not IsNothing(currentWS) Then
                                    CType(currentWS.Cells(aktzeile, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed
                                    saveNeeded = True
                                End If

                            End Try

                            appInstance.ActiveWorkbook.Close(SaveChanges:=saveNeeded)
                        Catch ex As Exception
                            appInstance.ActiveWorkbook.Close(SaveChanges:=saveNeeded)
                        End Try

                    End If

                End If


            Next i

        Catch ex As Exception

        End Try


        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True


    End Sub


    ''' <summary>
    ''' liest alle Dateien mit Kapazität und weist den Rollen die Kapa zu 
    ''' es werden nur Personen ausgelesen ! alle anderen werden ignoriert ...
    ''' </summary>
    Public Sub readMonthlyExternKapas(ByRef meldungen As Collection)

        Dim kapaFolder As String


        Dim ok As Boolean = True

        Dim summenZeile As Integer
        Dim spalte As Integer = 2
        Dim blattname As String = "Kapazität"
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim tmpDate As Date
        Dim tmpKapa As Double
        Dim lastSpalte As Integer
        Dim errMsg As String = ""


        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        kapaFolder = awinPath & projektRessOrdner

        Try
            Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(kapaFolder)

            For i = 0 To listOfImportfiles.Count - 1

                Dim dateiName As String = My.Computer.FileSystem.CombinePath(kapaFolder, listOfImportfiles.Item(i))

                If Not IsNothing(dateiName) Then

                    If My.Computer.FileSystem.FileExists(dateiName) And dateiName.Contains("Kapazität") Then

                        Try
                            appInstance.Workbooks.Open(dateiName)
                            ok = True

                            Try

                                currentWS = CType(appInstance.Worksheets(blattname), Global.Microsoft.Office.Interop.Excel.Worksheet)
                                Try
                                    summenZeile = currentWS.Range("intern_sum").Row
                                Catch ex As Exception
                                    ' wenn die Summenzeile nicht existiert, gehe ich davon aus, dass einfach jede Zeile ausgelesen werden soll 
                                    summenZeile = 0
                                    summenZeile = CType(currentWS.Cells(12000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row + 1
                                End Try

                                lastSpalte = CType(currentWS.Cells(1, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column

                                ' jetzt wird Zeile für Zeile nachgesehen, ob das eine Basic Role ist und dann die Kapas besetzt 

                                Dim aktzeile As Integer = 2
                                Do While aktzeile < summenZeile

                                    Dim subRoleName As String = CStr(CType(currentWS.Cells(aktzeile, 1), Excel.Range).Value)

                                    If Not IsNothing(subRoleName) Then
                                        subRoleName = subRoleName.Trim
                                        If subRoleName.Length > 0 And RoleDefinitions.containsName(subRoleName) Then

                                            Dim subRole As clsRollenDefinition = RoleDefinitions.getRoledef(subRoleName)

                                            ' nur weiter machen, wenn es keine SummenRollen ist ...
                                            If Not subRole.isCombinedRole Then

                                                Try
                                                    spalte = 2
                                                    tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)

                                                    ' erstmal dahin positionieren, wo das Datum auch mit oder nach StartOfCalendar beginnt  

                                                    Do While DateDiff(DateInterval.Month, StartofCalendar, tmpDate) < 0 And spalte <= lastSpalte
                                                        Try
                                                            spalte = spalte + 1
                                                            tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                                                        Catch ex As Exception

                                                        End Try
                                                    Loop

                                                    Do While spalte < 241 And spalte <= lastSpalte

                                                        Try
                                                            index = getColumnOfDate(tmpDate)
                                                            If index >= 1 Then
                                                                tmpKapa = CDbl(CType(currentWS.Cells(aktzeile, spalte), Excel.Range).Value)

                                                                If index <= 240 And index > 0 And tmpKapa >= 0 Then
                                                                    subRole.kapazitaet(index) = tmpKapa
                                                                End If
                                                            End If

                                                            spalte = spalte + 1
                                                            tmpDate = CDate(CType(currentWS.Cells(1, spalte), Excel.Range).Value)
                                                        Catch ex As Exception
                                                            errMsg = "File " & dateiName & ": error when setting value for " & subRoleName & " in row, column: " & aktzeile & ", " & spalte
                                                            meldungen.Add(errMsg)
                                                        End Try


                                                    Loop

                                                Catch ex As Exception

                                                End Try
                                            Else
                                                errMsg = "File " & dateiName & ": " & subRoleName & " is combinedRole; combinedRoles are calculated automatically"
                                                meldungen.Add(errMsg)
                                            End If
                                        Else
                                            If subRoleName.Length > 0 Then
                                                errMsg = "File " & dateiName & ": " & subRoleName & " does not exist ..."
                                                meldungen.Add(errMsg)
                                            End If
                                        End If

                                    End If

                                    aktzeile = aktzeile + 1
                                    ' jetzt spalte wieder auf 2 setzen 
                                    spalte = 2
                                Loop

                            Catch ex2 As Exception
                                errMsg = "File " & dateiName & ": unidentified error ... "
                                meldungen.Add(errMsg)
                            End Try

                            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                        Catch ex As Exception
                            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                        End Try

                    End If

                End If


            Next i

        Catch ex As Exception

        End Try


        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True

    End Sub


    ''' <summary>
    ''' liest das im Diretory ../ressource manager evt. liegende File 'Urlaubsplaner*.xlsx' File  aus
    ''' und hinterlegt an entsprechender Stelle im hrole.kapazitaet die verfügbaren Tage der entsprechenden Rolle
    ''' </summary>
    ''' <remarks></remarks>
    Friend Sub readAvailabilityOfRole(ByVal kapaFileName As String, ByRef oPCollection As Collection)

        Dim err As New clsErrorCodeMsg

        Dim ok As Boolean = True
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim msgtxt As String = ""
        Dim anzFehler As Integer = 0
        Dim fehler As Boolean = False

        Dim kapaWB As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim spalte As Integer = 2
        Dim firstUrlspalte As Integer = 5
        Dim noColor As Integer = -4142
        Dim whiteColor As Integer = 2
        Dim currentWS As Excel.Worksheet
        Dim index As Integer
        Dim tmpDate As Date

        Dim year As Integer = DatePart(DateInterval.Year, Date.Now)
        Dim anzMonthDays As Integer = 0
        Dim colDate As Integer = 0
        Dim anzDays As Integer = 0

        Dim lastZeile As Integer
        Dim lastSpalte As Integer
        Dim monthDays As New SortedList(Of Integer, Integer)

        Dim hrole As New clsRollenDefinition
        Dim rolename As String = ""

        Dim outPutCollection As New Collection

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        ' öffnen des Files 
        If My.Computer.FileSystem.FileExists(kapaFileName) Then

            Try
                kapaWB = appInstance.Workbooks.Open(kapaFileName)

                Try
                    For index = 1 To appInstance.Worksheets.Count

                        'If Not ok Then
                        '    Exit For
                        'End If


                        currentWS = CType(appInstance.Worksheets(index), Global.Microsoft.Office.Interop.Excel.Worksheet)
                        Dim hstr() As String = Split(currentWS.Name, "Halbjahr", , )
                        If hstr.Length > 1 Then

                            ok = True
                            ' Auslesen der Jahreszahl, falls vorhanden
                            If Not IsNothing(CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                year = CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).Value
                            End If

                            monthDays.Clear()
                            anzDays = 0


                            lastSpalte = CType(currentWS.Cells(4, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlToLeft).Column
                            lastZeile = CType(currentWS.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row

                            ' letzte Zeile bestimmen, wenn dies verbunden Zellen sind
                            ' -------------------------------------
                            Dim rng As Range
                            Dim rngEnd As Range

                            rng = CType(currentWS.Cells(lastZeile, 1), Global.Microsoft.Office.Interop.Excel.Range)

                            If rng.MergeCells Then

                                rng = rng.MergeArea
                                rngEnd = rng.Cells(rng.Rows.Count, rng.Columns.Count)

                                ' dann ist die lastZeile neu zu besetzen
                                lastZeile = rngEnd.Row
                            End If

                            ' nun hat die Variable lastZeile sicher den richtigen Wert
                            ' --------------------------------------


                            Dim vglColor As Integer = noColor         ' keine Farbe
                            Dim i As Integer = firstUrlspalte

                            While ok And i <= lastSpalte

                                If vglColor <> CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex Then
                                    ok = (anzDays = anzMonthDays) Or (anzDays = 0)
                                    vglColor = CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex
                                    anzDays = 1
                                Else
                                    If CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Text <> "" Then
                                        Dim monthName As String = CType(currentWS.Cells(1, i), Global.Microsoft.Office.Interop.Excel.Range).Text
                                        ' ''Dim strDate As String = "01." & monthName & " " & year
                                        ' ''Dim hdate As DateTime = DateValue(strDate)

                                        Dim isdate As Boolean = DateTime.TryParse(monthName & " " & year.ToString, tmpDate)
                                        If isdate Then
                                            colDate = getColumnOfDate(tmpDate)
                                            anzMonthDays = DateTime.DaysInMonth(year, Month(tmpDate))
                                            monthDays.Add(colDate, anzMonthDays)
                                        End If
                                    End If

                                    anzDays = anzDays + 1
                                End If

                                i = i + 1
                            End While


                            If Not ok Then

                                fehler = True

                                If awinSettings.englishLanguage Then
                                    msgtxt = "Error reading planning holidays: Please check the calendar in this file ..."
                                Else
                                    msgtxt = "Fehler beim Lesen der Urlaubsplanung: Bitte prüfen Sie die Korrektheit des Kalenders ..."
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If
                                'Call MsgBox(msgtxt)

                                Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)

                                If formerEE Then
                                    appInstance.EnableEvents = True
                                End If

                                If formerSU Then
                                    appInstance.ScreenUpdating = True
                                End If

                                enableOnUpdate = True
                                If awinSettings.englishLanguage Then
                                    msgtxt = "Your planning holidays couldn't be read, because of problems"
                                Else
                                    msgtxt = "Ihre Urlaubsplanung konnte nicht berücksichtigt werden"
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If

                                Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                'Call showOutPut(oPCollection, "Lesen Urlaubsplanung wurde mit Fehler abgeschlossen", "Meldungen zu Lesen Urlaubsplanung")
                                ' tk 12.2.19 ess oll alles gelesen werden - es wird nicht weitergemacht, wenn es Einträge in der outputCollection gibt 
                                'Throw New ArgumentException(msgtxt)
                            Else

                                For iZ = 5 To lastZeile


                                    rolename = CType(currentWS.Cells(iZ, 2), Global.Microsoft.Office.Interop.Excel.Range).Text
                                    If rolename <> "" Then
                                        hrole = RoleDefinitions.getRoledef(rolename)
                                        If Not IsNothing(hrole) Then

                                            Dim defaultHrsPerdayForThisPerson As Double = 8 * hrole.defaultKapa / nrOfDaysMonth

                                            Dim iSp As Integer = firstUrlspalte
                                            Dim anzArbTage As Double = 0
                                            Dim anzArbStd As Double = 0

                                            For Each kvp As KeyValuePair(Of Integer, Integer) In monthDays

                                                Dim colOfDate As Integer = kvp.Key
                                                anzDays = kvp.Value
                                                For sp = iSp + 0 To iSp + anzDays - 1

                                                    If iSp <= lastSpalte Then
                                                        Dim hint As Integer = CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex)

                                                        If CInt(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex) = noColor _
                                                            Or CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Interior.ColorIndex = whiteColor Then

                                                            If Not IsNothing(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                If IsNumeric(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                                                    Dim angabeInStd As Double = CType(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value, Double)

                                                                    If angabeInStd >= 0 And angabeInStd <= 24 Then
                                                                        anzArbStd = anzArbStd + CDbl(CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                                                    Else
                                                                        If awinSettings.englishLanguage Then
                                                                            msgtxt = "Error reading the amount of working hours for " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                        Else
                                                                            msgtxt = "Fehler beim Lesen der Anzahl zu leistenden Arbeitsstunden " & hrole.name & " : " & angabeInStd.ToString & " (!!)"
                                                                        End If
                                                                        If Not oPCollection.Contains(msgtxt) Then
                                                                            oPCollection.Add(msgtxt, msgtxt)
                                                                        End If
                                                                        'Call MsgBox(msgtxt)
                                                                        fehler = True
                                                                        Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                                    End If
                                                                Else
                                                                    ' Feld ist weiss, oder hat keine Farbe, keine Zahl: also ist es Arbeitstag mit Default-Std pro Tag 
                                                                    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                End If



                                                            Else

                                                                ' hier wird die Telair Variante gemacht 
                                                                ' das einfachste wäre eigentlich  
                                                                'anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson

                                                                Dim colorIndup As Integer = CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex

                                                                ' Wenn das Feld nicht durch einen Diagonalen Strich gekennzeichnet ist
                                                                If CType(currentWS.Cells(iZ, sp), Global.Microsoft.Office.Interop.Excel.Range).Borders(XlBordersIndex.xlDiagonalUp).ColorIndex = noColor Then
                                                                    'anzArbStd = anzArbStd + 8
                                                                    anzArbStd = anzArbStd + defaultHrsPerdayForThisPerson
                                                                Else
                                                                    ' freier Tag für Teilzeitbeschäftigte
                                                                    msgtxt = "Tag zählt nicht: Zeile " & iZ & ", Spalte " & sp
                                                                    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                                End If

                                                            End If
                                                        End If
                                                    Else
                                                        If awinSettings.englishLanguage Then
                                                            msgtxt = "Error reading the amount of working days of " & hrole.name & " ..."
                                                        Else
                                                            msgtxt = "Fehler beim Lesen der verfügbaren Arbeitstage von " & hrole.name & " ..."
                                                        End If
                                                        fehler = True
                                                        If Not oPCollection.Contains(msgtxt) Then
                                                            oPCollection.Add(msgtxt, msgtxt)
                                                        End If
                                                        Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                                    End If

                                                Next

                                                anzArbTage = anzArbStd / 8
                                                hrole.kapazitaet(colOfDate) = anzArbTage
                                                iSp = iSp + anzDays
                                                anzArbTage = 0              ' Anzahl Arbeitstage wieder zurücksetzen für den nächsten Monat
                                                anzArbStd = 0               ' Anzahl zu leistender Arbeitsstunden wieder zurücksetzen für den nächsten Monat

                                            Next

                                        Else

                                            If awinSettings.englishLanguage Then
                                                msgtxt = "Role " & rolename & " not defined ..."
                                            Else
                                                msgtxt = "Rolle " & rolename & " nicht definiert ..."
                                            End If
                                            If Not oPCollection.Contains(msgtxt) Then
                                                oPCollection.Add(msgtxt, msgtxt)
                                            End If
                                            'Call MsgBox(msgtxt)
                                            fehler = True
                                            Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                        End If
                                    Else

                                        If awinSettings.englishLanguage Then
                                            msgtxt = "No Name of role given ..."
                                        Else
                                            msgtxt = "kein Rollenname angegeben ..."
                                        End If
                                        If Not oPCollection.Contains(msgtxt) Then
                                            oPCollection.Add(msgtxt, msgtxt)
                                        End If
                                        Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                                    End If

                                Next iZ

                            End If   ' ende von if not OK
                        Else
                            If awinSettings.visboDebug Then

                                If awinSettings.englishLanguage Then
                                    msgtxt = "Worksheet " & hstr(0) & "doesn't belongs to planning holidays ..."
                                Else
                                    msgtxt = "Worksheet" & hstr(0) & " gehört nicht zum Urlaubsplaner ..."
                                End If
                                If Not oPCollection.Contains(msgtxt) Then
                                    oPCollection.Add(msgtxt, msgtxt)
                                End If
                                Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                            End If

                        End If

                    Next index


                Catch ex2 As Exception
                    'If fehler Then
                    '    'Call MsgBox(msgtxt)

                    '    RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(DateTime.Now, err)

                    '    msgtxt = "Es wurden nun die Kapazitäten aus der Datenbank gelesen ..."
                    '    If awinSettings.englishLanguage Then
                    '        msgtxt = "Therefore read the capacity of every Role from the DB  ..."
                    '    End If
                    '    If Not oPCollection.Contains(msgtxt) Then
                    '        oPCollection.Add(msgtxt, msgtxt)
                    '    End If
                    '    Call logfileSchreiben(msgtxt, kapaFileName, anzFehler)
                    'End If
                End Try

                'kapaWB.Close(SaveChanges:=False)
            Catch ex As Exception

            End Try

        End If


        If formerEE Then
            appInstance.EnableEvents = True
        End If

        If formerSU Then
            appInstance.ScreenUpdating = True
        End If

        enableOnUpdate = True
        kapaWB.Close(SaveChanges:=False)

        ' das wird jetzt an der übergeordneten Stelle gemacht
        'Call showOutPut(oPCollection, "Meldungen zu Lesen Urlaubsplanung", "Folgende Probleme sind beim Lesen der Urlaubsplanung aufgetreten")

        ' ''If outPutCollection.Count > 0 Then
        ' ''    Call showOutPut(outPutCollection, _
        ' ''                    "Meldungen Einlesevorgang Urlaubsdatei", _
        ' ''                    "zum Zeitpunkt " & storedAtOrBefore.ToString & " aufgeführte Rolle nicht definiert")
        ' ''End If


    End Sub

    ''' <summary>
    ''' liest die Name-Mapping Definitionen der Phasen bzw Meilensteine ein
    ''' </summary>
    ''' <param name="ws">Worksheet, in dem die Mappings stehen </param>
    ''' <param name="mappings">Klassen-Instanz, die die Mappings aufnimmt</param>
    ''' <remarks></remarks>
    Friend Sub readNameMappings(ByVal ws As Excel.Worksheet, ByRef mappings As clsNameMapping)

        Dim zeile As Integer, spalte As Integer

        With ws

            ' auslesen der Synonyme und Regular Expressions in Spalte 1, beginnend mit Zeile 3
            Dim ok As Boolean = False
            zeile = 3
            spalte = 1
            If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) And
                Not IsNothing(CType(.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value) Then
                If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 And
                    CStr(.Cells(zeile, spalte).offset(0, 1).Value).Trim.Length > 0 Then
                    ok = True
                End If
            End If

            Dim syn As String, stdName As String
            Do While ok

                syn = CStr(.Cells(zeile, spalte).Value).Trim
                stdName = CStr(.Cells(zeile, spalte).offset(0, 1).Value).Trim

                Dim regExpression As String = ""
                Dim isRegExpression As Boolean = False

                If syn.StartsWith("[") And syn.EndsWith("]") Then
                    isRegExpression = True
                    For i As Integer = 1 To syn.Length - 2
                        regExpression = regExpression & syn.Chars(i)
                    Next
                End If

                Try
                    If isRegExpression Then
                        mappings.addRegExpressName(regExpression, stdName)
                    Else
                        mappings.addSynonym(syn, stdName)
                    End If


                Catch ex As Exception

                End Try


                zeile = zeile + 1
                ok = False

                If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) And
                Not IsNothing(CType(.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value) Then
                    If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 And
                        CStr(.Cells(zeile, spalte).offset(0, 1).Value).Trim.Length > 0 Then
                        ok = True
                    End If
                End If
            Loop


            ' auslesen der Hierarchies Namens in Spalte 4, beginnend mit Zeile 3
            ok = False
            zeile = 3
            spalte = 4
            If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) Then
                If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 Then
                    ok = True
                End If
            End If

            Dim NameToC As String
            Do While ok

                NameToC = CStr(.Cells(zeile, spalte).Value).Trim

                Try

                    mappings.addNameToComplement(NameToC)

                Catch ex As Exception

                End Try


                zeile = zeile + 1
                ok = False

                If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) Then
                    If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 Then
                        ok = True
                    End If
                End If

            Loop

            ' auslesen der To-Ignore-Names in Spalte 6, beginnend mit Zeile 3
            ok = False
            zeile = 3
            spalte = 6
            If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) Then
                If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 Then
                    ok = True
                End If
            End If

            Dim ignoreName As String
            Do While ok

                ignoreName = CStr(.Cells(zeile, spalte).Value).Trim

                Try

                    mappings.addIgnoreName(ignoreName)

                Catch ex As Exception

                End Try


                zeile = zeile + 1
                ok = False

                If Not IsNothing(CType(.Cells(zeile, spalte), Excel.Range).Value) Then
                    If CStr(.Cells(zeile, spalte).Value).Trim.Length > 0 Then
                        ok = True
                    End If
                End If

            Loop

        End With

    End Sub


    ''' <summary>
    ''' initialisert im Inputfile die Tabelle 'Logbuch'
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Sub InitProtokoll(ByRef wslogbuch As Excel.Worksheet, ByVal tabblattname As String)

        ' diese Variable sagt, ob das Tabellenblatt existiert hat; wenn nein, müssen die Spalten-Breiten gesetzt werden 
        Dim didntExist As Boolean

        Try
            wslogbuch = CType(xlsLogfile.Worksheets(tabblattname),
               Global.Microsoft.Office.Interop.Excel.Worksheet)


            If Not IsNothing(wslogbuch) Then

                ' Änderung tk: 16.1.16
                ' es reicht die Inhalte zu löschen ...  
                wslogbuch.Cells.Clear()
                didntExist = False
                'xlsLogfile.Worksheets.Application.DisplayAlerts = False
                'wslogbuch.Delete()
                'xlsLogfile.Worksheets.Application.DisplayAlerts = True

                'wslogbuch = CType(xlsLogfile.Worksheets.Add(), _
                '   Global.Microsoft.Office.Interop.Excel.Worksheet)
                'wslogbuch.Name = tabblattname
            End If
        Catch ex As Exception
            'wsLogbuch = CType(xlsInput.Worksheets.Add(After:=xlsInput.Worksheets.Count), _
            '   Global.Microsoft.Office.Interop.Excel.Worksheet)
            wslogbuch = CType(xlsLogfile.Worksheets.Add(),
                Global.Microsoft.Office.Interop.Excel.Worksheet)
            wslogbuch.Name = tabblattname
            didntExist = True
        End Try


        With wslogbuch

            If didntExist Then
                .Rows.RowHeight = 15
                CType(.Rows(1), Excel.Range).RowHeight = 30
                CType(.Rows(1), Excel.Range).Font.Bold = True
            End If


            If awinSettings.fullProtocol Then
                CType(.Cells(1, 1), Excel.Range).Value() = "Datum"
                CType(.Cells(1, 2), Excel.Range).Value() = "Projekt"
                CType(.Cells(1, 3), Excel.Range).Value() = "Hierarchie"
                CType(.Cells(1, 4), Excel.Range).Value() = "Plan-Element"
                CType(.Cells(1, 5), Excel.Range).Value() = "Klasse"
                CType(.Cells(1, 6), Excel.Range).Value() = "Abkürzung"
                CType(.Cells(1, 7), Excel.Range).Value() = "Quelle"
                CType(.Cells(1, 8), Excel.Range).Value() = "Übernommen als"
                CType(.Cells(1, 9), Excel.Range).Value() = "Grund"
                CType(.Cells(1, 10), Excel.Range).Value() = "PT Hierarchie"
                CType(.Cells(1, 11), Excel.Range).Value() = "PT Klasse"

                ' nur verändern, wenn es nicht vorher schon existiert hat ... 
                ' falls der Anwender sich die Breiten so hingerichtet hat , wie er es gerne hätte, 
                ' sollte das nicht verändert werden 
                If didntExist Then
                    CType(.Columns(1), Excel.Range).ColumnWidth = 10
                    CType(.Columns(2), Excel.Range).ColumnWidth = 40
                    CType(.Columns(3), Excel.Range).ColumnWidth = 40
                    CType(.Columns(4), Excel.Range).ColumnWidth = 40
                    CType(.Columns(5), Excel.Range).ColumnWidth = 40
                    CType(.Columns(6), Excel.Range).ColumnWidth = 40
                    CType(.Columns(7), Excel.Range).ColumnWidth = 40
                    CType(.Columns(8), Excel.Range).ColumnWidth = 40
                    CType(.Columns(9), Excel.Range).ColumnWidth = 40
                    CType(.Columns(10), Excel.Range).ColumnWidth = 40
                    CType(.Columns(11), Excel.Range).ColumnWidth = 40
                End If

            Else
                CType(.Cells(1, 1), Excel.Range).Value() = "Datum"
                CType(.Cells(1, 2), Excel.Range).Value() = "Projekt"
                CType(.Cells(1, 4), Excel.Range).Value() = "Plan-Element"
                CType(.Cells(1, 8), Excel.Range).Value() = "Übernommen als"
                CType(.Cells(1, 9), Excel.Range).Value() = "Grund"

                ' nur verändern, wenn es nicht vorher schon existiert hat ... 
                ' falls der Anwender sich die Breiten so hingerichtet hat , wie er es gerne hätte, 
                ' sollte das nicht verändert werden 
                If didntExist Then
                    CType(.Columns(1), Excel.Range).ColumnWidth = 18
                    CType(.Columns(2), Excel.Range).ColumnWidth = 35
                    CType(.Columns(3), Excel.Range).ColumnWidth = 5
                    CType(.Columns(4), Excel.Range).ColumnWidth = 40
                    CType(.Columns(5), Excel.Range).ColumnWidth = 10
                    CType(.Columns(6), Excel.Range).ColumnWidth = 10
                    CType(.Columns(7), Excel.Range).ColumnWidth = 10
                    CType(.Columns(8), Excel.Range).ColumnWidth = 40
                    CType(.Columns(9), Excel.Range).ColumnWidth = 40
                End If

            End If

        End With

    End Sub

    ''' <summary>
    ''' schreibt das Protokoll in das Tabellenblatt
    ''' es wird eine Range definiert, die soviele Zeilen enthält wie öt 
    ''' </summary>
    ''' <param name="prtliste"></param>
    ''' <param name="tabblattname"></param>
    ''' <remarks></remarks>
    Sub writeProtokoll(ByRef prtliste As SortedList(Of Integer, clsProtokoll), ByVal tabblattname As String)

        Dim zelle As Excel.Range = Nothing
        Dim zeile As Integer

        Dim anzZeilen As Integer = prtliste.Count

        Dim wsLogbuch As Excel.Worksheet = Nothing

        Try
            Call InitProtokoll(wsLogbuch, tabblattname) ' Tabelle Logbuch wird initialisiert
            If Not IsNothing(xlsLogfile) Then
                xlsLogfile.Save()
            End If


        Catch ex As Exception

            Call MsgBox("Fehler beim Initialisieren des Protokolls")
        End Try

        Dim protokollRange As Excel.Range = wsLogbuch.Cells


        For Each prtline As KeyValuePair(Of Integer, clsProtokoll) In prtliste
            Try
                'rowOffset = CType(CType(xlsLogfile.Worksheets(Me.tabblattname), Excel.Worksheet).Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                zeile = prtline.Key

                With wsLogbuch

                    ' Änderung tk: das dauert sehr lange ... 
                    'If awinSettings.fullProtocol Then
                    '    CType(.Cells(zeile, 1), Excel.Range).Value() = prtline.Value.actDate
                    '    CType(.Cells(zeile, 2), Excel.Range).Value() = prtline.Value.Projekt
                    '    CType(.Cells(zeile, 3), Excel.Range).Value() = prtline.Value.hierarchie
                    '    CType(.Cells(zeile, 4), Excel.Range).Value() = prtline.Value.planelement
                    '    CType(.Cells(zeile, 5), Excel.Range).Value() = prtline.Value.klasse
                    '    CType(.Cells(zeile, 6), Excel.Range).Value() = prtline.Value.abkürzung
                    '    CType(.Cells(zeile, 7), Excel.Range).Value() = prtline.Value.quelle
                    '    CType(.Cells(zeile, 8), Excel.Range).Value() = prtline.Value.planeleÜbern
                    '    CType(.Cells(zeile, 8), Excel.Range).Interior.Color = prtline.Value.hgColor
                    '    CType(.Cells(zeile, 9), Excel.Range).Value() = prtline.Value.grund
                    '    CType(.Cells(zeile, 10), Excel.Range).Value() = prtline.Value.PThierarchie
                    '    CType(.Cells(zeile, 11), Excel.Range).Value() = prtline.Value.PTklasse
                    'Else
                    '    CType(.Cells(zeile, 1), Excel.Range).Value() = prtline.Value.actDate
                    '    CType(.Cells(zeile, 2), Excel.Range).Value() = prtline.Value.Projekt
                    '    CType(.Cells(zeile, 4), Excel.Range).Value() = prtline.Value.planelement
                    '    CType(.Cells(zeile, 8), Excel.Range).Value() = prtline.Value.planeleÜbern
                    '    CType(.Cells(zeile, 8), Excel.Range).Interior.Color = prtline.Value.hgColor
                    '    CType(.Cells(zeile, 9), Excel.Range).Value() = prtline.Value.grund
                    'End If

                    If awinSettings.fullProtocol Then
                        protokollRange.Cells(zeile, 1).Value = prtline.Value.actDate
                        protokollRange.Cells(zeile, 2).Value = prtline.Value.Projekt
                        protokollRange.Cells(zeile, 3).Value = prtline.Value.hierarchie
                        protokollRange.Cells(zeile, 4).Value = prtline.Value.planelement
                        protokollRange.Cells(zeile, 5).Value = prtline.Value.klasse
                        protokollRange.Cells(zeile, 6).Value = prtline.Value.abkürzung
                        protokollRange.Cells(zeile, 7).Value = prtline.Value.quelle
                        protokollRange.Cells(zeile, 8).Value = prtline.Value.planeleÜbern
                        protokollRange.Cells(zeile, 8).Interior.Color = prtline.Value.hgColor
                        protokollRange.Cells(zeile, 9).Value = prtline.Value.grund
                        protokollRange.Cells(zeile, 10).Value = prtline.Value.PThierarchie
                        protokollRange.Cells(zeile, 11).Value = prtline.Value.PTklasse
                    Else
                        protokollRange.Cells(zeile, 1).Value = prtline.Value.actDate
                        protokollRange.Cells(zeile, 2).Value = prtline.Value.Projekt
                        protokollRange.Cells(zeile, 4).Value = prtline.Value.planelement
                        protokollRange.Cells(zeile, 8).Value = prtline.Value.planeleÜbern
                        protokollRange.Cells(zeile, 8).Interior.Color = prtline.Value.hgColor
                        protokollRange.Cells(zeile, 9).Value = prtline.Value.grund
                    End If

                End With
            Catch ex As Exception

            End Try

        Next

        ' Logbuch sichern
        If Not IsNothing(xlsLogfile) Then
            xlsLogfile.Save()
        End If

    End Sub


    Public Sub XMLExportReportProfil(ByVal profil As clsReport)

        Dim dirname As String = awinPath & ReportProfileOrdner
        Dim xmlfilename As String = dirname & "\" & profil.name & ".xml"

        Try

            If Not My.Computer.FileSystem.DirectoryExists(dirname) Then
                Try
                    My.Computer.FileSystem.CreateDirectory(dirname)
                Catch ex As Exception

                End Try
            End If

            Dim serializer = New DataContractSerializer(GetType(clsReport))

            ' ''Dim file As New FileStream(xmlfilename, FileMode.Create)
            ' ''serializer.WriteObject(file, profil)
            ' ''file.Close()

            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = (ControlChars.Tab)
            settings.OmitXmlDeclaration = True

            Dim writer As XmlWriter = XmlWriter.Create(xmlfilename, settings)
            serializer.WriteObject(writer, profil)
            writer.Flush()
            writer.Close()

        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try

    End Sub


    Public Sub XMLExportReportProfil(ByVal profil As clsReportAll)

        Dim dirname As String = awinPath & ReportProfileOrdner
        Dim xmlfilename As String = dirname & "\" & profil.name & ".xml"

        Try

            If Not My.Computer.FileSystem.DirectoryExists(dirname) Then
                Try
                    My.Computer.FileSystem.CreateDirectory(dirname)
                Catch ex As Exception

                End Try
            End If

            Dim serializer = New DataContractSerializer(GetType(clsReportAll))

            ' ''Dim file As New FileStream(xmlfilename, FileMode.Create)
            ' ''serializer.WriteObject(file, profil)
            ' ''file.Close()

            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = (ControlChars.Tab)
            settings.OmitXmlDeclaration = True

            Dim writer As XmlWriter = XmlWriter.Create(xmlfilename, settings)
            serializer.WriteObject(writer, profil)
            writer.Flush()
            writer.Close()

        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try

    End Sub


    Public Function XMLImportReportProfil(ByVal profilName As String) As clsReportAll

        Dim ergprofil As New clsReportAll
        Dim aktfile As FileStream = Nothing

        Dim xmlfilename As String = awinPath & ReportProfileOrdner & "\" & profilName & ".xml"
        Try
            ' ur: 31.03.2017 von nun an wird auch bei BHTC mit der Struktur clsReportAll agiert.
            '                alte ReportProfile von BHTC können trotzdem noch gelesen werden, siehe Catch-fall
            Dim serializer = New DataContractSerializer(GetType(clsReportAll))
            Dim profil As New clsReportAll

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            aktfile = file
            profil = serializer.ReadObject(file)
            file.Close()
            ergprofil = profil

            XMLImportReportProfil = ergprofil


        Catch ex As Exception

            ' ur: 18.05.2017: oben geöffnete Datei zunächst schließen, da falsche Format
            If Not IsNothing(aktfile) Then
                aktfile.Close()
            End If


            ' ur: 31.03.2017 neu eingefügt
            Try
                Dim serializer = New DataContractSerializer(GetType(clsReport))
                Dim profil As New clsReport

                ' XML-Datei Öffnen
                ' A FileStream is needed to read the XML document.
                Dim file As New FileStream(xmlfilename, FileMode.Open)
                profil = serializer.ReadObject(file)
                file.Close()
                profil.CopyTo(ergprofil)

                XMLImportReportProfil = ergprofil


            Catch ex2 As Exception

                Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
                XMLImportReportProfil = Nothing
            End Try

        End Try

    End Function

    Public Function XMLImportReportAllProfil(ByVal profilName As String) As clsReportAll

        Dim profil As New clsReportAll

        Dim serializer = New DataContractSerializer(GetType(clsReportAll))
        Dim xmlfilename As String = awinPath & ReportProfileOrdner & "\" & profilName & ".xml"
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            profil = serializer.ReadObject(file)
            file.Close()

            XMLImportReportAllProfil = profil

        Catch ex As Exception

            Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            XMLImportReportAllProfil = Nothing
        End Try

    End Function


    ''' <summary>
    ''' synchronisiert die globalen mit den lokalen Konfigurations-Dateien 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub synchronizeGlobalToLocalFolder()


        Dim srcFile As String
        Dim destFile As String
        Dim destdir As String

        Try


            'Prüfen, ob der Globale Folder existiert
            If Not My.Computer.FileSystem.DirectoryExists(globalPath) Then

                Throw New ArgumentException("Globaler Requirementsordner " & globalPath & " existiert nicht!")

            Else
                ' '' ''Prüfen, ob der Lokale Folder existiert
                '' ''If Not My.Computer.FileSystem.DirectoryExists(awinPath) Then

                '' ''    Call MsgBox("lokaler Requirementsordner " & awinPath & " existiert nicht!")

                ' Lokaler Requirementsordner wird erzeugt, mit allen Unterdirectories
                Try


                    My.Computer.FileSystem.CreateDirectory(awinPath)
                    My.Computer.FileSystem.CreateDirectory(awinPath & requirementsOrdner)

                    For Each gdir In My.Computer.FileSystem.GetDirectories(globalPath & requirementsOrdner)

                        ' Name des lokalen Directories zusammensetzen
                        Dim hstr() As String
                        hstr = gdir.Split(New Char() {CChar("\")})

                        ' Name des destinationDirectories zusammen setzen
                        destdir = awinPath & requirementsOrdner

                        destdir = destdir & hstr(hstr.Length - 1)

                        My.Computer.FileSystem.CreateDirectory(destdir)

                    Next

                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & projektVorlagenOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & modulVorlagenOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & projektRessOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & RepProjectVorOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & RepPortfolioVorOrdner)
                    ' ''My.Computer.FileSystem.CreateDirectory(awinPath & ReportProfileOrdner)


                    '' ''importOrdnerNames(PTImpExp.visbo) = awinPath & "Import\VISBO Steckbriefe"
                    '' ''importOrdnerNames(PTImpExp.rplan) = awinPath & "Import\RPLAN-Excel"
                    '' ''importOrdnerNames(PTImpExp.msproject) = awinPath & "Import\MSProject"
                    '' ''importOrdnerNames(PTImpExp.simpleScen) = awinPath & "Import\einfache Szenarien"
                    '' ''importOrdnerNames(PTImpExp.modulScen) = awinPath & "Import\modulare Szenarien"
                    '' ''importOrdnerNames(PTImpExp.addElements) = awinPath & "Import\addOn Regeln"
                    '' ''importOrdnerNames(PTImpExp.rplanrxf) = awinPath & "Import\RXF Files"

                    '' ''exportOrdnerNames(PTImpExp.visbo) = awinPath & "Export\VISBO Steckbriefe"
                    '' ''exportOrdnerNames(PTImpExp.rplan) = awinPath & "Export\RPLAN-Excel"
                    '' ''exportOrdnerNames(PTImpExp.msproject) = awinPath & "Export\MSProject"
                    '' ''exportOrdnerNames(PTImpExp.simpleScen) = awinPath & "Export\einfache Szenarien"
                    '' ''exportOrdnerNames(PTImpExp.modulScen) = awinPath & "Export\modulare Szenarien"

                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.visbo))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.rplan))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.msproject))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.simpleScen))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.modulScen))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.addElements))
                    '' ''My.Computer.FileSystem.CreateDirectory(importOrdnerNames(PTImpExp.rplanrxf))

                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.visbo))
                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.rplan))
                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.msproject))
                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.simpleScen))
                    '' ''My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(PTImpExp.modulScen))

                Catch ex As Exception

                End Try
                '' ''Else


                Dim dirItem As String = globalPath & requirementsOrdner

                ' lokaler RequirementsOrdner existiert

                ' RequirementsOrdner:   alle Dateien , sofern sie im globalPath neuer als im awinPath sind kopieren

                For Each srcFile In My.Computer.FileSystem.GetFiles(dirItem)

                    ' Name des lokalen Files zusammensetzen
                    Dim hstr() As String
                    hstr = srcFile.Split(New Char() {CChar("\")})

                    ' Name des destinationDirectories zusammen setzen
                    destdir = awinPath & requirementsOrdner

                    destFile = destdir & "\" & hstr(hstr.Length - 1)

                    ' Test ob globales File neuer als lokales
                    Dim srcDate As Date = My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime
                    Dim destDate As Date = My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime
                    Dim ddiff As Long = DateDiff(DateInterval.Second,
                                                 My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime,
                                                 My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime)

                    ' Wenn globales neuer als lokales, dann von globalPath nach awinPath kopieren
                    If ddiff < 0 Then
                        ' Kopieren der Datei, mit Overwrite erzwingen
                        My.Computer.FileSystem.CopyFile(srcFile, destFile, True)
                        ' Debug Mode? 
                        If awinSettings.visboDebug Then
                            Call MsgBox("kopiert von global nach local:" & hstr(hstr.Length - 1))
                        End If
                    End If

                Next srcFile


                ' Unterdirectories von requirementsOrdner:      alle Dateien dieser  werden von globalPath nACH awinPath kopiert, sofern neueres Änderungsdatum

                For Each dirItem In My.Computer.FileSystem.GetDirectories(globalPath & requirementsOrdner)

                    For Each srcFile In My.Computer.FileSystem.GetFiles(dirItem)

                        ' Name des lokalen Files zusammensetzen
                        Dim hstr() As String
                        hstr = srcFile.Split(New Char() {CChar("\")})
                        ' Name des destinationDirectories zusammen setzen
                        Dim dirstr() As String
                        dirstr = dirItem.Split(New Char() {CChar("\")})
                        destdir = awinPath & requirementsOrdner & dirstr(dirstr.Length - 1)

                        destFile = destdir & "\" & hstr(hstr.Length - 1)

                        ' Test ob globales File neuer als lokales
                        Dim srcDate As Date = My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime
                        Dim destDate As Date = My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime
                        Dim ddiff As Long = DateDiff(DateInterval.Second,
                                                     My.Computer.FileSystem.GetFileInfo(srcFile).LastWriteTime,
                                                     My.Computer.FileSystem.GetFileInfo(destFile).LastWriteTime)

                        ' Wenn globales neuer als lokales, dann von globalPath nach awinPath kopieren
                        If ddiff < 0 Then
                            ' Kopieren der Datei, mit Overwrite erzwingen
                            My.Computer.FileSystem.CopyFile(srcFile, destFile, True)

                            ' Debug Mode? 
                            If awinSettings.visboDebug Then
                                Call MsgBox("kopiert von global nach local:" & hstr(hstr.Length - 2) & "/" & hstr(hstr.Length - 1))
                            End If
                        End If

                    Next srcFile

                Next dirItem

            End If


            ' ''End If


        Catch ex As Exception

        End Try
    End Sub

    Public Sub XMLExportLicences(ByVal lic As clsLicences, ByVal nameLicfile As String)


        Dim xmlfilename As String = awinPath & nameLicfile

        Try

            Dim serializer = New DataContractSerializer(GetType(clsLicences))

            ' ''Dim file As New FileStream(xmlfilename, FileMode.Create)
            ' ''serializer.WriteObject(file, lic)
            ' ''file.Close()

            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = (ControlChars.Tab)
            settings.OmitXmlDeclaration = True

            Dim writer As XmlWriter = XmlWriter.Create(xmlfilename, settings)
            serializer.WriteObject(writer, lic)
            writer.Flush()
            writer.Close()
        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try

    End Sub

    Public Function XMLImportLicences(ByVal licfile As String) As clsLicences

        Dim lic As New clsLicences

        Dim serializer = New DataContractSerializer(GetType(clsLicences))
        Dim xmlfilename As String = awinPath & licfile
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            lic = serializer.ReadObject(file)
            file.Close()

            XMLImportLicences = lic

        Catch ex As Exception
            'Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            Throw New ArgumentException("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            XMLImportLicences = Nothing
        End Try

    End Function

    Public Function XMLImportReportMsg(ByVal repMsgfile As String, ByVal language As String) As clsReportMessages

        Dim reportMessages As New clsReportMessages

        Dim serializer = New DataContractSerializer(GetType(clsReportMessages))
        Dim xmlfilename As String = awinPath & requirementsOrdner & repMsgfile & "_" & language & ".xml"
        Try

            ' XML-Datei Öffnen
            ' A FileStream is needed to read the XML document.
            Dim file As New FileStream(xmlfilename, FileMode.Open)
            reportMessages = serializer.ReadObject(file)
            file.Close()

            XMLImportReportMsg = reportMessages

        Catch ex As Exception

            Call MsgBox("Beim Lesen der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")
            XMLImportReportMsg = Nothing
        End Try

    End Function

    Public Sub XMLExportReportMsg(ByVal reportMsg As clsReportMessages, ByVal repMsgfile As String, ByVal language As String)



        Dim xmlfilename As String = awinPath & requirementsOrdner & repMsgfile & "_" & language & ".xml"
        Try
            Dim serializer = New DataContractSerializer(GetType(clsReportMessages))

            ' ''Dim file As New FileStream(xmlfilename, FileMode.Create)
            ' ''serializer.WriteObject(file, lic)
            ' ''file.Close()

            Dim settings As New XmlWriterSettings()
            settings.Indent = True
            settings.IndentChars = (ControlChars.Tab)
            settings.OmitXmlDeclaration = True

            Dim writer As XmlWriter = XmlWriter.Create(xmlfilename, settings)
            serializer.WriteObject(writer, reportMsg)
            writer.Flush()
            writer.Close()

        Catch ex As Exception

            Call MsgBox("Beim Schreiben der XML-Datei '" & xmlfilename & "' ist ein Fehler aufgetreten !")

        End Try
    End Sub
    ''' <summary>
    ''' importiert die ProjectboardConfig.xml
    ''' </summary>
    ''' <param name="cfgXMLfilename"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function XMLImportConfig(ByVal cfgXMLfilename As String) As configuration

        ' XML-Datei Öffnen
        ' A FileStream is needed to read the XML document.
        Dim fs As New FileStream(cfgXMLfilename, FileMode.Open)

        ' Declare an object variable of the type to be deserialized.
        Dim cfgs As New configuration           ' Class configuration erzeugt aus Projectboard.dll.config
        Try


            ' Create an instance of the XmlSerializer class;
            ' specify the type of object to be deserialized.
            Dim deserializer As New XmlSerializer(GetType(configuration))


            ' If the XML document has been altered with unknown
            ' nodes or attributes, handle them with the
            ' UnknownNode and UnknownAttribute events.
            AddHandler deserializer.UnknownNode, AddressOf deserializer_UnknownNode
            AddHandler deserializer.UnknownAttribute, AddressOf deserializer_UnknownAttribute


            ' Einlesen des kompletten XML-Dokument im die Klasse rxf
            ' Use the Deserialize method to restore the object's state with
            ' data from the XML document. 
            cfgs = CType(deserializer.Deserialize(fs), configuration)

            XMLImportConfig = cfgs

        Catch ex As Exception
            XMLImportConfig = Nothing
            Call MsgBox("Lesen der " & cfgXMLfilename & " fehlgeschlagen")
        End Try

        ' ProjectboardConfig.xml-Datei schließen
        fs.Close()

    End Function


    ''' <summary>
    ''' schreibt eine Datei mit den monatlichen Zuordnungen Rollenbedarfe / Kosten 
    ''' Diese Datei kann editiert werden , dann wieder importiert werden 
    ''' in Abhängigkeit vom Typ wird geschrieben: 
    ''' 0: alles
    ''' 1: nur Vergangenheit, von bestimmt den Start , Heute-1 das Ende 
    ''' 2: nur die Zukunft, Heute bestimmt den Start, bis  das Ende  
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub writeProjektBedarfeXLSX(ByVal von As Integer, ByVal bis As Integer, ByVal type As Integer)


        appInstance.EnableEvents = False

        Dim newWB As Excel.Workbook
        Dim rng As Excel.Range
        Dim ersteZeile As Excel.Range
        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try

            newWB = appInstance.Workbooks.Add()

        Catch ex As Exception
            Call MsgBox("Excel Datei konnte nicht erzeugt werden ... Abbruch ")
            appInstance.EnableEvents = True
            Exit Sub
        End Try

        ' jetzt schreiben der ersten Zeile 
        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        With newWB.ActiveSheet

            ersteZeile = CType(.Range(.cells(1, 1), .cells(1, 6 + bis - von)), Excel.Range)

            CType(.Cells(1, 1), Excel.Range).Value = "Projekt-Name"
            CType(.Cells(1, 2), Excel.Range).Value = "Varianten-Name"
            CType(.Cells(1, 3), Excel.Range).Value = "Phasen-Name"
            CType(.Cells(1, 4), Excel.Range).Value = "Ressourcen-Name"
            CType(.Cells(1, 5), Excel.Range).Value = "Kostenart-Name"


            ' jetzt wird die Zeile 1 geschrieben 
            CType(.Cells(1, 6), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(von - 1)
            CType(.Cells(1, 7), Global.Microsoft.Office.Interop.Excel.Range).Value = StartofCalendar.AddMonths(von)
            rng = .Range(.Cells(1, 6), .Cells(1, 7))

            '' Deutsches Format:
            'rng.NumberFormat = "[$-407]mmm yy;@"
            ' Englisches Format:
            rng.NumberFormat = "[$-409]mmm yy;@"

            Dim destinationRange As Excel.Range = .Range(.Cells(1, 6), .Cells(1, 6 + bis - von))
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
                .ReadingOrder = Excel.Constants.xlContext
                .MergeCells = False
            End With

            rng.AutoFill(Destination:=destinationRange, Type:=Excel.XlAutoFillType.xlFillMonths)

        End With



        zeile = 2

        Dim tmpName As String = ""
        Dim tmpValues() As Double
        Dim schnittmenge() As Double
        Dim usedRoles As Collection
        Dim usedCosts As Collection
        Dim pStart As Integer, pEnde As Integer

        Dim editRange As Excel.Range


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            pStart = getColumnOfDate(kvp.Value.startDate)
            pEnde = getColumnOfDate(kvp.Value.endeDate)

            usedRoles = kvp.Value.getRoleNames
            usedCosts = kvp.Value.getCostNames

            For r = 1 To usedRoles.Count
                tmpName = usedRoles.Item(r)
                tmpValues = kvp.Value.getRessourcenBedarf(tmpName)
                schnittmenge = calcArrayIntersection(von, bis, pStart, pEnde, tmpValues)

                If schnittmenge.Sum <> tmpValues.Sum Then
                    Dim a As Integer = 99
                End If

                ' Schreiben der Projekt-Informationen 
                With newWB.ActiveSheet
                    CType(.cells(zeile, 1), Excel.Range).Value = kvp.Value.name
                    CType(.cells(zeile, 2), Excel.Range).Value = kvp.Value.variantName
                    CType(.cells(zeile, 3), Excel.Range).Value = "."
                End With


                With newWB.ActiveSheet
                    CType(.cells(zeile, 4), Excel.Range).Value = tmpName
                    editRange = CType(.range(.cells(zeile, 6), .cells(zeile, 6 + bis - von)), Excel.Range)
                End With

                editRange.Value = schnittmenge
                zeile = zeile + 1

            Next



            For k As Integer = 1 To usedCosts.Count

                tmpName = usedCosts.Item(k)
                tmpValues = kvp.Value.getKostenBedarf(tmpName)
                schnittmenge = calcArrayIntersection(von, bis, pStart, pEnde, tmpValues)

                If schnittmenge.Sum <> tmpValues.Sum Then
                    Dim a As Integer = 99
                End If

                ' Schreiben der Projekt-Informationen 
                With newWB.ActiveSheet
                    CType(.cells(zeile, 1), Excel.Range).Value = kvp.Value.name
                    CType(.cells(zeile, 2), Excel.Range).Value = kvp.Value.variantName
                    CType(.cells(zeile, 3), Excel.Range).Value = "."
                End With

                With newWB.ActiveSheet
                    CType(.cells(zeile, 5), Excel.Range).Value = tmpName
                    editRange = CType(.range(.cells(zeile, 6), .cells(zeile, 6 + bis - von)), Excel.Range)
                End With

                editRange.Value = schnittmenge
                zeile = zeile + 1
            Next

        Next


        ' jetzt den Bereich markieren bzw. schützen 
        Dim startProtectedArea As Integer
        Dim endProtectedArea As Integer
        Dim protectedRange As Excel.Range = Nothing
        Dim wbName As String

        Select Case type
            Case 0
                startProtectedArea = 0
                endProtectedArea = 0
                wbName = "all"
            Case 1

                startProtectedArea = getColumnOfDate(Date.Now)
                endProtectedArea = bis
                wbName = "past"
            Case 2
                startProtectedArea = von
                endProtectedArea = getColumnOfDate(Date.Now)
                wbName = "future"
            Case Else
                Call MsgBox("Typ nicht erkannt, muss Werte 0, 1 oder 2 haben: ist aber" & type)
                appInstance.EnableEvents = True
                Exit Sub
        End Select

        Dim generalRange As Excel.Range = CType(newWB.ActiveSheet.Range(newWB.ActiveSheet.cells(1, 1),
                                                newWB.ActiveSheet.cells(zeile - 1, 5)),
                                                Excel.Range)
        Dim valueRange As Excel.Range = CType(newWB.ActiveSheet.Range(newWB.ActiveSheet.cells(1, 6),
                                                newWB.ActiveSheet.cells(zeile - 1, 6 + bis - von + 1)),
                                                Excel.Range)

        With generalRange
            .Columns.AutoFit()
        End With


        With ersteZeile
            .Interior.Color = awinSettings.AmpelGruen
        End With


        If type <> 0 Then

            With newWB.ActiveSheet
                protectedRange = CType(.Range(.cells(1, startProtectedArea),
                                                               .cells(zeile - 1, endProtectedArea)),
                                                                Excel.Range)

            End With
            protectedRange.Interior.Color = awinSettings.AmpelNichtBewertet
        End If



        Dim expFName As String = exportOrdnerNames(PTImpExp.visbo) & "\EditNeeds_" &
            Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Try
            appInstance.ActiveWorkbook.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
        Catch ex As Exception

        End Try

        Try
            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = True

        Call MsgBox("ok, Datei exportiert")

    End Sub

    ''' <summary>
    ''' schreibt die Projekt-Details in eine Excel Datei
    ''' dabei werden nur die Rollen / Kosten-Bedarfe rausgeschrieben, die in der myCollection angegeben sind
    ''' Bei Rollen werden auch alle Rollen rausgeschrieben, die Kind oder Kindeskind einer angegebenen Rolle sind    ''' 
    ''' </summary>
    ''' <param name="von">Start-Monat (showrangeleft)</param>
    ''' <param name="bis">Ende-Monat (showrangeright)</param>
    ''' <param name="roleNameIDCollection"></param>
    ''' <param name="costNameCollection"></param>
    Public Sub writeProjektDetailsToExcel(ByVal von As Integer, ByVal bis As Integer, ByVal roleNameIDCollection As Collection, ByVal costNameCollection As Collection)

        appInstance.EnableEvents = False

        Dim projectsToWork As New Collection
        Dim defDone As Boolean = False
        If Not IsNothing(selectedProjekte) Then
            If selectedProjekte.Count > 0 Then
                For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste
                    If Not projectsToWork.Contains(kvp.Key) Then
                        projectsToWork.Add(kvp.Key, kvp.Key)
                    End If
                Next
                defDone = True
            End If
        End If


        If Not defDone And ShowProjekte.getMarkedProjects.Count > 0 Then
            projectsToWork = ShowProjekte.getMarkedProjects
            defDone = True
        End If

        If Not defDone Then
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                projectsToWork.Add(kvp.Key, kvp.Key)
            Next
        End If


        Dim newWB As Excel.Workbook

        Dim considerAll As Boolean = True
        ' nur dann considerAll, wenn auch irgendwelche Rollen oder Kosten auch tatsächlich bekannt sind .. 
        considerAll = ((roleNameIDCollection.Count = 0) And (costNameCollection.Count = 0))


        Dim fNameExtension As String = ""
        ' den Dateinamen bestimmen ...
        If roleNameIDCollection.Count > 0 Then
            Dim teamID As Integer
            fNameExtension = RoleDefinitions.getRoleDefByIDKennung(roleNameIDCollection.Item(1), teamID).name
            If roleNameIDCollection.Count > 1 Or costNameCollection.Count > 0 Then
                fNameExtension = fNameExtension & " etc"
            End If

            If fNameExtension = "" And costNameCollection.Count > 0 Then
                fNameExtension = costNameCollection.Item(1)
                If costNameCollection.Count > 1 Then
                    fNameExtension = fNameExtension & " etc"
                End If
            End If
        End If


        Dim ressCostColumn As Integer

        Dim expFName As String = exportOrdnerNames(PTImpExp.massenEdit) & "\Details " & fNameExtension & ".xlsx"


        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try

            newWB = appInstance.Workbooks.Add()
            CType(newWB.Worksheets.Item(1), Excel.Worksheet).Name = "VISBO"
            newWB.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

        Catch ex As Exception
            Call MsgBox("Excel Datei konnte nicht erzeugt werden ... Abbruch ")
            appInstance.EnableEvents = True
            Exit Sub
        End Try

        ' jetzt schreiben der ersten Zeile 
        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        Dim startSpalteDaten As Integer = 9
        Dim roleCostNames As Excel.Range = Nothing
        Dim roleCostInput As Excel.Range = Nothing

        Dim tmpName As String = ""


        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
            Dim ersteZeile As Excel.Range
            ersteZeile = CType(.Range(.Cells(1, 1), .Cells(1, 6 + bis - von)), Excel.Range)

            CType(.Cells(1, 1), Excel.Range).Value = "Projekt-Name"
            CType(.Cells(1, 2), Excel.Range).Value = "Varianten-Name"
            CType(.Cells(1, 3), Excel.Range).Value = "Projekt-Nr"
            CType(.Cells(1, 4), Excel.Range).Value = "Verantwortlich"
            CType(.Cells(1, 5), Excel.Range).Value = "Phasen-Name"
            CType(.Cells(1, 6), Excel.Range).Value = "Ress./Kostenart-Name"
            CType(.Cells(1, 7), Excel.Range).Value = "Tagessatz"
            CType(.Cells(1, 8), Excel.Range).Value = "Summe [PT]"
            'CType(.Cells(1, 7), Excel.Range).Value = "Kostenart-Name"

            ' jetzt wird die Spalten-Nummer festgelegt, wo die Ressourcen/ Kosten später eingetragen werden
            ressCostColumn = 6
            ' jetzt wird die Zeile 1 geschrieben 
            Dim startMonat As Date = StartofCalendar.AddMonths(von - 1)


            ' jetzt werden die Überschriften des Datenbereichs geschrieben 
            For m As Integer = 0 To bis - von
                With CType(.Cells(1, startSpalteDaten + m), Global.Microsoft.Office.Interop.Excel.Range)
                    .Value = startMonat.AddMonths(m)
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                    .NumberFormat = "[$-409]mmm yy;@"
                    .WrapText = False
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ReadingOrder = Excel.Constants.xlContext
                End With
            Next


        End With

        zeile = 2

        Dim schnittmenge() As Double
        Dim zeilenWerte() As Double
        Dim zeilensumme As Double
        Dim pStart As Integer, pEnde As Integer

        Dim editRange As Excel.Range

        For Each pName As String In projectsToWork

            If ShowProjekte.contains(pName) Then
                Dim hproj As clsProjekt = ShowProjekte.getProject(pName)

                Dim atLeastOne As Boolean = False

                pStart = getColumnOfDate(hproj.startDate)
                pEnde = getColumnOfDate(hproj.endeDate)

                For p = 1 To hproj.CountPhases

                    Dim cphase As clsPhase = hproj.getPhase(p)
                    Dim phaseNameID As String = cphase.nameID
                    Dim phaseName As String = cphase.name
                    Dim chckNameID As String = calcHryElemKey(phaseName, False)

                    If phaseWithinTimeFrame(pStart, cphase.relStart, cphase.relEnde, von, bis) Then
                        ' nur wenn die Phase überhaupt im betrachteten Zeitraum liegt, muss das berücksichtigt werden 

                        ' jetzt müssen die Zellen, die zur Phase gehören , entsperrt werden  ...
                        Dim ixZeitraum As Integer
                        Dim ix As Integer, breite As Integer

                        Call awinIntersectZeitraum(pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, ixZeitraum, ix, breite)


                        For r = 1 To cphase.countRoles


                            Dim role As clsRolle = cphase.getRole(r)
                            Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(role.uid, role.teamID)
                            Dim relevant As Boolean = False

                            ' Prüfung: muss die Rolle überhaupt ausgegeben werden ? 
                            If roleNameIDCollection.Count = 0 Then
                                relevant = True
                            Else
                                Dim parentArray() As Integer = RoleDefinitions.getIDArray(roleNameIDCollection)
                                If RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, parentArray) Then
                                    relevant = True
                                End If
                            End If

                            ' nur weitermachen, wenn es relevant ist ..
                            If relevant Then
                                Dim teamID As Integer
                                Dim curRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(roleNameID, teamID)

                                If Not IsNothing(curRole) Then
                                    Dim roleUID As Integer = role.uid
                                    Dim xValues() As Double = role.Xwerte
                                    Dim tagessatz As Double = role.tagessatzIntern

                                    schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                                    zeilensumme = schnittmenge.Sum

                                    'ReDim zeilenWerte(2 * (bis - von + 1) - 1)
                                    ReDim zeilenWerte(bis - von)

                                    ' Schreiben der Projekt-Informationen 
                                    With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                                        CType(.Cells(zeile, 1), Excel.Range).Value = hproj.name
                                        CType(.Cells(zeile, 2), Excel.Range).Value = hproj.variantName
                                        CType(.Cells(zeile, 3), Excel.Range).Value = hproj.kundenNummer
                                        CType(.Cells(zeile, 4), Excel.Range).Value = hproj.leadPerson
                                        CType(.Cells(zeile, 5), Excel.Range).Value = cphase.name
                                        CType(.Cells(zeile, 6), Excel.Range).Value = role.name
                                        CType(.Cells(zeile, 7), Excel.Range).Value = tagessatz
                                        CType(.Cells(zeile, 8), Excel.Range).Value = zeilensumme

                                        editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
                                    End With

                                    ' zusammenmischen von Schnittmenge und Prozentual-Werte 
                                    For mis As Integer = 0 To bis - von
                                        zeilenWerte(mis) = schnittmenge(mis)
                                    Next

                                    editRange.Value = zeilenWerte
                                    atLeastOne = True

                                    zeile = zeile + 1

                                End If


                            End If

                        Next r

                        For c = 1 To cphase.countCosts
                            Dim cost As clsKostenart = cphase.getCost(c)
                            Dim xValues() As Double = cost.Xwerte
                            Dim costName As String = cost.name

                            Dim relevant As Boolean = False

                            ' Prüfung: muss die Rolle überhaupt ausgegeben werden ? 
                            If costNameCollection.Count = 0 Then
                                relevant = True
                            Else
                                ' If CostDefinitions.hasAnyChildParentRelationsship(costName, costCollection) Then
                                If costNameCollection.Contains(costName) Then
                                    relevant = True
                                End If

                            End If

                            ' nur weitermachen, wenn es relevant ist ..
                            If relevant Then


                                schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                                zeilensumme = schnittmenge.Sum

                                'ReDim zeilenWerte(2 * (bis - von + 1) - 1)
                                ReDim zeilenWerte(bis - von)

                                ' Schreiben der Projekt-Informationen 
                                With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

                                    CType(.Cells(zeile, 1), Excel.Range).Value = hproj.name
                                    CType(.Cells(zeile, 2), Excel.Range).Value = hproj.variantName
                                    CType(.Cells(zeile, 3), Excel.Range).Value = hproj.kundenNummer
                                    CType(.Cells(zeile, 4), Excel.Range).Value = hproj.leadPerson
                                    CType(.Cells(zeile, 5), Excel.Range).Value = cphase.name
                                    CType(.Cells(zeile, 6), Excel.Range).Value = costName
                                    CType(.Cells(zeile, 7), Excel.Range).Value = ""
                                    CType(.Cells(zeile, 8), Excel.Range).Value = zeilensumme

                                    editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
                                End With

                                ' zusammenmischen von Schnittmenge und Prozentual-Werte 
                                For mis As Integer = 0 To bis - von
                                    zeilenWerte(mis) = schnittmenge(mis)
                                    ' in auslastungsarray(r, 0) steht die Gesamt-Auslastung, spielt aber kein Kostenarten keine Rolle 
                                    'zeilenWerte(2 * mis) = schnittmenge(mis)
                                    'zeilenWerte(2 * mis + 1) = 0
                                Next

                                'editRange.Value = schnittmenge
                                editRange.Value = zeilenWerte
                                atLeastOne = True

                                zeile = zeile + 1

                            End If

                        Next c

                    End If

                Next p

                If Not atLeastOne Then
                    ' Schreiben der Projekt-Informationen 
                    With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

                        CType(.Cells(zeile, 1), Excel.Range).Value = hproj.name
                        CType(.Cells(zeile, 2), Excel.Range).Value = hproj.variantName
                        CType(.Cells(zeile, 3), Excel.Range).Value = hproj.kundenNummer
                        CType(.Cells(zeile, 4), Excel.Range).Value = hproj.leadPerson
                        CType(.Cells(zeile, 5), Excel.Range).Value = "-"
                        CType(.Cells(zeile, 6), Excel.Range).Value = "-"
                        CType(.Cells(zeile, 7), Excel.Range).Value = "-"
                        CType(.Cells(zeile, 8), Excel.Range).Value = "-"

                        zeile = zeile + 1
                    End With
                End If

            End If


        Next

        ' jetzt müssen die Spaltenbreiten und sonstigen Werte gesetzt werden ...
        Dim letzteSpalte As Integer = 9 + bis - von
        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
            For s As Integer = 1 To letzteSpalte

                If s = 1 Then
                    ' Projekt-Name
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 54
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                ElseIf s = 2 Then
                    ' Varianten-Name
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                ElseIf s = 3 Then
                    ' Projekt-Nummer
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    'CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                ElseIf s = 4 Then
                    ' Verantwortlich
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                ElseIf s = 5 Then
                    ' Phasen-Name
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                ElseIf s = 6 Then
                    ' Rolle / Kostenart 
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 36
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                ElseIf s = 7 Then
                    ' Tagessatz
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 12
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "#,##0.##"

                ElseIf s = 8 Then
                    ' Summe
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "#,##0.##"

                Else
                    ' die monatlichen Werte 
                    CType(.Columns.Item(s), Excel.Range).ColumnWidth = 9
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 1
                    CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "#,##0.##"
                End If
                ' 


            Next

            ' jetzt muss noch die erste Zeile formatiert werden 
            CType(.Rows.Item(1), Excel.Range).RowHeight = 45
            'CType(.Rows.Item(1), Excel.Range).VerticalAlignment = XlTopBottom.xlTop10Top
            CType(.Rows.Item(1), Excel.Range).VerticalAlignment = XlVAlign.xlVAlignTop
            CType(.Rows.Item(1), Excel.Range).Interior.Color = RGB(220, 220, 220)

        End With

        Try
            ' jetzt die Autofilter aktivieren ... 
            If Not CType(newWB.Worksheets("VISBO"), Excel.Worksheet).AutoFilterMode = True Then
                CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1).AutoFilter()
            End If

            newWB.Close(SaveChanges:=True)
        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Speichern" & ex.Message)
        End Try

        appInstance.EnableEvents = True

        Call MsgBox("ok, Datei exportiert")


    End Sub

    ''' <summary>
    ''' schreibt eine Datei mit den monatlichen Zuordnungen Projekt/Phase - Rollenbedarfe / Kosten 
    ''' Diese Datei kann editiert werden , dann wieder importiert werden 
    ''' in Abhängigkeit vom Typ wird geschrieben: 
    ''' 0: alles
    ''' 1: nur Vergangenheit, von bestimmt den Start , Heute-1 das Ende 
    ''' 2: nur die Zukunft, Heute bestimmt den Start, bis  das Ende  
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="type"></param>
    ''' <remarks></remarks>
    Public Sub writeProjektPhasenBedarfeXLSX(ByVal von As Integer, ByVal bis As Integer, ByVal type As Integer)


        appInstance.EnableEvents = False

        Dim newWB As Excel.Workbook

        Dim ressCostColumn As Integer

        Dim expFName As String = exportOrdnerNames(PTImpExp.massenEdit) & "\EditNeeds_" &
        Date.Now.ToString.Replace(":", ".") & ".xlsx"

        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try

            newWB = appInstance.Workbooks.Add()
            CType(newWB.Worksheets.Item(1), Excel.Worksheet).Name = "VISBO"
            newWB.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

        Catch ex As Exception
            Call MsgBox("Excel Datei konnte nicht erzeugt werden ... Abbruch ")
            appInstance.EnableEvents = True
            Exit Sub
        End Try

        ' jetzt schreiben der ersten Zeile 
        Dim zeile As Integer = 1
        Dim spalte As Integer = 1

        Dim startSpalteDaten As Integer = 9
        Dim roleCostNames As Excel.Range = Nothing
        Dim roleCostInput As Excel.Range = Nothing

        Dim tmpName As String = ""


        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
            Dim ersteZeile As Excel.Range
            ersteZeile = CType(.Range(.Cells(1, 1), .Cells(1, 6 + bis - von)), Excel.Range)

            CType(.Cells(1, 1), Excel.Range).Value = "Projekt-Name"
            CType(.Cells(1, 2), Excel.Range).Value = "Varianten-Name"
            CType(.Cells(1, 3), Excel.Range).Value = "Projekt-Nr"
            CType(.Cells(1, 4), Excel.Range).Value = "Verantwortlich"
            CType(.Cells(1, 5), Excel.Range).Value = "Phasen-Name"
            CType(.Cells(1, 6), Excel.Range).Value = "Ress./Kostenart-Name"
            CType(.Cells(1, 7), Excel.Range).Value = "Tagessatz"
            CType(.Cells(1, 8), Excel.Range).Value = "Summe [PT]"
            'CType(.Cells(1, 7), Excel.Range).Value = "Kostenart-Name"

            ' jetzt wird die Spalten-Nummer festgelegt, wo die Ressourcen/ Kosten später eingetragen werden
            ressCostColumn = 5
            ' jetzt wird die Zeile 1 geschrieben 
            Dim startMonat As Date = StartofCalendar.AddMonths(von - 1)

            ' jetzt wird der Name hinzugefügt
            Dim tmpRange1 As Excel.Range = CType(.Cells(1, startSpalteDaten), Global.Microsoft.Office.Interop.Excel.Range)
            Dim tmpRange2 As Excel.Range = CType(.Cells(1, startSpalteDaten + 2 * (bis - von)), Global.Microsoft.Office.Interop.Excel.Range)
            newWB.Names.Add(Name:="StartData", RefersToR1C1:=tmpRange1)
            newWB.Names.Add(Name:="EndData", RefersToR1C1:=tmpRange2)

            ' jetzt werden die Überschriften des Datenbereichs geschrieben 
            For m As Integer = 0 To bis - von
                With CType(.Cells(1, startSpalteDaten + m), Global.Microsoft.Office.Interop.Excel.Range)
                    .Value = startMonat.AddMonths(m)
                    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    .VerticalAlignment = Excel.XlVAlign.xlVAlignBottom
                    .NumberFormat = "[$-409]mmm yy;@"
                    .WrapText = False
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ReadingOrder = Excel.Constants.xlContext
                End With

                ' tk, 9.9.18 keine Ausgabe von Prz Werten mehr 
                'With CType(.Cells(1, startSpalteDaten + 2 * m + 1), Global.Microsoft.Office.Interop.Excel.Range)
                '    .Value = ""
                '    .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                '    .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                '    .Orientation = 0
                '    .AddIndent = False
                '    .IndentLevel = 0
                '    .ReadingOrder = Excel.Constants.xlContext
                'End With

            Next


        End With

        zeile = 2

        Dim schnittmenge() As Double
        Dim zeilenWerte() As Double
        Dim zeilensumme As Double
        Dim pStart As Integer, pEnde As Integer

        Dim editRange As Excel.Range


        ' zu Beginn werden die rollen-spezifischen Auslastungskennzahlen ermittelt, die sich über alle aktuell 
        ' betrachteten Projekte ergeben; 
        ' es werden sowohl die Gesamt-Auslastungs Werte im Zeitraum betrachtet als auch der einzelne monats-spezifische Wert   
        ' dazu wird ein Array angelegt mit der Dimension (anzahlRollen-1, bis-von+1) 
        ' tk 9.9.18 braucht man nicht mehr ..
        'Dim auslastungsArray(,) As Double

        'Try
        '    auslastungsArray = visboZustaende.getUpDatedAuslastungsArray(Nothing, von, bis, awinSettings.mePrzAuslastung)
        '    'auslastungsArray = ShowProjekte.getAuslastungsArray(von, bis)
        'Catch ex As Exception
        '    ReDim auslastungsArray(RoleDefinitions.Count - 1, bis - von + 1)
        'End Try




        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            pStart = getColumnOfDate(kvp.Value.startDate)
            pEnde = getColumnOfDate(kvp.Value.endeDate)

            For p = 1 To kvp.Value.CountPhases

                Dim cphase As clsPhase = kvp.Value.getPhase(p)
                Dim phaseNameID As String = cphase.nameID
                Dim phaseName As String = cphase.name
                Dim chckNameID As String = calcHryElemKey(phaseName, False)

                If phaseWithinTimeFrame(pStart, cphase.relStart, cphase.relEnde, von, bis) Then
                    ' nur wenn die Phase überhaupt im betrachteten Zeitraum liegt, muss das berücksichtigt werden 

                    ' jetzt müssen die Zellen, die zur Phase gehören , entsperrt werden  ...
                    Dim ixZeitraum As Integer
                    Dim ix As Integer, breite As Integer

                    Dim atLeastOne As Boolean = False

                    Call awinIntersectZeitraum(pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, ixZeitraum, ix, breite)


                    For r = 1 To cphase.countRoles


                        Dim role As clsRolle = cphase.getRole(r)
                        Dim roleName As String = role.name
                        Dim roleUID As Integer = RoleDefinitions.getRoledef(roleName).UID
                        Dim xValues() As Double = role.Xwerte
                        Dim tagessatz As Double = RoleDefinitions.getRoledef(roleName).tagessatzIntern

                        schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                        zeilensumme = schnittmenge.Sum

                        'ReDim zeilenWerte(2 * (bis - von + 1) - 1)
                        ReDim zeilenWerte(bis - von)

                        ' Schreiben der Projekt-Informationen 
                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                            CType(.Cells(zeile, 1), Excel.Range).Value = kvp.Value.businessUnit
                            CType(.Cells(zeile, 2), Excel.Range).Value = kvp.Value.name
                            CType(.Cells(zeile, 3), Excel.Range).Value = kvp.Value.variantName
                            CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                            Dim cellComment As Excel.Comment = CType(.Cells(zeile, 4), Excel.Range).Comment
                            If Not IsNothing(cellComment) Then
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                            End If
                            If chckNameID = phaseNameID Then
                                ' nichts weiter tun ... 
                                ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                            Else
                                CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                            End If

                            CType(.Cells(zeile, 5), Excel.Range).Value = roleName
                            CType(.Cells(zeile, 6), Excel.Range).Value = tagessatz
                            CType(.Cells(zeile, 7), Excel.Range).Value = zeilensumme.ToString("0")
                            'CType(.Cells(zeile, 7), Excel.Range).Value = auslastungsArray(roleUID - 1, 0).ToString("0%")
                            ' editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von + 1) - 1)), Excel.Range)
                            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
                        End With

                        ' zusammenmischen von Schnittmenge und Prozentual-Werte 
                        For mis As Integer = 0 To bis - von
                            zeilenWerte(mis) = schnittmenge(mis)
                            ' in auslastungsarray(r, 0) steht die Gesamt-Auslastung
                            'zeilenWerte(2 * mis) = schnittmenge(mis)
                            'zeilenWerte(2 * mis + 1) = auslastungsArray(roleUID - 1, mis + 1)
                        Next

                        'editRange.Value = schnittmenge
                        editRange.Value = zeilenWerte
                        atLeastOne = True
                        ' die Zellen entsperren, die editiert werden dürfen ...

                        ' tk 9.9.18 nicht mehr nötig 
                        ''With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

                        ''    For l = 0 To bis - von

                        ''        If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                        ''            'CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Locked = False
                        ''            CType(.Range(.Cells(zeile, l + startSpalteDaten),
                        ''                         .Cells(zeile, l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                        ''            ' CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten),
                        ''            '.Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet

                        ''        Else
                        ''            CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                        ''        End If

                        ''    Next

                        ''    ' vorheriger Code
                        ''    ''For l As Integer = ixZeitraum To ixZeitraum + breite - 1
                        ''    ''    CType(.cell(zeile, l + 6), Excel.Range).Locked = False
                        ''    ''    CType(.cell(zeile, l + 6), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                        ''    ''Next
                        ''End With


                        ''With newWB.ActiveSheet
                        ''    For l As Integer = ixZeitraum To ixZeitraum + breite - 1
                        ''        CType(.cells(zeile, l + 6), Excel.Range).Locked = False
                        ''    Next
                        ''End With

                        zeile = zeile + 1

                    Next r

                    For c = 1 To cphase.countCosts
                        Dim cost As clsKostenart = cphase.getCost(c)
                        Dim xValues() As Double = cost.Xwerte
                        Dim costName As String = cost.name
                        schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                        zeilensumme = schnittmenge.Sum

                        'ReDim zeilenWerte(2 * (bis - von + 1) - 1)
                        ReDim zeilenWerte(bis - von)

                        ' Schreiben der Projekt-Informationen 
                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                            CType(.Cells(zeile, 1), Excel.Range).Value = kvp.Value.businessUnit
                            CType(.Cells(zeile, 2), Excel.Range).Value = kvp.Value.name
                            CType(.Cells(zeile, 3), Excel.Range).Value = kvp.Value.variantName
                            CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                            Dim cellComment As Excel.Comment = CType(.Cells(zeile, 4), Excel.Range).Comment
                            If Not IsNothing(cellComment) Then
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                            End If
                            If chckNameID = phaseNameID Then
                                ' nichts weiter tun ... 
                                ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                            Else
                                CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                            End If

                            CType(.Cells(zeile, 5), Excel.Range).Value = costName
                            CType(.Cells(zeile, 7), Excel.Range).Value = zeilensumme.ToString("0")
                            'editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von + 1) - 1)), Excel.Range)
                            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
                        End With

                        ' zusammenmischen von Schnittmenge und Prozentual-Werte 
                        For mis As Integer = 0 To bis - von
                            zeilenWerte(mis) = schnittmenge(mis)
                            ' in auslastungsarray(r, 0) steht die Gesamt-Auslastung, spielt aber kein Kostenarten keine Rolle 
                            'zeilenWerte(2 * mis) = schnittmenge(mis)
                            'zeilenWerte(2 * mis + 1) = 0
                        Next

                        'editRange.Value = schnittmenge
                        editRange.Value = zeilenWerte
                        atLeastOne = True
                        ' die Zellen entsperren, die editiert werden dürfen ...

                        ' die Zellen entsperren, die editiert werden dürfen ...

                        'With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

                        '    For l = 0 To bis - von

                        '        If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                        '            'CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Locked = False
                        '            CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten),
                        '                         .Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                        '            CType(.Cells(zeile, 2 * l + 1 + startSpalteDaten), Excel.Range).Value = ""
                        '        Else
                        '            CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                        '            CType(.Cells(zeile, 2 * l + 1 + startSpalteDaten), Excel.Range).Value = ""
                        '        End If

                        '    Next

                        'End With

                        zeile = zeile + 1

                    Next c

                    If Not atLeastOne Then
                        ' jetzt sollte eine leere Projekt-Phasen-Information geschrieben werden, quasi ein Platzhalter
                        ' in diesem Platzhalter kann dann später die Ressourcen Information aufgenommen werden  
                        ' Schreiben der Projekt-Informationen 
                        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                            CType(.Cells(zeile, 1), Excel.Range).Value = kvp.Value.businessUnit
                            CType(.Cells(zeile, 2), Excel.Range).Value = kvp.Value.name
                            CType(.Cells(zeile, 3), Excel.Range).Value = kvp.Value.variantName
                            CType(.Cells(zeile, 4), Excel.Range).Value = cphase.name

                            Dim cellComment As Excel.Comment = CType(.Cells(zeile, 4), Excel.Range).Comment
                            If Not IsNothing(cellComment) Then
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Delete()
                            End If
                            If chckNameID = phaseNameID Then
                                ' nichts weiter tun ... 
                                ' denn dann kann die PhaseNameID aus der PhaseName konstruiert werden
                                ' wenn es eine laufende Nummer 2, 3 etc ist, dann muss explizit die PhaseNameID in den Kommentarbereich geschreiben werden 
                            Else
                                CType(.Cells(zeile, 4), Excel.Range).AddComment(Text:=cphase.nameID)
                                CType(.Cells(zeile, 4), Excel.Range).Comment.Visible = False
                            End If

                            CType(.Cells(zeile, 5), Excel.Range).Value = ""
                            CType(.Cells(zeile, 6), Excel.Range).Value = ""
                            CType(.Cells(zeile, 7), Excel.Range).Value = ""
                            'editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + 2 * (bis - von))), Excel.Range)
                            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
                        End With

                        ' die Zellen entsperren, die editiert werden dürfen ...
                        'With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

                        '    For l = 0 To bis - von

                        '        If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                        '            'CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Locked = False
                        '            CType(.Range(.Cells(zeile, 2 * l + startSpalteDaten),
                        '                         .Cells(zeile, 2 * l + 1 + startSpalteDaten)), Excel.Range).Interior.Color = awinSettings.AmpelNichtBewertet
                        '        Else
                        '            CType(.Cells(zeile, 2 * l + startSpalteDaten), Excel.Range).Value = ""
                        '        End If

                        '    Next

                        'End With

                        zeile = zeile + 1

                    End If

                End If



            Next p



        Next


        ' jetzt die Größe der Spalten anpassen 
        Dim infoBlock As Excel.Range
        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
            infoBlock = CType(.Range(.Columns(1), .Columns(startSpalteDaten - 1)), Excel.Range)
            infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            infoBlock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            infoBlock.AutoFit()
        End With

        Dim tmpRange As Excel.Range
        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)

            'Dim isPrz As Boolean = False
            'For mis As Integer = 0 To 2 * (bis - von + 1) - 1
            For mis As Integer = 0 To bis - von
                tmpRange = CType(.Range(.Cells(2, startSpalteDaten + mis), .Cells(zeile, startSpalteDaten + mis)), Excel.Range)
                ''If isPrz Then
                ''    tmpRange.Columns.ColumnWidth = 3.1
                ''    tmpRange.Font.Size = 6
                ''    tmpRange.NumberFormat = "0%"
                ''    tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                ''Else
                ''    tmpRange.Columns.ColumnWidth = 5
                ''    tmpRange.Font.Size = 10
                ''    tmpRange.NumberFormat = "0"
                ''    tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                ''End If
                ''isPrz = Not isPrz
                tmpRange.Columns.ColumnWidth = 8
                'tmpRange.Font.Size = 10
                tmpRange.NumberFormat = "#0.##"
                tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            Next

        End With



        ' jetzt wird der RoleCostInput Bereich festgelegt 
        'With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
        '    Dim maxRows As Integer = .Rows.Count
        '    roleCostInput = CType(.Range(.Cells(2, ressCostColumn), .Cells(maxRows, ressCostColumn)), Excel.Range)
        'End With

        'With roleCostInput
        '    .Validation.Delete()
        '    .Validation.Add(Type:=XlDVType.xlValidateList, AlertStyle:=XlDVAlertStyle.xlValidAlertStop,
        '                                   Formula1:="=RollenKostenNamen")
        'End With



        Try
            '' jetzt die Autofilter aktivieren ... 
            'If Not CType(newWB.Worksheets("VISBO"), Excel.Worksheet).AutoFilterMode = True Then
            '    'CType(CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
            '    CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1).AutoFilter()
            'End If

            ' ExcelFile abspeichern und schließen
            newWB.Close(SaveChanges:=True)
        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Speichern" & ex.Message)
        End Try

        appInstance.EnableEvents = True

        Call MsgBox("ok, Datei exportiert")

    End Sub

    ''' <summary>
    ''' schreibt eine Datei, die zur Priorisierung / Analyse verwendet werden kann 
    ''' Diese Datei kann editiert werden , dann wieder importiert werden 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub writeProjektsForSequencing(ByVal roleCostCollection As Collection)

        Dim err As New clsErrorCodeMsg

        appInstance.EnableEvents = False

        Dim newWB As Excel.Workbook
        Dim considerAll As Boolean = (roleCostCollection.Count = 0)

        Dim roleCollection As New Collection
        Dim costCollection As New Collection
        Dim vorgabeValue As Double = 0.0
        Dim aktuellValue As Double = 0.0
        Dim vorgabeProj As clsProjekt = Nothing
        Dim roleNames As String = ""
        Dim costNames As String = ""

        If Not considerAll Then
            For Each itemName As String In roleCostCollection
                If RoleDefinitions.containsName(itemName) Then
                    If Not roleCollection.Contains(itemName) Then
                        roleCollection.Add(itemName, itemName)
                        If roleNames = "" Then
                            roleNames = vbLf & itemName
                        Else
                            roleNames = roleNames & "; " & itemName
                        End If
                    End If
                ElseIf CostDefinitions.containsName(itemName) Then
                    If Not costCollection.Contains(itemName) Then
                        costCollection.Add(itemName, itemName)
                        If costNames = "" Then
                            costNames = vbLf & itemName
                        Else
                            costNames = costNames & "; " & itemName
                        End If
                    End If
                End If
            Next
            ' nur dann considerAll, wenn auch irgendwelche Rollen oder Kosten auch tatsächlich bekannt sind .. 
            considerAll = ((roleCollection.Count = 0) And (costCollection.Count = 0))
        End If

        Dim fNameExtension As String = ""
        ' den Dateinamen bestimmen ...
        If roleCollection.Count > 0 Then
            fNameExtension = roleCollection.Item(1)
            If roleCollection.Count > 1 Or costCollection.Count > 0 Then
                fNameExtension = fNameExtension & " etc"
            End If

            If fNameExtension = "" And costCollection.Count > 0 Then
                fNameExtension = costCollection.Item(1)
                If costCollection.Count > 1 Then
                    fNameExtension = fNameExtension & " etc"
                End If
            End If
        End If

        Dim expFName As String = ""
        If considerAll Then
            expFName = exportOrdnerNames(PTImpExp.scenariodefs) & "\" & currentConstellationName & "_Prio.xlsx"
        Else
            expFName = exportOrdnerNames(PTImpExp.massenEdit) & "\Overview " & fNameExtension & ".xlsx"
        End If
        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try

            newWB = appInstance.Workbooks.Add()
            CType(newWB.Worksheets.Item(1), Excel.Worksheet).Name = "VISBO"
            newWB.SaveAs(Filename:=expFName, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

        Catch ex As Exception
            Call MsgBox("Excel Datei konnte nicht erzeugt werden ... Abbruch ")
            appInstance.EnableEvents = True
            Exit Sub
        End Try

        ' jetzt schreiben der ersten Zeile 
        Dim zeile As Integer = 1
        Dim spalte As Integer = 1


        Dim startOfCustomFields As Integer = 15
        Dim ersteZeile As Excel.Range


        With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)


            If awinSettings.englishLanguage Then
                CType(.Cells(1, 1), Excel.Range).Value = "Project-Name"
                CType(.Cells(1, 2), Excel.Range).Value = "Variant-Name"
                CType(.Cells(1, 3), Excel.Range).Value = "Project-Nr"
                CType(.Cells(1, 4), Excel.Range).Value = "Responsible"
                CType(.Cells(1, 5), Excel.Range).Value = "Business-Unit"
                CType(.Cells(1, 6), Excel.Range).Value = "Project-Start"
                CType(.Cells(1, 7), Excel.Range).Value = "Project-End"

                If considerAll Then
                    CType(.Cells(1, 8), Excel.Range).Value = "Budget [T€]"
                    CType(.Cells(1, 11), Excel.Range).Value = "Profit/Loss [T€]"
                Else
                    CType(.Cells(1, 8), Excel.Range).Value = "First Version [T€]"
                    CType(.Cells(1, 11), Excel.Range).Value = "Difference [T€]"
                End If

                CType(.Cells(1, 9), Excel.Range).Value = "Sum Personnel-Cost [T€]" & roleNames
                CType(.Cells(1, 10), Excel.Range).Value = "Sum Other Cost [T€]" & costNames

                CType(.Cells(1, 12), Excel.Range).Value = "Strategy"
                CType(.Cells(1, 13), Excel.Range).Value = "Risk"
                CType(.Cells(1, 14), Excel.Range).Value = "Description"
            Else

                CType(.Cells(1, 1), Excel.Range).Value = "Projekt-Name"
                CType(.Cells(1, 2), Excel.Range).Value = "Varianten-Name"
                CType(.Cells(1, 3), Excel.Range).Value = "Projekt-Nr"
                CType(.Cells(1, 4), Excel.Range).Value = "Verantwortlich"
                CType(.Cells(1, 5), Excel.Range).Value = "Business-Unit"
                CType(.Cells(1, 6), Excel.Range).Value = "Projekt-Start"
                CType(.Cells(1, 7), Excel.Range).Value = "Projekt-Ende"

                If considerAll Then
                    CType(.Cells(1, 8), Excel.Range).Value = "Budget [T€]"
                    CType(.Cells(1, 11), Excel.Range).Value = "Gewinn/Verlust [T€]"
                Else
                    CType(.Cells(1, 8), Excel.Range).Value = "Erste Planung [T€]"
                    CType(.Cells(1, 11), Excel.Range).Value = "Differenz [T€]"
                End If

                CType(.Cells(1, 9), Excel.Range).Value = "Summe Personalkosten [T€]" & roleNames
                CType(.Cells(1, 10), Excel.Range).Value = "Summe sonst. Kosten [T€]" & costNames

                CType(.Cells(1, 12), Excel.Range).Value = "Strategie"
                CType(.Cells(1, 13), Excel.Range).Value = "Risiko"
                CType(.Cells(1, 14), Excel.Range).Value = "Beschreibung"


            End If



            spalte = startOfCustomFields
            For Each cstField As KeyValuePair(Of Integer, clsCustomFieldDefinition) In customFieldDefinitions.liste
                .Cells(zeile, spalte).value = cstField.Value.name
                spalte = spalte + 1
            Next

            ersteZeile = CType(.Range(.Cells(1, 1), .Cells(1, startOfCustomFields + customFieldDefinitions.liste.Count - 1)), Excel.Range)

        End With


        '' jetzt den AutoFit machen 
        'Try
        '    ersteZeile.AutoFit()
        'Catch ex As Exception

        'End Try

        zeile = 2
        Dim hproj As clsProjekt = Nothing
        Try
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                hproj = kvp.Value
                Dim budget As Double, pk As Double, ok As Double, rk As Double, pl As Double
                Dim alterPlanStand As Date = Date.MinValue
                Dim standVom As String = ""

                If considerAll Then
                    Call kvp.Value.calculateRoundedKPI(budget, pk, ok, rk, pl)
                Else
                    ' jetzt müssen budget, pk, ok, rk, pl anhand der Rollen-/Kosten-Vorgaben bestimmt werden 
                    Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString
                    vorgabeProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(kvp.Value.name, vorgabeVariantName, err)

                    ' Berechnung budget/Vorgabe 
                    budget = 0.0
                    pk = 0.0
                    ok = 0.0
                    alterPlanStand = Date.MinValue

                    If Not IsNothing(vorgabeProj) Then
                        For Each itemName As String In roleCollection
                            budget = budget + vorgabeProj.getRessourcenBedarf(itemName, inclSubRoles:=True, outPutInEuro:=True).Sum
                            pk = pk + kvp.Value.getRessourcenBedarf(itemName, inclSubRoles:=True, outPutInEuro:=True).Sum
                        Next

                        For Each itemName As String In costCollection
                            budget = budget + vorgabeProj.getKostenBedarfNew(itemName).Sum
                            ok = ok + kvp.Value.getKostenBedarfNew(itemName).Sum
                        Next

                        ' welcher Planungs-Stand ist das ? 
                        alterPlanStand = vorgabeProj.timeStamp
                        standVom = alterPlanStand.ToShortDateString
                    Else
                        ' es gibt kein Vorgabe Proj
                        budget = 0
                        standVom = "n.a"

                        For Each itemName As String In roleCollection
                            pk = pk + kvp.Value.getRessourcenBedarf(itemName, inclSubRoles:=True, outPutInEuro:=True).Sum
                        Next

                        For Each itemName As String In costCollection
                            ok = ok + kvp.Value.getKostenBedarfNew(itemName).Sum
                        Next

                    End If

                    ' Berechnung Personalkosten
                    pl = budget - (pk + ok)

                End If


                With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                    CType(.Cells(zeile, 1), Excel.Range).Value = kvp.Value.name
                    CType(.Cells(zeile, 2), Excel.Range).Value = kvp.Value.variantName
                    CType(.Cells(zeile, 3), Excel.Range).Value = kvp.Value.kundenNummer
                    CType(.Cells(zeile, 4), Excel.Range).Value = kvp.Value.leadPerson
                    CType(.Cells(zeile, 5), Excel.Range).Value = kvp.Value.businessUnit
                    CType(.Cells(zeile, 6), Excel.Range).Value = kvp.Value.startDate
                    CType(.Cells(zeile, 7), Excel.Range).Value = kvp.Value.endeDate

                    CType(.Cells(zeile, 8), Excel.Range).Value = budget
                    CType(.Cells(zeile, 8), Excel.Range).NumberFormat = "0.00"
                    If Not considerAll Then
                        ' damit wird klar, von wann diese Version ist
                        CType(.Cells(zeile, 8), Excel.Range).AddComment(standVom)
                    End If


                    CType(.Cells(zeile, 9), Excel.Range).Value = pk
                    CType(.Cells(zeile, 9), Excel.Range).NumberFormat = "0.00"

                    CType(.Cells(zeile, 10), Excel.Range).Value = ok
                    CType(.Cells(zeile, 10), Excel.Range).NumberFormat = "0.00"

                    CType(.Cells(zeile, 11), Excel.Range).Value = pl
                    CType(.Cells(zeile, 11), Excel.Range).NumberFormat = "0.00"

                    CType(.Cells(zeile, 12), Excel.Range).Value = kvp.Value.StrategicFit
                    CType(.Cells(zeile, 13), Excel.Range).Value = kvp.Value.Risiko
                    CType(.Cells(zeile, 14), Excel.Range).Value = kvp.Value.fullDescription

                    spalte = startOfCustomFields
                    For Each cstField As KeyValuePair(Of Integer, clsCustomFieldDefinition) In customFieldDefinitions.liste

                        Dim qualifier As String = cstField.Value.name
                        Dim ausgabe As String = ""
                        If cstField.Value.type = ptCustomFields.Str Then
                            ausgabe = kvp.Value.getCustomSField(qualifier)
                        ElseIf cstField.Value.type = ptCustomFields.Dbl Then
                            ausgabe = kvp.Value.getCustomDField(qualifier).ToString
                        ElseIf cstField.Value.type = ptCustomFields.bool Then
                            ausgabe = kvp.Value.getCustomBField(qualifier).ToString
                        End If

                        If IsNothing(ausgabe) Then
                            ausgabe = ""
                        End If

                        CType(.Cells(zeile, spalte), Excel.Range).Value = ausgabe
                        spalte = spalte + 1
                    Next

                End With
                zeile = zeile + 1
            Next
        Catch ex As Exception
            Call MsgBox("Problems with " & hproj.name)
        End Try


        ' jetzt müssen die Spaltenbreiten und sonstigen Werte gesetzt werden ...
        Dim letzteSpalte As Integer = startOfCustomFields + customFieldDefinitions.liste.Count - 1
        Dim curSpalte As Integer = -1

        Try
            With CType(newWB.Worksheets("VISBO"), Excel.Worksheet)
                For s As Integer = 1 To letzteSpalte
                    curSpalte = s
                    If s = 1 Then
                        ' Projekt-Name
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 54
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                    ElseIf s = 2 Then
                        ' Varianten-Name
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                    ElseIf s = 3 Then
                        ' Projekt-Nummer
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        'CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                    ElseIf s = 4 Then
                        ' Verantwortlich
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                    ElseIf s = 5 Then
                        ' Business Unit
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                    ElseIf s = 6 Then
                        ' Projekt-Start
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                    ElseIf s = 7 Then
                        ' Projekt-Ende
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                    ElseIf s = 8 Then
                        ' Budget bzw. erste Planung 
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "0.00"
                    ElseIf s = 9 Then
                        ' summe Personalkosten
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 28
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "0.00"

                    ElseIf s = 10 Then
                        ' summe Sonst Kosten
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 28
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "0.00"
                    ElseIf s = 11 Then
                        ' Profit/Loss bzw. Differenz
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).IndentLevel = 2
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "0.00"
                    ElseIf s = 12 Then
                        ' Strategie 
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 12
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "0"
                    ElseIf s = 13 Then
                        ' Risiko 
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 12
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).NumberFormat = "0"
                    ElseIf s = 14 Then
                        ' Beschreibung
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 36
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                    Else
                        ' customFields
                        CType(.Columns.Item(s), Excel.Range).ColumnWidth = 18
                        CType(.Range(.Cells(2, s), .Cells(zeile - 1, s)), Excel.Range).WrapText = False
                    End If
                    ' 


                Next

                ' jetzt muss noch die erste Zeile formatiert werden 
                CType(.Rows.Item(1), Excel.Range).RowHeight = 45
                CType(.Rows.Item(1), Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
                CType(.Rows.Item(1), Excel.Range).Interior.Color = RGB(220, 220, 220)

            End With
        Catch ex As Exception
            Call MsgBox("Problem with Column: " & curSpalte)
        End Try





        Try
            ' jetzt die Autofilter aktivieren ... 
            If Not CType(newWB.Worksheets("VISBO"), Excel.Worksheet).AutoFilterMode = True Then
                'CType(CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
                CType(newWB.Worksheets("VISBO"), Excel.Worksheet).Cells(1, 1).AutoFilter()
            End If

            ' ExcelFile abspeichern und schließen
            newWB.Close(SaveChanges:=True)
        Catch ex As Exception
            Throw New ArgumentException("Fehler beim Filtersetzen und Speichern" & ex.Message)
        End Try

        appInstance.EnableEvents = True

        Call MsgBox("ok, Datei exportiert")

    End Sub

    '' tk 26.12.18 deprecated
    '''' <summary>
    '''' erstellt für alles, alleRollen, alleKosten und die einzelnen Sammel-Rollen die ValidationStrings
    '''' die dann im Mass-Edit verwendet werden können 
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function createMassEditRcValidations() As SortedList(Of String, String)
    '    Dim validationStrings As New SortedList(Of String, String)
    '    Dim validationName As String

    '    ' Aufbau Alles
    '    validationName = "alles"
    '    Dim sortedRCListe As New SortedList(Of String, String)
    '    Dim rcDefinition As String = ""
    '    Dim tmpName As String

    '    For iz As Integer = 1 To RoleDefinitions.Count
    '        tmpName = RoleDefinitions.getRoledef(iz).name
    '        If Not sortedRCListe.ContainsKey(tmpName) Then
    '            sortedRCListe.Add(tmpName, tmpName)
    '        End If
    '    Next

    '    For iz As Integer = 1 To sortedRCListe.Count
    '        If rcDefinition.Length = 0 Then
    '            rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
    '        Else
    '            rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
    '        End If
    '    Next

    '    sortedRCListe.Clear()

    '    For iz As Integer = 1 To CostDefinitions.Count - 1
    '        tmpName = CostDefinitions.getCostdef(iz).name
    '        If Not sortedRCListe.ContainsKey(tmpName) Then
    '            sortedRCListe.Add(tmpName, tmpName)
    '        End If
    '    Next

    '    For iz As Integer = 1 To sortedRCListe.Count
    '        If rcDefinition.Length = 0 Then
    '            rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
    '        Else
    '            rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
    '        End If
    '    Next

    '    If Not validationStrings.ContainsKey(validationName) Then
    '        validationStrings.Add(validationName, rcDefinition)
    '    End If


    '    '
    '    ' jetzt kommen alleRollen 

    '    validationName = "alleRollen"
    '    sortedRCListe = New SortedList(Of String, String)
    '    rcDefinition = ""

    '    For iz As Integer = 1 To RoleDefinitions.Count
    '        tmpName = RoleDefinitions.getRoledef(iz).name
    '        If Not sortedRCListe.ContainsKey(tmpName) Then
    '            sortedRCListe.Add(tmpName, tmpName)
    '        End If
    '    Next

    '    For iz As Integer = 1 To sortedRCListe.Count
    '        If rcDefinition.Length = 0 Then
    '            rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
    '        Else
    '            rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
    '        End If
    '    Next

    '    If Not validationStrings.ContainsKey(validationName) Then
    '        validationStrings.Add(validationName, rcDefinition)
    '    End If

    '    ' Ende alleRollen
    '    '

    '    '
    '    ' jetzt kommen alleKosten 

    '    validationName = "alleKosten"
    '    sortedRCListe = New SortedList(Of String, String)
    '    rcDefinition = ""

    '    For iz As Integer = 1 To CostDefinitions.Count - 1
    '        tmpName = CostDefinitions.getCostdef(iz).name
    '        If Not sortedRCListe.ContainsKey(tmpName) Then
    '            sortedRCListe.Add(tmpName, tmpName)
    '        End If
    '    Next

    '    For iz As Integer = 1 To sortedRCListe.Count
    '        If rcDefinition.Length = 0 Then
    '            rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
    '        Else
    '            rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
    '        End If
    '    Next

    '    If Not validationStrings.ContainsKey(validationName) Then
    '        validationStrings.Add(validationName, rcDefinition)
    '    End If

    '    ' Ende alleKosten
    '    '

    '    '
    '    ' jetzt kommen die einzelnen Sammelrollen, unter Angabe ihres Namens 

    '    Dim sammelrollenNamen As Collection = RoleDefinitions.getSummaryRoles

    '    For iz As Integer = 1 To sammelrollenNamen.Count
    '        tmpName = CStr(sammelrollenNamen.Item(iz))
    '        If Not sortedRCListe.ContainsKey(tmpName) Then
    '            sortedRCListe.Add(tmpName, tmpName)
    '        End If
    '    Next

    '    For Each validationName In sammelrollenNamen

    '        sortedRCListe = New SortedList(Of String, String)
    '        rcDefinition = ""

    '        For iz As Integer = 1 To sammelrollenNamen.Count
    '            tmpName = CStr(sammelrollenNamen.Item(iz))
    '            If Not sortedRCListe.ContainsKey(tmpName) Then
    '                sortedRCListe.Add(tmpName, tmpName)
    '            End If
    '        Next

    '        Dim subRoleIDs As SortedList(Of Integer, Double) = RoleDefinitions.getSubRoleIDsOf(validationName, PTcbr.all)

    '        For Each srKvP As KeyValuePair(Of Integer, Double) In subRoleIDs
    '            Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(srKvP.Key)
    '            If Not IsNothing(tmpRole) Then
    '                tmpName = tmpRole.name
    '                If Not sortedRCListe.ContainsKey(tmpName) Then
    '                    sortedRCListe.Add(tmpName, tmpName)
    '                End If
    '            End If

    '        Next


    '        For iz As Integer = 1 To sortedRCListe.Count
    '            If rcDefinition.Length = 0 Then
    '                rcDefinition = sortedRCListe.ElementAt(iz - 1).Value
    '            Else
    '                rcDefinition = rcDefinition & ";" & sortedRCListe.ElementAt(iz - 1).Value
    '            End If
    '        Next

    '        ' jetzt den Validation String hinzufügen 
    '        If Not validationStrings.ContainsKey(validationName) Then
    '            validationStrings.Add(validationName, rcDefinition)
    '        End If

    '    Next

    '    createMassEditRcValidations = validationStrings
    'End Function

    Sub massEditZeileLoeschen(ByVal ID As String)

        Dim currentCell As Excel.Range
        Dim meWS As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
        appInstance.EnableEvents = False

        Dim ok As Boolean = True

        Try

            currentCell = CType(appInstance.ActiveCell, Excel.Range)
            Dim zeile As Integer = currentCell.Row

            If zeile >= 2 And zeile <= visboZustaende.meMaxZeile Then

                Dim columnEndData As Integer = visboZustaende.meColED
                Dim columnStartData As Integer = visboZustaende.meColSD
                Dim columnRC As Integer = visboZustaende.meColRC


                Dim pName As String = CStr(meWS.Cells(zeile, 2).value)
                Dim vName As String = CStr(meWS.Cells(zeile, 3).value)

                Dim phaseName As String = CStr(meWS.Cells(zeile, 4).value)
                Dim phaseNameID As String = getPhaseNameIDfromExcelCell(CType(meWS.Cells(zeile, 4), Excel.Range))


                Dim rcName As String = CStr(meWS.Cells(zeile, columnRC).value)
                Dim rcNameID As String = getRCNameIDfromExcelCell(CType(meWS.Cells(zeile, columnRC), Excel.Range))

                Dim isRole As Boolean = RoleDefinitions.containsName(rcName)

                ' Überprüfen, ob es actualData gibt ... 
                Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                Dim actualDataExists As Boolean = hproj.getPhaseRCActualValues(phaseNameID, rcNameID, isRole, False).Sum > 0

                ' jetzt wird gelöscht, wenn es noch keine Ist-Daten gibt ..
                If Not actualDataExists Then
                    Call meRCZeileLoeschen(currentCell.Row, pName, phaseNameID, rcNameID, isRole)
                Else
                    Call MsgBox("zur Phase gibt es bereits Ist-Daten - deshalb kann die Rolle " & rcName & vbLf &
                                    " nicht gelöscht werden ...")
                End If


            Else
                Call MsgBox(" es können nur Zeilen aus dem Datenbereich gelöscht werden ...")
            End If

        Catch ex As Exception
            Call MsgBox("Fehler beim Löschen einer Zeile ..." & vbLf & ex.Message)
        End Try

        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' löscht eine Zeile im Massen-Edit; dabei ist bereits überprüft, ob sie gelöscht werden darf ... 
    ''' dass heisst, actualDataExists = false
    ''' </summary>
    ''' <param name="zeile"></param>
    ''' <param name="pName"></param>
    ''' <param name="phNameID"></param>
    ''' <param name="rcNameID"></param>
    ''' <param name="isRole"></param>
    Sub meRCZeileLoeschen(ByVal zeile As Integer,
                          ByVal pName As String,
                          ByVal phNameID As String,
                          ByVal rcNameID As String,
                          ByVal isRole As Boolean)

        Dim meWS As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
        appInstance.EnableEvents = False

        Dim ok As Boolean = True
        Dim nothingHappened As Boolean = True

        Try
            Dim teamID As Integer = -1
            Dim currentRole As clsRollenDefinition = Nothing
            If isRole Then
                currentRole = RoleDefinitions.getRoleDefByIDKennung(rcNameID, teamID)
            End If

            If zeile >= 2 And zeile <= visboZustaende.meMaxZeile Then
                Dim columnEndData As Integer = visboZustaende.meColED
                Dim columnStartData As Integer = visboZustaende.meColSD
                Dim columnRC As Integer = visboZustaende.meColRC

                ' hier wird die Rolle- bzw. Kostenart aus der Projekt-Phase gelöscht 
                Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                Dim cphase As clsPhase = hproj.getPhaseByID(phNameID)




                If IsNothing(rcNameID) Then
                    ' nichts tun
                ElseIf rcNameID.Trim.Length = 0 Then
                    ' nichts tun ... 
                ElseIf Not IsNothing(currentRole) Then
                    ' es handelt sich um eine Rolle
                    ' das darf aber nur gelöscht werden, wenn die Phase komplett im showrangeleft / showrangeright liegt 
                    ' gibt es Ist-Daten ? 

                    If phaseWithinTimeFrame(hproj.Start, cphase.relStart, cphase.relEnde,
                                             showRangeLeft, showRangeRight, True) Then
                        cphase.removeRoleByNameID(rcNameID)
                        nothingHappened = False
                    Else
                        Dim rcName As String = RoleDefinitions.getRoleDefByIDKennung(rcNameID, teamID).name

                        Call MsgBox("die Phase wird nicht vollständig angezeigt - deshalb kann die Rolle " & rcName & vbLf &
                                    " nicht gelöscht werden ...")

                        ok = False
                    End If

                ElseIf CostDefinitions.containsName(rcNameID) Then
                    ' es handelt sich um eine Kostenart 
                    If phaseWithinTimeFrame(hproj.Start, cphase.relStart, cphase.relEnde,
                                             showRangeLeft, showRangeRight, True) Then
                        cphase.removeCostByName(rcNameID)
                        nothingHappened = False
                    Else

                        Call MsgBox("die Phase wird nicht vollständig angezeigt - deshalb kann die Kostenart " & rcNameID & vbLf &
                                    " nicht gelöscht werden ...")

                        ok = False
                    End If


                End If


                If ok Then
                    ' jetzt wird die Zeile gelöscht, wenn sie nicht die letzte ihrer Art ist
                    ' denn es sollte für weitere Eingaben immer wenigstens ein Projekt-/Phasen-Repräsentant da sein 
                    If noDuplicatesInSheet(pName, phNameID, Nothing, zeile) Then
                        ' diese Zeile nicht löschen, soll weiter als Platzhalter für diese Projekt-Phase dienen können 
                        ' aber die Werte müssen alle gelöscht werden 
                        For ix As Integer = columnRC To columnEndData + 1
                            CType(meWS.Cells(zeile, ix), Excel.Range).Value = ""
                        Next
                    Else
                        CType(meWS.Rows(zeile), Excel.Range).Delete()
                        zeile = zeile - 1
                        If zeile < 2 Then
                            zeile = 2
                        End If
                    End If

                    ' jetzt wird auf die Ressourcen-/Kosten-Spalte positioniert 
                    CType(meWS.Cells(zeile, columnRC), Excel.Range).Select()

                    ' jetzt wird der Old-Value gesetzt 
                    With visboZustaende
                        .oldRow = zeile
                        .oldValue = CStr(CType(meWS.Cells(zeile, columnRC), Excel.Range).Value)
                        .meMaxZeile = CType(meWS.UsedRange, Excel.Range).Rows.Count
                    End With

                Else
                    ' nichts tun 
                End If


            Else
                Call MsgBox(" es können nur Zeilen aus dem Datenbereich gelöscht werden ...")
            End If

            If Not nothingHappened Then
                Try

                    If Not IsNothing(formProjectInfo1) Then
                        Call updateProjectInfo1(visboZustaende.lastProject, visboZustaende.lastProjectDB)
                    End If
                    Call aktualisiereCharts(visboZustaende.lastProject, True)
                    Call awinNeuZeichnenDiagramme(typus:=6, roleCost:=currentRole.name)

                Catch ex As Exception

                End Try

            End If

        Catch ex As Exception
            Call MsgBox("Fehler beim Löschen einer Zeile ..." & vbLf & ex.Message)
        End Try

        appInstance.EnableEvents = True
    End Sub

    ''' <summary>
    ''' 
    ''' fügt nach der Zeile eine Zeile ein ...     ''' 
    ''' die neue Zeile bekommt die gleichen Inhalte wie die kopierte Zelle bis auf rcName, rcNameID, Summe und alle Werte  
    ''' Vorbedingung: enableEvents ist false ..
    ''' vorab gecheckt: hat die Phase überhaupt Planungs-Monate oder liegt sie vollständig in der Vergangenheit ? 
    ''' </summary>
    ''' <param name="zeile"></param>
    Public Sub meRCZeileEinfuegen(ByVal zeile As Integer, ByVal rcNameID As String, ByVal isRole As Boolean)

        Dim ws As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim currentRow As Excel.Range
        Dim currentRowPlus1 As Excel.Range
        Dim insertRow As Boolean = True
        Dim newZeile As Integer
        appInstance.EnableEvents = False

        Try

            currentRow = CType(ws.Rows(zeile), Excel.Range)
            Dim columnRC As Integer = visboZustaende.meColRC

            Dim currentValue As String = getStringFromExcelCell(ws.Cells(zeile, columnRC))
            insertRow = (currentValue <> "" Or rcNameID = "")


            Dim columnEndData As Integer = visboZustaende.meColED
            Dim columnStartData As Integer = visboZustaende.meColSD



            Dim hoehe As Double = CDbl(currentRow.Height)

            If insertRow Then
                currentRowPlus1 = CType(ws.Cells(currentRow.Row + 1, currentRow.Column), Excel.Range)
                currentRowPlus1.EntireRow.Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown)
                newZeile = zeile + 1
            Else
                newZeile = zeile
            End If


            ' Blattschutz aufheben ... 
            If Not awinSettings.meEnableSorting Then
                ' es muss der Blattschutz aufgehoben werden, nachher wieder aktiviert werden ...
                With CType(appInstance.ActiveSheet, Excel.Worksheet)
                    .Unprotect(Password:="x")
                End With
            End If


            With CType(appInstance.ActiveSheet, Excel.Worksheet)


                If insertRow Then
                    Dim copySource As Excel.Range = CType(.Range(.Cells(zeile, 1), .Cells(zeile, 1).offset(0, columnStartData - 3)), Excel.Range)
                    Dim copyDestination As Excel.Range = CType(.Range(.Cells(zeile + 1, 1), .Cells(zeile + 1, 1).offset(0, columnStartData - 3)), Excel.Range)

                    copySource.Copy(Destination:=copyDestination)

                    CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Rows(zeile + 1), Excel.Range).RowHeight = hoehe
                End If

                ' hier wird jetzt der Rollen- bzw Kostenart-NAme eingetragen 
                Dim rcName As String = rcNameID
                Dim islocked As Boolean = False

                If isRole And rcNameID <> "" Then
                    ' der rcname muss erst noch bestimmt werden 
                    Dim teamID As Integer = -1
                    Dim roleID As Integer = RoleDefinitions.parseRoleNameID(rcNameID, teamID)
                    If roleID > 0 Then
                        rcName = RoleDefinitions.getRoleDefByID(roleID).name
                    End If
                End If

                Call writeMECellWithRoleNameID(CType(.Cells(newZeile, columnRC), Excel.Range), islocked, rcName, rcNameID, isRole)

                For c As Integer = columnStartData - 1 To columnEndData
                    With CType(.Cells(newZeile, c), Excel.Range)
                        .Value = Nothing
                        If c = columnStartData - 2 Or c = columnStartData - 1 Then
                            .ClearComments()
                        End If
                    End With

                Next

            End With

            ' jetzt wird auf die Ressourcen-/Kosten-Spalte positioniert 
            CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Cells(newZeile, columnRC), Excel.Range).Select()

            With CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Cells(newZeile, columnRC), Excel.Range)

                ' wenn eine neue Zeile eingefügt ist  müssen die jetzt wieder auf frei gesetzt werden 
                .Locked = False

                ' jetzt für die Zelle die Validation neu bestimmen, der Blattschutz muss aufgehoben sein ...  
                Try
                    If Not IsNothing(.Validation) Then
                        .Validation.Delete()
                    End If

                Catch ex As Exception

                End Try

            End With

            ' jetzt wird der Old-Value gesetzt 
            With visboZustaende
                'If CStr(CType(appInstance.ActiveCell, Excel.Range).Value) <> "" Then
                '    Call MsgBox("Fehler 099 in PTzeileEinfügen")
                'End If
                .oldRow = newZeile
                .oldValue = rcNameID
                .meMaxZeile = CType(CType(appInstance.ActiveSheet, Excel.Worksheet).UsedRange, Excel.Range).Rows.Count
            End With

            ' tk 14.12.18 wird an aufrufender Stelle gemacht 
            '' jetzt den Blattschutz wiederherstellen ... 
            'If Not awinSettings.meEnableSorting Then
            '    ' es muss der Blattschutz wieder aktiviert werden ... 
            '    With CType(appInstance.ActiveSheet, Excel.Worksheet)
            '        .Protect(Password:="x", UserInterfaceOnly:=True,
            '                 AllowFormattingCells:=True,
            '                 AllowFormattingColumns:=True,
            '                 AllowInsertingColumns:=False,
            '                 AllowInsertingRows:=True,
            '                 AllowDeletingColumns:=False,
            '                 AllowDeletingRows:=True,
            '                 AllowSorting:=True,
            '                 AllowFiltering:=True)
            '        .EnableSelection = Excel.XlEnableSelection.xlUnlockedCells
            '        .EnableAutoFilter = True
            '    End With
            'End If

        Catch ex As Exception
            Call MsgBox("Fehler beim Kopieren einer Zeile ...")
        End Try

        appInstance.EnableEvents = True

    End Sub


    ''' <summary>
    ''' fügt eine Zeile im MassEdit ein 
    ''' </summary>
    ''' <param name="controlID"></param>
    Sub massEditZeileEinfügen(ByVal controlID As String)


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        Dim ws As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)
        Dim currentCell As Excel.Range
        'Dim currentCellPlus1 As Excel.Range

        Try


            ' hier nicht benötigt
            '' '' jetzt werden die Validation-Strings für alles, alleRollen, alleKosten und die einzelnen SammelRollen aufgebaut 
            ' ''Dim validationStrings As SortedList(Of String, String) = createMassEditRcValidations()

            currentCell = CType(appInstance.ActiveCell, Excel.Range)
            Dim zeile As Integer = currentCell.Row
            Dim spalte As Integer = currentCell.Column

            Call meRCZeileEinfuegen(zeile, "", True)


            ' jetzt den Blattschutz wiederherstellen ... 
            If Not awinSettings.meEnableSorting Then
                ' es muss der Blattschutz wieder aktiviert werden ... 
                With CType(appInstance.ActiveSheet, Excel.Worksheet)
                    .Protect(Password:="x", UserInterfaceOnly:=True,
                             AllowFormattingCells:=True,
                             AllowFormattingColumns:=True,
                             AllowInsertingColumns:=False,
                             AllowInsertingRows:=True,
                             AllowDeletingColumns:=False,
                             AllowDeletingRows:=True,
                             AllowSorting:=True,
                             AllowFiltering:=True)
                    .EnableSelection = Excel.XlEnableSelection.xlUnlockedCells
                    .EnableAutoFilter = True
                End With
            End If

        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try

        'appInstance.EnableEvents = True
        appInstance.EnableEvents = formerEE
    End Sub



    ''' <summary>    ''' 
    ''' schreibt die Daten der in einer todoListe übergebenen Projekt-Namen in ein extra Tabellenblatt 
    ''' die Info-Daten werden in einer Range mit Name informationColumns zusammengefasst   
    ''' ur: 05.06.2018: nicht mehr:Dabei wird überprüft, was der längste mögliche Ressourcen und Kosten-Namen überhaupt ist 
    ''' und was der längste eingetragene Namen ist ... Am Schluss wird notfalls die Spaltenbreite verlängert, damit auch der längste Namen reingeht ... 
    ''' </summary>
    ''' <param name="todoListe">enthält die pvNames der Projekte</param>"
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <remarks>seit 15.9 ohne die Mahle Spalten für Auslastung und Freie Tage. 
    ''' das ist in dem Branch MahleSaveMassEdit festgelaten </remarks>
    Public Sub writeOnlineMassEditRessCost(ByVal todoListe As Collection,
                                           ByVal von As Integer, ByVal bis As Integer)
        Dim err As New clsErrorCodeMsg

        Dim maxRCLengthAbsolut As Integer = 0
        Dim maxRCLengthVorkommen As Integer = 0

        If todoListe.Count = 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("no projects for mass-edit available ..")
            Else
                Call MsgBox("keine Projekte für den Massen-Edit vorhanden ..")
            End If

            Exit Sub
        End If

        Try

            appInstance.EnableEvents = False

            ' jetzt die selectedProjekte Liste zurücksetzen ... ohne die currentConstellation zu verändern ...
            selectedProjekte.Clear(False)

            Dim currentWS As Excel.Worksheet
            Dim currentWB As Excel.Workbook
            Dim ressCostColumn As Integer
            Dim tmpName As String


            Try
                currentWB = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                currentWS = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)

                Try
                    ' off setzen des AutoFilter Modus ... 
                    If CType(currentWS, Excel.Worksheet).AutoFilterMode = True Then
                        'CType(CType(currentWS, Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
                        CType(currentWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
                    End If
                Catch ex As Exception

                End Try

                ' braucht man eigentlich nicht mehr, aber sicher ist sicher ...
                Try
                    With CType(currentWS.Cells, Excel.Range)
                        .Clear()
                        .Value = Nothing
                        '.Interior.Color = RGB(242, 242, 242)
                    End With
                Catch ex As Exception

                End Try


            Catch ex As Exception
                Call MsgBox("es gibt Probleme mit dem Mass-Edit Worksheet ...")
                appInstance.EnableEvents = True
                Exit Sub
            End Try


            ' jetzt schreiben der ersten Zeile 
            Dim zeile As Integer = 1
            Dim spalte As Integer = 1


            Dim startSpalteDaten As Integer = 7
            Dim roleCostInput As Excel.Range = Nothing

            tmpName = ""

            With CType(currentWS, Excel.Worksheet)

                If .ProtectContents Then
                    .Unprotect(Password:="x")
                End If

                If awinSettings.englishLanguage Then
                    CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
                    CType(.Cells(1, 2), Excel.Range).Value = "Project-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Variant-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Phase-Name"
                    CType(.Cells(1, 5), Excel.Range).Value = "Res./Cost-Name"
                    maxRCLengthVorkommen = 14
                    CType(.Cells(1, 6), Excel.Range).Value = "Sum" & vbLf & "[FTE]"

                Else
                    CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
                    CType(.Cells(1, 2), Excel.Range).Value = "Projekt-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Varianten-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Phasen-Name"
                    CType(.Cells(1, 5), Excel.Range).Value = "Ress./Kostenart-Name"
                    maxRCLengthVorkommen = 20
                    CType(.Cells(1, 6), Excel.Range).Value = "Summe" & vbLf & "[PT]"

                End If

                ' das Erscheinungsbild der Zeile 1 bestimmen  
                Call massEditZeile1Appearance(ptTables.meRC)


                ' jetzt wird die Spalten-Nummer festgelegt, wo die Ressourcen/ Kosten später eingetragen werden
                ressCostColumn = 5
                ' jetzt wird die Zeile 1 geschrieben 


                ' jetzt wird der Name hinzugefügt
                Dim tmpRange1 As Excel.Range = CType(.Cells(1, startSpalteDaten), Global.Microsoft.Office.Interop.Excel.Range)
                Dim tmpRange2 As Excel.Range = CType(.Cells(1, startSpalteDaten + (bis - von)), Global.Microsoft.Office.Interop.Excel.Range)
                Dim tmpRange3 As Excel.Range = CType(.Cells(1, 5), Global.Microsoft.Office.Interop.Excel.Range)


                Try
                    If Not IsNothing(CType(currentWB.Names.Item("StartData"), Excel.Name)) Then
                        currentWB.Names.Item("StartData").Delete()
                    End If
                Catch ex As Exception

                End Try

                Try
                    If Not IsNothing(CType(currentWB.Names.Item("EndData"), Excel.Name)) Then
                        currentWB.Names.Item("EndData").Delete()
                    End If
                Catch ex As Exception

                End Try

                Try
                    If Not IsNothing(CType(currentWB.Names.Item("RoleCost"), Excel.Name)) Then
                        currentWB.Names.Item("RoleCost").Delete()
                    End If
                Catch ex As Exception

                End Try

                currentWB.Names.Add(Name:="StartData", RefersToR1C1:=tmpRange1)
                currentWB.Names.Add(Name:="EndData", RefersToR1C1:=tmpRange2)
                currentWB.Names.Add(Name:="RoleCost", RefersToR1C1:=tmpRange3)

            End With


            zeile = 2

            Dim schnittmenge() As Double

            Dim zeilensumme As Double
            Dim pStart As Integer, pEnde As Integer

            Dim editRange As Excel.Range



            For Each pvName As String In todoListe

                Dim hproj As clsProjekt = Nothing
                If AlleProjekte.Containskey(pvName) Then
                    hproj = AlleProjekte.getProject(pvName)
                End If

                If Not IsNothing(hproj) Then

                    Dim projectWithActualData As Boolean = False
                    Dim actualDataRelColumn As Integer = -1
                    Dim summeEditierenErlaubt As Boolean = awinSettings.allowSumEditing

                    ' jetzt wird geprüft, ob es bereits Ist-Daten geben könnte 
                    If DateDiff(DateInterval.Month, StartofCalendar, hproj.actualDataUntil) > 0 Then
                        projectWithActualData = (getColumnOfDate(hproj.actualDataUntil) - getColumnOfDate(hproj.startDate) >= 0)
                        actualDataRelColumn = getColumnOfDate(hproj.actualDataUntil) - von
                    End If

                    ' ist das Projekt geschützt ? 
                    ' wenn nein, dann temporär schützen 
                    Dim protectionText As String = ""
                    Dim wpItem As clsWriteProtectionItem
                    Dim isProtectedbyOthers As Boolean

                    If awinSettings.visboServer Then
                        isProtectedbyOthers = Not (CType(databaseAcc, DBAccLayer.Request).checkChgPermission(hproj.name, hproj.variantName, dbUsername, err, ptPRPFType.project))
                    Else
                        isProtectedbyOthers = Not tryToprotectProjectforMe(hproj.name, hproj.variantName)
                    End If


                    If isProtectedbyOthers Then

                        ' nicht erfolgreich, weil durch anderen geschützt ... 
                        ' oder aber noch gar nicht in Datenbank: aber das ist noch nicht berücksichtigt  
                        wpItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                        writeProtections.upsert(wpItem)

                        protectionText = writeProtections.getProtectionText(calcProjektKey(hproj.name, hproj.variantName))

                    End If

                    If actualDataRelColumn >= 0 And Not isProtectedbyOthers Then
                        If awinSettings.englishLanguage Then
                            protectionText = "Actual Data until " & hproj.actualDataUntil.Month & "/" & hproj.actualDataUntil.Year
                        Else
                            protectionText = "Ist-Daten bis " & hproj.actualDataUntil.Month & "/" & hproj.actualDataUntil.Year
                        End If
                    End If


                    pStart = getColumnOfDate(hproj.startDate)
                    pEnde = getColumnOfDate(hproj.endeDate)
                    'Dim defaultEmptyValidation As String = validationStrings(rcValidation(anzahlRollen + 1)) ' alle Rollen und Kostenarten 

                    For p = 1 To hproj.CountPhases

                        Dim cphase As clsPhase = hproj.getPhase(p)
                        Dim phaseNameID As String = cphase.nameID
                        Dim phaseName As String = cphase.name
                        Dim chckNameID As String = calcHryElemKey(phaseName, False)


                        Dim indentlevel As Integer = hproj.hierarchy.getIndentLevel(phaseNameID)

                        If phaseWithinTimeFrame(pStart, cphase.relStart, cphase.relEnde, von, bis) Then
                            ' nur wenn die Phase überhaupt im betrachteten Zeitraum liegt, muss das berücksichtigt werden 

                            ' jetzt müssen die Zellen, die zur Phase gehören , geschrieben werden ...
                            Dim ixZeitraum As Integer
                            Dim ix As Integer, breite As Integer

                            Dim atLeastOne As Boolean = False

                            Call awinIntersectZeitraum(pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, ixZeitraum, ix, breite)


                            For r = 1 To cphase.countRoles

                                Dim role As clsRolle = cphase.getRole(r)

                                Dim roleName As String = role.name
                                Dim roleUID As Integer = role.uid
                                Dim teamID As Integer = role.teamID

                                Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(roleUID, teamID)
                                Dim validRole As Boolean = True

                                If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Then
                                    If myCustomUserRole.specifics.Length > 0 Then
                                        If RoleDefinitions.containsNameID(myCustomUserRole.specifics) Then
                                            Dim trTeamID As Integer = -1
                                            Dim restrictedTopRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, trTeamID)

                                            If RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, restrictedTopRole.UID) Then
                                                validRole = True
                                            Else
                                                validRole = False
                                            End If
                                        End If
                                    End If
                                End If


                                If validRole Then
                                    Dim xValues() As Double = role.Xwerte

                                    ' hier muss bestimmt werden, ob das Projekt in dieser Phase mit dieser Rolle schon actualdata hat ...
                                    Dim hasActualData As Boolean = hproj.getPhaseRCActualValues(phaseNameID, roleNameID, True, False).Sum > 0

                                    summeEditierenErlaubt = (awinSettings.allowSumEditing And Not hasActualData)

                                    schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                                    zeilensumme = schnittmenge.Sum

                                    'ReDim zeilenWerte(bis - von)

                                    Dim ok As Boolean = massEditWrite1Zeile(currentWS.Name, hproj, cphase, indentlevel, isProtectedbyOthers, zeile, roleName, roleNameID, True,
                                                                            protectionText, von, bis,
                                                                            actualDataRelColumn, hasActualData, summeEditierenErlaubt,
                                                                            ixZeitraum, breite, startSpalteDaten, maxRCLengthVorkommen)

                                    If ok Then

                                        With currentWS
                                            CType(.Cells(zeile, 6), Excel.Range).Value = zeilensumme
                                            editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
                                        End With

                                        If schnittmenge.Sum > 0 Then
                                            For l As Integer = 0 To bis - von

                                                If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                                                    editRange.Cells(1, l + 1).value = schnittmenge(l)
                                                Else
                                                    editRange.Cells(1, l + 1).value = ""
                                                End If

                                            Next
                                        Else
                                            editRange.Value = ""
                                        End If

                                        atLeastOne = True

                                        zeile = zeile + 1
                                    Else
                                        Call MsgBox("not ok")
                                    End If
                                End If

                            Next r

                            ' jetzt kommt die Behandlung der Kostenarten

                            For c = 1 To cphase.countCosts

                                Dim hasActualData As Boolean = False
                                Dim cost As clsKostenart = cphase.getCost(c)
                                Dim costName As String = cost.name
                                Dim xValues() As Double = cost.Xwerte

                                ' neu 12.12.18 
                                ' hier muss bestimmt werden, ob das Projekt in dieser Phase mit dieser Kostenart schon actualdata hat ...
                                hasActualData = hproj.getPhaseRCActualValues(phaseNameID, costName, False, True).Sum > 0

                                ' ist Summe Editieren erlaubt ? 
                                If projectWithActualData Then
                                    summeEditierenErlaubt = (awinSettings.allowSumEditing And Not hasActualData)
                                End If


                                schnittmenge = calcArrayIntersection(von, bis, pStart + cphase.relStart - 1, pStart + cphase.relEnde - 1, xValues)
                                zeilensumme = schnittmenge.Sum

                                'ReDim zeilenWerte(bis - von)

                                Dim ok As Boolean = massEditWrite1Zeile(currentWS.Name, hproj, cphase, indentlevel, isProtectedbyOthers, zeile, costName, "", False,
                                                                            protectionText, von, bis,
                                                                            actualDataRelColumn, hasActualData, summeEditierenErlaubt,
                                                                            ixZeitraum, breite, startSpalteDaten, maxRCLengthVorkommen)

                                If ok Then

                                    With currentWS
                                        CType(.Cells(zeile, 6), Excel.Range).Value = zeilensumme
                                        editRange = CType(.Range(.Cells(zeile, startSpalteDaten), .Cells(zeile, startSpalteDaten + bis - von)), Excel.Range)
                                    End With

                                    If schnittmenge.Sum > 0 Then
                                        For l As Integer = 0 To bis - von

                                            If l >= ixZeitraum And l <= ixZeitraum + breite - 1 Then
                                                editRange.Cells(1, l + 1).value = schnittmenge(l)
                                            Else
                                                editRange.Cells(1, l + 1).value = ""
                                            End If

                                        Next
                                    Else
                                        editRange.Value = ""
                                    End If

                                    atLeastOne = True

                                    zeile = zeile + 1
                                Else
                                    Call MsgBox("not ok")
                                End If

                            Next c

                            If Not atLeastOne Then

                                ' in diesem Fall sollte eine leere Projekt-Phasen-Information geschrieben werden, quasi ein Platzhalter
                                ' in diesem Platzhalter kann dann später die Ressourcen Information aufgenommen werden  


                                Dim ok As Boolean = massEditWrite1Zeile(currentWS.Name, hproj, cphase, indentlevel, isProtectedbyOthers, zeile, "", "", False,
                                                                            protectionText, von, bis,
                                                                            actualDataRelColumn, False, summeEditierenErlaubt,
                                                                            ixZeitraum, breite, startSpalteDaten, maxRCLengthVorkommen)

                                If ok Then
                                    zeile = zeile + 1
                                Else
                                    Call MsgBox("not ok")
                                End If

                            End If

                        End If

                    Next p

                End If

            Next

            ' für Testzwecke only 
            ' last Check - jetzt letzte
            ''Dim checkRange As Excel.Range = currentWS.UsedRange
            ''Dim anzZ As Integer = checkRange.Rows.Count
            ''Dim anzSp As Integer = checkRange.Columns.Count

            ''CType(currentWS.Cells(anzZ + 1, 1), Range).Value = "Anzahl Zeilen " & anzZ.ToString
            ''CType(currentWS.Cells(1, anzSp + 1), Range).Value = "Anzahl Spalten " & anzSp.ToString
            ' Ende für Testzwecke only 

            ' tk 7.12.16 kommt immer auf Fehler, weil nur 1 Zeile und eine Auswahl von Spalten .... 
            '' jetzt die erste Zeile so groß wie nötig machen 
            'Try
            '    ersteZeile.AutoFit()
            'Catch ex As Exception

            'End Try

            ' jetzt die Größe der Spalten für BU, pName, vName, Phasen-Name, RC-Name anpassen 

            Dim infoBlock As Excel.Range
            Dim infoDatablock As Excel.Range

            Try

                With CType(currentWS, Excel.Worksheet)
                    infoBlock = CType(.Range(.Columns(1), .Columns(startSpalteDaten - 2)), Excel.Range)
                    infoDatablock = CType(.Range(.Cells(2, 1), .Cells(zeile, startSpalteDaten - 2)), Excel.Range)

                    infoDatablock.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                    infoDatablock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

                    ' hier prüfen, ob es bereits Werte für massColValues gibt ..
                    If massColFontValues(0, 0) > 4 Then
                        ' diese Werte übernehmen 
                        infoDatablock.Font.Size = CInt(massColFontValues(0, 0))
                        For ik As Integer = 1 To 5
                            CType(infoBlock.Columns(ik), Excel.Range).ColumnWidth = massColFontValues(0, ik)
                        Next


                    Else
                        ' hier jetzt prüfen, ob nicht zu viel Platz eingenommen wird
                        infoBlock.AutoFit()

                        Try
                            'Dim availableScreenWidth As Double = appInstance.ActiveWindow.UsableWidth
                            'Dim availableScreenWidth As Double = CType(projectboardWindows(PTwindows.massEdit), Window).UsableWidth
                            Dim availableScreenWidth As Double = maxScreenWidth
                            If infoBlock.Width > 0.6 * availableScreenWidth Then

                                infoDatablock.Font.Size = CInt(CType(infoBlock.Cells(2, 2), Excel.Range).Font.Size) - 2
                                ' BU bekommt 5%
                                'CType(infoBlock.Columns(1), Excel.Range).ColumnWidth = 0.05 * 0.4 * availableScreenWidth
                                CType(infoBlock.Columns(1), Excel.Range).ColumnWidth = 3
                                ' pName bekomt 30%
                                'CType(infoBlock.Columns(2), Excel.Range).ColumnWidth = 0.3 * 0.4 * availableScreenWidth
                                CType(infoBlock.Columns(2), Excel.Range).ColumnWidth = 16
                                ' vName bekomt 5%
                                'CType(infoBlock.Columns(3), Excel.Range).ColumnWidth = 0.05 * 0.4 * availableScreenWidth
                                CType(infoBlock.Columns(3), Excel.Range).ColumnWidth = 3
                                ' phaseName bekomt 30%
                                'CType(infoBlock.Columns(4), Excel.Range).ColumnWidth = 0.3 * 0.4 * availableScreenWidth
                                CType(infoBlock.Columns(4), Excel.Range).ColumnWidth = 16
                                ' RoleCost Name bekomt 30%
                                'CType(infoBlock.Columns(5), Excel.Range).ColumnWidth = 0.3 * 0.4 * availableScreenWidth
                                CType(infoBlock.Columns(5), Excel.Range).ColumnWidth = 16
                            End If
                        Catch ex As Exception

                        End Try

                        ' Werte setzen ...
                        massColFontValues(0, 0) = CDbl(CType(infoBlock.Cells(2, 2), Excel.Range).Font.Size)
                        For ik As Integer = 1 To 5
                            massColFontValues(0, ik) = CType(infoBlock.Columns(ik), Excel.Range).ColumnWidth
                        Next

                    End If



                End With
            Catch ex As Exception

            End Try


            ' die Breite der Summen-Spalte festlegen 
            Try
                With CType(currentWS, Excel.Worksheet)
                    ' nur die Überschrift der Summe ...
                    infoBlock = CType(.Columns(startSpalteDaten - 1), Excel.Range)
                    infoBlock.ColumnWidth = 14
                    'infoBlock.AutoFit()
                End With
            Catch ex As Exception

            End Try

            Try
                With CType(currentWS, Excel.Worksheet)
                    ' nur die Überschrift der Summe ...
                    infoBlock = CType(.Cells(1, startSpalteDaten - 1), Excel.Range)
                    infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                    'infoBlock.AutoFit()
                End With
            Catch ex As Exception

            End Try

            Try
                With CType(currentWS, Excel.Worksheet)
                    ' nur den Datenbereich der Summe ...
                    infoBlock = CType(.Range(.Cells(2, startSpalteDaten - 1), .Cells(zeile, startSpalteDaten - 1)), Excel.Range)
                    infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                    infoBlock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                    infoBlock.Font.Size = CInt(massColFontValues(0, 0))
                    'infoBlock.AutoFit()
                End With
            Catch ex As Exception

            End Try
            ' Summe Datenbereich formatieren 


            ' die Breite der Daten festlegen 
            Try
                Dim tmpRange As Excel.Range
                With CType(currentWS, Excel.Worksheet)

                    For mis As Integer = 0 To bis - von
                        tmpRange = CType(.Range(.Cells(2, startSpalteDaten + mis), .Cells(zeile, startSpalteDaten + mis)), Excel.Range)

                        tmpRange.Columns.ColumnWidth = 5
                        'tmpRange.Font.Size = 10
                        If CInt(massColFontValues(0, 0)) > 3 Then
                            CType(tmpRange.Font, Excel.Font).Size = CInt(massColFontValues(0, 0) - 1)
                        Else
                            CType(tmpRange.Font, Excel.Font).Size = 9
                        End If

                        tmpRange.NumberFormat = "##,##0.#"
                        tmpRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

                    Next


                    Dim startMonat As Date = StartofCalendar.AddMonths(von - 1)
                    ' jetzt werden die Überschriften des Datenbereichs geschrieben 

                    For m As Integer = 0 To bis - von
                        With CType(.Cells(1, startSpalteDaten + m), Global.Microsoft.Office.Interop.Excel.Range)
                            .Value = startMonat.AddMonths(m)
                            If massColFontValues(0, 0) > 4 Then
                                .Font.Size = CInt(massColFontValues(0, 0))
                            Else
                                .Font.Size = 10
                            End If

                            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                            .NumberFormat = "[$-409]mmm yy;@"
                            .WrapText = False
                            .Orientation = 90
                            .ShrinkToFit = False
                            .AddIndent = False
                            .IndentLevel = 0
                            .ReadingOrder = Excel.Constants.xlContext
                        End With

                    Next

                End With
            Catch ex As Exception

            End Try


            appInstance.EnableEvents = True

        Catch ex As Exception
            appInstance.EnableEvents = True
        End Try




    End Sub

    ''' <summary>
    ''' schreibt eine Zeile für den MassenEditOnlien Ressourcen oder Kosten 
    ''' </summary>
    ''' <param name="wsName"></param>
    ''' <param name="hproj"></param>
    ''' <param name="cphase"></param>
    ''' <param name="indentLevel"></param>
    ''' <param name="isProtectedbyOthers"></param>
    ''' <param name="zeile"></param>
    ''' <param name="rcName"></param>
    ''' <param name="protectiontext"></param>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="actualdataRelColumn"></param>
    ''' <param name="summeEditierenErlaubt"></param>
    ''' <param name="ixZeitraum"></param>
    ''' <param name="breite"></param>
    ''' <param name="startSpalteDaten"></param>
    ''' <returns></returns>
    Public Function massEditWrite1Zeile(ByVal wsName As String, ByVal hproj As clsProjekt, ByVal cphase As clsPhase, ByVal indentLevel As Integer,
                                         ByVal isProtectedbyOthers As Boolean, ByVal zeile As Integer,
                                         ByVal rcName As String, ByVal rcNameID As String, ByVal isRole As Boolean,
                                         ByVal protectiontext As String,
                                         ByVal von As Integer, ByVal bis As Integer,
                                         ByVal actualdataRelColumn As Integer, ByVal hasActualdata As Boolean, ByVal summeEditierenErlaubt As Boolean,
                                         ByVal ixZeitraum As Integer, ByVal breite As Integer, ByVal startSpalteDaten As Integer, ByRef maxRcLength As Integer) As Boolean

        Dim currentWS As Excel.Worksheet = Nothing
        Dim writeResult As Boolean = False


        Try
            currentWS = appInstance.ActiveWorkbook.Worksheets(wsName)
        Catch ex As Exception
            writeResult = False
            massEditWrite1Zeile = writeResult
            Exit Function
        End Try

        Try

            ' Schreiben der Projekt-Informationen 
            With CType(currentWS, Excel.Worksheet)

                ' Business Unit schreiben 
                CType(.Cells(zeile, 1), Excel.Range).Value = hproj.businessUnit

                ' Name schreiben
                Call writeMEcellWithProjectName(CType(.Cells(zeile, 2), Excel.Range), hproj.name, isProtectedbyOthers, protectiontext)

                ' den Varianten-Namen schreiben
                CType(.Cells(zeile, 3), Excel.Range).Value = hproj.variantName

                ' Phase und ggf PhaseNameID schreiben
                Call writeMEcellWithPhaseNameID(CType(.Cells(zeile, 4), Excel.Range), indentLevel, cphase.name, cphase.nameID)


                ' Rolle oder Kostenart schreiben 
                Dim isLocked As Boolean = (isProtectedbyOthers Or hasActualdata)
                Call writeMECellWithRoleNameID(CType(.Cells(zeile, 5), Excel.Range), isLocked, rcName, rcNameID, isRole)


                ' das Format der Zeile mit der Summe
                CType(.Cells(zeile, 6), Excel.Range).NumberFormat = Format("##,##0.#")
                CType(.Cells(zeile, 6), Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

                If summeEditierenErlaubt Then
                    With CType(.Cells(zeile, 6), Excel.Range)

                        If Not isProtectedbyOthers Then
                            .Locked = False
                            '.Interior.Color = awinSettings.AmpelNichtBewertet
                            Try
                                If Not IsNothing(.Validation) Then
                                    .Validation.Delete()
                                End If
                                ' jetzt wird die ValidationList aufgebaut 
                                .Validation.Add(Type:=XlDVType.xlValidateDecimal,
                                                AlertStyle:=XlDVAlertStyle.xlValidAlertStop,
                                                Operator:=XlFormatConditionOperator.xlGreaterEqual,
                                                Formula1:="0")

                                ' jetzt wird der eventuell vorhandene Kommentar von Kostenart gelöscht
                                .ClearComments()

                            Catch ex As Exception

                            End Try
                        End If

                    End With
                End If

            End With


            ' jetzt werden die Monats-Werte formatiert 
            With CType(currentWS, Excel.Worksheet)

                For spix = 0 To bis - von

                    With CType(.Cells(zeile, spix + startSpalteDaten), Excel.Range)
                        If spix >= ixZeitraum And spix <= ixZeitraum + breite - 1 Then

                            If (Not isProtectedbyOthers) And (spix > actualdataRelColumn) Then
                                .Locked = False
                                Try
                                    If Not IsNothing(.Validation) Then
                                        .Validation.Delete()
                                    End If
                                    .Validation.Add(Type:=XlDVType.xlValidateDecimal,
                                                AlertStyle:=XlDVAlertStyle.xlValidAlertStop,
                                                Operator:=XlFormatConditionOperator.xlGreaterEqual,
                                                Formula1:="0")
                                Catch ex As Exception

                                End Try
                            End If

                            ' jetzt kommt die Farbsetzung ... die hängt nur von actualDataRelColumn ab
                            If spix <= actualdataRelColumn Then
                                .Interior.Color = awinSettings.AmpelNichtBewertet
                                .Font.Color = XlRgbColor.rgbBlack
                            Else
                                .Interior.Color = visboFarbeBlau
                                .Font.Color = XlRgbColor.rgbWhite
                            End If

                        Else
                            ' hier muss nichts getan werden ...
                        End If
                    End With

                Next

            End With

            writeResult = True
        Catch ex As Exception
            writeResult = False
        End Try

        massEditWrite1Zeile = writeResult

    End Function

    ''' <summary>
    ''' bestimmt das Erscheinungsbild der ersten Zeile in einem Mass-Edit Fenster Ressourcen, Termine, Attribute
    ''' </summary>
    ''' <param name="tableTyp"></param>
    Private Sub massEditZeile1Appearance(ByVal tableTyp As Integer)

        Dim currentWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(tableTyp)), Excel.Worksheet)
        Dim ersteZeile As Excel.Range = CType(currentWS.Rows(1), Excel.Range)

        With ersteZeile
            .RowHeight = awinSettings.zeilenhoehe1 + 5
            .Interior.Color = visboFarbeOrange
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = XlRgbColor.rgbWhite
            .VerticalAlignment = Excel.XlVAlign.xlVAlignTop
        End With

    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="todoListe"></param>
    Public Sub writeOnlineMassEditTermine(ByVal todoListe As Collection)

        Dim err As New clsErrorCodeMsg

        If todoListe.Count = 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("no projects for mass-edit available ..")
            Else
                Call MsgBox("keine Projekte für den Massen-Edit vorhanden ..")
            End If

            Exit Sub
        End If

        Try

            appInstance.EnableEvents = False

            ' jetzt die selectedProjekte Liste zurücksetzen ... ohne die currentConstellation zu verändern ...
            selectedProjekte.Clear(False)

            Dim currentWS As Excel.Worksheet
            Dim currentWB As Excel.Workbook
            Dim startDateColumn As Integer = 5
            Dim tmpName As String


            ' hier muss jetzt das entsprechende File aufgemacht werden ...
            ' das File 
            Try
                currentWB = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                currentWS = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meTE)), Excel.Worksheet)

                Try
                    ' off setzen des AutoFilter Modus ... 
                    If CType(currentWS, Excel.Worksheet).AutoFilterMode = True Then
                        'CType(CType(currentWS, Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
                        CType(currentWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
                    End If
                Catch ex As Exception

                End Try

                ' braucht man eigentlich nicht mehr, aber sicher ist sicher ...
                Try
                    currentWS.UsedRange.Clear()
                Catch ex As Exception

                End Try


            Catch ex As Exception
                Call MsgBox("es gibt Probleme mit dem Mass-Edit Worksheet ...")
                appInstance.EnableEvents = True
                Exit Sub
            End Try


            ' jetzt schreiben der ersten Zeile 
            Dim zeile As Integer = 1
            Dim spalte As Integer = 1

            Dim startSpalteDaten As Integer = 4
            'Dim roleCostNames As Excel.Range = Nothing
            Dim datesInput As Excel.Range = Nothing

            tmpName = ""

            ' Schreiben der Überschriften
            With CType(currentWS, Excel.Worksheet)

                If .ProtectContents Then
                    .Unprotect(Password:="x")
                End If


                If awinSettings.englishLanguage Then
                    CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
                    CType(.Cells(1, 2), Excel.Range).Value = "Project-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Variant-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Element-Name"
                    CType(.Cells(1, 5), Excel.Range).Value = "Start-Date"
                    CType(.Cells(1, 6), Excel.Range).Value = "End-Date"
                    CType(.Cells(1, 7), Excel.Range).Value = "Trafficlight"
                    CType(.Cells(1, 8), Excel.Range).Value = "Explanation"
                    CType(.Cells(1, 9), Excel.Range).Value = "Deliverables"
                    CType(.Cells(1, 10), Excel.Range).Value = "Responsible"
                    CType(.Cells(1, 11), Excel.Range).Value = "% Done"

                Else
                    CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
                    CType(.Cells(1, 2), Excel.Range).Value = "Projekt-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Varianten-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Element-Name"
                    CType(.Cells(1, 5), Excel.Range).Value = "Start-Datum"
                    CType(.Cells(1, 6), Excel.Range).Value = "End-Datum"
                    CType(.Cells(1, 7), Excel.Range).Value = "Ampel"
                    CType(.Cells(1, 8), Excel.Range).Value = "Erläuterung"
                    CType(.Cells(1, 9), Excel.Range).Value = "Lieferumfänge"
                    CType(.Cells(1, 10), Excel.Range).Value = "Verantwortlich"
                    CType(.Cells(1, 11), Excel.Range).Value = "% abgeschlossen"
                End If

                ' das Erscheinungsbild der Zeile 1 bestimmen  
                Call massEditZeile1Appearance(ptTables.meTE)


            End With


            zeile = 2


            For Each pvName As String In todoListe

                Dim hproj As clsProjekt = Nothing
                If AlleProjekte.Containskey(pvName) Then
                    hproj = AlleProjekte.getProject(pvName)
                End If

                If Not IsNothing(hproj) Then

                    ' ist das Projekt geschützt ? 
                    ' wenn nein, dann temporär schützen 
                    Dim protectionText As String = ""
                    Dim wpItem As clsWriteProtectionItem
                    Dim isProtectedbyOthers As Boolean

                    If awinSettings.visboServer Then
                        isProtectedbyOthers = Not (CType(databaseAcc, DBAccLayer.Request).checkChgPermission(hproj.name, hproj.variantName, dbUsername, err, ptPRPFType.project))
                    Else
                        isProtectedbyOthers = Not tryToprotectProjectforMe(hproj.name, hproj.variantName)
                    End If


                    If isProtectedbyOthers Then

                        ' nicht erfolgreich, weil durch anderen geschützt ... 
                        ' oder aber noch gar nicht in Datenbank: aber das ist noch nicht berücksichtigt  
                        wpItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                        writeProtections.upsert(wpItem)

                        protectionText = writeProtections.getProtectionText(calcProjektKey(hproj.name, hproj.variantName))

                    End If

                    ' jetzt wird für jedes Element in der Hierarchy eine Zeile rausgeschrieben 
                    ' das ist jetzt die rootphase-NameID
                    Dim curElemID As String = rootPhaseName
                    Dim indentLevel As Integer = 0
                    ' abbruchlevel gibt, wo die Funktion getNextIdOfId aufhört: erst an der Rootphase(=0) oder beim Element 
                    Dim abbruchLevel As Integer = 0
                    Dim indentOffset As Integer = 1

                    ' jetzt wird die Hierarchy abgeklappert .. beginnend mit dem ersten Element, der RootPhase
                    Do While curElemID <> ""

                        Dim cPhase As clsPhase = Nothing
                        Dim cMilestone As clsMeilenstein = Nothing
                        Dim isMilestone As Boolean = elemIDIstMeilenstein(curElemID)

                        If isMilestone Then
                            cMilestone = hproj.getMilestoneByID(curElemID)
                            ' schreibe den Meilenstein
                            With CType(currentWS, Excel.Worksheet)
                                ' Business-Unit
                                CType(.Cells(zeile, 1), Excel.Range).Value = hproj.businessUnit
                                ' Projekt-Name
                                CType(.Cells(zeile, 2), Excel.Range).Value = hproj.name
                                ' Varianten-Name
                                CType(.Cells(zeile, 3), Excel.Range).Value = hproj.variantName
                                ' Element-Name Meilenstein bzw. Phase inkl Indentlevel schreiben 
                                CType(.Cells(zeile, 4), Excel.Range).Value = cMilestone.name
                                CType(.Cells(zeile, 4), Excel.Range).IndentLevel = indentLevel
                                ' Startdatum, gibt es bei Meilensteinen nicht, deswegen sperren  
                                CType(.Cells(zeile, 5), Excel.Range).Value = ""
                                ' Ende-Datum 
                                CType(.Cells(zeile, 6), Excel.Range).Value = cMilestone.getDate.ToShortDateString
                                ' Ampel-Farbe
                                CType(.Cells(zeile, 7), Excel.Range).Value = cMilestone.ampelStatus
                                ' Ampel-Erläuterung
                                CType(.Cells(zeile, 8), Excel.Range).Value = cMilestone.ampelErlaeuterung
                                ' Lieferumfänge
                                CType(.Cells(zeile, 9), Excel.Range).Value = cMilestone.getAllDeliverables
                                ' wer ist verantwortlich
                                CType(.Cells(zeile, 10), Excel.Range).Value = cMilestone.verantwortlich
                                ' wieviel ist erledigt ? 
                                CType(.Cells(zeile, 11), Excel.Range).Value = cMilestone.percentDone.ToString("0#%")
                            End With
                        Else
                            cPhase = hproj.getPhaseByID(curElemID)
                            ' schreibe die Phase
                            With CType(currentWS, Excel.Worksheet)
                                ' Business-Unit
                                CType(.Cells(zeile, 1), Excel.Range).Value = hproj.businessUnit
                                ' Projekt-Name
                                CType(.Cells(zeile, 2), Excel.Range).Value = hproj.name
                                ' Varianten-Name
                                CType(.Cells(zeile, 3), Excel.Range).Value = hproj.variantName
                                ' Element-Name Meilenstein bzw. Phase
                                CType(.Cells(zeile, 4), Excel.Range).Value = cPhase.name
                                CType(.Cells(zeile, 4), Excel.Range).IndentLevel = indentLevel
                                ' Startdatum 
                                CType(.Cells(zeile, 5), Excel.Range).Value = cPhase.getStartDate.ToShortDateString
                                ' Ende-Datum 
                                CType(.Cells(zeile, 6), Excel.Range).Value = cPhase.getEndDate.ToShortDateString
                                ' Ampel-Farbe
                                CType(.Cells(zeile, 7), Excel.Range).Value = cPhase.ampelStatus
                                ' Ampel-Erläuterung
                                CType(.Cells(zeile, 8), Excel.Range).Value = cPhase.ampelErlaeuterung
                                ' Lieferumfänge
                                CType(.Cells(zeile, 9), Excel.Range).Value = cPhase.getAllDeliverables
                                ' wer ist verantwortlich
                                CType(.Cells(zeile, 10), Excel.Range).Value = cPhase.verantwortlich
                                ' wieviel ist erledigt ? 
                                CType(.Cells(zeile, 11), Excel.Range).Value = cPhase.percentDone.ToString("0#%")
                            End With
                        End If

                        ' jetzt müssen die locked Attribute gesetzt werden entsprechend der isProtectedbyOthers ...

                        ' jetzt muss geprüft werden, ob es durch jdn anders geschützt wurde ... 
                        If isProtectedbyOthers Then

                            Dim kompletteZeile As Excel.Range = CType(currentWS.Rows(zeile), Excel.Range)

                            With CType(currentWS, Excel.Worksheet)
                                CType(.Cells(zeile, 2), Excel.Range).Font.Color = awinSettings.protectedByOtherColor
                                ' Kommentar einfügen 
                                Dim cellComment As Excel.Comment = CType(.Cells(zeile, 2), Excel.Range).Comment
                                If Not IsNothing(cellComment) Then
                                    CType(.Cells(zeile, 2), Excel.Range).Comment.Delete()
                                End If
                                CType(.Cells(zeile, 2), Excel.Range).AddComment(Text:=protectionText)
                                CType(.Cells(zeile, 2), Excel.Range).Comment.Visible = False
                            End With

                            kompletteZeile.Locked = True

                        Else
                            Dim protectArea As Excel.Range = Nothing
                            Dim editArea As Excel.Range = Nothing
                            If isMilestone Then
                                With currentWS
                                    protectArea = CType(.Range(.Cells(zeile, 1), .Cells(zeile, 5)), Excel.Range)
                                    editArea = CType(.Range(.Cells(zeile, 6), .Cells(zeile, 11)), Excel.Range)
                                End With
                            Else
                                With currentWS
                                    protectArea = CType(.Range(.Cells(zeile, 1), .Cells(zeile, 4)), Excel.Range)
                                    editArea = CType(.Range(.Cells(zeile, 5), .Cells(zeile, 11)), Excel.Range)
                                End With
                            End If
                            protectArea.Locked = True
                            editArea.Locked = False
                        End If

                        ' Zeile eins weiter ... 
                        zeile = zeile + 1
                        curElemID = hproj.hierarchy.getNextIdOfId(curElemID, indentLevel, abbruchLevel)

                    Loop

                End If

            Next


            ' jetzt die Größe der Spalten für BU, pName, vName, Phasen-Name, RC-Name anpassen 
            Dim infoBlock As Excel.Range
            Dim infoDataBlock As Excel.Range
            With CType(currentWS, Excel.Worksheet)
                infoBlock = CType(.Range(.Columns(1), .Columns(11)), Excel.Range)
                infoDataBlock = CType(.Range(.Cells(2, 1), .Cells(zeile + 100, 11)), Excel.Range)
                infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                infoBlock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter


                ' die Besonderheiten abbilden 

                ' Phasen bzw. Meilenstein Name
                With CType(infoDataBlock.Columns(4), Excel.Range)
                    .WrapText = True
                End With

                ' Erläuterung
                With CType(infoDataBlock.Columns(8), Excel.Range)
                    .WrapText = True
                End With

                ' Lieferumfänge 
                With CType(infoDataBlock.Columns(9), Excel.Range)
                    .WrapText = True
                End With

                ' percent Done 
                With CType(infoBlock.Columns(11), Excel.Range)
                    .NumberFormat = "0#%"
                End With

                ' hier prüfen, ob es bereits Werte für massColValues gibt ..
                If massColFontValues(1, 0) > 4 Then
                    ' diese Werte übernehmen 
                    infoDataBlock.Font.Size = CInt(massColFontValues(1, 0))
                    For ik As Integer = 1 To 11
                        If massColFontValues(1, ik) > 0 Then
                            CType(infoBlock.Columns(ik), Excel.Range).ColumnWidth = massColFontValues(1, ik)
                        End If

                    Next


                Else
                    ' hier jetzt prüfen, ob nicht zu viel Platz eingenommen wird
                    Try
                        infoDataBlock.AutoFit()
                    Catch ex As Exception

                    End Try


                    Try
                        'Dim availableScreenWidth As Double = appInstance.ActiveWindow.UsableWidth
                        'Dim availableScreenWidth As Double = CType(projectboardWindows(PTwindows.massEdit), Window).UsableWidth
                        Dim availableScreenWidth As Double = maxScreenWidth
                        If infoDataBlock.Width > availableScreenWidth Then

                            infoDataBlock.Font.Size = CInt(CType(infoBlock.Cells(2, 2), Excel.Range).Font.Size) - 2
                            infoDataBlock.AutoFit()

                        End If
                    Catch ex As Exception

                    End Try

                    ' Werte setzen ...
                    massColFontValues(1, 0) = CDbl(CType(infoBlock.Cells(2, 2), Excel.Range).Font.Size)
                    For ik As Integer = 1 To 11
                        massColFontValues(1, ik) = CType(infoBlock.Columns(ik), Excel.Range).ColumnWidth
                    Next

                End If

                ' jetzt noch die Spalte 7 bedingt formatieren .. 
                Dim trafficLightRange As Excel.Range = CType(.Range(.Cells(2, 7), .Cells(zeile, 7)), Excel.Range)
                With trafficLightRange
                    .Interior.Color = visboFarbeNone

                    Dim trafficLightColorScale As Excel.ColorScale = .FormatConditions.AddColorScale(3)

                    CType(trafficLightColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).Value = "1"
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeGreen

                    CType(trafficLightColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).Value = "2"
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeYellow

                    CType(trafficLightColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).Value = "3"
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeRed

                End With

            End With

            appInstance.EnableEvents = True

        Catch ex As Exception
            Call MsgBox("Fehler in Aufbereitung Termine" & vbLf & ex.Message)
            appInstance.EnableEvents = True
        End Try


    End Sub

    ''' <summary>
    ''' massen-Editieren von Projekt-Attributen
    ''' </summary>
    ''' <param name="todoListe"></param>
    Public Sub writeOnlineMassEditAttribute(ByVal todoListe As Collection)

        Dim err As New clsErrorCodeMsg

        If todoListe.Count = 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("no projects for mass-edit available ..")
            Else
                Call MsgBox("keine Projekte für den Massen-Edit vorhanden ..")
            End If

            Exit Sub
        End If

        Try

            appInstance.EnableEvents = False

            ' jetzt die selectedProjekte Liste zurücksetzen ... ohne die currentConstellation zu verändern ...
            selectedProjekte.Clear(False)

            Dim currentWS As Excel.Worksheet
            Dim currentWB As Excel.Workbook
            Dim startDateColumn As Integer = 5

            Dim anzahlSpalten As Integer = 13 + customFieldDefinitions.count

            ' hier muss jetzt das entsprechende File aufgemacht werden ...
            ' das File 
            Try
                currentWB = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                currentWS = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meAT)), Excel.Worksheet)

                Try
                    ' off setzen des AutoFilter Modus ... 
                    If CType(currentWS, Excel.Worksheet).AutoFilterMode = True Then
                        'CType(CType(currentWS, Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
                        CType(currentWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
                    End If
                Catch ex As Exception

                End Try

                ' braucht man eigentlich nicht mehr, aber sicher ist sicher ...
                Try
                    currentWS.UsedRange.Clear()
                Catch ex As Exception

                End Try


            Catch ex As Exception
                Call MsgBox("es gibt Probleme mit dem Mass-Edit Worksheet ...")
                appInstance.EnableEvents = True
                Exit Sub
            End Try


            ' jetzt schreiben der ersten Zeile 
            Dim zeile As Integer = 1
            Dim spalte As Integer = 1


            ' Schreiben der Überschriften
            With CType(currentWS, Excel.Worksheet)

                If .ProtectContents Then
                    .Unprotect(Password:="x")
                End If


                If awinSettings.englishLanguage Then
                    CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
                    CType(.Cells(1, 2), Excel.Range).Value = "Project-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Variant-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Start"
                    CType(.Cells(1, 5), Excel.Range).Value = "End"
                    CType(.Cells(1, 6), Excel.Range).Value = "Goals"
                    CType(.Cells(1, 7), Excel.Range).Value = "Description Variant"
                    CType(.Cells(1, 8), Excel.Range).Value = "Responsible"
                    CType(.Cells(1, 9), Excel.Range).Value = "Traffic-Light"
                    CType(.Cells(1, 10), Excel.Range).Value = "Explanation"
                    CType(.Cells(1, 11), Excel.Range).Value = "Strategic Fit"
                    CType(.Cells(1, 12), Excel.Range).Value = "Risk"
                    CType(.Cells(1, 13), Excel.Range).Value = "Risk Description"


                Else
                    CType(.Cells(1, 1), Excel.Range).Value = "Business-Unit"
                    CType(.Cells(1, 2), Excel.Range).Value = "Projekt-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Varianten-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Start"
                    CType(.Cells(1, 5), Excel.Range).Value = "Ende"
                    CType(.Cells(1, 6), Excel.Range).Value = "Ziele"
                    CType(.Cells(1, 7), Excel.Range).Value = "Beschreibung (Variante)"
                    CType(.Cells(1, 8), Excel.Range).Value = "Verantwortlich"
                    CType(.Cells(1, 9), Excel.Range).Value = "Projekt-Ampel"
                    CType(.Cells(1, 10), Excel.Range).Value = "Erläuterung"
                    CType(.Cells(1, 11), Excel.Range).Value = "Strategischer Fit"
                    CType(.Cells(1, 12), Excel.Range).Value = "Risiko"
                    CType(.Cells(1, 13), Excel.Range).Value = "Risiko-Beschreibung"


                End If

                ' jetzt noch die CustomFields
                For i As Integer = 1 To customFieldDefinitions.count

                    Dim cfType As Integer = customFieldDefinitions.getDef(i).type
                    Dim tmpName As String = customFieldDefinitions.getDef(i).name

                    Try
                        If cfType = 0 Then
                            ' String
                            CType(.Cells(1, 13 + i), Excel.Range).Value = tmpName & " (S)"
                        ElseIf cfType = 1 Then
                            ' Double
                            CType(.Cells(1, 13 + i), Excel.Range).Value = tmpName & " (D)"
                        ElseIf cfType = 2 Then
                            ' boolean 
                            CType(.Cells(1, 13 + i), Excel.Range).Value = tmpName & " (B)"
                        End If

                    Catch ex As Exception

                    End Try

                Next


                ' das Erscheinungsbild der Zeile 1 bestimmen  
                Call massEditZeile1Appearance(ptTables.meAT)


            End With


            zeile = 2


            For Each pvName As String In todoListe

                Dim hproj As clsProjekt = Nothing
                If AlleProjekte.Containskey(pvName) Then
                    hproj = AlleProjekte.getProject(pvName)
                End If

                If Not IsNothing(hproj) Then

                    ' ist das Projekt geschützt ? 
                    ' wenn nein, dann temporär schützen 
                    Dim protectionText As String = ""
                    Dim wpItem As clsWriteProtectionItem
                    Dim isProtectedbyOthers As Boolean

                    If awinSettings.visboServer Then
                        isProtectedbyOthers = Not (CType(databaseAcc, DBAccLayer.Request).checkChgPermission(hproj.name, hproj.variantName, dbUsername, err, ptPRPFType.project))
                    Else
                        isProtectedbyOthers = Not tryToprotectProjectforMe(hproj.name, hproj.variantName)
                    End If


                    If isProtectedbyOthers Then

                        ' nicht erfolgreich, weil durch anderen geschützt ... 
                        ' oder aber noch gar nicht in Datenbank: aber das ist noch nicht berücksichtigt  
                        wpItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                        writeProtections.upsert(wpItem)

                        protectionText = writeProtections.getProtectionText(calcProjektKey(hproj.name, hproj.variantName))

                    End If

                    ' jetzt wird für jedes Projekt genau eine Zeile geschrieben 
                    With CType(currentWS, Excel.Worksheet)
                        CType(.Cells(zeile, 1), Excel.Range).Value = hproj.businessUnit
                        CType(.Cells(zeile, 2), Excel.Range).Value = hproj.name
                        CType(.Cells(zeile, 3), Excel.Range).Value = hproj.variantName
                        CType(.Cells(zeile, 4), Excel.Range).Value = hproj.startDate.ToShortDateString
                        CType(.Cells(zeile, 5), Excel.Range).Value = hproj.endeDate.ToShortDateString
                        CType(.Cells(zeile, 6), Excel.Range).Value = hproj.description
                        CType(.Cells(zeile, 7), Excel.Range).Value = hproj.variantDescription
                        CType(.Cells(zeile, 8), Excel.Range).Value = hproj.leadPerson
                        CType(.Cells(zeile, 9), Excel.Range).Value = hproj.ampelStatus
                        CType(.Cells(zeile, 10), Excel.Range).Value = hproj.ampelErlaeuterung
                        CType(.Cells(zeile, 11), Excel.Range).Value = hproj.StrategicFit
                        CType(.Cells(zeile, 12), Excel.Range).Value = hproj.Risiko
                        CType(.Cells(zeile, 13), Excel.Range).Value = ""

                        For i As Integer = 1 To customFieldDefinitions.count
                            Dim cfType As Integer = customFieldDefinitions.getDef(i).type
                            Dim uid As Integer = customFieldDefinitions.getDef(i).uid

                            Try
                                If cfType = 0 Then
                                    ' String
                                    CType(.Cells(zeile, 13 + i), Excel.Range).Value = CStr(hproj.getCustomSField(uid))
                                ElseIf cfType = 1 Then
                                    ' Double
                                    CType(.Cells(zeile, 13 + i), Excel.Range).Value = CDbl(hproj.getCustomDField(uid))
                                ElseIf cfType = 2 Then
                                    ' boolean 
                                    CType(.Cells(zeile, 13 + i), Excel.Range).Value = CBool(hproj.getCustomBField(uid))
                                End If

                            Catch ex As Exception

                            End Try

                        Next
                    End With



                    ' jetzt müssen die locked Attribute gesetzt werden entsprechend der isProtectedbyOthers ...

                    ' jetzt muss geprüft werden, ob es durch jdn anders geschützt wurde ... 
                    If isProtectedbyOthers Then

                        Dim kompletteZeile As Excel.Range = CType(currentWS.Rows(zeile), Excel.Range)

                        With CType(currentWS, Excel.Worksheet)
                            CType(.Cells(zeile, 2), Excel.Range).Font.Color = awinSettings.protectedByOtherColor
                            ' Kommentar einfügen 
                            Dim cellComment As Excel.Comment = CType(.Cells(zeile, 2), Excel.Range).Comment
                            If Not IsNothing(cellComment) Then
                                CType(.Cells(zeile, 2), Excel.Range).Comment.Delete()
                            End If
                            CType(.Cells(zeile, 2), Excel.Range).AddComment(Text:=protectionText)
                            CType(.Cells(zeile, 2), Excel.Range).Comment.Visible = False
                        End With

                        kompletteZeile.Locked = True

                    Else
                        Dim protectArea As Excel.Range = Nothing
                        Dim editArea As Excel.Range = Nothing

                        With currentWS
                            protectArea = CType(.Range(.Cells(zeile, 1), .Cells(zeile, 5)), Excel.Range)
                            editArea = CType(.Range(.Cells(zeile, 6), .Cells(zeile, anzahlSpalten)), Excel.Range)
                        End With

                        protectArea.Locked = True
                        editArea.Locked = False
                    End If

                    ' Zeile eins weiter ... 
                    zeile = zeile + 1

                End If

            Next


            ' jetzt die Größe der Spalten für BU, pName, vName, Phasen-Name, RC-Name anpassen 
            Dim infoBlock As Excel.Range
            Dim infoDataBlock As Excel.Range
            With CType(currentWS, Excel.Worksheet)

                infoBlock = CType(.Range(.Columns(1), .Columns(anzahlSpalten)), Excel.Range)
                infoDataBlock = CType(.Range(.Cells(2, 1), .Cells(zeile + 100, anzahlSpalten)), Excel.Range)
                infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                infoBlock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter


                ' hier prüfen, ob es bereits Werte für massColValues gibt ..
                If massColFontValues(1, 0) > 4 Then
                    ' diese Werte übernehmen 
                    infoDataBlock.Font.Size = CInt(massColFontValues(2, 0))
                    For ik As Integer = 1 To anzahlSpalten
                        If massColFontValues(1, ik) > 0 Then
                            CType(infoBlock.Columns(ik), Excel.Range).ColumnWidth = massColFontValues(1, ik)
                        End If

                    Next


                Else

                    'Dim ersteZeile As Excel.Range = CType(currentWS.Rows(1), Excel.Range)
                    'Try
                    '    ersteZeile.AutoFit()
                    'Catch ex As Exception

                    'End Try

                    ' die Besonderheiten abbilden 
                    ' BU
                    ' Description

                    Try
                        With CType(infoDataBlock.Columns(1), Excel.Range)
                            .ColumnWidth = 13
                        End With

                        '' Projekt-Name
                        With CType(infoDataBlock.Columns(2), Excel.Range)
                            .ColumnWidth = 15
                        End With

                        '' Varianten-Name
                        With CType(infoDataBlock.Columns(3), Excel.Range)
                            .ColumnWidth = 5
                        End With

                        '' Start
                        With CType(infoDataBlock.Columns(4), Excel.Range)
                            .ColumnWidth = 10
                        End With

                        '' Ende
                        With CType(infoDataBlock.Columns(5), Excel.Range)
                            .ColumnWidth = 10
                        End With

                        '' Description
                        With CType(infoDataBlock.Columns(6), Excel.Range)
                            .WrapText = True
                            .ColumnWidth = 20
                        End With

                        '' Variant Description
                        With CType(infoDataBlock.Columns(7), Excel.Range)
                            .WrapText = True
                            .ColumnWidth = 20
                        End With

                        '' Verantwortlich 
                        With CType(infoDataBlock.Columns(8), Excel.Range)
                            .ColumnWidth = 10
                        End With

                        '' Ampel-Farbe
                        With CType(infoDataBlock.Columns(9), Excel.Range)
                            .ColumnWidth = 2
                        End With

                        '' Ampel-Erläuterung
                        With CType(infoDataBlock.Columns(10), Excel.Range)
                            .ColumnWidth = 20
                            .WrapText = True
                        End With

                        ' Strategic Fit 
                        With CType(infoDataBlock.Columns(11), Excel.Range)
                            .ColumnWidth = 14
                            .HorizontalAlignment = HorizontalAlignment.Center
                        End With

                        ' Risiko  
                        With CType(infoDataBlock.Columns(12), Excel.Range)
                            .ColumnWidth = 42
                            .HorizontalAlignment = HorizontalAlignment.Center
                        End With

                        ' Risiko-Beschreibung  
                        With CType(infoDataBlock.Columns(13), Excel.Range)
                            .ColumnWidth = 20
                        End With

                        For i As Integer = 1 To customFieldDefinitions.count

                            With CType(infoDataBlock.Columns(13 + i), Excel.Range)
                                .ColumnWidth = 12
                            End With

                        Next

                    Catch ex As Exception

                    End Try


                End If

                ' jetzt noch die Spalte 9 bedingt formatieren .. 
                Dim trafficLightRange As Excel.Range = CType(.Range(.Cells(2, 9), .Cells(zeile, 9)), Excel.Range)
                With trafficLightRange
                    .Interior.Color = visboFarbeNone

                    Dim trafficLightColorScale As Excel.ColorScale = .FormatConditions.AddColorScale(3)

                    CType(trafficLightColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).Value = "1"
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(1), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeGreen

                    CType(trafficLightColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).Value = "2"
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(2), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeYellow

                    CType(trafficLightColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).Type = XlConditionValueTypes.xlConditionValueNumber
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).Value = "3"
                    CType(trafficLightColorScale.ColorScaleCriteria.Item(3), Excel.ColorScaleCriterion).FormatColor.Color = visboFarbeRed

                End With

            End With

            appInstance.EnableEvents = True

        Catch ex As Exception
            Call MsgBox("Fehler in Aufbereitung Termine" & vbLf & ex.Message)
            appInstance.EnableEvents = True
        End Try



    End Sub

    ''' <summary>
    ''' liest, falls vorhanden aus ProjectboardConfig.xml die Settings
    ''' wenn nicht vorhanden, gibt false zurück 
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns>ob erfolgreich oder nicht </returns>
    ''' <remarks></remarks>
    Public Function readawinSettings(ByVal path As String) As Boolean


        Dim cfgs As New configuration
        Dim cfgFile As String = path & "\ProjectboardConfig.xml"

        Dim erg As Boolean = My.Computer.FileSystem.FileExists(cfgFile)

        Try

            cfgs = XMLImportConfig(cfgFile)

            If Not IsNothing(cfgs) Then

                Dim anzahlSettings As Integer = cfgs.applicationSettings.ExcelWorkbook1MySettings.Length

                For i = 0 To anzahlSettings - 1

                    Select Case cfgs.applicationSettings.ExcelWorkbook1MySettings(i).name
                        Case "mongoDBURL"
                            awinSettings.databaseURL = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "mongoDBname"
                            awinSettings.databaseName = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "mongoDBWithSSL"
                            awinSettings.DBWithSSL = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "proxyServerURL"
                            awinSettings.proxyURL = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "globalPath"
                            awinSettings.globalPath = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "awinPath"
                            awinSettings.awinPath = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "TaskClass"
                            awinSettings.visboTaskClass = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOAbbreviation"
                            awinSettings.visboAbbreviation = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOAmpel"
                            awinSettings.visboAmpel = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOAmpelText"
                            awinSettings.visboAmpelText = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOdeliverables"
                            awinSettings.visbodeliverables = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOresponsible"
                            awinSettings.visboresponsible = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOpercentDone"
                            awinSettings.visbopercentDone = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOMapping"
                            awinSettings.visboMapping = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "userNamePWD"
                            awinSettings.userNamePWD = cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value
                        Case "VISBOServer"
                            awinSettings.visboServer = CType(cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value, Boolean)
                        Case "mongoDBWithSSL"
                            awinSettings.DBWithSSL = CType(cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value, Boolean)
                        Case "VISBODebug"
                            awinSettings.visboDebug = CType(cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value, Boolean)
                        Case "rememberUserPWD"
                            awinSettings.rememberUserPwd = CType(cfgs.applicationSettings.ExcelWorkbook1MySettings(i).value, Boolean)

                    End Select
                Next

                readawinSettings = True

            Else

                readawinSettings = False

            End If

        Catch ex As Exception

            readawinSettings = False

        End Try

    End Function



    ''' <summary>
    ''' liest das Customization File aus und initialisiert die globalen Variablen entsprechend
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinsetTypen(ByVal special As String)
        Try
            Dim err As New clsErrorCodeMsg

            ' neu 9.11.2016
            Dim formerSU As Boolean = True
            Dim needToBeSaved As Boolean = False
            '  um dahinter temporär die Darstellungsklassen kopieren zu können , nur für ProjectBoard nötig 
            Dim projectBoardSheet As Excel.Worksheet = Nothing

            Dim xlsCustomization As Excel.Workbook = Nothing


            Dim anzIEOrdner As Integer = [Enum].GetNames(GetType(PTImpExp)).Length
            ReDim importOrdnerNames(anzIEOrdner - 1)
            ReDim exportOrdnerNames(anzIEOrdner - 1)


            ' Auslesen des Window Namens 
            Dim accountToken As IntPtr = WindowsIdentity.GetCurrent().Token
            Dim myUser As New WindowsIdentity(accountToken)
            myWindowsName = myUser.Name


            globalPath = awinSettings.globalPath


            ' Debug-Mode?
            If awinSettings.visboDebug Then
                If Not IsNothing(globalPath) Then
                    If globalPath.Length > 0 Then
                        Call MsgBox("GlobalPath:" & globalPath & vbLf &
                                    "existiert: " & My.Computer.FileSystem.DirectoryExists(globalPath).ToString)
                    Else
                        Call MsgBox("GlobalPath: leerer String")
                    End If
                Else
                    Call MsgBox("GlobalPath: Nothing")
                End If


            End If

            ' tk 12.12.18 damit wird sichergestellt, dass bei einer Installation die Demo Daten einfach im selben Directory liegen können
            ' im ProjectBoardConfig kann demnach entweder der leere String stehen oder aber ein relativer Pfad, der vom User/Home Directory ausgeht ... 
            Dim locationOfProjectBoard = My.Computer.FileSystem.GetParentPath(appInstance.ActiveWorkbook.FullName)
            Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments

            Dim stdDemoDataName As String = "VISBO Demo-Daten"

            If awinSettings.awinPath = "" Then
                awinPath = My.Computer.FileSystem.CombinePath(locationOfProjectBoard, stdDemoDataName)
                If My.Computer.FileSystem.DirectoryExists(awinPath) Then
                    ' alles ok
                Else
                    awinPath = My.Computer.FileSystem.CombinePath(curUserDir, stdDemoDataName)
                    If My.Computer.FileSystem.DirectoryExists(awinPath) Then
                        ' alles ok 
                    End If
                End If
            Else
                awinPath = My.Computer.FileSystem.CombinePath(curUserDir, awinSettings.awinPath)
            End If


            If Not awinPath.EndsWith("\") Then
                awinPath = awinPath & "\"
            End If


            ' Debug-Mode?
            If awinSettings.visboDebug Then
                Call MsgBox("awinPath:" & vbLf & awinPath)
                Call MsgBox("globalPath:" & vbLf & globalPath)

            End If
            If awinSettings.visboDebug And special <> "BHTC" Then
                Call MsgBox("Betriebssystem: " & appInstance.OperatingSystem & Chr(10) &
                           "Excel-Version: " & appInstance.Version, vbInformation, "Info")
            End If

            If awinPath = "" And (globalPath <> "" And My.Computer.FileSystem.DirectoryExists(globalPath)) Then
                awinPath = globalPath
            ElseIf globalPath = "" And (awinPath <> "" And My.Computer.FileSystem.DirectoryExists(awinPath)) Then
                globalPath = awinPath
            ElseIf globalPath = "" Or awinPath = "" Then
                Throw New ArgumentException("Globaler Ordner " & awinSettings.globalPath & " und Lokaler Ordner " & awinSettings.awinPath & " existieren nicht")
            End If

            If My.Computer.FileSystem.DirectoryExists(globalPath) And (Dir(globalPath, vbDirectory) = "") Then
                Throw New ArgumentException("Requirementsordner " & awinSettings.globalPath & " existiert nicht")
            End If



            If Not globalPath.EndsWith("\") Then
                globalPath = globalPath & "\"
            End If

            ' Synchronization von Globalen und Lokalen Pfad

            If awinPath <> globalPath And My.Computer.FileSystem.DirectoryExists(globalPath) Then

                If awinSettings.visboDebug Then
                    Call MsgBox("jetzt wird synchronisiert ...")
                End If

                Call synchronizeGlobalToLocalFolder()

            Else

                If awinSettings.visboDebug Then
                    If awinPath = globalPath Then
                        Call MsgBox("awinPath = globalPath: keine Synchronisierung ...")
                    Else
                        Call MsgBox("globalPath existiert nicht: " & vbLf & globalPath)
                    End If

                End If

                If My.Computer.FileSystem.DirectoryExists(awinPath) And (Dir(awinPath, vbDirectory) = "") Then
                    Throw New ArgumentException("Requirementsordner " & awinSettings.awinPath & " existiert nicht")
                End If

            End If


            ' Erzeugen des Report Ordners, wenn er nicht schon existiert ..

            reportOrdnerName = awinPath & "Reports\"
            Try
                My.Computer.FileSystem.CreateDirectory(reportOrdnerName)
            Catch ex As Exception

            End Try

            importOrdnerNames(PTImpExp.visbo) = awinPath & "Import\VISBO Steckbriefe"
            importOrdnerNames(PTImpExp.rplan) = awinPath & "Import\RPLAN-Excel"
            importOrdnerNames(PTImpExp.msproject) = awinPath & "Import\MSProject"
            importOrdnerNames(PTImpExp.simpleScen) = awinPath & "Import\einfache Szenarien"
            importOrdnerNames(PTImpExp.modulScen) = awinPath & "Import\modulare Szenarien"
            importOrdnerNames(PTImpExp.addElements) = awinPath & "Import\addOn Regeln"
            importOrdnerNames(PTImpExp.rplanrxf) = awinPath & "Import\RXF Files"
            importOrdnerNames(PTImpExp.massenEdit) = awinPath & "Import\massEdit"
            importOrdnerNames(PTImpExp.offlineData) = awinPath & "Import\massEdit"
            importOrdnerNames(PTImpExp.scenariodefs) = awinPath & "Import\Scenario Definitions"
            importOrdnerNames(PTImpExp.Orga) = awinPath & "requirements"
            importOrdnerNames(PTImpExp.customUserRoles) = awinPath & "requirements"
            importOrdnerNames(PTImpExp.actualData) = awinPath & "Import\einfache Szenarien"

            exportOrdnerNames(PTImpExp.visbo) = awinPath & "Export\VISBO Steckbriefe"
            exportOrdnerNames(PTImpExp.rplan) = awinPath & "Export\RPLAN-Excel"
            exportOrdnerNames(PTImpExp.msproject) = awinPath & "Export\MSProject"
            exportOrdnerNames(PTImpExp.simpleScen) = awinPath & "Export\einfache Szenarien"
            exportOrdnerNames(PTImpExp.modulScen) = awinPath & "Export\modulare Szenarien"
            exportOrdnerNames(PTImpExp.massenEdit) = awinPath & "Export\massEdit"
            exportOrdnerNames(PTImpExp.scenariodefs) = awinPath & "Export\Scenario Definitions"

            If special = "ProjectBoard" Then

                ' jetzt werden die Directories alle angelegt, sofern Sie nicht schon existieren ... 
                For di As Integer = 0 To importOrdnerNames.Length - 1
                    Try
                        My.Computer.FileSystem.CreateDirectory(importOrdnerNames(di))
                    Catch ex As Exception

                    End Try
                Next

                For di As Integer = 0 To exportOrdnerNames.Length - 1
                    Try
                        My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(di))
                    Catch ex As Exception

                    End Try
                Next

            End If ' if special

            StartofCalendar = StartofCalendar.Date

            LizenzKomponenten(PTSWKomp.ProjectAdmin) = "ProjectAdmin"
            LizenzKomponenten(PTSWKomp.Swimlanes2) = "Swimlanes2"
            LizenzKomponenten(PTSWKomp.Premium) = "Premium"
            LizenzKomponenten(PTSWKomp.SWkomp2) = "SWkomp2"
            LizenzKomponenten(PTSWKomp.SWkomp3) = "SWkomp3"
            LizenzKomponenten(PTSWKomp.SWkomp4) = "SWkomp4"

            ' 14.11.16 tk nicht mehr notwenig , wird in Module initial gesetzt 
            ''ProjektStatus(0) = "geplant"
            ''ProjektStatus(1) = "beauftragt"
            ''ProjektStatus(2) = "beauftragt, Änderung noch nicht freigegeben"
            ''ProjektStatus(3) = "beendet" ' ein Projekt wurde in seinem Verlauf beendet, ohne es plangemäß abzuschliessen
            ''ProjektStatus(4) = "abgeschlossen"


            DiagrammTypen(0) = "Phase"
            DiagrammTypen(1) = "Rolle"
            DiagrammTypen(2) = "Kostenart"
            DiagrammTypen(3) = "Portfolio"
            DiagrammTypen(4) = "Ergebnis"
            DiagrammTypen(5) = "Meilenstein"
            DiagrammTypen(6) = "Meilenstein Trendanalyse"
            DiagrammTypen(7) = "Phasen-Kategorie"
            DiagrammTypen(8) = "Meilenstein-Kategorie"


            Try
                repMessages = XMLImportReportMsg(repMsgFileName, awinSettings.ReportLanguage)
                Call setLanguageMessages()
            Catch ex As Exception

            End Try

            autoSzenarioNamen(0) = "vor Optimierung"
            autoSzenarioNamen(1) = "1. Optimum"
            autoSzenarioNamen(2) = "2. Optimum"
            autoSzenarioNamen(3) = "3. Optimum"

            '
            ' die Namen der Worksheets Ressourcen und Portfolio verfügbar machen
            ' die Zahlen müssen korrespondieren mit der globalen Enumeration ptTables 
            arrWsNames(1) = "repCharts" ' Tabellenblatt zur Aufnahme der Charts für Reports 
            arrWsNames(2) = "Vorlage" ' depr
            ' arrWsNames(3) = 
            arrWsNames(ptTables.MPT) = "MPT"                          ' Multiprojekt-Tafel 
            arrWsNames(4) = "Einstellungen"                ' in Customization File 
            ' arrWsNames(5) = 
            arrWsNames(ptTables.meRC) = "meRC"                          ' Edit Ressourcen
            arrWsNames(6) = "meTE"                          ' Edit Termine
            arrWsNames(7) = "Darstellungsklassen"           ' wird in awinsettypen hinter MPT kopiert; nimmt für die Laufzeit die Darstellungsklassen auf 
            arrWsNames(8) = "Phasen-Mappings"               ' in Customization
            arrWsNames(9) = "meAT"                          ' Edit Attribute 
            arrWsNames(10) = "Meilenstein-Mappings"         ' in Customization
            ' arrWsNames(11) = 
            arrWsNames(ptTables.meCharts) = "meCharts"                     ' Massen-Edit Charts 
            arrWsNames(ptTables.mptPfCharts) = "mptPfCharts"                     ' vorbereitet: Portfolio Charts 
            arrWsNames(ptTables.mptPrCharts) = "mptPrCharts"                     ' vorbereitet: Projekt Charts 
            arrWsNames(14) = "Objekte" ' depr
            arrWsNames(15) = "missing Definitions"          ' in Customization File 


            awinSettings.applyFilter = False

            showRangeLeft = 0
            showRangeRight = 0

            'selectedRoleNeeds = 0
            'selectedCostNeeds = 0


            If special = "ProjectBoard" Then


                '' Versuch, awinsetTypen allgemeingültiger zu machen

                '  bestimmen der maximalen Breite und Höhe 
                formerSU = appInstance.ScreenUpdating
                appInstance.ScreenUpdating = False

                ' 9.11.2016: wird nun ganz am Anfang von awinsetTypen definiert
                '
                '' ''  um dahinter temporär die Darstellungsklassen kopieren zu können  
                ' ''Dim projectBoardSheet As Excel.Worksheet = CType(appInstance.ActiveSheet, _
                ' ''                                        Global.Microsoft.Office.Interop.Excel.Worksheet)
                projectBoardSheet = CType(appInstance.ActiveSheet,
                                                    Global.Microsoft.Office.Interop.Excel.Worksheet)

                Call setWindowParameters()

                Call logfileOpen(clear:=True)

                Call logfileSchreiben("Windows-User: ", myWindowsName, anzFehler)


                '' '--------------------------------------------------------------------------------
                '   Testen, ob der User die passende Lizenz besitzt
                '' '--------------------------------------------------------------------------------

                ' -----------------------------------------------------
                ' Speziell für Pilot-Kunden
                ' -----------------------------------------------------
                ' ab jetzt braucht man keine Lizenzen mehr ... 
                Dim pilot As Date = "15.11.2118"

                If special = "BHTC" Then

                    Dim user As String = myWindowsName
                    Dim komponente As String = LizenzKomponenten(PTSWKomp.Premium)     ' Lizenz für Projectboard notwendig

                    ' Lesen des Lizenzen-Files

                    Dim lizenzen As clsLicences = XMLImportLicences(licFileName)

                    ' Prüfen der Lizenzen
                    If Not lizenzen.validLicence(user, komponente) Then

                        Call logfileSchreiben("Aktueller User " & myWindowsName & " hat keine passende Lizenz", myWindowsName, anzFehler)

                        ''Call MsgBox("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                        ''            & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")
                        Throw New ArgumentException("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                                    & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")

                    End If

                    ' Lizenz ist ok

                Else
                    ' Für Pilotkunden soll keine Lizenz erforderlich sein

                    ' also:
                    ' Lizenz ist ok
                End If





            End If ' if special = "ProjectBoard"

            If special = "BHTC" Or special = "ReportGen" Then

                appInstance = New Excel.Application

                ' hier muss jetzt das Customization File aufgemacht werden ...
                Try
                    xlsCustomization = appInstance.Workbooks.Open(Filename:=awinPath & customizationFile, [ReadOnly]:=True, Editable:=False)
                    myCustomizationFile = appInstance.ActiveWorkbook.Name

                    Call logfileOpen()

                    Call logfileSchreiben("Windows-User: ", myWindowsName, anzFehler)

                    If awinSettings.visboDebug Then
                        Call MsgBox("Windows-User: " & myWindowsName)
                    End If


                Catch ex As Exception
                    Throw New ArgumentException("Customization File nicht gefunden - Abbruch")
                End Try

            ElseIf special = "ProjectBoard" Then

                ' hier muss jetzt das Customization File aufgemacht werden ...
                Try
                    xlsCustomization = appInstance.Workbooks.Open(awinPath & customizationFile)
                    myCustomizationFile = appInstance.ActiveWorkbook.Name
                Catch ex As Exception
                    appInstance.ScreenUpdating = formerSU

                End Try
            Else
                Throw New ArgumentException("Fehler: awinsettypen wurde mit Parameter '" & special & "' aufgerufen!")

            End If

            'Dim wsName4 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(4)), _
            '                                        Global.Microsoft.Office.Interop.Excel.Worksheet)

            Dim wsName4 As Excel.Worksheet = CType(xlsCustomization.Worksheets(arrWsNames(4)),
                                                    Global.Microsoft.Office.Interop.Excel.Worksheet
                                                    )
            If awinSettings.visboDebug Then
                Call MsgBox("wsName4 angesprochen")
            End If

            If special = "ProjectBoard" Then

                If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

                    noDB = False

                    '' ur: 23.01.2015: Abfragen der Login-Informationen
                    'loginErfolgreich = loginProzedur()

                    loginErfolgreich = logInToMongoDB(True)

                    ' ur:02012019: eigentlich wird das mit setUserRole erledigt!!!
                    '' ' hier muss jetzt ggf das Formular zur Bestimmung der CustomUser Role aufgeschaltet werden
                    ''Dim allMyCustomUserRoles As New clsCustomUserRoles
                    ''allMyCustomUserRoles = CType(databaseAcc, DBAccLayer.Request).retrieveCustomUserRolesOf(dbUsername, err)

                    ''If allMyCustomUserRoles.count > 1 Then
                    ''    Call MsgBox("hier muss eine Auswahl der Rollen getroffen werden")
                    ''Else
                    ''    myCustomUserRole = allMyCustomUserRoles.elementAt(0)
                    ''End If


                    If Not loginErfolgreich Then
                        ' Customization-File wird geschlossen
                        xlsCustomization.Close(SaveChanges:=False)
                        Call logfileSchreiben("LOGIN cancelled ...", "", -1)
                        Call logfileSchliessen()
                        If awinSettings.englishLanguage Then
                            Throw New ArgumentException("LOGIN cancelled ...")
                        Else
                            Throw New ArgumentException("LOGIN abgebrochen ...")
                        End If

                    End If

                End If

            End If 'if special="ProjectBoard"


            ''Dim wsName7810 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(7)), _
            ''                                        Global.Microsoft.Office.Interop.Excel.Worksheet)

            Dim wsName7810 As Excel.Worksheet = CType(xlsCustomization.Worksheets(arrWsNames(7)),
                                                    Global.Microsoft.Office.Interop.Excel.Worksheet
                                                    )

            If awinSettings.visboDebug Then
                Call MsgBox("wsName7810 angesprochen")
            End If

            Try
                ' Aufbauen der Darstellungsklassen  
                Call aufbauenAppearanceDefinitions(wsName7810)

                ' Auslesen der BusinessUnit Definitionen
                Call readBusinessUnitDefinitions(wsName4)

                ' Auslesen der Phasen Definitionen 
                Call readPhaseDefinitions(wsName4)

                ' Auslesen der Meilenstein Definitionen 
                Call readMilestoneDefinitions(wsName4)

                If awinSettings.visboDebug Then
                    Call MsgBox("readMilestoneDefinitions")
                End If

                ' auslesen der anderen Informationen 
                Call readOtherDefinitions(wsName4)

                If awinSettings.visboDebug Then
                    Call MsgBox("readOtherDefinitions")
                End If


                ' Kosten und Rollen sollen nur bei Initialisierung des system vom CustomizationFile gelsen werden,
                ' sonst von der DB

                ' jetzt die CurrentOrga definieren
                Dim currentOrga As New clsOrganisation

                If Not awinSettings.readCostRolesFromDB Then

                    ' tlk 15.2.19 Orga soll nur noch aus Import Orga geholt werden .. 
                    'Dim outputCollection As New Collection

                    '' Auslesen der Rollen Definitionen 
                    'Call readRoleDefinitions(wsName4, RoleDefinitions, outputCollection)

                    'If awinSettings.visboDebug Then
                    '    Call MsgBox("readRoleDefinitions")
                    'End If

                    '' Auslesen der Kosten Definitionen 
                    'Call readCostDefinitions(wsName4, CostDefinitions, outputCollection)


                    '' und jetzt werden noch die Gruppen-Definitionen ausgelesen 
                    'Call readRoleDefinitions(wsName4, RoleDefinitions, outputCollection, readingGroups:=True)

                    'If RoleDefinitions.Count > 0 Then
                    '    ' jetzt sind die Rollen alle aufgebaut und auch die Teams definiert 
                    '    ' jetzt kommt der Validation-Check 

                    '    Dim TeamsAreNotOK As Boolean = checkTeamDefinitions(RoleDefinitions, outputCollection)
                    '    Dim existingOverloads As Boolean = checkTeamMemberOverloads(RoleDefinitions, outputCollection)

                    '    If outputCollection.Count > 0 Then
                    '        Call showOutPut(outputCollection, "Organisations-Definition", "")
                    '    End If

                    'End If

                    '' jetzt sind die Rollen alle aus CustomizationFile aufgebaut und auch die Teams definiert 
                    'RoleDefinitions.buildTopNodes()
                    'With currentOrga
                    '    .validFrom = StartofCalendar
                    '    .allRoles = RoleDefinitions
                    '    .allCosts = CostDefinitions
                    'End With

                Else

                    ' 
                    ' initiales Auslesen der Rollen und Kosten aus der Datenbank ! 
                    ' das Organisations-Setting auslesen  mit heutigem Datum ...

                    currentOrga = CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", Date.Now, False, err)

                    If Not IsNothing(currentOrga) Then
                        CostDefinitions = currentOrga.allCosts
                        RoleDefinitions = currentOrga.allRoles
                    Else
                        If awinSettings.englishLanguage Then
                            Call MsgBox("You don't have any organization in your system!")
                        Else
                            Call MsgBox("Es existiert keine Organisation im System!")
                        End If
                    End If

                    'RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(Date.Now, err)
                    'CostDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCostsFromDB(Date.Now, err)

                End If

                ' tk 17.2.19 - da mehrere Organisationen aktuell noch nicht ausgewertet werden, wird das before und nextOrga erst noch rausgenommen ... 
                ' 
                If Not IsNothing(currentOrga) And awinSettings.readCostRolesFromDB Then

                    If currentOrga.count > 0 Then
                        validOrganisations.addOrga(currentOrga)
                    End If


                    ' Auslesen der Orga, die vor der currentOrga gültig war
                    ' also mit validFrom aus currentOrga lesen - 1 Tag

                    'Dim validBefore As Date = currentOrga.validFrom.AddDays(-1)
                    ' tk 17.2.19 - da mehrere Organisationen aktuell noch nicht ausgewertet werden, wird das before und nextOrga erst noch rausgenommen ... 
                    'Dim beforeOrga As clsOrganisation = CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", validBefore, False, err)

                    'If Not IsNothing(beforeOrga) Then

                    '    If beforeOrga.count > 0 Then
                    '        validOrganisations.addOrga(beforeOrga)
                    '    End If

                    'End If


                    ' Auslesen der Orga, die nach der currentOrga gültig sein  wird
                    ' also mit validFrom aus currentOrga lesen +  1 Tag

                    'Dim validNext As Date = currentOrga.validFrom.AddDays(1)

                    ' tk 15.2.19 Fehler - deshalb auskommentiert ... 
                    'Dim nextOrga As clsOrganisation =
                    'CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", validNext, True, err)


                    'If Not IsNothing(nextOrga) Then

                    '    If nextOrga.count > 0 Then
                    '        validOrganisations.addOrga(nextOrga)
                    '    End If

                    'End If

                    If awinSettings.visboDebug Then
                        Call MsgBox("Ende Lesen der Organisationen vorher-aktuell-nachher")
                    End If

                End If

                ' Lesen der Custom Field Definitions

                If Not awinSettings.readCostRolesFromDB Then

                    ' Auslesen der Custom Field Definitions aus Customization-File
                    Try
                        Call readCustomFieldDefinitions(wsName4)
                    Catch ex As Exception

                    End Try

                Else

                    ' Auslesen der Custom Field Definitions aus den VCSettings über ReST-Server
                    Try
                        customFieldDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCustomFieldsFromDB("", Date.Now, err)

                        If IsNothing(customFieldDefinitions) Then
                            ' nochmal versuchen, denn beim Lesen werden sie dann auch in die Datenbank geschrieben ... 
                            Try
                                Call readCustomFieldDefinitions(wsName4)
                            Catch ex As Exception

                            End Try
                        ElseIf customFieldDefinitions.count = 0 Then
                            Try
                                Call readCustomFieldDefinitions(wsName4)
                            Catch ex As Exception

                            End Try
                        End If
                    Catch ex As Exception

                    End Try

                End If

                ' jetzt kommt die Prüfung , ob die awinsettings.allianzdelroles korrekt sind ... 
                If awinSettings.allianzI2DelRoles <> "" And awinSettings.readCostRolesFromDB Then
                    Dim idArray() As Integer = RoleDefinitions.getIDArray(awinSettings.allianzI2DelRoles)
                    Dim tmpstr() As String = awinSettings.allianzI2DelRoles.Split(New Char() {CChar(";")})
                    If idArray.Length <> tmpstr.Length Then
                        Dim errMsg As String = "Fehler bei Angabe Ist-Daten Orga-Einheiten : " & vbLf & awinSettings.allianzI2DelRoles
                        Call MsgBox(errMsg)
                        Throw New ArgumentException(errMsg)
                    End If

                End If


                '' auslesen der anderen Informationen 
                'Call readOtherDefinitions(wsName4)

                'If awinSettings.visboDebug Then
                '    Call MsgBox("readOtherDefinitions")
                'End If


                If special = "ProjectBoard" Then

                    Try
                        ' die Info, welche Sprache gelten soll, ist in ReadOtherDefinitions ...

                        repMessages = XMLImportReportMsg(repMsgFileName, repCult.Name)
                        Call setLanguageMessages()

                    Catch ex As Exception

                    End Try

                    ' sollen die missingDefinitions gelesen / geschrieben werden 

                    If awinSettings.readWriteMissingDefinitions Then
                        Try
                            Dim wsName15 As Excel.Worksheet
                            Try

                                wsName15 = CType(appInstance.Worksheets(arrWsNames(15)),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

                                ' Auslesen der MissingPhase Definitionen 
                                Call readPhaseDefinitions(wsName15, True)

                                ' Auslesen der Meilenstein Definitionen 
                                Call readMilestoneDefinitions(wsName15, True)
                            Catch ex1 As Exception

                                ' wenn das Sheet nicht existiert, muss es angelegt werden 
                                needToBeSaved = True
                                wsName15 = appInstance.Worksheets.Add(Count:=appInstance.Worksheets.Count + 1)
                                wsName15.Name = arrWsNames(15)
                                With wsName15

                                    Dim tmpRange As Excel.Range = .Range(.Cells(1, 2), .Cells(2, 2))
                                    tmpRange.Offset(0, -1).Value = "unbekannte Phasen-/Vorgangs-Namen"
                                    .Names.Add(Name:="Missing_Phasen_Definition", RefersToR1C1:=tmpRange)

                                    tmpRange = .Range(.Cells(4, 2), .Cells(5, 2))
                                    tmpRange.Offset(0, -1).Value = "unbekannte Meilenstein-Namen"
                                    .Names.Add(Name:="Missing_Meilenstein_Definition", RefersToR1C1:=tmpRange)
                                End With

                            End Try



                        Catch ex As Exception

                        End Try
                    End If

                End If ' if special="ProjectBoard"



                ' hier muss jetzt das Worksheet Phasen-Mappings aufgemacht werden, das ist in arrwsnames(8) abgelegt 
                ''wsName7810 = CType(appInstance.Worksheets(arrWsNames(8)), _
                ''                                        Global.Microsoft.Office.Interop.Excel.Worksheet)

                wsName7810 = CType(xlsCustomization.Worksheets(arrWsNames(8)),
                                                        Global.Microsoft.Office.Interop.Excel.Worksheet
                                                        )

                Call readNameMappings(wsName7810, phaseMappings)
                If awinSettings.visboDebug Then
                    Call MsgBox("readNameMappings Phases")
                End If



                ' hier muss jetzt das Worksheet Milestone-Mappings aufgemacht werden, das ist in arrwsnames(10) abgelegt 
                'wsName7810 = CType(appInstance.Worksheets(arrWsNames(10)), _
                '                                        Global.Microsoft.Office.Interop.Excel.Worksheet)

                wsName7810 = CType(xlsCustomization.Worksheets(arrWsNames(10)),
                                                       Global.Microsoft.Office.Interop.Excel.Worksheet
                                                       )

                Call readNameMappings(wsName7810, milestoneMappings)

                If awinSettings.visboDebug Then
                    Call MsgBox("readNameMappings Milestones")
                End If

                ' hier werden nur für VISBO 1-Click PPT die vorlagen gelesen
                If special = "BHTC" Then
                    If awinSettings.visboDebug Then
                        Call MsgBox("readVorlagen: BHTC")
                    End If
                    Call readVorlagen(False)
                End If

                If special = "ProjectBoard" Then

                    ' jetzt muss die Seite mit den Appearance-Shapes kopiert werden 
                    appInstance.EnableEvents = False
                    CType(appInstance.Workbooks(myCustomizationFile).Worksheets(arrWsNames(7)),
                    Global.Microsoft.Office.Interop.Excel.Worksheet).Copy(After:=projectBoardSheet)

                    ' hier wird die Datei Projekt Tafel Customizations als aktives workbook wieder geschlossen ....
                    appInstance.Workbooks(myCustomizationFile).Close(SaveChanges:=needToBeSaved) ' ur: 6.5.2014 savechanges hinzugefügt; tk 1.3.16 needtobesaved hinzugefügt
                    appInstance.EnableEvents = True


                    ' jetzt muss die apperanceDefinitions wieder neu aufgebaut werden 
                    appearanceDefinitions.Clear()
                    wsName7810 = CType(appInstance.Workbooks(myProjektTafel).Worksheets(arrWsNames(7)),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
                    Call aufbauenAppearanceDefinitions(wsName7810)

                    ' tk 12.2.19 im awinsettypen sollen die Kapas überhaupt nicht mehr gelesen werden ... 
                    ' das Ganze soll nur noch über Menupunkt Import-Kapazitäten passieren ...
                    'Dim meldungen As New Collection
                    'If Not awinSettings.readCostRolesFromDB Then

                    '    ' jetzt werden die ggf vorhandenen detaillierten Ressourcen Kapazitäten ausgelesen 
                    '    Call readRessourcenDetails(meldungen)

                    '    ' jetzt werden die ggf vorhandenen  Urlaubstage berücksichtigt 
                    '    Call readRessourcenDetails2(meldungen)

                    '    If meldungen.Count > 0 Then
                    '        Call showOutPut(meldungen, "Errors Reading Capacities", "")
                    '        Call logfileSchreiben(meldungen)
                    '    End If

                    '    '    RoleDefinitions.buildTopNodes()

                    '    'Else

                    '    '    '' Auslesen der Rollen und Kosten ausschließlich  aus der Datenbank ! 

                    '    '    'RoleDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveRolesFromDB(Date.Now)
                    '    '    'CostDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCostsFromDB(Date.Now)

                    '    '    RoleDefinitions.buildTopNodes()

                    '    '    If awinSettings.visboDebug Then
                    '    '        Call MsgBox("Anzahl gelesene Rolen Definitionen: " & RoleDefinitions.Count.ToString)
                    '    '        Call MsgBox("Anzahl gelesene Kosten Definitionen: " & CostDefinitions.Count.ToString)
                    '    '    End If

                    'End If

                    '
                    ' ur: 07.01.2019: RoleDefinitions.buildTopNodes() wurde ersetzt durch Aufruf in .addOrga 

                    If awinSettings.visboDebug Then
                        Call MsgBox("Anzahl gelesene Rolen Definitionen: " & RoleDefinitions.Count.ToString)
                        Call MsgBox("Anzahl gelesene Kosten Definitionen: " & CostDefinitions.Count.ToString)
                    End If


                    ' jetzt werden die Modul-Vorlagen ausgelesen 
                    Call readVorlagen(True)

                    ' jetzt werden die Projekt-Vorlagen ausgelesen 
                    Call readVorlagen(False)

                    Dim a As Integer = Projektvorlagen.Count
                    Dim b As Integer = ModulVorlagen.Count

                    ' jetzt wird die Projekt-Tafel präpariert - Spaltenbreite und -Höhe
                    ' Beschriftung des Kalenders
                    appInstance.EnableEvents = False
                    Call prepareProjektTafel()

                    If awinSettings.visboDebug Then
                        Call MsgBox("prepareProjektTafel , ok")
                    End If

                    projectBoardSheet.Activate()
                    appInstance.EnableEvents = True

                    If Not noDB And awinSettings.readCostRolesFromDB Then

                        ' ur: 31.08.2017: Initialisierung
                        beforeFilterConstellation = Nothing

                        ' jetzt werden aus der Datenbank die Konstellationen und Dependencies gelesen 
                        Call readInitConstellations()

                        currentSessionConstellation.constellationName = calcLastSessionScenarioName()

                        If awinSettings.visboDebug Then
                            Call MsgBox("readInitConstellations , ok")
                        End If

                    End If

                    Dim meldungen As Collection = New Collection

                    ' jetzt werden die Rollen besetzt 
                    If awinSettings.readCostRolesFromDB Then
                        Call setUserRoles(meldungen)

                        If meldungen.Count > 0 Then
                            Call showOutPut(meldungen, "Error: setUserRoles", "")
                            Call logfileSchreiben(meldungen)
                        End If
                    Else
                        myCustomUserRole = New clsCustomUserRole

                        With myCustomUserRole
                            .customUserRole = ptCustomUserRoles.OrgaAdmin
                            .specifics = ""
                            .userName = dbUsername
                        End With
                        ' jetzt gibt es eine currentUserRole: myCustomUserRole
                        Call myCustomUserRole.setNonAllowances()
                    End If


                    ' Logfile wird geschlossen
                    Call logfileSchliessen()

                End If ' if special ="ProjectBoard"

            Catch ex As Exception
                If special = "ProjectBoard" Then
                    appInstance.ScreenUpdating = formerSU
                End If
                appInstance.EnableEvents = True
                Throw New ArgumentException(ex.Message)
            End Try

            ' jetzt werden die windowNames noch gesetzt 


            If awinSettings.englishLanguage Then
                windowNames(PTwindows.mpt) = "VISBO Multiproject-Board"
                windowNames(PTwindows.massEdit) = "edit projects: "
                windowNames(PTwindows.meChart) = "project and portfolio Charts: "
                windowNames(PTwindows.mptpf) = "Portfolio Charts: "
                windowNames(PTwindows.mptpr) = "Project Charts"
            Else
                windowNames(PTwindows.mpt) = "VISBO Multiprojekt-Tafel"
                windowNames(PTwindows.massEdit) = "Projekte editieren: "
                windowNames(PTwindows.meChart) = "Projekt und Portfolio Charts: "
                windowNames(PTwindows.mptpf) = "Portfolio Charts: "
                windowNames(PTwindows.mptpr) = "Projekt Charts"
            End If


            projectboardViews(PTview.mpt) = Nothing
            projectboardViews(PTview.mptpr) = Nothing
            projectboardViews(PTview.mptprpf) = Nothing
            projectboardViews(PTview.meOnly) = Nothing
            projectboardViews(PTview.meChart) = Nothing

            projectboardWindows(PTwindows.mpt) = Nothing
            projectboardWindows(PTwindows.mptpr) = Nothing
            projectboardWindows(PTwindows.mptpf) = Nothing
            projectboardWindows(PTwindows.massEdit) = Nothing
            projectboardWindows(PTwindows.meChart) = Nothing


        Catch ex As Exception
            Dim msg As String = ""
            If Not ex.Message.StartsWith("LOGIN cancelled") Then
                ' wird an der aufrufenden Stelle gemacht 
                'Call MsgBox("Fehler in awinsettypen " & special & vbLf & ex.Message)
                msg = "Fehler in awinsettypen " & special & vbLf & ex.Message
            Else
                msg = ex.Message
            End If
            Throw New ArgumentException(msg)
        End Try

        ' english?
        If awinSettings.englishLanguage Then
            autoSzenarioNamen(0) = "before Optimization"
        End If


    End Sub

    ''' <summary>
    ''' im Visual Board: Auswahl der vorhandenen User Roles durch den Nutzer, wenn er mehrere hat
    ''' im SmartInfo: Auswahl über das , was im PPT vorgegeben ist ; wenn der 
    ''' danach ist die globale Variable myCustomUserRole gesetzt 
    ''' </summary>
    ''' <param name="meldungen"></param>
    Public Sub setUserRoles(ByRef meldungen As Collection, Optional ByVal encryptedUserRole As String = "")

        Dim err As New clsErrorCodeMsg

        Dim allCustomUserRoles As clsCustomUserRoles = CType(databaseAcc, DBAccLayer.Request).retrieveCustomUserRoles(err)

        If Not IsNothing(allCustomUserRoles) Then

            ' hier muss jetzt ggf das Formular zur Bestimmung der CustomUser Role aufgeschaltet werden
            Dim allMyCustomUserRoles As Collection = allCustomUserRoles.getCustomUserRoles(dbUsername)

            If encryptedUserRole.Length > 0 Then
                ' bestimme die UserRole 
                Dim chkUserRole As New clsCustomUserRole
                Call chkUserRole.decrypt(encryptedUserRole)

            Else
                If allMyCustomUserRoles.Count > 1 Then
                    Dim chooseUserRole As New frmChooseCustomUserRole

                    With chooseUserRole
                        .myUserRoles = allMyCustomUserRoles
                    End With
                    ' Formular zur Auswahl der User Rolle anzeigen 
                    Dim returnResult As DialogResult = chooseUserRole.ShowDialog()


                    If returnResult = DialogResult.OK Then
                        myCustomUserRole = allMyCustomUserRoles.Item(chooseUserRole.selectedIndex)
                    Else
                        myCustomUserRole = CType(allMyCustomUserRoles.Item(1), clsCustomUserRole)
                    End If

                ElseIf allMyCustomUserRoles.Count = 1 Then
                    myCustomUserRole = CType(allMyCustomUserRoles.Item(1), clsCustomUserRole)

                Else
                    myCustomUserRole = New clsCustomUserRole
                    With myCustomUserRole

                        If awinSettings.readCostRolesFromDB Then
                            .customUserRole = ptCustomUserRoles.OrgaAdmin
                        Else
                            .customUserRole = ptCustomUserRoles.ProjektLeitung
                        End If

                        .specifics = ""
                        .userName = dbUsername
                    End With
                End If

                ' jetzt gibt es eine currentUserRole: myCustomUserRole
                Call myCustomUserRole.setNonAllowances()
            End If

        Else
            ' muss ins logfile
            meldungen.Add(err.errorMsg)
            Call MsgBox(err.errorMsg)
        End If

    End Sub

    ''' <summary>
    ''' schreibt evtl neu hinzugekommene Phasen und Meilensteine in 
    ''' das Customization File 
    ''' ausserdem werden Auswahl Validation Dropboxes gesetzt 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinWritePhaseMilestoneDefinitions(Optional ByVal writeMappings As Boolean = False)

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False



        ' hier muss jetzt das File Projekt Tafel Definitions.xlsx aufgemacht werden ...
        ' das File 
        Try
            appInstance.Workbooks.Open(awinPath & customizationFile)

        Catch ex As Exception
            Call MsgBox("Customization File nicht gefunden - Abbruch")
            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = formerSU
            Throw New ArgumentException("Customization File nicht gefunden - Abbruch")
        End Try

        appInstance.Workbooks(myCustomizationFile).Activate()

        ' schreibe die Phase- und MilestoneDefinitions
        Call WriteDefinitions(False)
        ' schreibe - in Abhängigkeit von dem Parameter . die MissingPhase- und MissingMilestone-Definitions
        If awinSettings.readWriteMissingDefinitions Then
            Call WriteDefinitions(True)
        End If


        ' prüfen , ob die Mappings-Behandlung auch gemacht werden soll ...
        If writeMappings Then

            '
            ' jetzt werden erstmal die Phase Mappings geschrieben  
            '
            Dim wsName8 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(8)),
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
            Dim area As Excel.Range
            Dim letzteZeile As Integer
            Dim aktuelleZeile As Integer

            With wsName8

                ' Synonyme schreiben 
                letzteZeile = System.Math.Max(CInt(CType(.Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row),
                                                CInt(CType(.Cells(20000, 6), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row))

                If letzteZeile >= 3 Then
                    area = CType(.Range(.Cells(3, 1), .Cells(letzteZeile, 2)), Excel.Range)
                    ' alte Synonym / regEX Area löschen
                    area.Clear()
                End If


                ' neue Area definieren
                area = CType(.Range(.Cells(3, 1), .Cells(phaseMappings.countSynonyms + phaseMappings.countRegEx + 4, 2)), Excel.Range)

                aktuelleZeile = 1
                For ix As Integer = 1 To phaseMappings.countSynonyms
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = phaseMappings.getSynonymMapping(ix - 1).Key
                    CType(area.Cells(aktuelleZeile, 2), Excel.Range).Value = phaseMappings.getSynonymMapping(ix - 1).Value
                    aktuelleZeile = aktuelleZeile + 1
                Next

                ' regular expressions schreiben 
                For ix As Integer = 1 To phaseMappings.countRegEx
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = phaseMappings.getRegExMapping(ix - 1).Key
                    CType(area.Cells(aktuelleZeile, 2), Excel.Range).Value = phaseMappings.getRegExMapping(ix - 1).Value
                    aktuelleZeile = aktuelleZeile + 1
                Next


                ' ignoreNames schreiben 
                letzteZeile = CType(.Cells(20000, 6), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row

                ' alte area löschen
                If letzteZeile >= 3 Then
                    area = CType(.Range(.Cells(3, 6), .Cells(letzteZeile, 6)), Excel.Range)
                    area.Clear()
                End If

                area = CType(.Range(.Cells(3, 6), .Cells(phaseMappings.countIgnore + 4, 6)), Excel.Range)
                aktuelleZeile = 1

                For ix As Integer = 1 To phaseMappings.countIgnore
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = phaseMappings.getIgnoreElement(ix - 1)
                    aktuelleZeile = aktuelleZeile + 1
                Next
            End With

            '
            ' jetzt werden erstmal die Phase Mappings geschrieben  
            '
            Dim wsName10 As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(10)),
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

            With wsName10

                ' Synonyme schreiben 
                letzteZeile = System.Math.Max(CInt(CType(.Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row),
                                                CInt(CType(.Cells(20000, 6), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row))


                If letzteZeile >= 3 Then
                    area = CType(.Range(.Cells(3, 1), .Cells(letzteZeile, 2)), Excel.Range)
                    ' alte Area löschen
                    area.Clear()
                End If


                ' neue Area definieren
                area = CType(.Range(.Cells(3, 1), .Cells(milestoneMappings.countSynonyms + milestoneMappings.countRegEx + 4, 2)), Excel.Range)

                aktuelleZeile = 1
                For ix As Integer = 1 To milestoneMappings.countSynonyms
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = milestoneMappings.getSynonymMapping(ix - 1).Key
                    CType(area.Cells(aktuelleZeile, 2), Excel.Range).Value = milestoneMappings.getSynonymMapping(ix - 1).Value
                    aktuelleZeile = aktuelleZeile + 1
                Next

                ' regular expressions schreiben 
                For ix As Integer = 1 To milestoneMappings.countRegEx
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = milestoneMappings.getRegExMapping(ix - 1).Key
                    CType(area.Cells(aktuelleZeile, 2), Excel.Range).Value = milestoneMappings.getRegExMapping(ix - 1).Value
                    aktuelleZeile = aktuelleZeile + 1
                Next

                ' ignoreNames schreiben 
                If letzteZeile >= 3 Then
                    area = CType(.Range(.Cells(3, 6), .Cells(letzteZeile, 6)), Excel.Range)
                    ' alte Area löschen
                    area.Clear()
                End If

                area = CType(.Range(.Cells(3, 6), .Cells(milestoneMappings.countIgnore + 4, 6)), Excel.Range)
                aktuelleZeile = 1

                For ix As Integer = 1 To milestoneMappings.countIgnore
                    CType(area.Cells(aktuelleZeile, 1), Excel.Range).Value = milestoneMappings.getIgnoreElement(ix - 1)
                    aktuelleZeile = aktuelleZeile + 1
                Next
            End With


        End If


        appInstance.ActiveWorkbook.Close(SaveChanges:=True)
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = formerSU

    End Sub

    ''' <summary>
    ''' schreibt die Phase-/MilestoneDefinitions bzw. die missingPhase- Milestone-Definitions
    ''' missingDefinitions werden in WorkSheet MissingDefinitions geschrieben 
    ''' </summary>
    ''' <param name="writeMissingDefinitions"></param>
    ''' <remarks></remarks>
    Private Sub WriteDefinitions(Optional ByVal writeMissingDefinitions As Boolean = False)

        Dim phaseDefs As Excel.Range
        Dim milestoneDefs As Excel.Range
        'Dim foundRow As Integer
        Dim phName As String
        Dim lastrow As Excel.Range
        Dim firstrow As Excel.Range
        Dim tmpAnzahl As Integer

        Dim msName As String
        Dim shortName As String

        Dim darstellungsKlasse As String

        Dim wsName4 As Excel.Worksheet


        ' beim Starten der Projekt-Tafel wird sichergestellt, dass auch das Worksheet MissingDefinitions = arrwsnames(15) existiert ...
        ' inkl der Namen der Phase- und MilestoneDefinitions
        If writeMissingDefinitions Then
            Try
                wsName4 = CType(CType(appInstance.Workbooks.Item(myCustomizationFile), Excel.Workbook).Worksheets(arrWsNames(15)),
                                               Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                Exit Sub
            End Try

        Else
            wsName4 = CType(CType(appInstance.Workbooks.Item(myCustomizationFile), Excel.Workbook).Worksheets(arrWsNames(4)),
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)
        End If

        If writeMissingDefinitions Then
            Try
                phaseDefs = wsName4.Range("Missing_Phasen_Definition")
            Catch ex As Exception
                Exit Sub
            End Try

        Else
            phaseDefs = wsName4.Range("awin_Phasen_Definition")
        End If


        ' diese Range sollte auf alle Fälle mindestens eine Zeile haben 
        Dim anzZeilen As Integer = phaseDefs.Rows.Count
        lastrow = CType(phaseDefs.Rows(anzZeilen), Excel.Range)
        firstrow = CType(phaseDefs.Rows(1), Excel.Range)


        ' das folgende muss nur gemacht werden, wenn die PhaseDefinitions geschrieben werden ... 
        ' jetzt wird geprüft, ob die missingPhaseDefinitions in PhaseDefinitions übertragen werden 
        If awinSettings.addMissingPhaseMilestoneDef Then

            For ix As Integer = 1 To missingPhaseDefinitions.Count
                Try
                    PhaseDefinitions.Add(missingPhaseDefinitions.getPhaseDef(ix))
                Catch ex As Exception

                End Try

            Next

            missingPhaseDefinitions.Clear()

            ' jetzt die Meilensteine
            For ix As Integer = 1 To missingMilestoneDefinitions.Count
                Try
                    MilestoneDefinitions.Add(missingMilestoneDefinitions.getMilestoneDef(ix))
                Catch ex As Exception

                End Try

            Next

            missingMilestoneDefinitions.Clear()

        End If

        ' jetzt können erst die PhaseDefinitions, dann die MilestoneDefinitions geschrieben werden 

        ' hier werden die Validation-String aufgebaut 
        ' hier muss jetzt noch die Validierung rein .... damit der Anwender in einem nächsten Schritt sehr bequem die verschiedenen Darstellungsklassen zuweisen kann 
        Dim milestoneAppearanceClasses As String = ""
        Dim phaseAppearanceClasses As String = ""
        Dim msAppearanceRng As Excel.Range = Nothing
        Dim phAppearanceRng As Excel.Range = Nothing
        Try
            Dim wsAppearances As Excel.Worksheet = CType(CType(appInstance.Workbooks.Item(myCustomizationFile), Excel.Workbook).Worksheets(arrWsNames(7)),
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)

            For Each cl As Excel.Range In wsAppearances.Range("MeilensteinKlassen")
                If milestoneAppearanceClasses = "" Then
                    milestoneAppearanceClasses = cl.Value
                Else
                    milestoneAppearanceClasses = milestoneAppearanceClasses & ";" & cl.Value
                End If
            Next

            For Each cl As Excel.Range In wsAppearances.Range("PhasenKlassen")
                If phaseAppearanceClasses = "" Then
                    phaseAppearanceClasses = cl.Value
                Else
                    phaseAppearanceClasses = phaseAppearanceClasses & ";" & cl.Value
                End If
            Next
        Catch ex As Exception

        End Try

        ' hier muss erst mal geprüft werden, ob Zeilen eingefügt oder gelöscht werden müssen 
        ' anzZeilen muss immer um 2 größer sein als die Anzahl der Definitionen ; 
        ' die erste und letzte Zeile des Bereichs sind leer  

        Dim anzDefinitions As Integer = PhaseDefinitions.Count

        If writeMissingDefinitions Then
            anzDefinitions = missingPhaseDefinitions.Count
        Else
            anzDefinitions = PhaseDefinitions.Count
        End If


        If anzZeilen = anzDefinitions + 2 Then
        ElseIf anzZeilen < anzDefinitions + 2 Then
            ' Zeilen einfügen 

            tmpAnzahl = anzDefinitions + 2 - anzZeilen
            For ix As Integer = 1 To tmpAnzahl
                CType(lastrow.EntireRow, Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
            Next

            ' anzZeilen und phaseDefinitions.count müssen jetzt genau gleich sein 
            anzZeilen = phaseDefs.Rows.Count

        Else
            ' Zeilen löschen
            tmpAnzahl = anzZeilen - (anzDefinitions + 2)

            For ix As Integer = 1 To tmpAnzahl
                CType(phaseDefs.Rows(2).EntireRow, Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            Next

            ' jetzt sind mindestens zwei Zeilen übrig , und zwar genau dann wenn phaseDefinitions.count = 0 
            anzZeilen = phaseDefs.Rows.Count

        End If

        ' jetzt können die Phase-Definitions in den Range geschrieben werden 
        ' und zwar so, dass sie mit der 2. Zeile beginnen 


        For ix As Integer = 1 To anzDefinitions

            If writeMissingDefinitions Then
                With missingPhaseDefinitions.getPhaseDef(ix)
                    phName = .name
                    shortName = .shortName
                    darstellungsKlasse = .darstellungsKlasse
                End With
            Else
                With PhaseDefinitions.getPhaseDef(ix)
                    phName = .name
                    shortName = .shortName
                    darstellungsKlasse = .darstellungsKlasse
                End With
            End If


            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 0).Value = phName.ToString
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 5).Value = shortName
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6).Value = darstellungsKlasse

            Try
                If phaseAppearanceClasses.Length > 0 Then
                    With CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6)
                        .Validation.Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                                                                               Formula1:=phaseAppearanceClasses)
                    End With
                End If

            Catch ex As Exception

            End Try



        Next ix

        '
        ' jetzt werden die Meilensteine geschrieben 
        '

        ' erste , letzte Zeile des Meilenstein Ranges setzen 
        ' diese Range sollte auf alle Fälle mindestens eine Zeile haben 

        If writeMissingDefinitions Then
            Try
                milestoneDefs = wsName4.Range("Missing_Meilenstein_Definition")
            Catch ex As Exception
                Exit Sub
            End Try

        Else
            milestoneDefs = wsName4.Range("awin_Meilenstein_Definition")
        End If

        anzZeilen = milestoneDefs.Rows.Count
        lastrow = CType(milestoneDefs.Rows(anzZeilen), Excel.Range)
        firstrow = CType(milestoneDefs.Rows(1), Excel.Range)


        ' hier muss erst mal geprüft werden, ob Zeilen eingefügt oder gelöscht werden müssen 
        If writeMissingDefinitions Then
            anzDefinitions = missingMilestoneDefinitions.Count
        Else
            anzDefinitions = MilestoneDefinitions.Count
        End If

        If anzZeilen = anzDefinitions + 2 Then
        ElseIf anzZeilen < anzDefinitions + 2 Then
            ' Zeilen einfügen 

            tmpAnzahl = anzDefinitions + 2 - anzZeilen

            For ix As Integer = 1 To tmpAnzahl
                CType(lastrow.EntireRow, Excel.Range).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
            Next

            ' anzZeilen und phaseDefinitions.count müssen jetzt genau gleich sein 
            anzZeilen = milestoneDefs.Rows.Count


        Else
            ' Zeilen löschen
            tmpAnzahl = anzZeilen - (anzDefinitions + 2)

            For ix As Integer = 1 To tmpAnzahl
                CType(milestoneDefs.Rows(2).EntireRow, Excel.Range).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
            Next

            ' jetzt sind mindestens zwei Zeilen übrig , und zwar genau dann wenn phaseDefinitions.count = 0 
            anzZeilen = milestoneDefs.Rows.Count

        End If

        ' jetzt können die Meilenstein-Definitions in den Range geschrieben werden 


        For ix As Integer = 1 To anzDefinitions

            If writeMissingDefinitions Then
                With missingMilestoneDefinitions.getMilestoneDef(ix)
                    msName = .name
                    shortName = .shortName
                    darstellungsKlasse = .darstellungsKlasse
                End With
            Else
                With MilestoneDefinitions.getMilestoneDef(ix)
                    msName = .name
                    shortName = .shortName
                    darstellungsKlasse = .darstellungsKlasse
                End With
            End If


            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 0).Value = msName.ToString
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 5).Value = shortName
            CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6).Value = darstellungsKlasse

            Try
                If milestoneAppearanceClasses.Length > 0 Then
                    With CType(firstrow.Cells(ix, 1), Excel.Range).Offset(1, 6)
                        .Validation.Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                                                                               Formula1:=milestoneAppearanceClasses)
                    End With
                End If

            Catch ex As Exception

            End Try

        Next ix



        '
        ' Ende der Behandlung der Phasen-/Meilenstein Behandlung 
    End Sub


    ''' <summary>
    '''speziell auf BMW Rplan Output angepasstes Inventur Import File 
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
        Dim cresult As clsMeilenstein
        Dim cbewertung As clsBewertung
        Dim ix As Integer
        Dim tmpStr(20) As String
        Dim aktuelleZeile As String
        Dim nameSopTyp As String = " "
        Dim nameBU As String
        Dim sopDate As Date
        Dim tmpStartSop As Date ' wird benutzt , um eine Hilfsphase zu machen 
        Dim startDate As Date, endDate As Date
        Dim startoffset As Long, duration As Long
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
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                Dim tstStr As String
                Try
                    tstStr = CStr(CType(activeWSListe.Cells(2, 1), Excel.Range).Value)
                    projektFarbe = CType(activeWSListe.Cells(2, 1), Excel.Range).Interior.Color
                Catch ex As Exception
                    projektFarbe = CType(activeWSListe.Cells(2, 1), Excel.Range).Interior.ColorIndex
                End Try

                ' hier werden jetzt die Columns bestimmt 


                lastRow = System.Math.Max(CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row,
                                          CType(.Cells(2000, 2), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row)

                While zeile <= lastRow

                    anfang = zeile + 1
                    ix = anfang


                    Do While CBool((CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color IsNot projektFarbe)) And (ix <= lastRow)
                        ix = ix + 1
                    Loop

                    ende = ix - 1

                    ' hier wird Name, Typ, SOP, Business Unit, vname, Start-Datum, Dauer der Phase(1) ausgelesen  
                    aktuelleZeile = CStr(CType(activeWSListe.Cells(zeile, 2), Excel.Range).Value).Trim
                    startDate = CDate(CType(activeWSListe.Cells(zeile, 3), Excel.Range).Value)
                    endDate = CDate(CType(activeWSListe.Cells(zeile, 4), Excel.Range).Value)
                    farbKennung = CInt(CType(activeWSListe.Cells(zeile, 12), Excel.Range).Value)
                    responsible = CStr(CType(activeWSListe.Cells(zeile, 9), Excel.Range).Value)


                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    If duration < 0 Then
                        startDate = endDate
                        duration = -1 * duration
                        endDate = startDate.AddDays(duration)
                    End If

                    tmpStr = aktuelleZeile.Trim.Split(New Char() {CChar("["), CChar("]")}, 5)


                    nameSopTyp = tmpStr(0).Trim
                    If Not isValidProjectName(nameSopTyp) Then
                        nameSopTyp = makeValidProjectName(nameSopTyp)
                    End If
                    pName = nameSopTyp
                    Try
                        nameBU = tmpStr(1)
                        tmpStr = nameBU.Split(New Char() {CChar(" ")}, 3)
                        nameBU = tmpStr(0)
                    Catch ex1 As Exception
                        nameBU = ""
                    End Try


                    Dim foundIX As Integer = -1

                    tmpStr = nameSopTyp.Trim.Split(New Char() {CChar(" ")}, 15)
                    Dim k As Integer = 0

                    Do While foundIX < 0 And k <= tmpStr.Length - 2
                        If tmpStr(k).Trim = "SOP" And k < tmpStr.Length - 1 Then
                            Try
                                sopDate = CDate(tmpStr(k + 1)).AddMonths(1).AddDays(-1)
                                tmpStartSop = CDate(tmpStr(k + 1))
                            Catch ex As Exception
                                Dim tmp1Str(3) As String
                                tmp1Str = tmpStr(k + 1).Split(New Char() {CChar("/")}, 8)

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
                        Throw New Exception("es gibt keine entsprechende Vorlage mit Namen  " & vorlagenName & vbLf & ex.Message)
                    End Try


                    Try

                        hproj.name = pName
                        hproj.startDate = startDate
                        hproj.earliestStartDate = hproj.startDate.AddMonths(hproj.earliestStart)
                        hproj.latestStartDate = hproj.startDate.AddMonths(hproj.latestStart)
                        ' immer als beauftragtes PRojekt importieren 

                        hproj.Status = getStatusOfBaseVariant(pName, ProjektStatus(PTProjektStati.geplant))
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
                    cphase.nameID = rootPhaseName
                    startoffset = 0
                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    cphase.changeStartandDauer(startoffset, duration)

                    cresult = New clsMeilenstein(parent:=cphase)
                    cresult.nameID = calcHryElemKey("SOP", True)
                    cresult.setDate = sopDate

                    cbewertung = New clsBewertung
                    cbewertung.colorIndex = farbKennung
                    cbewertung.description = " .. es wurde  keine Erläuterung abgegeben .. "
                    cresult.addBewertung(cbewertung)

                    Try
                        cphase.addMilestone(cresult)
                    Catch ex As Exception

                    End Try


                    hproj.AddPhase(cphase)


                    Dim phaseIX As Integer = PhaseDefinitions.Count + 1


                    Dim pStartDate As Date
                    Dim pEndDate As Date
                    Dim ok As Boolean = True
                    Dim lastPhaseName As String = cphase.nameID

                    Dim i As Integer
                    For i = anfang To ende

                        Try
                            itemName = CStr(CType(.Cells(i, 2), Excel.Range).Value).Trim
                        Catch ex As Exception
                            itemName = ""
                            ok = False
                        End Try

                        If ok Then

                            pStartDate = CDate(CType(.Cells(i, 3), Excel.Range).Value)
                            pEndDate = CDate(CType(.Cells(i, 4), Excel.Range).Value)
                            startoffset = DateDiff(DateInterval.Day, hproj.startDate, pStartDate)
                            duration = DateDiff(DateInterval.Day, pStartDate, pEndDate) + 1

                            If duration > 1 Then
                                ' es handelt sich um eine Phase 
                                phaseName = itemName
                                cphase = New clsPhase(parent:=hproj)
                                cphase.nameID = hproj.hierarchy.findUniqueElemKey(phaseName, False)

                                If PhaseDefinitions.Contains(phaseName) Then
                                    ' nichts tun 
                                Else
                                    ' in die Phase-Definitions aufnehmen 

                                    Dim hphase As clsPhasenDefinition
                                    hphase = New clsPhasenDefinition

                                    'hphase.farbe = CLng(CType(.Cells(i, 1), Excel.Range).Interior.Color)
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
                                lastPhaseName = cphase.nameID

                            ElseIf duration = 1 Then

                                Try
                                    ' es handelt sich um einen Meilenstein 

                                    Dim bewertungsAmpel As Integer
                                    Dim explanation As String

                                    bewertungsAmpel = CInt(CType(.Cells(i, 12), Excel.Range).Value)
                                    explanation = CStr(CType(.Cells(i, 1), Excel.Range).Value)

                                    cphase = hproj.getPhaseByID(lastPhaseName)
                                    cresult = New clsMeilenstein(parent:=cphase)
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
                                        .nameID = hproj.hierarchy.findUniqueElemKey(itemName, True)
                                        .setDate = pEndDate
                                        If Not cbewertung Is Nothing Then
                                            .addBewertung(cbewertung)
                                        End If
                                    End With

                                    Try
                                        With cphase
                                            .addMilestone(cresult)
                                        End With
                                    Catch ex As Exception

                                    End Try

                                Catch ex As Exception

                                End Try




                            End If




                            ' handelt es sich um eine Phase oder um einen Meilenstein ? 


                        End If


                    Next


                    ' jetzt muss das Projekt eingetragen werden 
                    ImportProjekte.Add(hproj, updateCurrentConstellation:=False, checkOnConflicts:=False)
                    myCollection.Add(hproj.name)


                    zeile = ende + 1

                    Do While CBool(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color IsNot projektFarbe) And zeile <= lastRow
                        zeile = zeile + 1
                    Loop

                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei BMW Projekt-Inventur " & vbLf & ex.Message & vbLf & pName)
        End Try



    End Sub


    ''' <summary>
    ''' gibt den numerischen Wert einer einzigen Excel Zelle zurück, sofern der Wert .GE. minvalue und .LE. maxValue ist 
    ''' und sofern er überhaupt numerisch ist. 
    ''' Bei Fehler wird der defaultValue zurückgegeben 
    ''' </summary>
    ''' <param name="zelle"></param>
    ''' <param name="defaultValue"></param>
    ''' <param name="minValue"></param>
    ''' <param name="maxValue"></param>
    ''' <returns></returns>
    Private Function getNumericValueFromExcelCell(ByVal zelle As Excel.Range, ByVal defaultValue As Double,
                                                  ByVal minValue As Double, ByVal maxValue As Double) As Double

        Dim tmpResult As Double = defaultValue

        Try
            If zelle.Rows.Count > 1 Or zelle.Columns.Count > 1 Then
                ' nichts tun , tmpResult hat schon den Default Wert 
            Else
                Dim tmpValue As String = zelle.Value
                If IsNothing(tmpValue) Then
                    ' nichts tun, tmpResult hat schon den Default Wert
                ElseIf tmpValue.Trim.Length = 0 Then
                    ' nichts tun, tmpResult hat schon den Default Wert
                ElseIf Not IsNumeric(tmpValue) Then
                    ' nichts tun, tmpResult hat schon den Default Wert
                ElseIf CDbl(tmpValue) < minValue Or CDbl(tmpValue) > maxValue Then
                    ' nichts tun, tmpResult hat schon den Default Wert
                Else
                    tmpResult = CDbl(tmpValue)
                End If
            End If


        Catch ex As Exception

        End Try

        getNumericValueFromExcelCell = tmpResult
    End Function


    ''' <summary>
    ''' liest die optional vorhandenen Custom Field Definitionen aus 
    ''' </summary>
    ''' <param name="wsname"></param>
    ''' <remarks></remarks>
    Private Sub readCustomFieldDefinitions(wsname As Excel.Worksheet)

        '
        ' Custom Field Definitions Definitionen auslesen - im bereich awin_CustomField_Definitions
        '

        Try


            With wsname

                Dim customFieldRange As Excel.Range = .Range("awin_CustomField_Definitions")
                Dim anzZeilen As Integer = customFieldRange.Rows.Count
                Dim c As Excel.Range


                For i = 2 To anzZeilen - 1
                    c = CType(customFieldRange.Cells(i, 1), Excel.Range)

                    Dim uid As Integer = i - 1
                    Dim cfType As Integer = -1
                    Dim cfName As String = ""
                    Dim ok As Boolean = False
                    Try
                        cfName = CStr(CType(customFieldRange.Cells(i, 1), Excel.Range).Value)
                        cfType = CInt(CType(customFieldRange.Cells(i, 2), Excel.Range).Value)
                        ok = True
                    Catch ex As Exception

                    End Try

                    If ok And cfName <> "" And isValidCustomField(cfType) Then

                        ' jetzt die CustomField Definition hinzufügen 
                        Try
                            customFieldDefinitions.add(cfName, cfType, uid)
                        Catch ex As Exception
                            Call MsgBox(ex.Message)
                        End Try


                    End If

                Next

            End With


            Dim err As New clsErrorCodeMsg
            Dim ts As Date = CDate("1.1.1900")
            Dim customFieldsName As String = CStr(settingTypes(ptSettingTypes.customfields))
            Dim result As Boolean = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(customFieldDefinitions,
                                                                                           CStr(settingTypes(ptSettingTypes.customfields)),
                                                                                           customFieldsName,
                                                                                           ts,
                                                                                           err)
            If Not result Then
                Call MsgBox("Fehler beim Speichern der Customfields: " & err.errorCode & err.errorMsg)
            End If


        Catch ex As Exception
            Throw New ArgumentException("Fehler im Customization-File: Custom Field Definitions")
        End Try



    End Sub

    ''' <summary>
    ''' gibt zurück, ob die übergebene Zahl ein gültiger CustomField Typ ist
    ''' </summary>
    ''' <param name="id"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function isValidCustomField(ByVal id As Integer) As Boolean

        If id = ptCustomFields.bool Or
            id = ptCustomFields.Str Or
            id = ptCustomFields.Dbl Then
            isValidCustomField = True
        Else
            isValidCustomField = False
        End If

    End Function


    ''' <summary>
    ''' liest die Phasen Definitionen aus 
    ''' </summary>
    ''' <param name="wsname">Name des Worksheets, aus dem die Infos ausgelesen werden</param>
    ''' <remarks></remarks>
    Private Sub readMilestoneDefinitions(ByVal wsname As Excel.Worksheet, Optional ByVal missingDefinitions As Boolean = False)

        Dim i As Integer = 0
        Dim hMilestone As clsMeilensteinDefinition
        Dim tmpStr As String


        Try

            With wsname

                Dim milestoneRange As Excel.Range

                If missingDefinitions Then
                    Try
                        milestoneRange = .Range("Missing_Meilenstein_Definition")
                    Catch ex As Exception
                        Exit Sub
                    End Try

                Else
                    milestoneRange = .Range("awin_Meilenstein_Definition")
                End If

                Dim anzZeilen As Integer = milestoneRange.Rows.Count
                Dim c As Excel.Range

                For iZeile As Integer = 2 To anzZeilen - 1

                    c = CType(milestoneRange.Cells(iZeile, 1), Excel.Range)

                    ' hier muss das Aufbauen der MilestoneDefinitions gemacht werden  
                    If Not IsNothing(c.Value) Then

                        If CStr(c.Value) <> "" Then
                            i = i + 1
                            tmpStr = CType(c.Value, String)
                            ' das neue ...
                            hMilestone = New clsMeilensteinDefinition
                            With hMilestone
                                .name = tmpStr.Trim
                                .UID = i

                                ' hat der Milestone einen Schwellwert ? 

                                If IsNothing(c.Offset(0, 1).Value) Then
                                ElseIf IsNumeric(c.Offset(0, 1).Value) Then
                                    If CInt(c.Offset(0, 1).Value) > 0 Then
                                        .schwellWert = CInt(c.Offset(0, 1).Value)
                                    End If
                                End If


                                ' hat der Milestone einen Bezug ? 
                                Dim bezug As String = ""
                                If Not IsNothing(c.Offset(0, 4).Value) Then

                                    bezug = CStr(c.Offset(0, 4).Value).Trim

                                    If PhaseDefinitions.Contains(bezug) Then
                                    Else
                                        bezug = ""
                                    End If

                                End If

                                .belongsTo = bezug

                                ' hat der Milestone eine Abkürzung ? 
                                Dim abbrev As String = ""
                                If Not IsNothing(c.Offset(0, 5).Value) Then
                                    abbrev = CStr(c.Offset(0, 5).Value).Trim
                                End If

                                .shortName = abbrev


                                ' hat der Milestone eine Darstellungsklasse ? 

                                Dim darstellungsklasse As String = ""
                                If Not IsNothing(c.Offset(0, 6).Value) Then

                                    If CStr(c.Offset(0, 6).Value).Trim.Length > 0 Then
                                        darstellungsklasse = CStr(c.Offset(0, 6).Value).Trim
                                        If appearanceDefinitions.ContainsKey(darstellungsklasse) Then
                                            .darstellungsKlasse = darstellungsklasse
                                        Else
                                            .darstellungsKlasse = ""
                                        End If
                                    End If

                                End If



                            End With

                            Try
                                If missingDefinitions Then
                                    missingMilestoneDefinitions.Add(hMilestone)
                                Else
                                    MilestoneDefinitions.Add(hMilestone)
                                End If

                            Catch ex As Exception

                            End Try


                        End If

                    End If

                Next

            End With

        Catch ex As Exception

            Throw New ArgumentException("Fehler in Customization File: Meilensteine")

        End Try


    End Sub

    ''' <summary>
    ''' liest die Rollen Definitionen ein 
    ''' wird in der globalen Variablen RoleDefinitions abgelegt 
    ''' in der Spalte 1 stehen jetzt ggf die ID der Rollen, dann berücksichtigen ...
    ''' </summary>
    ''' <param name="wsname"></param>
    ''' <remarks></remarks>
    Private Sub readRoleDefinitions(ByVal wsname As Excel.Worksheet, ByRef neueRollendefinitionen As clsRollen, ByRef meldungen As Collection,
                                    Optional ByVal readingGroups As Boolean = False)

        '
        ' Rollen Definitionen auslesen - im bereich awin_Rollen_Definition
        '
        Dim index As Integer = 0
        Dim tmpStr As String
        Dim hrole As clsRollenDefinition
        Dim roleUID As Integer = 0
        Dim roleUidsDefined As Boolean = False
        Dim przSatz As Double = 1.0
        Dim defaultTagessatz As Double = 800.0
        Dim errMsg As String = ""

        Try
            Dim hasHierarchy As Boolean = False
            Dim atleastOneWithIndent As Boolean = False
            Dim maxIndent As Integer = 0
            Dim rolesRange As Excel.Range = Nothing

            If readingGroups Then
                Try
                    errMsg = "Range <awin_Gruppen_Definition> nicht definiert ! Abbruch ..."
                    rolesRange = wsname.Range("awin_Gruppen_Definition")
                Catch ex As Exception
                    rolesRange = Nothing
                End Try

            Else
                Try
                    errMsg = "Range <awin_Rollen_Definition> nicht definiert ! Abbruch ..."
                    rolesRange = wsname.Range("awin_Rollen_Definition")
                    przSatz = 1.0
                Catch ex As Exception
                    rolesRange = Nothing
                End Try

            End If

            ' Exit, wenn es keine Definitionen gibt ... 
            If IsNothing(rolesRange) Then
                meldungen.Add(errMsg)
                Exit Sub
            Else
                errMsg = ""
                Dim anzZeilen As Integer = rolesRange.Rows.Count
                Dim c As Excel.Range

                ' jetzt wird erst mal gecheckt, ob alle Rollen entweder keine Integer Kennzahl haben: dann wird die aus der Position errechnet 
                ' oder ob sie eine haben und ob keine Mehrfachnennungen vorkommen 
                ' ausserdem wird gleich mal gecheckt ob die erste Rolle indent = 0 hat und sonstige Indent-Level vorkommen
                ' ausserdem wird hier gecheckt, ob jeder NAme auch nur genau einmal vorkommt 

                Dim anzWithID As Integer = 0
                Dim anzWithoutID As Integer = 0
                Dim IDCollection As New Collection
                Dim groupDefinitionIsOk As Boolean = True
                Dim uniqueNames As New Collection

                For i = 2 To anzZeilen - 1

                    Try
                        Dim tmpIDValue As String = CType(rolesRange.Cells(i, 1), Excel.Range).Offset(0, -1).Value
                        Dim tmpOrgaName As String = getStringFromExcelCell(CType(rolesRange.Cells(i, 1), Excel.Range))

                        c = CType(rolesRange.Cells(i, 1), Excel.Range)

                        ' checken, ob nachher die Rollen-Hierarchie aufgebaut werden soll .. 
                        ' 1.Rolle muss bei Indent 0 anfangen, alle anderen dann entsprechend ihrer Hierarchie eingerückt sein 
                        If i = 2 Then
                            If CType(rolesRange.Cells(i, 1), Excel.Range).IndentLevel = 0 Then
                                hasHierarchy = True
                            End If
                        Else
                            Dim tmpIndent As Integer = CType(rolesRange.Cells(i, 1), Excel.Range).IndentLevel
                            If tmpIndent > 0 Then
                                atleastOneWithIndent = True
                                maxIndent = System.Math.Max(maxIndent, tmpIndent)
                            End If
                        End If

                        Dim isWithoutID As Boolean = True

                        If CStr(c.Value) <> "" Then
                            If Not IsNothing(tmpIDValue) Then
                                If tmpIDValue.Trim <> "" Then
                                    If IsNumeric(tmpIDValue.Trim) Then
                                        If CInt(tmpIDValue.Trim) > 0 Then
                                            If Not IDCollection.Contains(tmpIDValue.Trim) Then
                                                IDCollection.Add(tmpIDValue.Trim, tmpIDValue.Trim)
                                                isWithoutID = False
                                            Else
                                                errMsg = "roles with identical IDs are not allowed: " & tmpIDValue.Trim
                                                meldungen.Add(errMsg)
                                                CType(rolesRange.Cells(i, 1), Excel.Range).Offset(0, -1).Interior.Color = XlRgbColor.rgbOrangeRed
                                            End If
                                        Else
                                            anzWithoutID = anzWithoutID + 1
                                        End If
                                    Else
                                        anzWithoutID = anzWithoutID + 1
                                    End If
                                Else
                                    anzWithoutID = anzWithoutID + 1
                                End If
                            Else
                                anzWithoutID = anzWithoutID + 1
                            End If
                        End If

                        ' jetzt auf identisch vorkommende Namen checken ... aber nur im Modus not readingGroups
                        If Not readingGroups Then
                            If tmpOrgaName = "" Then
                                errMsg = "roles with empty string are not allowed "
                                meldungen.Add(errMsg)
                                CType(rolesRange.Cells(i, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed
                            Else
                                If Not uniqueNames.Contains(tmpOrgaName) Then
                                    uniqueNames.Add(tmpOrgaName, tmpOrgaName)
                                Else
                                    errMsg = "roles with same name are not allowed: " & tmpOrgaName
                                    meldungen.Add(errMsg)
                                    CType(rolesRange.Cells(i, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed
                                End If
                            End If
                        Else
                            ' readingGroups
                            If tmpIDValue <> "" Then
                                If Not uniqueNames.Contains(tmpOrgaName) Then
                                    uniqueNames.Add(tmpOrgaName, tmpOrgaName)
                                    If neueRollendefinitionen.containsName(tmpOrgaName) Then
                                        errMsg = "groups with same Name as certain orga-element are not allowed: " & tmpOrgaName
                                        meldungen.Add(errMsg)
                                        CType(rolesRange.Cells(i, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed
                                    End If
                                Else
                                    errMsg = "roles with same name are not allowed: " & tmpOrgaName
                                    meldungen.Add(errMsg)
                                    CType(rolesRange.Cells(i, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed
                                End If
                            Else
                                If neueRollendefinitionen.containsNameID(tmpIDValue) Then
                                    errMsg = "group must not have same ID than other Orga-Unit: " & tmpOrgaName
                                    meldungen.Add(errMsg)
                                    CType(rolesRange.Cells(i, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed
                                End If
                            End If
                        End If


                        ' jetzt checken 
                        If readingGroups And isWithoutID Then
                            ' c.value muss in RoleDefinitions vorkommen, sonst Fehler ...
                            Dim roleName As String = CStr(c.Value.trim)

                            If Not neueRollendefinitionen.containsName(roleName) Then
                                errMsg = "Team-Role " & roleName & " does not exist ..."
                                meldungen.Add(errMsg)
                                CType(rolesRange.Cells(i, 1), Excel.Range).Interior.Color = XlRgbColor.rgbOrangeRed

                                groupDefinitionIsOk = False
                            End If

                        End If

                    Catch ex As Exception
                        anzWithoutID = anzWithoutID + 1
                    End Try

                Next

                anzWithID = IDCollection.Count
                If anzWithID > 0 And anzWithoutID > 0 And Not readingGroups Then
                    errMsg = "some roles do contain IDs, others not ..."
                    meldungen.Add(errMsg)
                    Exit Sub
                ElseIf Not groupDefinitionIsOk Then
                    errMsg = "Group Definitions not correct ..."
                    meldungen.Add(errMsg)
                    Exit Sub
                Else
                    ' jetzt ist sichergestellt, dass alle Rollen eine ID haben oder keine ; dann wird sie generiert .. 
                    ' oder aber man ist im Reading Group Modus, wo ja nur die Gruppen eine ID benötigen
                    For i = 2 To anzZeilen - 1

                        c = CType(rolesRange.Cells(i, 1), Excel.Range)

                        If CStr(c.Value).Trim <> "" Then

                            index = index + 1
                            If anzWithID > 0 Then
                                roleUID = CInt(CType(rolesRange.Cells(i, 1), Excel.Range).Offset(0, -1).Value)
                            Else
                                roleUID = index
                            End If

                            tmpStr = CType(c.Value, String)
                            If isValidRoleName(tmpStr, errMsg) Then
                                If readingGroups Then
                                    przSatz = getNumericValueFromExcelCell(CType(c.Offset(0, 4), Excel.Range), 1.0, 0.0, 1.0)
                                Else
                                    przSatz = 1.0
                                End If

                                ' jetzt kommt die Rollen Definition 
                                hrole = New clsRollenDefinition
                                Dim cp As Integer
                                With hrole
                                    .name = tmpStr.Trim

                                    .defaultKapa = CDbl(c.Offset(0, 1).Value)

                                    .tagessatzIntern = CDbl(c.Offset(0, 2).Value)
                                    If .tagessatzIntern <= 0 Then
                                        .tagessatzIntern = defaultTagessatz
                                    End If

                                    ' tk 5.12 Aufnahme extern
                                    Dim tmpValue As String = CStr(c.Offset(0, 3).Value)

                                    If Not IsNothing(tmpValue) Then
                                        tmpValue = tmpValue.Trim
                                        Dim positiveCriterias() As String = {"J", "j", "ja", "Ja", "Y", "y", "yes", "Yes", "1"}

                                        If positiveCriterias.Contains(tmpValue) Then
                                            .isExternRole = True
                                        End If
                                    End If


                                    ' Änderung 29.5.14: von StartofCalendar 240 Monate nach vorne kucken ... 
                                    For cp = 1 To 240

                                        .kapazitaet(cp) = .defaultKapa
                                        '.externeKapazitaet(cp) = 0.0

                                    Next
                                    .farbe = c.Interior.Color
                                    .UID = roleUID
                                End With

                                ' wenn readingGroups, dann kann die Rolle bereits enthalten sein 
                                If readingGroups And neueRollendefinitionen.containsName(hrole.name) Then
                                    ' nichts tun, alles gut : 
                                Else
                                    ' im anderen Fall soll die Rolle aufgenommen werden; wenn readinggroups = false und Rolle existiert schon, dann gibt es Fehler 
                                    If Not neueRollendefinitionen.containsName(hrole.name) Then
                                        neueRollendefinitionen.Add(hrole)
                                    End If

                                End If
                            Else
                                meldungen.Add(errMsg)
                            End If



                            'hrole = Nothing

                        End If

                    Next

                End If

                ' tk Änderung 25.5.18 Auslesen der Hierarchie - dann sind keine Ressourcen Manager Dateien mehr notwendig .. 
                ' jetzt checken ob eine Hierarchie aufgebaut werden soll ..
                hasHierarchy = hasHierarchy And atleastOneWithIndent

                If hasHierarchy Then
                    ' Hierarchie aufbauen

                    Dim parents(maxIndent) As String

                    Dim ix As Integer
                    parents(0) = CStr(CType(rolesRange.Cells(2, 1), Excel.Range).Value).Trim


                    Dim lastLevel As Integer = 0
                    Dim curLevel As Integer = 0

                    Dim curRoleName As String = ""

                    ix = 3

                    Do While ix <= anzZeilen - 1

                        Try
                            curLevel = CType(rolesRange.Cells(ix, 1), Excel.Range).IndentLevel
                            curRoleName = CStr(CType(rolesRange.Cells(ix, 1), Excel.Range).Value).Trim

                            If readingGroups Then
                                przSatz = getNumericValueFromExcelCell(CType(rolesRange.Cells(ix, 1), Excel.Range).Offset(0, 4), 1.0, 0.0, 1.0)
                            Else
                                przSatz = 1.0
                            End If

                            Do While curLevel = lastLevel And ix <= anzZeilen - 1

                                If curLevel > 0 Then
                                    ' als Child aufnehmen 
                                    ' hier, wenn maxIndent = curlevel, auf alle Fälle Team-Member
                                    Dim parentRole As clsRollenDefinition = neueRollendefinitionen.getRoledef(parents(curLevel - 1))
                                    Dim subRole As clsRollenDefinition = neueRollendefinitionen.getRoledef(curRoleName)
                                    parentRole.addSubRole(subRole.UID, przSatz)

                                    If curLevel = maxIndent And readingGroups Then
                                        If Not parentRole.isTeam Then
                                            parentRole.isTeam = True
                                        End If
                                        subRole.addTeam(parentRole.UID, przSatz)
                                    End If

                                    ' 29.6.18 auch hier den Parent weiterschalten 
                                    parents(curLevel) = curRoleName
                                Else
                                    ' hier den Parent weiterschalten  
                                    parents(curLevel) = curRoleName
                                End If

                                ' weiterschalten ..
                                ix = ix + 1

                                ' hat sich der Indentlevel immer noch nicht geändert ? 
                                If ix <= anzZeilen - 1 Then
                                    curLevel = CType(rolesRange.Cells(ix, 1), Excel.Range).IndentLevel
                                    curRoleName = CStr(CType(rolesRange.Cells(ix, 1), Excel.Range).Value).Trim
                                    If readingGroups Then
                                        przSatz = getNumericValueFromExcelCell(CType(rolesRange.Cells(ix, 1), Excel.Range).Offset(0, 4), 1.0, 0.0, 1.0)
                                    Else
                                        przSatz = 1.0
                                    End If

                                Else
                                    ' das Abbruch Kriterium schlägt gleich zu ... 
                                End If

                            Loop

                            If curLevel <> lastLevel And ix <= anzZeilen - 1 Then

                                parents(curLevel) = curRoleName

                                If curLevel < lastLevel Then
                                    ' in der Hierarchie zurück 
                                    For i As Integer = curLevel + 1 To maxIndent
                                        parents(i) = ""
                                    Next
                                End If

                                If curLevel > 0 Then
                                    ' als Child aufnehmen 
                                    Dim parentRole As clsRollenDefinition = neueRollendefinitionen.getRoledef(parents(curLevel - 1))
                                    Dim subRole As clsRollenDefinition = neueRollendefinitionen.getRoledef(curRoleName)
                                    parentRole.addSubRole(subRole.UID, przSatz)

                                    ' hier kann er eigentlich nie hinkommen ...
                                    If curLevel = maxIndent And readingGroups Then
                                        If Not parentRole.isTeam Then
                                            parentRole.isTeam = True
                                        End If
                                        subRole.addTeam(parentRole.UID, przSatz)
                                    End If

                                Else
                                    ' nichts tun 
                                End If

                                ' alle alten löschen 
                                lastLevel = curLevel
                                ix = ix + 1

                            End If
                        Catch ex As Exception
                            errMsg = "zeile: " & ix.ToString & " : " & ex.Message
                            meldungen.Add(errMsg)
                            CType(rolesRange.Cells(ix, 1), Excel.Range).Offset(0, -1).Interior.Color = XlRgbColor.rgbOrangeRed
                        End Try


                    Loop

                End If

            End If



        Catch ex As Exception
            errMsg = "general, unidentified error: " & ex.Message
            meldungen.Add(errMsg)
        End Try



    End Sub


    ''' <summary>
    ''' liest die Business Unit Definitionen aus der awinsetTypen
    ''' die globale Variable businessUnitDefinitions wird dabei befüllt
    ''' die erste und letzte Zeile des Range wird ignoriert 
    ''' </summary>
    ''' <param name="wsname">Name des Excel Worksheets, das die Infos im aktuellen Workbook enthält</param>
    ''' <remarks></remarks>
    Private Sub readBusinessUnitDefinitions(ByVal wsname As Excel.Worksheet)

        ' hier werden jetzt die Business Unit Informationen ausgelesen 
        businessUnitDefinitions = New SortedList(Of Integer, clsBusinessUnit)

        Try

            With wsname
                '
                ' Business Unit Definitionen auslesen - im bereich awin_BusinessUnit_Definitions
                '
                Dim index As Integer = 1
                Dim tmpBU As clsBusinessUnit

                Dim BURange As Excel.Range = CType(.Range("awin_BusinessUnit_Definitions"), Excel.Range)
                Dim anzZeilen As Integer = BURange.Rows.Count

                For i As Integer = 2 To anzZeilen - 1

                    tmpBU = New clsBusinessUnit

                    Try
                        tmpBU.name = CStr(BURange.Cells(i, 1).value).Trim
                        tmpBU.color = CLng(BURange.Cells(i, 1).Interior.color)

                        If tmpBU.name.Length > 0 Then
                            businessUnitDefinitions.Add(i - 1, tmpBU)
                        End If

                    Catch ex As Exception
                        ' nichts tun ...

                    End Try

                Next

            End With

        Catch ex As Exception
            Throw New ArgumentException("Fehler im Customization-File: BU Definition")
        End Try


    End Sub

    ''' <summary>
    ''' liest die Phasen Definitionen aus 
    ''' baut die globale Variable PhaseDefinitions auf 
    ''' </summary>
    ''' <param name="wsname">Name des Worksheets, aus dem die Infos ausgelesen werden</param>
    ''' <remarks></remarks>
    Private Sub readPhaseDefinitions(ByVal wsname As Excel.Worksheet, Optional ByVal missingDefinitions As Boolean = False)

        Dim hphase As clsPhasenDefinition
        Dim tmpStr As String = ""

        Try

            With wsname

                Dim phaseRange As Excel.Range

                If missingDefinitions Then
                    Try
                        phaseRange = .Range("Missing_Phasen_Definition")
                    Catch ex As Exception
                        Exit Sub
                    End Try

                Else
                    phaseRange = .Range("awin_Phasen_Definition")
                End If

                Dim anzZeilen As Integer = phaseRange.Rows.Count
                Dim c As Excel.Range

                For iZeile As Integer = 2 To anzZeilen - 1

                    c = CType(phaseRange.Cells(iZeile, 1), Excel.Range)

                    If Not IsNothing(c.Value) Then

                        If CStr(c.Value) <> "" Then
                            tmpStr = CType(c.Value, String)
                            ' das neue ...
                            hphase = New clsPhasenDefinition
                            With hphase
                                '.farbe = CLng(c.Interior.Color)
                                .name = tmpStr.Trim
                                .UID = iZeile - 1

                                ' hat die Phase einen Schwellwert ? 
                                Try
                                    If CInt(c.Offset(0, 1).Value) > 0 Then
                                        .schwellWert = CInt(c.Offset(0, 1).Value)
                                    End If
                                Catch ex As Exception

                                End Try

                                ' ist die Phase eine special Phase ? 
                                Try
                                    If Not IsNothing(CType(c.Offset(0, 2), Excel.Range).Value) Then
                                        If CStr(c.Offset(0, 2).Value).Trim = "LeLe" Then
                                            specialListofPhases.Add(hphase.name, hphase.name)
                                        End If
                                    End If
                                Catch ex As Exception
                                End Try



                                ' hat die Phase eine Abkürzung ? 
                                Dim abbrev As String = ""
                                If Not IsNothing(c.Offset(0, 5).Value) Then
                                    abbrev = CStr(c.Offset(0, 5).Value).Trim
                                End If

                                .shortName = abbrev


                                ' hat die Phase eine Darstellungsklasse ? 
                                Try
                                    Dim darstellungsklasse As String
                                    If Not IsNothing(c.Offset(0, 6).Value) Then

                                        If CStr(c.Offset(0, 6).Value).Trim.Length > 0 Then
                                            darstellungsklasse = CStr(c.Offset(0, 6).Value).Trim
                                            If appearanceDefinitions.ContainsKey(darstellungsklasse) Then
                                                .darstellungsKlasse = darstellungsklasse
                                            Else
                                                .darstellungsKlasse = ""
                                            End If
                                        End If

                                    End If

                                Catch ex As Exception
                                    .darstellungsKlasse = ""
                                End Try



                            End With

                            Try
                                If missingDefinitions Then

                                    missingPhaseDefinitions.Add(hphase)

                                Else

                                    PhaseDefinitions.Add(hphase)

                                End If

                            Catch ex As Exception

                            End Try


                        End If

                    End If


                Next


            End With

        Catch ex As Exception

            Throw New ArgumentException("Fehler in Customization File: Phasen")

        End Try


    End Sub


    ''' <summary>
    ''' liest die sonstigen Einstellungen wie Farben, Spaltenbreite, Spaltenhöhe etc aus
    ''' wird in entsprechenden globalen Variablen abgelegt  
    ''' </summary>
    ''' <param name="wsname">Name des Worksheets, aus dem die Infos ausgelesen werden</param>
    ''' <remarks></remarks>
    Private Sub readOtherDefinitions(ByVal wsname As Excel.Worksheet)


        With wsname
            Try
                'showRangeLeft = CInt(.Range("Linker_Rand_Ressourcen_Diagramme").Value)
                'showRangeRight = CInt(.Range("Rechter_Rand_Ressourcen_Diagramme").Value)
                showtimezone_color = .Range("Show_Time_Zone_Color").Interior.Color
                noshowtimezone_color = .Range("NoShow_Time_Zone_Color").Interior.Color
                calendarFontColor = .Range("NoShow_Time_Zone_Color").Font.Color
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

                    Try
                        awinSettings.timeSpanColor = CLng(.Range("FarbeZeitraum").Interior.Color)
                        awinSettings.showTimeSpanInPT = CBool(.Range("FarbeZeitraum").Value)
                    Catch ex2 As Exception
                        ' ansonsten wird die Voreinstellung verwendet 
                    End Try
                    Try
                        awinSettings.gridLineColor = CLng(.Range("FarbeGridLine").Interior.Color)
                    Catch ex As Exception

                    End Try

                Catch ex As Exception
                    Throw New ArgumentException("Customization File fehlerhaft - Farben fehlen ... " & vbLf & ex.Message)
                End Try

                Try
                    awinSettings.missingDefinitionColor = CLng(.Range("MissingDefinitionColor").Interior.Color)
                    ' ''If awinSettings.missingDefinitionColor = XlRgbColor.rgbWhite Then
                    ' ''    Call MsgBox("leeres missingDefinitionColor - Feld in customizationfile " & awinSettings.missingDefinitionColor.ToString)
                    ' ''End If
                Catch ex As Exception

                End Try

                Try
                    awinSettings.allianzI2DelRoles = CStr(.Range("allianzI2DelRoles").Value).Trim

                Catch ex As Exception
                    awinSettings.allianzI2DelRoles = ""
                End Try

                Try
                    awinSettings.autoSetActualDataDate = CBool(.Range("autoSetActualDataDate").Value)
                Catch ex As Exception
                    awinSettings.autoSetActualDataDate = False
                End Try

                Try
                    awinSettings.actualDataMonth = CDate(.Range("ActualDataMonth").Value)
                Catch ex As Exception
                    awinSettings.actualDataMonth = Date.MinValue
                End Try

                ' tk 23.12.18 deprecated
                'Try
                '    awinSettings.isRestrictedToOrgUnit = CStr(.Range("roleAccessIsRestrictedTo").Value)

                '    If IsNothing(awinSettings.isRestrictedToOrgUnit) Then
                '        awinSettings.isRestrictedToOrgUnit = ""
                '    End If
                'Catch ex As Exception
                '    awinSettings.isRestrictedToOrgUnit = ""
                'End Try


                ergebnisfarbe1 = .Range("Ergebnisfarbe1").Interior.Color
                ergebnisfarbe2 = .Range("Ergebnisfarbe2").Interior.Color
                weightStrategicFit = CDbl(.Range("WeightStrategicFit").Value)
                ' jetzt wird KalenderStart, Zeiteinheit und Datenbank Name ausgelesen 
                awinSettings.kalenderStart = CDate(.Range("Start_Kalender").Value)
                awinSettings.zeitEinheit = CStr(.Range("Zeiteinheit").Value)
                awinSettings.kapaEinheit = CStr(.Range("kapaEinheit").Value)
                If awinSettings.kapaEinheit <> "PT" And
                    awinSettings.kapaEinheit <> "PD" Then
                    awinSettings.kapaEinheit = "PT"
                    Call MsgBox("Kapa-Einheit: Personen-Tage")
                End If
                awinSettings.offsetEinheit = CStr(.Range("offsetEinheit").Value)
                'ur: 6.08.2015: umgestellt auf Settings in app.config ''awinSettings.databaseName = CStr(.Range("Datenbank").Value)
                awinSettings.EinzelRessExport = CInt(.Range("EinzelRessourcenExport").Value)
                awinSettings.zeilenhoehe1 = CDbl(.Range("Zeilenhoehe1").Value)
                awinSettings.zeilenhoehe2 = CDbl(.Range("Zeilenhoehe2").Value)
                awinSettings.spaltenbreite = CDbl(.Range("Spaltenbreite").Value)
                awinSettings.autoCorrectBedarfe = True
                awinSettings.propAnpassRess = False
                awinSettings.showValuesOfSelected = False
            Catch ex As Exception
                Throw New ArgumentException("fehlende Einstellung im Customization-File ... Abbruch " & vbLf & ex.Message)
            End Try

            ' gibt es die Einstellung für ProjectWithNoMPmayPass

            Try
                awinSettings.mppProjectsWithNoMPmayPass = CBool(.Range("passFilterWithNoMPs").Value)
            Catch ex As Exception
                awinSettings.mppProjectsWithNoMPmayPass = False
            End Try


            ' ist Einstellung für volles Protokoll vorhanden ? 
            Try

                awinSettings.fullProtocol = CBool(.Range("volles_Protokol").Value)
            Catch ex As Exception
                awinSettings.fullProtocol = False
            End Try

            ' Einstellung für addMissingDefinitions
            Try
                awinSettings.addMissingPhaseMilestoneDef = CBool(.Range("addMissingDefinitions").Value)
            Catch ex As Exception
                awinSettings.addMissingPhaseMilestoneDef = False
            End Try

            ' Einstellung für alwaysAcceptTemplate Names 
            Try
                awinSettings.alwaysAcceptTemplateNames = CBool(.Range("alywaysAcceptTemplateDefs").Value)
            Catch ex As Exception
                awinSettings.alwaysAcceptTemplateNames = False
            End Try

            ' Einstellungen, um Duplikate zu eliminieren ; 
            Try
                awinSettings.eliminateDuplicates = CBool(.Range("eliminate_Duplicates").Value)
            Catch ex As Exception
                awinSettings.eliminateDuplicates = True
            End Try

            ' Einstellungen, um unbekannte Namen zu importieren 
            Try
                awinSettings.importUnknownNames = CBool(.Range("importUnknownNames").Value)
            Catch ex As Exception
                awinSettings.importUnknownNames = True
            End Try

            ' Einstellung, um Geschwister-Namen immer eindeutig zu machen
            Try
                awinSettings.createUniqueSiblingNames = CBool(.Range("uniqueSiblingNames").Value)
            Catch ex As Exception
                awinSettings.createUniqueSiblingNames = True
            End Try

            ' Einstellung, um das Lesen / Schreiben von MissingDefinitions zu steuern 
            Try
                awinSettings.readWriteMissingDefinitions = CBool(.Range("RW_MissingDefinitions").Value)
            Catch ex As Exception
                awinSettings.readWriteMissingDefinitions = False
            End Try

            ' Einstellung, um für MassEdit zu steuern, ob %-tuale Spalte auch angezeigt werden soll
            Try
                awinSettings.meExtendedColumnsView = CBool(.Range("meExtendedView").Value)
            Catch ex As Exception
                awinSettings.meExtendedColumnsView = False
            End Try
            ' Einstellung, um im MassEdit für AutoReduce zu steuern, ob nachgefragt wird, bevor von Folge- oder VorgängerMonaten die Ressource zu holen
            Try
                awinSettings.meDontAskWhenAutoReduce = CBool(.Range("meDontAskWhenAutoReduce").Value)
            Catch ex As Exception
                awinSettings.meDontAskWhenAutoReduce = True
            End Try

            ' Einstellung, um zu signalisieren, dass Rollen und Kosten ausschließlich von der DB gelesen werden sollen ; 
            Try
                awinSettings.readCostRolesFromDB = CBool(.Range("Cost_Roles_fromDB").Value)
            Catch ex As Exception
                awinSettings.readCostRolesFromDB = True
            End Try




            StartofCalendar = awinSettings.kalenderStart
            'StartofCalendar = StartofCalendar.ToLocalTime()

            historicDate = StartofCalendar

            ' Import Typ regelt, um welche DateiFormate es sich bei dem Import handelt
            ' 1: Standard
            ' 2: BMW Rplan Export in Excel 
            Try
                awinSettings.importTyp = CInt(.Range("Import_Typ").Value)
            Catch ex As Exception
                awinSettings.importTyp = 1
            End Try


            ' sollen im Massen-Edit bei der Berechnung der auslastungsWerte die externen aus der Kapa-Datei mitberücksichtigt werden ? 
            Try
                awinSettings.meAuslastungIsInclExt = CBool(.Range("KapaIstMitExt").Value)
            Catch ex As Exception
                awinSettings.meAuslastungIsInclExt = True
            End Try

            ' welche Sprache soll verwendet werden: wenn english, alles andere ist deutsch
            Try
                awinSettings.englishLanguage = CBool(.Range("englishLanguage").Value)
                If awinSettings.englishLanguage Then
                    menuCult = ReportLang(PTSprache.englisch)
                    repCult = menuCult
                    awinSettings.kapaEinheit = "PD"
                Else
                    awinSettings.kapaEinheit = "PT"
                    menuCult = ReportLang(PTSprache.deutsch)
                    repCult = menuCult
                End If
            Catch ex As Exception
                awinSettings.englishLanguage = False
                awinSettings.kapaEinheit = "PT"
                menuCult = ReportLang(PTSprache.deutsch)
                repCult = menuCult
            End Try

            ' sollen Sammelrollen immer nur in Summe dargestellt werden, oder aufgeteilt in Platzhalter / Assigned 
            Try
                awinSettings.showPlaceholderAndAssigned = CBool(.Range("ShowPlaceHolderAndAssigned").Value)
            Catch ex As Exception
                awinSettings.showPlaceholderAndAssigned = False
            End Try

            ' sollen die Risiko Kennzahlen bei der Berechnung der Portfolio / Projekt-Ergebnisse mitgerechnet werden ?  
            Try
                awinSettings.considerRiskFee = CBool(.Range("considerRiskFee").Value)
            Catch ex As Exception
                awinSettings.considerRiskFee = False
            End Try

            '
            ' ende Auslesen Einstellungen in Sheet "Einstellungen"
        End With


    End Sub

    ''' <summary>
    ''' liest für die definierten Rollen ggf vorhandene detaillierte Ressourcen Kapazitäten ein 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub readRessourcenDetails(ByRef meldungen As Collection)

        ' tk 28.5.18 hier werden, sofern es was gibt die monatlichen Details für die Rollen ausgelesen 
        Call readMonthlyExternKapas(meldungen)

    End Sub
    ''' <summary>
    ''' liest für die definierten Rollen ggf vorhandene Urlaubsplanung ein 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub readInterneAnwesenheitslisten(ByRef meldungen As Collection)

        Dim kapaFileName As String
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = Nothing
        Dim anzFehler As Integer = 0

        If formerEE Then
            appInstance.EnableEvents = False
        End If

        If formerSU Then
            appInstance.ScreenUpdating = False
        End If

        enableOnUpdate = False

        kapaFileName = "Urlaubsplaner*.xlsx"

        ' Dateien mit WildCards lesen
        listOfFiles = My.Computer.FileSystem.GetFiles(awinPath & projektRessOrdner,
                     FileIO.SearchOption.SearchTopLevelOnly, kapaFileName)

        ''listOfFiles = My.Computer.FileSystem.GetFiles(awinPath & projektRessOrdner,
        ''              FileIO.SearchOption.SearchTopLevelOnly, "Urlaubsplaner*.xlsx")

        If listOfFiles.Count >= 1 Then

            For Each tmpDatei As String In listOfFiles
                Call logfileSchreiben("Einlesen Verfügbarkeiten " & tmpDatei, "", anzFehler)
                Call readAvailabilityOfRole(tmpDatei, meldungen)
            Next

        Else
            Dim errMsg As String = "Es gibt keine Datei zur Urlaubsplanung" & vbLf _
                         & "Es wurde daher jetzt keine berücksichtigt"

            ' das sollte nicht dazu führen, dass nichts gemacht wird 
            'meldungen.Add(errMsg)

            Call logfileSchreiben(errMsg, "", anzFehler)
        End If

    End Sub

    ''' <summary>
    ''' liest die Projekt- bzw. Modul-Vorlagen ein 
    ''' </summary>
    ''' <param name="isModulVorlage"></param>
    ''' <remarks></remarks>
    Private Sub readVorlagen(ByVal isModulVorlage As Boolean)

        Dim dirName As String
        Dim dateiName As String

        If isModulVorlage Then
            dirName = awinPath & modulVorlagenOrdner
        Else
            dirName = awinPath & projektVorlagenOrdner
        End If

        If My.Computer.FileSystem.DirectoryExists(dirName) Then

            Dim listOfFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirName)

            For i As Integer = 1 To listOfFiles.Count

                dateiName = listOfFiles.Item(i - 1)
                Dim fInf As FileInfo = My.Computer.FileSystem.GetFileInfo(dateiName)
                If fInf.Attributes = FileAttributes.Archive Or fInf.Attributes = FileAttributes.Normal Then
                    ' dann was machen, sonst nicht - dann geht das auch mit Tilde$.. Dateien 
                    If dateiName.Contains(".xls") Or dateiName.Contains(".xlsx") Then
                        Try

                            appInstance.Workbooks.Open(dateiName)


                            If awinSettings.importTyp = 1 Then



                                Dim projVorlage As New clsProjektvorlage

                                ' Auslesen der Projektvorlage wird wie das Importieren eines Projekts behandelt, nur am Ende in die Liste der Projektvorlagen eingehängt
                                ' Kennzeichen für Projektvorlage ist der 3.Parameter im Aufruf (isTemplate)

                                Call awinImportProjectmitHrchy(Nothing, projVorlage, True, Date.Now)

                                ' ur: 21.05.2015: Vorlagen nun neues Format, mit Hierarchie
                                ' Call awinImportProject(Nothing, projVorlage, True, Date.Now)

                                If isModulVorlage Then
                                    ModulVorlagen.Add(projVorlage)
                                Else
                                    Projektvorlagen.Add(projVorlage)
                                End If



                            ElseIf awinSettings.importTyp = 2 Then

                                ' hier muss die Datei ausgelesen werden
                                Dim myCollection As New Collection
                                Dim ok As Boolean
                                Dim hproj As clsProjekt = Nothing

                                Try
                                    Call planExcelImport(myCollection, True, dateiName)
                                    'Call bmwImportProjekteITO15(myCollection, True)

                                    ' jetzt muss für jeden Eintrag in ImportProjekte eine Vorlage erstellt werden  
                                    For Each pName As String In myCollection

                                        ok = True

                                        Try

                                            hproj = ImportProjekte.getProject(pName)

                                        Catch ex As Exception
                                            Call MsgBox("Projekt " & pName & " ist kein gültiges Projekt ... es wird ignoriert ...")
                                            ok = False
                                        End Try

                                        If ok Then

                                            ' hier müssen die Werte für die Vorlage übergeben werden.
                                            ' Änderung tk 19.4.15 Übernehmen der Hierarchie 
                                            Dim projVorlage As New clsProjektvorlage
                                            projVorlage.VorlagenName = hproj.name
                                            projVorlage.Schrift = hproj.Schrift
                                            projVorlage.Schriftfarbe = hproj.Schriftfarbe
                                            projVorlage.farbe = hproj.farbe
                                            projVorlage.earliestStart = -6
                                            projVorlage.latestStart = 6
                                            projVorlage.AllPhases = hproj.AllPhases

                                            projVorlage.hierarchy = hproj.hierarchy

                                            If isModulVorlage Then
                                                ModulVorlagen.Add(projVorlage)
                                            Else
                                                Projektvorlagen.Add(projVorlage)
                                            End If

                                        End If

                                    Next
                                Catch ex As Exception

                                    Call MsgBox(ex.Message & vbLf & dateiName)

                                End Try




                            End If
                            ' ur: Test
                            Dim anzphase As Integer = PhaseDefinitions.Count

                            appInstance.ActiveWorkbook.Close(SaveChanges:=True)


                        Catch ex As Exception
                            appInstance.ActiveWorkbook.Close(SaveChanges:=True)
                            Call MsgBox(ex.Message)
                        End Try
                    End If
                End If

            Next

            Try
                If isModulVorlage Then
                    If ModulVorlagen.Count > 0 Then
                        awinSettings.lastModulTyp = ModulVorlagen.Liste.ElementAt(0).Value.VorlagenName
                        ' Änderung tk 26.11.15 muss doch hier gar nicht gemacht werden .. erst mit Beenden des Wörterbuchs bzw. Beenden der Applikation
                        'Call awinWritePhaseDefinitions()
                        'Call awinWritePhaseMilestoneDefinitions 
                    End If

                Else
                    If Projektvorlagen.Count > 0 Then
                        awinSettings.lastProjektTyp = Projektvorlagen.Liste.ElementAt(0).Value.VorlagenName
                        'Call awinWritePhaseDefinitions()
                        'Call awinWritePhaseMilestoneDefinitions 
                    End If

                End If

            Catch ex As Exception
                awinSettings.lastProjektTyp = ""
            End Try



        Else
            If isModulVorlage Then
                ' nichts tun - kein Problem, wenn es keine Vorlagen gibt 
            Else
                Throw New ArgumentException("der Vorlagen Ordner fehlt:" & vbLf & dirName)
            End If
        End If


    End Sub


    ''' <summary>
    ''' liest die Kosten Definitionen ein 
    ''' wird in der globalen Variablen CostDefinitions abgelegt 
    ''' </summary>
    ''' <param name="wsname"></param>
    ''' <remarks></remarks>
    Private Sub readCostDefinitions(ByVal wsname As Excel.Worksheet, ByRef kostendefinitionen As clsKostenarten, ByRef outputCollection As Collection)


        Dim index As Integer = 0
        Dim hcost As clsKostenartDefinition
        Dim tmpStr As String
        Dim errmsg As String = ""


        Try

            With wsname

                Dim costRange As Excel.Range = .Range("awin_Kosten_Definition")

                If Not IsNothing(costRange) Then
                    Dim anzZeilen As Integer = costRange.Rows.Count
                    Dim c As Excel.Range

                    For i As Integer = 2 To anzZeilen - 1

                        c = CType(costRange.Cells(i, 1), Excel.Range)
                        If CStr(c.Value) <> "" Or index > 0 Then
                            index = index + 1

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
                                .UID = index
                            End With

                            kostendefinitionen.Add(hcost)
                        End If

                    Next
                Else
                    errmsg = "Range <awin_Kosten_Definition> not defined - exit ..."
                    outputCollection.Add(errmsg)
                    kostendefinitionen = New clsKostenarten
                End If


            End With


        Catch ex As Exception
            errmsg = "Range <awin_Kosten_Definition> not defined - exit ..."
            outputCollection.Add(errmsg)
            kostendefinitionen = New clsKostenarten
        End Try


    End Sub



End Module
