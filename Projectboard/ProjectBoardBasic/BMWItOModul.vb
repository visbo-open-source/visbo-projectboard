Imports ProjectBoardDefinitions
Imports Excel = Microsoft.Office.Interop.Excel
Public Module BMWItOModul


    ' spezifisch für BMW Export 

    Friend bmwFC52Vorlage As String = "FC52 Vorlage.xlsx"

    ''' <summary>
    ''' speziell auf BMW Mpp Anforderungen angepasstes BMW Import File
    ''' Status Dezember 2014/Jan 2015
    ''' </summary>
    ''' <param name="myCollection">gibt die Namen der importierten Fahrzeug Projekt zurück</param>
    ''' <remarks></remarks>
    ''' 
    Public Sub bmwImportProjekteITO15(ByRef myCollection As Collection, ByVal isVorlage As Boolean)

        Dim phaseHierarhy(9) As String
        Dim currentHierarchy As Integer = 0
        Dim zeile As Integer, spalte As Integer
        Dim pName As String = " "
        Dim currentDateiName As String
        Dim isMilestone As Boolean

        Dim lastRow As Integer

        Dim hproj As clsProjekt
        Dim vproj As clsProjektvorlage
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

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 




        Dim colName As Integer
        Dim colAnfang As Integer
        Dim colEnde As Integer
        Dim colDauer As Integer
        Dim colProduktlinie As Integer
        Dim colAbbrev As Integer = -1
        Dim colVorgangsKlasse As Integer = -1
        Dim firstZeile As Excel.Range
        Dim protocolRange As Excel.Range


        Dim suchstr(7) As String
        suchstr(ptPlanNamen.Name) = "Name"
        suchstr(ptPlanNamen.Anfang) = "Anfang"
        suchstr(ptPlanNamen.Ende) = "Ende"
        suchstr(ptPlanNamen.Beschreibung) = "Beschreibung"
        suchstr(ptPlanNamen.Vorgangsklasse) = "Vorgangsklasse"
        suchstr(ptPlanNamen.BusinessUnit) = "Spalte A"
        suchstr(ptPlanNamen.Protocol) = "Übernommen als"
        suchstr(ptPlanNamen.Dauer) = "Dauer"


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



        Dim aktivesSheet As Excel.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet, _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

        With aktivesSheet
            firstZeile = CType(.Rows(1), Excel.Range)
        End With



        ' diese Daten müssen vorhanden sein - andernfalls Abbruch 
        Try
            colName = firstZeile.Find(What:=suchstr(ptPlanNamen.Name), LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole).Column
            colAnfang = firstZeile.Find(What:=suchstr(ptPlanNamen.Anfang), LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole).Column
            colEnde = firstZeile.Find(What:=suchstr(ptPlanNamen.Ende), LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole).Column

        Catch ex As Exception
            Throw New ArgumentException("Fehler im Datei Aufbau ..." & vbLf & ex.Message)
        End Try

        Try
            colDauer = firstZeile.Find(What:=suchstr(ptPlanNamen.Dauer), LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole).Column
        Catch ex As Exception
            colDauer = -1
        End Try


        Try
            colProduktlinie = firstZeile.Find(What:=suchstr(ptPlanNamen.BusinessUnit), LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole).Column
        Catch ex As Exception
            colProduktlinie = -1
        End Try

        ' diese Daten können vorhanden sein - wenn nicht, weitermachen ...  
        Try
            colAbbrev = firstZeile.Find(What:=suchstr(ptPlanNamen.Beschreibung), LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole).Column
            colVorgangsKlasse = firstZeile.Find(What:=suchstr(ptPlanNamen.Vorgangsklasse), LookAt:=Microsoft.Office.Interop.Excel.XlLookAt.xlWhole).Column
        Catch ex As Exception

        End Try


        With aktivesSheet

            lastRow = System.Math.Max(CType(.Cells(40000, colName), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row, _
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
            End If

            ' wird immer geschrieben 
            CType(.Cells(1, colProtocol + 6), Excel.Range).Value = suchstr(ptPlanNamen.Protocol)
            CType(.Cells(1, colProtocol + 7), Excel.Range).Value = "Grund"

        End With







        Try

            With aktivesSheet

                Try
                    projektFarbe = CType(aktivesSheet.Cells(zeile, 1), Excel.Range).Interior.Color
                    ' das Folgende wird nur für die Projekt-Vorlagen benötigt (isVorlage = true) 
                    schriftfarbe = CLng(CType(aktivesSheet.Cells(zeile, 1), Excel.Range).Font.Color)
                    schriftGroesse = CInt(CType(aktivesSheet.Cells(zeile, 1), Excel.Range).Font.Size)

                Catch ex As Exception
                    projektFarbe = CType(aktivesSheet.Cells(zeile, 1), Excel.Range).Interior.ColorIndex
                End Try


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
                    completeName = CStr(CType(aktivesSheet.Cells(zeile, colName), Excel.Range).Value).Trim
                    startDate = CDate(CType(aktivesSheet.Cells(zeile, colAnfang), Excel.Range).Value)
                    endDate = CDate(CType(aktivesSheet.Cells(zeile, colEnde), Excel.Range).Value)

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

                    If Not isVorlage Then
                        If tmpStr(0).Trim.EndsWith("eA") Then
                            typKennung = "eA"

                        ElseIf tmpStr(0).Trim.EndsWith("wA") Then
                            typKennung = "wA"

                        ElseIf tmpStr(0).Trim.EndsWith("E") Then
                            typKennung = "E"

                        Else
                            typKennung = "?"
                        End If
                    Else
                        vorlagenName = ""
                    End If



                    ' prüfen, ob das Projekt überhaupt vollständig im Kalender liegt 
                    ' wenn nein, dann nicht importieren 
                    If DateDiff(DateInterval.Day, StartofCalendar, startDate) < 0 Then

                        Call MsgBox("Projekt liegt vor dem Kalender-Anfang und wird deshalb nicht importiert")

                    Else
                        '
                        ' jetzt wird das Projekt angelegt 
                        '
                        hproj = New clsProjekt




                        Try

                            hproj.name = pName
                            hproj.startDate = startDate
                            ' Projekte sollten erstmal nicht verschoben werden können
                            ' dazu muss eine Variante erzeugt werden , die kann dann verschoben werden 
                            hproj.Status = ProjektStatus(1)

                            If DateDiff(DateInterval.Month, startDate, Date.Now) <= 0 Then
                                hproj.earliestStartDate = hproj.startDate.AddMonths(hproj.earliestStart)
                                hproj.latestStartDate = hproj.startDate.AddMonths(hproj.latestStart)
                            Else
                                hproj.earliestStartDate = startDate
                                hproj.latestStartDate = startDate
                            End If

                            hproj.StrategicFit = zufall.NextDouble * 10
                            hproj.Risiko = zufall.NextDouble * 10
                            hproj.volume = zufall.NextDouble * 1000000
                            hproj.complexity = zufall.NextDouble
                            hproj.businessUnit = defaultBU
                            hproj.description = ""

                            hproj.Erloes = 0.0


                        Catch ex As Exception
                            Throw New Exception("in erstelle Import BMW Projekte: " & vbLf & ex.Message)
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

                        hproj.AddPhase(cphase)

                        Try
                            pHierarchy.add(cphase, rootPhaseName, 0)
                            origHierarchy.add(cphase, rootPhaseName, 0)
                        Catch ex As Exception

                        End Try

                        Dim itemStartDate As Date
                        Dim itemEndDate As Date
                        Dim ok As Boolean = True

                        Dim curZeile As Integer
                        Dim txtVorgangsKlasse As String
                        Dim origVorgangsKlasse As String
                        Dim txtAbbrev As String
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

                            Dim indentLevel As Integer

                            Try

                                origItem = CStr(CType(.Cells(curZeile, colName), Excel.Range).Value)

                                If origItem.Trim.Length = 0 Then

                                    'CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                    '            "leerer String wird ignoriert .."
                                    logMessage = "leerer String wird ignoriert .."
                                    ok = False

                                Else

                                    ' bestimme den Indent-Level 
                                    indentLevel = pHierarchy.getLevel(origItem)
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



                                        ' Änderung 26.1.15 Ignorieren 

                                        itemStartDate = CDate(CType(.Cells(curZeile, colAnfang), Excel.Range).Value)
                                        itemEndDate = CDate(CType(.Cells(curZeile, colEnde), Excel.Range).Value)

                                        If DateDiff(DateInterval.Day, itemStartDate, itemEndDate) = 0 Then
                                            isMilestone = True
                                        Else
                                            isMilestone = False
                                        End If

                                        If itemName = "Projektphasen" Then
                                            Try
                                                Dim tmpBU As String
                                                If colProduktlinie > 0 Then
                                                    tmpBU = CStr(CType(.Cells(curZeile, colProduktlinie), Excel.Range).Value).Trim
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
                                                            hproj.businessUnit = tmpBU

                                                            If awinSettings.fullProtocol Then

                                                                CType(aktivesSheet.Cells(curZeile, colProtocol - 1), Excel.Range).Value = tmpBU
                                                            End If


                                                        Else
                                                            bix = bix + 1
                                                        End If
                                                    End While
                                                End If


                                                If Not found And awinSettings.fullProtocol Then

                                                    CType(aktivesSheet.Cells(curZeile, colProtocol - 1), Excel.Range).Value = hproj.businessUnit

                                                End If

                                            Catch ex1 As Exception

                                            End Try
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
                                    If elemNameOfElemID(parentNodeID) <> parentElemName Then
                                        Call MsgBox("nicht konsistent in bmwImportProjekteITO15, zeile 663")
                                    End If


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
                                            Dim duplicateSiblingID As String = hproj.getDuplicatePhaseSiblingID(stdName, parentPhase.nameID, _
                                                                                                                 itemStartDate, itemEndDate, 0.97)

                                            If duplicateSiblingID = "" Then
                                                ok1 = True
                                            Else
                                                ok1 = False
                                                logMessage = stdName & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) & _
                                                             " und wird ignoriert "
                                            End If
                                        End If



                                    End If



                                    ' jetzt muss geprüft werden, ob das Element in Std Definitions aufgenommen werden muss 
                                    Dim ok2 As Boolean = True

                                    If Not PhaseDefinitions.Contains(stdName) And ok1 Then

                                        If isMissingDefinitionOK(stdName, isVorlage, False) Then
                                            ok2 = True
                                        Else
                                            ok2 = False
                                            logMessage = "ist nicht in der Liste der zugelassenen Elemente enthalten"
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

                                        ' der Aufbau der Hierarchie erfolgt in addphase
                                        hproj.AddPhase(cphase, origName:=origItem.Trim, _
                                                       parentID:=pHierarchy.getIDBeforeLevel(indentLevel))

                                        ' wird übernommen als 
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Value = stdName

                                        Dim PTBreadCrumb As String = hproj.hierarchy.getBreadCrumb(elemID)


                                        If awinSettings.fullProtocol Then

                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 8), Excel.Range).Value = PTBreadCrumb
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 9), Excel.Range).Value = txtVorgangsKlasse
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

                                        Dim bewertungsAmpel As Integer = 0
                                        Dim explanation As String = ""

                                        ' hole die Parentphase
                                        cphase = pHierarchy.getPhaseBeforeLevel(indentLevel)
                                        cmilestone = New clsMeilenstein(parent:=cphase)
                                        cbewertung = New clsBewertung


                                        ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                        With cbewertung
                                            '.bewerterName = resultVerantwortlich
                                            .colorIndex = bewertungsAmpel
                                            .datum = Date.Now
                                            .description = explanation
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
                                            Dim duplicateSiblingID As String = hproj.getDuplicateMsSiblingID(stdName, cphase.nameID, _
                                                                                                                 itemStartDate, 0)

                                            If duplicateSiblingID = "" Then
                                                ok1 = True
                                            Else
                                                ok1 = False
                                                logMessage = stdName & " ist Duplikat zu Geschwister " & elemNameOfElemID(duplicateSiblingID) & _
                                                             " und wird ignoriert "
                                            End If

                                        End If


                                        ' jetzt muss geprüft werden, ob stdName bereits aufgenommen ist
                                        Dim ok2 As Boolean = True
                                        If Not MilestoneDefinitions.Contains(stdName) And ok1 Then

                                            If isMissingDefinitionOK(stdName, isVorlage, True) Then
                                                ok2 = True
                                            Else
                                                ok2 = False
                                                logMessage = "ist nicht in der Liste der zugelassenen Elemente enthalten"
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
                                                CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                                        stdName.Trim & " existiert bereits: Datum 1: " & cphase.getMilestone(stdName).getDate.ToShortDateString & _
                                                        "   , Datum 2: " & cmilestone.getDate.ToShortDateString

                                            End If
                                        Else

                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = logMessage
                                            CType(aktivesSheet.Cells(curZeile, colProtocol + 6), Excel.Range).Interior.Color = awinSettings.AmpelRot
                                            anzIgnored = anzIgnored + 1

                                        End If


                                    Catch ex As Exception
                                        CType(aktivesSheet.Cells(curZeile, colProtocol + 7), Excel.Range).Value = _
                                                            "Fehler in Zeile " & zeile & ", Item-Name: " & itemName
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

                            Try
                                Dim sopDate As Date = hproj.getMilestone("SOP").getDate

                                If DateDiff(DateInterval.Month, StartofCalendar, sopDate) > 0 Then
                                    Dim sopMonth As Integer = sopDate.Month
                                    If sopMonth >= 3 And sopMonth <= 6 Then
                                        anlaufKennung = "03"
                                    ElseIf sopMonth >= 7 And sopMonth <= 10 Then
                                        anlaufKennung = "07"
                                    Else
                                        anlaufKennung = "11"
                                    End If
                                Else
                                    anlaufKennung = "?"
                                End If

                            Catch ex As Exception
                                anlaufKennung = "?"
                            End Try

                            ' jetzt wird die Vorlagen Kennung bestimmt 
                            Dim tstphase As clsPhase = Nothing
                            Dim relNr As String
                            tstphase = hproj.getPhase("Systemgestaltung")

                            If IsNothing(tstphase) Then
                                tstphase = hproj.getPhase("I500")
                                If IsNothing(tstphase) Then
                                    tstphase = hproj.getPhase("I300")
                                    If IsNothing(tstphase) Then
                                        relNr = "rel 4 "
                                    Else
                                        relNr = "rel 5 "
                                    End If
                                Else
                                    relNr = "rel 5 "
                                End If
                            Else
                                relNr = "rel 5 "
                            End If

                            vorlagenName = relNr & typKennung & "-" & anlaufKennung
                            Try
                                vorlagenName = vorlagenName.Trim
                            Catch ex As Exception
                                vorlagenName = "unknown"
                            End Try

                            If Projektvorlagen.Contains(vorlagenName) Then
                                hproj.VorlagenName = vorlagenName
                            Else
                                hproj.VorlagenName = vorlagenName & "*"
                            End If

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
                                    hproj.farbe = awinSettings.AmpelNichtBewertet
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
                        ImportProjekte.Add(hproj, False)
                        myCollection.Add(calcProjektKey(hproj))

                    End If

                    zeile = ende + 1



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
            Throw New Exception("Fehler in Datei BMW Import ITO15 " & vbLf & ex.Message & vbLf & _
                                 pName & vbLf)
        End Try


    End Sub


    ''' <summary>
    ''' exportiert das angegebene Projekt in die bereits geöffnete Datei 
    ''' Das Schreiben beginnt ab "zeile"
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="zeile"></param>
    ''' <remarks></remarks>
    Public Sub bmwExportProject(ByVal hproj As clsProjekt, ByRef zeile As Integer)

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


        color = CLng(CType(ws.Cells(2, 1), Excel.Range).Interior.Color)

        ' jetzt wird das Projekt geschrieben 
        CType(ws.Cells(zeile, spalte), Excel.Range).Value = hproj.getShapeText
        CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = hproj.startDate.ToShortDateString
        CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = hproj.endeDate.ToShortDateString
        CType(ws.Rows(zeile), Excel.Range).Interior.Color = color

        Dim indentPhase As String = "   "
        'Dim indentMS As String = "      "

        ' die erste Phase kann auch Meilensteine haben !
        cphase = hproj.getPhase(1)
        indentlevel = hproj.hierarchy.getIndentLevel(cphase.nameID)

        For im = 1 To cphase.countMilestones
            zeile = zeile + 1
            cmilestone = cphase.getMilestone(im)
            startdate = cmilestone.getDate
            ' Änderung 20.4.15
            ' alt: 
            'If cmilestone.nameID.StartsWith(cphase.name & "+") Then

            '    Dim parentName As String = cphase.name & "+"
            '    curName = ""
            '    Dim posStart As Integer = parentName.Length

            '    For posX As Integer = posStart + 1 To cmilestone.nameID.Length
            '        curName = curName & cmilestone.nameID.Chars(posX)
            '    Next

            '    ' hier den Original Name verwenden !? nein, aktuell noch nicht 

            'Else
            '    curName = cmilestone.nameID
            'End If
            ' neu:
            curName = cmilestone.name

            indentlevel = hproj.hierarchy.getIndentLevel(cmilestone.nameID)
            CType(ws.Cells(zeile, spalte), Excel.Range).Value = erzeugeIndent(indentlevel) & curName

            If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = startdate.ToShortDateString
                CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = startdate.ToShortDateString
            Else
                CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = "Fehler !"
                CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = "Fehler !"
            End If
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
                CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = startdate.ToShortDateString
            Else
                CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = "Fehler !"
            End If

            If DateDiff(DateInterval.Day, StartofCalendar, endDate) > 0 Then
                CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = endDate.ToShortDateString
            Else
                CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = "Fehler !"
            End If

            For im = 1 To cphase.countMilestones
                zeile = zeile + 1
                cmilestone = cphase.getMilestone(im)
                startdate = cmilestone.getDate
                'If cmilestone.nameID.StartsWith(cphase.name & "+") Then

                '    Dim parentName As String = cphase.name & "+"
                '    curName = ""
                '    Dim posStart As Integer = parentName.Length

                '    For posX As Integer = posStart + 1 To cmilestone.nameID.Length
                '        curName = curName & cmilestone.nameID.Chars(posX)
                '    Next

                '    ' hier den Original Name verwenden !? nein, aktuell noch nicht 

                'Else
                '    curName = cmilestone.nameID
                'End If

                curName = cmilestone.name
                indentlevel = hproj.hierarchy.getIndentLevel(cmilestone.nameID)
                CType(ws.Cells(zeile, spalte), Excel.Range).Value = erzeugeIndent(indentlevel) & curName

                If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                    CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = startdate.ToShortDateString
                    CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = startdate.ToShortDateString
                Else
                    CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = "Fehler !"
                    CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = "Fehler !"
                End If
            Next

        Next

        ' jetzt muss um eine Zeile weitergeschaltet werden, damit immer auf eine freie Zeile geschrieben wird
        zeile = zeile + 1

    End Sub


    ''' <summary>
    ''' schreibt gemäß der FC-52 Vorlage die aktuell geladenen Projekte in eine Datei im Export Directory
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinWriteFC52()


        appInstance.EnableEvents = False


        ' hier muss jetzt das entsprechende File aufgemacht werden ...
        ' das File 
        Try
            appInstance.Workbooks.Open(awinPath & requirementsOrdner & bmwFC52Vorlage)

        Catch ex As Exception
            Call MsgBox("FC52 Vorlage nicht gefunden - Abbruch")
            Throw New ArgumentException("FC52 Vorlage nicht gefunden - Abbruch")
        End Try

        'appInstance.Workbooks(myCustomizationFile).Activate()
        Dim wsName As Excel.Worksheet = CType(appInstance.Worksheets("Report"), _
                                                Global.Microsoft.Office.Interop.Excel.Worksheet)


        Dim zeile As Integer = 2
        Dim spalte As Integer = 1
        Dim tmpdate As Date
        Dim milestone As clsMeilenstein = Nothing



        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            With wsName
                ' Hauptkategorie nicht in RPLAN Export vorhanden  

                If kvp.Value.businessUnit.Length > 0 Then
                    CType(.Cells(zeile, spalte), Excel.Range).Value = kvp.Value.businessUnit
                Else
                    CType(.Cells(zeile, spalte), Excel.Range).Value = "-"
                End If


                ' Name schreiben 
                CType(.Cells(zeile, spalte + 1), Excel.Range).Value = kvp.Value.name

                ' Zielvereinbarung schreiben 
                Try

                    milestone = kvp.Value.getMilestone("Zielvereinbarung")
                    If Not IsNothing(milestone) Then
                        tmpdate = milestone.getDate
                        CType(.Cells(zeile, spalte + 2), Excel.Range).Value = tmpdate.ToShortDateString
                    Else
                        milestone = kvp.Value.getMilestone("SP ZVA")
                        If Not IsNothing(milestone) Then
                            tmpdate = milestone.getDate
                            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = tmpdate.ToShortDateString
                        Else
                            CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                        End If
                    End If


                Catch ex As Exception
                    CType(.Cells(zeile, spalte + 2), Excel.Range).Value = "-"
                End Try

                'SOP schreiben
                Try

                    milestone = kvp.Value.getMilestone("SOP")
                    If Not IsNothing(milestone) Then
                        tmpdate = milestone.getDate
                        CType(.Cells(zeile, spalte + 3), Excel.Range).Value = tmpdate.ToShortDateString
                    Else
                        CType(.Cells(zeile, spalte + 3), Excel.Range).Value = "-"
                    End If

                Catch ex As Exception
                    CType(.Cells(zeile, spalte + 3), Excel.Range).Value = "-"
                End Try

                ' MEPS schreiben - Markteinführung 

                Try

                    milestone = kvp.Value.getMilestone("Bestätigung Markteinführung & Prozess-Sicherheit")
                    If Not IsNothing(milestone) Then
                        tmpdate = milestone.getDate
                        CType(.Cells(zeile, spalte + 4), Excel.Range).Value = tmpdate.ToShortDateString
                    Else
                        CType(.Cells(zeile, spalte + 4), Excel.Range).Value = "-"
                    End If

                Catch ex As Exception
                    CType(.Cells(zeile, spalte + 4), Excel.Range).Value = "-"
                End Try


                ' End of Production ist nicht im RPLAN abgelegt 
                CType(.Cells(zeile, spalte + 5), Excel.Range).Value = "-"



            End With

            zeile = zeile + 1

        Next

        'Dim expFName As String = awinPath & exportFilesOrdner & "\Report_" & _
        '    Date.Now.ToString.Replace(":", ".") & ".xlsx"

        Dim expFName As String = exportOrdnerNames(PTImpExp.rplan) & "\Report_" & _
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

        Call MsgBox("ok, Report exportiert")

    End Sub

    ''' <summary>
    ''' testet die Konsistenz der Positionierungs-Informationen in projectboardshapes, im Projekt und in Showprojekte 
    ''' allerdings ist Showprojekte aktuell nicht wirklich relevant ... 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub testConsistencyOfPT1()
        Dim atleastOne As Boolean = False
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            Dim zeile1 As Integer = kvp.Value.tfZeile
            'Dim zeile2 As Integer = ShowProjekte.getPTZeile(kvp.Key)
            Dim zeile2 As Integer = currentSessionConstellation.getBoardZeile(kvp.Key)
            If zeile1 <> zeile2 Then
                Call MsgBox("ungleich 1/2: " & kvp.Key & ": " & zeile1.ToString & " , " & zeile2.ToString)
                atleastOne = True
            End If

            Dim coord() As Double = projectboardShapes.getCoord(kvp.Key)
            Dim zeile3 As Integer = calcYCoordToZeile(coord(0))

            If zeile1 <> zeile3 Then
                Call MsgBox("ungleich 1/3: " & kvp.Key & ": " & zeile1.ToString & " , " & zeile3.ToString)
                atleastOne = True
            End If

        Next

        If Not atleastOne Then
            Call MsgBox("alles i.O")
        End If
    End Sub
End Module
