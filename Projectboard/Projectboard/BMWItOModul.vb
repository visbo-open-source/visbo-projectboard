Imports ProjectBoardDefinitions
Imports Excel = Microsoft.Office.Interop.Excel
Module BMWItOModul
    

    Private Enum ptNamen
        Name = 0
        Anfang = 1
        Ende = 2
        Beschreibung = 3
        Vorgangsklasse = 4
        Produktlinie = 5
    End Enum

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
        Dim completeName As String
        Dim nameSopTyp As String = " "
        Dim nameProduktlinie As String = ""

        Dim startDate As Date, endDate As Date
        Dim startoffset As Long, duration As Long
        Dim vorlagenName As String

        Dim itemName As String
        Dim zufall As New Random(10)
        Dim protocolColumn As Integer = 20

        Dim schriftGroesse As Integer
        Dim schriftfarbe As Long


        Dim milestoneIX As Integer = MilestoneDefinitions.Count + 1
        Dim phaseIX As Integer = PhaseDefinitions.Count + 1

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 




        Dim colName As Integer
        Dim colAnfang As Integer
        Dim colEnde As Integer
        Dim colAbbrev As Integer = -1
        Dim colVorgangsKlasse As Integer = -1
        Dim firstZeile As Excel.Range

        Dim suchstr(5) As String
        suchstr(ptNamen.Name) = "Name"
        suchstr(ptNamen.Anfang) = "Anfang"
        suchstr(ptNamen.Ende) = "Ende"
        suchstr(ptNamen.Beschreibung) = "Beschreibung"
        suchstr(ptNamen.Vorgangsklasse) = "Vorgangsklasse"
        suchstr(ptNamen.Produktlinie) = "Spalte A"


        zeile = 2
        spalte = 5
        geleseneProjekte = 0


        Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet, _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

        firstZeile = CType(activeWSListe.Rows(1), Excel.Range)
        ' jetzt die wichtigen Spalten bestimmen 

        ' diese Daten müssen vorhanden sein - andernfalls Abbruch 
        Try
            colName = firstZeile.Find(What:=suchstr(ptNamen.Name)).Column
            colAnfang = firstZeile.Find(What:=suchstr(ptNamen.Anfang)).Column
            colEnde = firstZeile.Find(What:=suchstr(ptNamen.Ende)).Column
        Catch ex As Exception
            Throw New ArgumentException("Fehler im Datei Aufbau ..." & vbLf & ex.Message)
        End Try

        ' diese Daten können vorhanden sein - wenn nicht, weitermachen ...  
        Try
            colAbbrev = firstZeile.Find(What:=suchstr(ptNamen.Beschreibung)).Column
            colVorgangsKlasse = firstZeile.Find(What:=suchstr(ptNamen.Vorgangsklasse)).Column
        Catch ex As Exception

        End Try


        Try




            With activeWSListe




                Try
                    projektFarbe = CType(activeWSListe.Cells(zeile, 1), Excel.Range).Interior.Color
                    ' das Folgende wird nur für die Projekt-Vorlagen benötigt (isVorlage = true) 
                    schriftfarbe = CLng(CType(activeWSListe.Cells(zeile, 1), Excel.Range).Font.Color)
                    schriftGroesse = CInt(CType(activeWSListe.Cells(zeile, 1), Excel.Range).Font.Size)

                Catch ex As Exception
                    projektFarbe = CType(activeWSListe.Cells(zeile, 1), Excel.Range).Interior.ColorIndex
                End Try


                lastRow = System.Math.Max(CType(.Cells(40000, colName), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row, _
                                          CType(.Cells(40000, colAnfang), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row)

                While zeile <= lastRow


                    ix = zeile + 1

                    Dim zellenFarbe As Long = CLng(CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color)
                    Do While zellenFarbe <> CLng(projektFarbe) And (ix <= lastRow)
                        ix = ix + 1
                        zellenFarbe = CLng(CType(.Cells(ix, 1), Global.Microsoft.Office.Interop.Excel.Range).Interior.Color)
                    Loop

                    anfang = zeile + 1
                    ende = ix - 1

                    ' hier wird Name, Typ, SOP, Business Unit, vname, Start-Datum, Dauer der Phase(1) ausgelesen  
                    completeName = CStr(CType(activeWSListe.Cells(zeile, colName), Excel.Range).Value).Trim
                    startDate = CDate(CType(activeWSListe.Cells(zeile, colAnfang), Excel.Range).Value)
                    endDate = CDate(CType(activeWSListe.Cells(zeile, colEnde), Excel.Range).Value)



                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    If duration < 0 Then
                        startDate = endDate
                        duration = -1 * duration
                        endDate = startDate.AddDays(duration)
                    End If

                    tmpStr = completeName.Trim.Split(New Char() {CChar("["), CChar("]")}, 5)
                    If tmpStr(0).Contains("SOP") Then
                        Dim positionIX As Integer = tmpStr(0).IndexOf("SOP") - 1
                        pName = ""
                        For ih As Integer = 0 To positionIX
                            pName = pName & tmpStr(0).Chars(ih)
                        Next
                        pName = pName.Trim
                    Else
                        pName = tmpStr(0).Trim
                    End If

                    If Not isVorlage Then
                        If tmpStr(0).Trim.EndsWith("eA") Then
                            vorlagenName = "enge Ableitung"
                        ElseIf tmpStr(0).Trim.EndsWith("wA") Then
                            vorlagenName = "weite Ableitung"
                        ElseIf tmpStr(0).Trim.EndsWith("E") Then
                            vorlagenName = "Erstanläufer"
                        Else
                            vorlagenName = ""
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
                            If isVorlage Then
                                hproj.farbe = projektFarbe
                                hproj.Schrift = schriftGroesse
                                hproj.Schriftfarbe = schriftfarbe
                            Else
                                vproj = Projektvorlagen.getProject(vorlagenName)
                                hproj.farbe = vproj.farbe
                                hproj.Schrift = vproj.Schrift
                                hproj.Schriftfarbe = vproj.Schriftfarbe
                                hproj.name = ""
                                hproj.VorlagenName = vorlagenName
                                hproj.earliestStart = vproj.earliestStart
                                hproj.latestStart = vproj.latestStart
                                hproj.ampelStatus = PTfarbe.none
                                hproj.leadPerson = ""
                                hproj.businessUnit = ""
                            End If




                        Catch ex As Exception
                            Throw New Exception("es gibt keine entsprechende Vorlage ..  " & vbLf & ex.Message)
                        End Try


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
                            hproj.businessUnit = ""
                            hproj.description = ""

                            hproj.Erloes = 0.0


                        Catch ex As Exception
                            Throw New Exception("in erstelle Import BMW Projekte: " & vbLf & ex.Message)
                        End Try

                        ' jetzt wird die Import Hierarchie angelegt 
                        Dim pHierarchy As New clsImportFileHierarchy


                        ' jetzt werden all die Phasen angelegt , beginnend mit der ersten 
                        cphase = New clsPhase(parent:=hproj)
                        cphase.name = pName
                        startoffset = 0
                        duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                        cphase.changeStartandDauer(startoffset, duration)


                        hproj.AddPhase(cphase)

                        Try
                            pHierarchy.add(cphase, 0)
                        Catch ex As Exception

                        End Try

                        Dim pStartDate As Date
                        Dim pEndDate As Date
                        Dim ok As Boolean = True

                        Dim curZeile As Integer
                        Dim txtVorgangsKlasse As String
                        Dim txtAbbrev As String

                        ' hier werden jetzt die einzelnen Zeilen = Phasen oder Meilensteine ausgelesen 
                        For curZeile = anfang To ende

                            txtVorgangsKlasse = ""
                            txtAbbrev = ""

                            Try
                                itemName = CStr(CType(.Cells(curZeile, colName), Excel.Range).Value)


                                If itemName.Trim = "Projektphasen" Then
                                    Try
                                        Dim tmpBU As String = CStr(CType(.Cells(curZeile, colName), Excel.Range).Value).Trim
                                        If tmpBU.Length > 0 Then
                                            If businessUnit.Contains(tmpBU) Then
                                                hproj.businessUnit = tmpBU
                                                CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                                "Wert für Business Unit erkannt: " & tmpBU
                                            End If
                                        End If
                                    Catch ex1 As Exception

                                    End Try
                                End If

                                ' jetzt prüfen, ob es sich um ein grundsätzlich zu ignorierendes Element handelt .. 
                                If phaseMappings.tobeIgnored(itemName) Or milestoneMappings.tobeIgnored(itemName) Then
                                    CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                                    "Element wird ignoriert: " & itemName.Trim
                                    ok = False
                                Else
                                    ok = True
                                End If

                            Catch ex As Exception
                                itemName = ""
                                ok = False
                            End Try

                            If ok Then


                                pStartDate = CDate(CType(.Cells(curZeile, colAnfang), Excel.Range).Value)
                                pEndDate = CDate(CType(.Cells(curZeile, colEnde), Excel.Range).Value)


                                startoffset = DateDiff(DateInterval.Day, hproj.startDate, pStartDate)
                                duration = DateDiff(DateInterval.Day, pStartDate, pEndDate) + 1


                                ' jetzt werden vorgangsklasse und Abkürzung rausgelesen 
                                Try

                                    txtVorgangsKlasse = CStr((CType(.Cells(curZeile, colVorgangsKlasse), Excel.Range).Value)).Trim
                                    If duration > 1 Then
                                        txtVorgangsKlasse = mapToAppearance(txtVorgangsKlasse, False)
                                        CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                                "auf folgende Phasen Darstellungsklasse abgebildet: " & txtVorgangsKlasse.Trim
                                    Else
                                        txtVorgangsKlasse = mapToAppearance(txtVorgangsKlasse, True)
                                        CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                                "auf folgende Meilenstein Darstellungsklasse abgebildet: " & txtVorgangsKlasse.Trim
                                    End If




                                Catch ex As Exception

                                    CType(activeWSListe.Cells(curZeile, protocolColumn + 2), Excel.Range).Value = _
                                                "Fehler bei Abbildung auf Darstellungsklasse ... " & txtVorgangsKlasse.Trim

                                End Try

                                ' jetzt wird die Abkürzung rausgelesen 
                                Try

                                    txtAbbrev = CStr((CType(.Cells(curZeile, colAbbrev), Excel.Range).Value)).Trim

                                Catch ex As Exception

                                End Try

                                Dim realName As String

                                If duration > 1 Then
                                    ' es handelt sich um eine Phase 
                                    'phaseName = itemName

                                    ' erstmal prüfen, ob es sich um eine "Phasen Dopplung" handelt - dann soll das Element ignoriert werden 
                                    If pHierarchy.dopplung(itemName) Then
                                        ' nichts tun - das Element soll ignoriert werden
                                        CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                                "Ignoriert wegen Dopplung: " & itemName.Trim

                                    ElseIf itemName.Trim.Length = 0 Then
                                        CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                                "leerer String wurde ignoriert  " & itemName.Trim


                                    Else
                                        Dim indentLevel As Integer
                                        ' bestimme den Indent-Level , damit die Hierarchie
                                        indentLevel = pHierarchy.getLevel(itemName)

                                        Dim parentPhaseName As String = pHierarchy.getPhaseBeforeLevel(indentLevel).name

                                        ' jetzt den tatsächlichen Namen bestimmen , ggf wird dazu der Parent Phase Name benötigt 
                                        Try

                                            If Not PhaseDefinitions.Contains(itemName) Then
                                                realName = phaseMappings.mapToRealName(parentPhaseName, itemName)
                                            Else
                                                realName = itemName.Trim
                                            End If

                                        Catch ex As Exception
                                            realName = itemName.Trim
                                        End Try


                                        If realName.Trim <> itemName.Trim Then
                                            CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                                    itemName.Trim & " --> " & realName.Trim
                                        End If


                                        If PhaseDefinitions.Contains(realName) Then
                                            ' nichts tun 
                                        Else
                                            ' in die Phase-Definitions aufnehmen 


                                            Dim hphase As clsPhasenDefinition
                                            hphase = New clsPhasenDefinition

                                            'hphase.farbe = CLng(CType(.Cells(curZeile, 1), Excel.Range).Interior.Color)
                                            hphase.darstellungsKlasse = txtVorgangsKlasse
                                            hphase.shortName = txtAbbrev
                                            hphase.name = realName
                                            hphase.UID = phaseIX
                                            phaseIX = phaseIX + 1

                                            Try
                                                PhaseDefinitions.Add(hphase)
                                            Catch ex As Exception

                                            End Try

                                        End If

                                        cphase = New clsPhase(parent:=hproj)
                                        cphase.name = realName

                                        cphase.changeStartandDauer(startoffset, duration)

                                        hproj.AddPhase(cphase)



                                        Try
                                            pHierarchy.add(cphase, indentLevel)
                                        Catch ex As Exception
                                            'Call MsgBox("Phase " & cphase.name & ", Level = " & indentLevel)
                                        End Try
                                        'lastPhaseName = cphase.name
                                    End If






                                ElseIf duration = 1 Then

                                    If itemName.Trim.Length > 0 Then

                                        Dim indentLevel As Integer
                                        ' bestimme den Indent-Level 
                                        indentLevel = pHierarchy.getLevel(itemName)



                                        Try
                                            ' es handelt sich um einen Meilenstein 

                                            Dim bewertungsAmpel As Integer = 0
                                            Dim explanation As String = ""

                                            'bewertungsAmpel = CInt(CType(.Cells(curZeile, 12), Excel.Range).Value)
                                            'explanation = CStr(CType(.Cells(curZeile, 1), Excel.Range).Value)

                                            'cphase = hproj.getPhase(lastPhaseName)
                                            cphase = pHierarchy.getPhaseBeforeLevel(indentLevel)
                                            cresult = New clsMeilenstein(parent:=cphase)
                                            cbewertung = New clsBewertung

                                            'If bewertungsAmpel < 0 Or bewertungsAmpel > 3 Then
                                            '    ' es gibt keine Bewertung
                                            '    bewertungsAmpel = 0
                                            'End If

                                            ' damit Kriterien auch eingelesen werden, wenn noch keine Bewertung existiert ...
                                            With cbewertung
                                                '.bewerterName = resultVerantwortlich
                                                .colorIndex = bewertungsAmpel
                                                .datum = Date.Now
                                                .description = explanation
                                            End With


                                            Dim parentPhaseName As String = cphase.name
                                            ' jetzt den tatsächlichen Namen bestimmen , ggf wird dazu der Parent Phase Name benötigt 

                                            Try
                                                If Not MilestoneDefinitions.Contains(itemName) Then
                                                    realName = milestoneMappings.mapToRealName(parentPhaseName, itemName)
                                                Else
                                                    realName = itemName
                                                End If

                                            Catch ex As Exception
                                                realName = itemName.Trim
                                            End Try


                                            If realName.Trim <> itemName.Trim Then
                                                CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                                        itemName.Trim & " --> " & realName.Trim
                                            End If

                                            With cresult
                                                .name = realName
                                                .setDate = pEndDate
                                                If Not cbewertung Is Nothing Then
                                                    .addBewertung(cbewertung)
                                                End If
                                            End With

                                            ' Meilenstein in aufnehmen, 
                                            If MilestoneDefinitions.Contains(realName) Then
                                                ' nichts tun 
                                            Else
                                                ' in die Milestone-Definitions aufnehmen 

                                                Dim hMilestone As New clsMeilensteinDefinition

                                                With hMilestone
                                                    .name = realName
                                                    .belongsTo = parentPhaseName
                                                    .shortName = txtAbbrev
                                                    .darstellungsKlasse = txtVorgangsKlasse
                                                    .UID = milestoneIX
                                                End With

                                                milestoneIX = milestoneIX + 1

                                                Try
                                                    MilestoneDefinitions.Add(hMilestone)
                                                Catch ex As Exception

                                                End Try

                                            End If



                                            With cphase
                                                .addresult(cresult)
                                            End With
                                        Catch ex As Exception

                                        End Try
                                    Else

                                        CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                                "leerer String wurde ignoriert  " & itemName.Trim

                                    End If





                                End If




                                ' handelt es sich um eine Phase oder um einen Meilenstein ? 


                            End If


                        Next

                        '' wenn es sich um eine Vorlage handelt: 

                        'If awinSettings.importTyp = 2 Then
                        '    hproj.farbe = 0
                        '    hproj.Schrift = schriftGroesse
                        '    hproj.Schriftfarbe = schriftfarbe
                        'End If


                        ' jetzt muss das Projekt eingetragen werden 
                        ImportProjekte.Add(hproj)
                        myCollection.Add(hproj.name)

                    End If

                    zeile = ende + 1


                End While





            End With
        Catch ex As Exception
            Throw New Exception("Fehler in Datei BMW Projekt-Inventur " & vbLf & ex.Message & vbLf & pName)
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

        ' diese Datei muss offen sein und das aktive Workbook
        ' wenn nein, dann aktivieren ! 
        Try
            If appInstance.ActiveWorkbook.Name <> bmwExportVorlage Then
                appInstance.Workbooks(bmwExportVorlage).Activate()
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

        Dim indentPhase As Integer = 3
        Dim indentMS As Integer = 6

        ' die erste Phase kann auch Meilensteine haben !
        cphase = hproj.getPhase(1)
        For im = 1 To cphase.CountResults
            zeile = zeile + 1
            cmilestone = cphase.getResult(im)
            startdate = cmilestone.getDate
            If cmilestone.name.StartsWith(cphase.name & "+") Then

                Dim parentName As String = cphase.name & "+"
                curName = ""
                Dim posStart As Integer = parentName.Length

                For posX As Integer = posStart + 1 To cmilestone.name.Length
                    curName = curName & cmilestone.name.Chars(posX)
                Next

                ' hier den Original Name verwenden !? nein, aktuell noch nicht 

            Else
                curName = cmilestone.name
            End If

            CType(ws.Cells(zeile, spalte), Excel.Range).Value = indentMS & curName

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
            curName = indentPhase & cphase.name

            CType(ws.Cells(zeile, spalte), Excel.Range).Value = curName

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

            For im = 1 To cphase.CountResults
                zeile = zeile + 1
                cmilestone = cphase.getResult(im)
                startdate = cmilestone.getDate
                If cmilestone.name.StartsWith(cphase.name & "+") Then

                    Dim parentName As String = cphase.name & "+"
                    curName = ""
                    Dim posStart As Integer = parentName.Length

                    For posX As Integer = posStart + 1 To cmilestone.name.Length
                        curName = curName & cmilestone.name.Chars(posX)
                    Next

                    ' hier den Original Name verwenden !? nein, aktuell noch nicht 

                Else
                    curName = cmilestone.name
                End If

                CType(ws.Cells(zeile, spalte), Excel.Range).Value = indentMS & curName

                If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                    CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = startdate.ToShortDateString
                    CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = startdate.ToShortDateString
                Else
                    CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = "Fehler !"
                    CType(ws.Cells(zeile, spalte).offset(0, 2), Excel.Range).Value = "Fehler !"
                End If
            Next

        Next



    End Sub
End Module
