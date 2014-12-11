Imports ProjectBoardDefinitions
Imports Excel = Microsoft.Office.Interop.Excel
Module BMWItOModul
    ''' <summary>
    ''' speziell auf BMW Mpp Anforderungen angepasstes BMW Import File
    ''' Status Dezember 2014/Jan 2015
    ''' </summary>
    ''' <param name="myCollection">gibt die Namen der importierten Fahrzeug Projekt zurück</param>
    ''' <remarks></remarks>
    Public Sub bmwImportProjekteITO15(ByRef myCollection As Collection)

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
        Dim aktuelleZeile As String
        Dim nameSopTyp As String = " "
        Dim nameBU As String = ""
        
        Dim startDate As Date, endDate As Date
        Dim startoffset As Long, duration As Long
        Dim vorlagenName As String

        Dim itemName As String
        Dim zufall As New Random(10)
        Dim protocolColumn As Integer = 20
        

        Dim milestoneIX As Integer = MilestoneDefinitions.Count + 1
        Dim phaseIX As Integer = PhaseDefinitions.Count + 1

        ' Vorbedingung: das Excel File. das importiert werden soll , ist bereits geöffnet 

        zeile = 2
        spalte = 1
        geleseneProjekte = 0


        Try
            
            Dim activeWSListe As Excel.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet, _
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            With activeWSListe

                Dim tstStr As String
                Try
                    tstStr = CStr(CType(activeWSListe.Cells(zeile, 1), Excel.Range).Value)
                    projektFarbe = CType(activeWSListe.Cells(zeile, 1), Excel.Range).Interior.Color
                Catch ex As Exception
                    projektFarbe = CType(activeWSListe.Cells(zeile, 1), Excel.Range).Interior.ColorIndex
                End Try


                lastRow = System.Math.Max(CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row, _
                                          CType(.Cells(2000, 2), Global.Microsoft.Office.Interop.Excel.Range).End(Excel.XlDirection.xlUp).Row)

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
                    aktuelleZeile = CStr(CType(activeWSListe.Cells(zeile, 1), Excel.Range).Value).Trim
                    startDate = CDate(CType(activeWSListe.Cells(zeile, 2), Excel.Range).Value)
                    endDate = CDate(CType(activeWSListe.Cells(zeile, 3), Excel.Range).Value)
                    


                    duration = DateDiff(DateInterval.Day, startDate, endDate) + 1
                    If duration < 0 Then
                        startDate = endDate
                        duration = -1 * duration
                        endDate = startDate.AddDays(duration)
                    End If

                    tmpStr = aktuelleZeile.Trim.Split(New Char() {CChar("["), CChar("]")}, 5)
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


                    If tmpStr(0).Contains("eA") Then
                        vorlagenName = "enge Ableitung"
                    ElseIf tmpStr(0).Contains("wA") Then
                        vorlagenName = "weite Ableitung"
                    ElseIf tmpStr(0).Contains("E") Then
                        vorlagenName = "Erstanläufer"
                    Else
                        vorlagenName = "Erstanläufer"
                    End If



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
                        hproj.ampelStatus = PTfarbe.none
                        hproj.leadPerson = ""
                        'hproj.ampelStatus = farbKennung
                        'hproj.leadPerson = responsible

                    Catch ex As Exception
                        Throw New Exception("es gibt keine entsprechende Vorlage ..  " & vbLf & ex.Message)
                    End Try


                    Try

                        hproj.name = pName
                        hproj.startDate = startDate

                        If DateDiff(DateInterval.Month, startDate, Date.Now) <= 0 Then
                            hproj.Status = ProjektStatus(0)
                            hproj.earliestStartDate = hproj.startDate.AddMonths(hproj.earliestStart)
                            hproj.latestStartDate = hproj.startDate.AddMonths(hproj.latestStart)
                        Else
                            hproj.Status = ProjektStatus(1)
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

                    For curZeile = anfang To ende

                        Try
                            itemName = CStr(CType(.Cells(curZeile, spalte), Excel.Range).Value)
                            ' jetzt prüfen, ob es sich um ein grundsätzlich zu ignorierendes Element handelt .. 
                            'If itemName.Trim = "Projektphasen" Then
                            '    CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                            '                "Phase wird ignoriert: " & itemName.Trim
                            '    ok = False
                            'Else
                            '    ok = True
                            'End If

                        Catch ex As Exception
                            itemName = ""
                            ok = False
                        End Try

                        If ok Then

                            pStartDate = CDate(CType(.Cells(curZeile, spalte + 1), Excel.Range).Value)
                            pEndDate = CDate(CType(.Cells(curZeile, spalte + 2), Excel.Range).Value)
                            startoffset = DateDiff(DateInterval.Day, hproj.startDate, pStartDate)
                            duration = DateDiff(DateInterval.Day, pStartDate, pEndDate) + 1

                            Dim realName As String

                            If duration > 1 Then
                                ' es handelt sich um eine Phase 
                                'phaseName = itemName

                                ' erstmal prüfen, ob es sich um eine "Phasen Dopplung" handelt - dann soll das Element ignoriert werden 
                                If pHierarchy.dopplung(itemName) Then
                                    ' nichts tun - das Element soll ignoriert werden
                                    CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                            "Ignoriert wegen Dopplung: " & itemName.Trim

                                Else
                                    Dim indentLevel As Integer
                                    ' bestimme den Indent-Level , damit die Hierarchie
                                    indentLevel = pHierarchy.getLevel(itemName)

                                    Dim parentPhaseName As String = pHierarchy.getPhaseBeforeLevel(indentLevel).name

                                    ' jetzt den tatsächlichen Namen bestimmen , ggf wird dazu der Parent Phase Name benötigt 

                                    Try
                                        realName = phaseMappings.mapToRealName(parentPhaseName, itemName)
                                    Catch ex As Exception
                                        realName = itemName.Trim
                                    End Try


                                    If realName.Trim <> itemName.Trim Then
                                        CType(activeWSListe.Cells(curZeile, protocolColumn), Excel.Range).Value = _
                                                itemName.Trim & " --> " & realName.Trim
                                    End If

                                    cphase = New clsPhase(parent:=hproj)
                                    cphase.name = realName

                                    If PhaseDefinitions.Contains(realName) Then
                                        ' nichts tun 
                                    Else
                                        ' in die Phase-Definitions aufnehmen 

                                        Dim hphase As clsPhasenDefinition
                                        hphase = New clsPhasenDefinition

                                        hphase.farbe = CLng(CType(.Cells(curZeile, 1), Excel.Range).Interior.Color)
                                        hphase.name = realName
                                        hphase.UID = phaseIX
                                        phaseIX = phaseIX + 1

                                        Try
                                            PhaseDefinitions.Add(hphase)
                                        Catch ex As Exception

                                        End Try

                                    End If

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
                                        realName = milestoneMappings.mapToRealName(parentPhaseName, itemName)
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
                                            .shortName = ""
                                            .darstellungsKlasse = ""
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




                            End If




                            ' handelt es sich um eine Phase oder um einen Meilenstein ? 


                        End If


                    Next


                    ' jetzt muss das Projekt eingetragen werden 
                    ImportProjekte.Add(hproj)
                    myCollection.Add(hproj.name)


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
                If cmilestone.name.Contains(cphase.name & "+") Then
                    ' hier den Original Name verwenden !? 
                    ' und den VISBO - Name in die Spalte "Visbo" schreben !? 
                Else
                    curName = cmilestone.name
                End If
                If DateDiff(DateInterval.Day, StartofCalendar, startdate) > 0 Then
                    CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = startdate.ToShortDateString
                Else
                    CType(ws.Cells(zeile, spalte).offset(0, 1), Excel.Range).Value = "Fehler !"
                End If
            Next

        Next



    End Sub
End Module
