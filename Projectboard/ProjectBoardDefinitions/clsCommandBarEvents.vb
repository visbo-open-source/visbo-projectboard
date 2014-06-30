Imports Microsoft.Office.Interop
Imports System.Math

Public Class clsCommandBarEvents
    Public WithEvents cmdbars As Microsoft.Office.Core.CommandBars

    ' appInstance.OnKey("{clear}", MsgBox("delelte gedrückt")) eventuell für bestimmte Tastenkombi Entf


    Private Sub cmdbars_OnUpdate() Handles cmdbars.OnUpdate
        'Dim ws As Excel.Worksheet = appInstance.ActiveSheet

        Dim shpelement As Excel.Shape
        Dim tmpShapes As Excel.Shapes
        Dim i As Integer
        Dim pname As String
        Dim hproj As clsProjekt
        Dim tmpShpListe As New Collection
        Dim tmpDelListe As New Collection
        Dim top As Double, left As Double, width As Double, height As Double

        Dim SID As String
        Dim zeile As Integer
        'Dim spalte As Integer
        Dim laengeInMon As Integer
        Dim six As Integer = 1
        Dim anzahlShapes As Integer
        Dim key As String
        Dim updateKennung As Integer = 8
        Dim shapeType As Integer


        If Me.cmdbars.ActiveMenuBar.Index <> 1 Then
            Exit Sub
        End If

        If Not enableOnUpdate Then
            Exit Sub
        End If

        Dim somethingChanged As Boolean = False
        Dim ChartsNeedUpdate As Boolean = False

        Dim selCollection As New Collection


        ' awinSelection enthält alle selektierten Shapes 
        Dim awinSelection As Excel.ShapeRange
        Try
            awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            If awinSelection.Count > 0 Then

                
                ' in selCollection kommen jetzt alle Shapes, die Projekte sind ... 
                For i = 1 To awinSelection.Count
                    ' es dürfen nur solche in die Collection aufgenommen werden, die schon existiert haben; also wenn shpelement.id = hproj.shpuid 
                    shpelement = awinSelection.Item(i)


                    With shpelement
                        Try
                            If .ID = ShowProjekte.getProject(.Name).shpUID Then
                                selCollection.Add(.Name, .Name)
                            End If
                        Catch ex1 As Exception

                        End Try

                    End With

                Next

                ' jetzt heraus finden, ob sich die Selektion der Projekte gegenüber vorher verändert hat 

                If projectSelectionChanged(selCollection) Then

                    selectedProjekte.Clear()

                    For Each tmpName As String In selCollection
                        Try
                            selectedProjekte.Add(ShowProjekte.getProject(tmpName))
                        Catch ex As Exception

                        End Try
                    Next

                    ' wegen Anzeigen der selektierten Projekte in PRCCollection Diagrammen 

                    ChartsNeedUpdate = True


                End If


            End If
        Catch ex As Exception
            awinSelection = Nothing
            Exit Sub
        End Try

        ' jetzt stehen in selectedProjekte alle selektierten Projekte 
        ' in selCollection ist die Liste aller Shapes, die Projekte sind

        anzahlCalls = anzahlCalls + 1


        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False




        With appInstance.Worksheets(arrWsNames(3))
            tmpShapes = .shapes
            'tmpShapes = awinSelection



            If awinSelection.Count > 0 Then

                six = 1
                'in Anzahl shapes ist die Anzahl der selektierten Shapes 
                anzahlShapes = awinSelection.Count


                For six = 1 To anzahlShapes

                    shpelement = awinSelection.Item(six)
                    shapeType = kindOfShape(shpelement)

                    ' Änderung 5.11: prüfung auf hasChart ist notwendig, um zusammengesetztes Projekt-Shape von Chart zu unterscheiden ...
                    ' Änderung 17.11: prüfung auf Connector ist notwendig, um zusammengesetztes Shape von Connector = Phasen-Shape zu unterscheiden

                    If shapeType = PTshty.projektE Or shapeType = PTshty.projektN Then


                        SID = shpelement.ID.ToString


                        laengeInMon = shpelement.Width / boxWidth


                        '
                        ' ist das Shape schon bekannt ? Existiert also schon in der ShowProjekte Liste  
                        '
                        If ShowProjekte.shpListe.ContainsKey(SID) Then

                            ' wenn es nur gerigfügige Änderugen waren, wird das Shape in dieser Funktion 
                            ' wieder auf den alten Zustand eingerüttelt
                            somethingChanged = projectboardShapes.hasAchanged(shpelement)
                            hproj = ShowProjekte.getProjectS(SID)
                            pname = hproj.name

                            If somethingChanged Then
                                Dim movedToNoshow As Boolean = False
                                ' hole das betreffende Projekt

                                zeile = calcYCoordToZeile(shpelement.Top)

                                ' wenn Röntgen Blick an ist: 
                                If roentgenBlick.isOn Then
                                    Call NoshowNeedsofProject(hproj.name)
                                End If

                                If zeile = 0 Or shpelement.Rotation <> 0 Then
                                    ' ins Noshow stecken ... 
                                    movedToNoshow = True
                                    enableOnUpdate = False
                                    Call awinShowNoShowProject(pname:=hproj.name)
                                    enableOnUpdate = True

                                    updateKennung = 3
                                    ChartsNeedUpdate = True

                                ElseIf hproj.Status = ProjektStatus(0) Then
                                    ' nur dann kann verschoben/gedehnt/gestaucht werden - das wird in der Property sync gemacht
                                    ' hier wird das Projekt mit den Shape Werten "synchronisiert" 
                                    projectboardShapes.sync(shpelement, selCollection)
                                    updateKennung = 1
                                    ChartsNeedUpdate = True

                                Else
                                    ' hier werden die verschobenen Shapes wieder zurückgesetzt 
                                    ' Charts müssen nicht aktualisiert werden 
                                    projectboardShapes.sync(shpelement, selCollection)

                                    ChartsNeedUpdate = False

                                End If


                                ' wenn Röntgen Blick an ist und notMovedToNoshow
                                If roentgenBlick.isOn And Not movedToNoshow Then
                                    With roentgenBlick
                                        Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=hproj.name)
                                    End With

                                End If

                                
                            End If

                            Try
                                If selCollection.Contains(pname) Then
                                    selCollection.Remove(pname)
                                End If
                            Catch ex As Exception
                                Call MsgBox("interner Fehler .... (selCollection) " & pname)
                            End Try


                        Else
                            '
                            ' das Shape ist neu dazu gekommen, also kopiert worden und muss in die Liste aufgenommen werden 
                            updateKennung = 2
                            ChartsNeedUpdate = True


                            Dim shpName As String = shpelement.Name
                            Dim zaehler As Integer = 1

                            pname = ""
                            ' Änderung 25.3.14
                            ' ein kopiertes Projekt sollte jetzt in der nächsten Zeile platziert werden 
                            'zeile = findeMagicBoardPosition(selCollection, pname, zeile, spalte, laengeInMon)


                            pname = shpelement.Name & " - Kopie " & zaehler

                            Dim anzahlZeilen As Integer
                            Dim oldproj As clsProjekt
                            Try
                                oldproj = ShowProjekte.getProject(shpName) ' der shpName ist identisch mit dem Projekt-Namen aus dem kopiert wurde

                                ' Änderung 25.3.14 wegen Kopiertes Projekt soll einfach in der nächsten Zeile gezeichnet werden  
                                anzahlZeilen = getNeededSpace(oldproj)
                                zeile = oldproj.tfZeile + 1
                                Dim tmpCollection As New Collection
                                Call moveShapesDown(tmpCollection, zeile, anzahlZeilen, 0) ' Stoppzeile 0: alle Elemente werden verschoben  
                            Catch ex As Exception
                                Throw New ArgumentException("Projekt in OnUpdate nicht gefunden: " & shpName)
                                Exit Sub
                            End Try

                            hproj = New clsProjekt
                            oldproj.CopyTo(hproj)

                            With hproj
                                .name = pname
                                .shpUID = shpelement.ID.ToString
                                '.dauer = laenge
                                .tfZeile = zeile + anzahlZeilen - 1
                                '.tfSpalte = spalte
                                .Status = ProjektStatus(0)
                                ' Änderung 8.11 : ein neues Projekt sollte in der Zukunft angelegt werden, wenn oldproj.startdate in der Vergangenheit liegt - 
                                ' ansonsten ein Termin ein Monat nach oldproj ..

                                If DateDiff(DateInterval.Month, oldproj.startDate, Date.Now) >= 0 Then
                                    ' das Datum von Oldproj liegt in der Vergangenheit 
                                    .startDate = Date.Now.AddMonths(1)
                                Else
                                    ' das Ursprungsprojekt leigt auch noch in der Zukunft - deswegen wird es zeitlich in der Nähe positioniert 
                                    .startDate = oldproj.startDate.AddMonths(1)
                                End If

                                .variantName = ""
                                ' Änderung 19.8 : Bewertungen löschen in dem kopierten Projekt, ausserdem Status auf <geplant> setzen 
                                .clearBewertungen()

                                ' das Budget wird auf Null gesetzt , ebenso die monatlichen Budget-Werte
                                .Erloes = 0
                                Dim budgetvalues() As Double
                                ReDim budgetvalues(.Dauer - 1)
                                .budgetWerte = budgetvalues

                            End With

                            Dim successful As Boolean = False
                            While Not successful
                                Try
                                    key = pname & "#"
                                    AlleProjekte.Add(key, hproj)
                                    successful = True
                                Catch ex As Exception
                                    zaehler = zaehler + 1
                                    pname = shpelement.Name & " - Kopie " & zaehler
                                    hproj.name = pname
                                End Try
                            End While

                            Try
                                If successful Then
                                    ShowProjekte.Add(hproj)
                                    selectedProjekte.Add(hproj)
                                    ChartsNeedUpdate = True
                                End If
                            Catch ex As Exception

                            End Try


                            If roentgenBlick.isOn Then
                                With roentgenBlick
                                    Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
                                End With
                            End If

                            If shapeType = PTshty.projektE Then
                                '
                                ' zusammengesetztes Shape 
                                '
                                ' jetzt wird das kopierte Shape gelöscht 
                                enableOnUpdate = False
                                ' das gibt es noch gar nicht in der Liste, deswegen ist der Befehl an dieser 
                                ' stelle nicht nötig 
                                ' projectboardShapes.remove(shpelement.Name)
                                shpelement.Delete()

                                ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                                ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                                Dim tmpCollection As New Collection
                                Call ZeichneProjektinPlanTafel(tmpCollection, pname, hproj.tfZeile)

                                enableOnUpdate = True


                            Else
                                hproj.CalculateShapeCoord(top, left, width, height)

                                With shpelement
                                    .Name = pname
                                    .TextFrame2.TextRange.Text = pname
                                    .Top = top
                                    .Left = left
                                    .Width = width
                                    .Height = height
                                    .Rotation = 0.0
                                End With

                                projectboardShapes.add(shpelement)

                                ' Änderung 19.8 - falls ein beauftragtes Projekt kopiert wurde, muss es entsprechend Status=<geplant> 
                                ' visualisiert werden 
                                Call defineShapeAppearance(hproj, shpelement)
                            End If



                        End If

                        ' Änderung 8.6.14 hier werden jetzt die Projekt Charts aktualisiert, sofern welche da sind und die Time Machine nicht aktiv ist
                        If Not timeMachineIsOn Then
                            Call aktualisiereCharts(hproj, True)
                        End If



                    ElseIf shapeType = PTshty.phaseE Or shapeType = PTshty.phaseN Then


                        ' es darf nur ein Shape von diesem Typ selektiert sein ... 
                        If awinSelection.Count = 1 Then

                            Dim phaseName As String
                            somethingChanged = projectboardShapes.hasAchanged(shpelement)

                            pname = extractName(shpelement.Name, PTshty.projektN)
                            phaseName = extractName(shpelement.Name, PTshty.phaseN)
                            hproj = ShowProjekte.getProject(pname)

                            If somethingChanged Then

                                projectboardShapes.sync(shpelement, selCollection)

                                If hproj.Status = ProjektStatus(0) Then
                                    ' Charts müssen aktualisiert werden 
                                    updateKennung = 2
                                    ChartsNeedUpdate = True

                                Else
                                    ' Charts müssen nicht aktualisiert werden, da beauftragte Projekte nicht verändert werden können

                                    ChartsNeedUpdate = False

                                End If


                            End If


                            If formPhase Is Nothing Then
                                formPhase = New frmPhaseInformation
                            End If

                            If Not formPhase.Visible Then
                                If formPhase.IsDisposed Then
                                    formPhase = New frmPhaseInformation
                                End If
                            End If

                            'Call updatePhaseInformation(shpelement)
                            Call updatePhaseInformation(hproj, phaseName)

                        Else
                            Call MsgBox("bitte nur ein Element bzw. eine Phase selektieren ...")
                            Call awinDeSelect()
                            Exit For
                        End If





                    ElseIf shapeType = PTshty.milestoneN Or shapeType = PTshty.milestoneE Then

                        ' es darf nur ein Shape von diesem Typ selektiert sein ... 
                        If awinSelection.Count = 1 Then

                            Dim phaseName As String
                            Dim resultName As String
                            somethingChanged = projectboardShapes.hasAchanged(shpelement)

                            pname = extractName(shpelement.Name, PTshty.projektN)
                            phaseName = extractName(shpelement.Name, PTshty.phaseN)
                            resultName = shpelement.Title
                            hproj = ShowProjekte.getProject(pname)

                            If somethingChanged Then

                                If hproj.Status = ProjektStatus(0) Then
                                    ' nur dann kann verschoben werden - das wird in der Property sync gemacht
                                    ' hier wird das Projekt mit den Shape Werten "synchronisiert" 
                                    projectboardShapes.sync(shpelement, selCollection)
                                    ChartsNeedUpdate = True

                                Else
                                    ' hier werden die verschobenen Shapes wieder zurückgesetzt 
                                    ' Charts müssen nicht aktualisiert werden 
                                    projectboardShapes.sync(shpelement, selCollection)
                                    ChartsNeedUpdate = False

                                End If


                            End If

                            If formMilestone Is Nothing Then
                                formMilestone = New frmMilestoneInformation
                            End If

                            If Not formMilestone.Visible Then
                                If formMilestone.IsDisposed Then
                                    formMilestone = New frmMilestoneInformation
                                End If
                            End If

                            'Call updateMilestoneInformation(shpelement)
                            Call updateMilestoneInformation(hproj, phaseName, resultName)

                        Else

                            Call MsgBox("bitte nur ein Element bzw. einen Meilenstein selektieren ...")
                            Call awinDeSelect()
                            Exit For

                        End If
                        

                    ElseIf shapeType = PTshty.status Then
                        ' war vorher: 
                        'ElseIf shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval Then
                        'Print("not recognized" & shpelement.autoType & "," & shpelement.Name)

                        ' es darf nur ein Shape von diesem Typ selektiert sein ... 
                        If awinSelection.Count = 1 Then

                            If formStatus Is Nothing Then
                                formStatus = New frmStatusInformation
                            End If

                            If Not formStatus.Visible Then
                                If formStatus.IsDisposed Then
                                    formStatus = New frmStatusInformation
                                End If
                            End If

                            Call updateStatusInformation(shpelement)

                        Else

                            Call MsgBox("bitte nur ein Element bzw. einen Status selektieren ...")
                            Call awinDeSelect()
                            Exit For

                        End If


                    End If

                Next ' six
            End If

            ' jetzt muss noch geprüft werden, ob Shapes gelöscht wurden
            ' also: existieren in Projekte Einträge, die keine Entsprechung in Shapes Auflistung haben 

            For Each kvp As KeyValuePair(Of String, String) In ShowProjekte.shpListe

                Try

                    shpelement = tmpShapes.Item(kvp.Value)

                Catch ex As Exception
                    tmpDelListe.Add(kvp.Value)
                End Try



            Next

            If tmpDelListe.Count > 0 Then

                ChartsNeedUpdate = True
                updateKennung = 3

                For i = 1 To tmpDelListe.Count
                    pname = tmpDelListe.Item(i)

                    If roentgenBlick.isOn Then
                        Call NoshowNeedsofProject(pname)
                        somethingChanged = False
                    End If


                    ' Änderung 18.6.2013: notwendig, weil durch Drücken der Del Taste das Shape gelöscht wurde; 
                    ' anschließendes sofortiges Eintragen eines neuen Projektes hat dann zu einem Fehler geführt 
                    Try
                        hproj = ShowProjekte.getProject(pname)
                        key = calcProjektKey(hproj)
                        Try
                            ShowProjekte.Remove(pname)
                            AlleProjekte.Remove(key)
                            DeletedProjekte.Add(hproj)
                        Catch ex1 As Exception

                        End Try
                    Catch ex As Exception

                    End Try


                Next
                tmpDelListe.Clear()

            End If


            ' wenn die Charts geupdated werden müssen ...  

            If ChartsNeedUpdate Then
                enableOnUpdate = False
                Call awinNeuZeichnenDiagramme(updateKennung)
                enableOnUpdate = True
            End If


        End With

        appInstance.EnableEvents = formerEE

    End Sub


    Private Function projectSelectionChanged(ByVal selCollection As Collection) As Boolean

        Dim tmpVar As Boolean = False



        If selCollection.Count <> selectedProjekte.Count Then
            tmpVar = True
        Else
            If selectedProjekte.Count = 0 Then
                tmpVar = False
            Else
                For Each kvp As KeyValuePair(Of String, clsProjekt) In selectedProjekte.Liste
                    If Not selCollection.Contains(kvp.Key) Then
                        tmpVar = True
                    End If
                Next
            End If
            

        End If
        projectSelectionChanged = tmpVar
    End Function


    'Private Sub cmdbars_OnUpdate() Handles cmdbars.OnUpdate
    '    'Dim ws As Excel.Worksheet = appInstance.ActiveSheet

    '    Dim shpelement As Excel.Shape
    '    Dim tmpShapes As Excel.Shapes
    '    Dim i As Integer
    '    Dim pname As String
    '    Dim hproj As clsProjekt
    '    Dim tmpShpListe As New Collection
    '    Dim tmpDelListe As New Collection
    '    Dim top As Double, left As Double, width As Double, height As Double

    '    Dim SID As String
    '    Dim zeile As Integer
    '    Dim spalte As Integer
    '    Dim laengeInMon As Integer
    '    Dim six As Integer = 1
    '    Dim anzahlShapes As Integer
    '    Dim key As String
    '    Dim updateKennung As Integer = 8
    '    Dim shapeType As Integer


    '    If Me.cmdbars.ActiveMenuBar.Index <> 1 Then
    '        Exit Sub
    '    End If

    '    If Not enableOnUpdate Then
    '        Exit Sub
    '    End If

    '    Dim somethingChanged As Boolean = False
    '    Dim ChartsNeedUpdate As Boolean = False

    '    Dim selCollection As New Collection


    '    ' awinSelection enthält alle selektierten Shapes 
    '    Dim awinSelection As Excel.ShapeRange
    '    Try
    '        awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
    '        If awinSelection.Count > 0 Then


    '            ' in selCollection kommen jetzt alle Shapes, die Projekte sind ... 
    '            For i = 1 To awinSelection.Count
    '                ' es dürfen nur solche in die Collection aufgenommen werden, die schon existiert haben; also wenn shpelement.id = hproj.shpuid 
    '                shpelement = awinSelection.Item(i)


    '                With shpelement
    '                    Try
    '                        If .ID = ShowProjekte.getProject(.Name).shpUID Then
    '                            selCollection.Add(.Name, .Name)
    '                        End If
    '                    Catch ex1 As Exception

    '                    End Try

    '                End With

    '            Next

    '            ' jetzt heraus finden, ob sich die Selektion der Projekte gegenüber vorher verändert hat 

    '            If projectSelectionChanged(selCollection) Then

    '                selectedProjekte.Clear()

    '                For Each tmpName As String In selCollection
    '                    Try
    '                        selectedProjekte.Add(ShowProjekte.getProject(tmpName))
    '                    Catch ex As Exception

    '                    End Try
    '                Next

    '                ' wegen Anzeigen der selektierten Projekte in PRCCollection Diagrammen 

    '                ChartsNeedUpdate = True


    '            End If


    '        End If
    '    Catch ex As Exception
    '        awinSelection = Nothing
    '        Exit Sub
    '    End Try

    '    ' jetzt stehen in selectedProjekte alle selektierten Projekte 
    '    ' in selCollection ist die Liste aller Shapes, die Projekte sind

    '    anzahlCalls = anzahlCalls + 1


    '    Dim formerEE As Boolean = appInstance.EnableEvents
    '    appInstance.EnableEvents = False




    '    With appInstance.Worksheets(arrWsNames(3))
    '        tmpShapes = .shapes
    '        'tmpShapes = awinSelection



    '        If awinSelection.Count > 0 Then

    '            six = 1
    '            'in Anzahl shapes ist die Anzahl der selektierten Shapes 
    '            anzahlShapes = awinSelection.Count


    '            For six = 1 To anzahlShapes

    '                shpelement = awinSelection.Item(six)
    '                shapeType = kindOfShape(shpelement)

    '                ' Änderung 5.11: prüfung auf hasChart ist notwendig, um zusammengesetztes Projekt-Shape von Chart zu unterscheiden ...
    '                ' Änderung 17.11: prüfung auf Connector ist notwendig, um zusammengesetztes Shape von Connector = Phasen-Shape zu unterscheiden

    '                If shapeType = PTshty.projektE Or shapeType = PTshty.projektN Then


    '                    'If Not shpelement.AlternativeText = "Phase" And _
    '                    '    (shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle Or _
    '                    '    (shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMixed And Not shpelement.HasChart _
    '                    '     And Not shpelement.Connector = Microsoft.Office.Core.MsoTriState.msoTrue)) Then

    '                    SID = shpelement.ID.ToString

    '                    zeile = 1 + (shpelement.Top - topOfMagicBoard) / boxHeight
    '                    'Dim precisevalue As Double = shpelement.Left / boxWidth

    '                    spalte = System.Math.Truncate(shpelement.Left / boxWidth) + 1


    '                    laengeInMon = shpelement.Width / boxWidth


    '                    '
    '                    ' ist das Shape schon bekannt ? Existiert also schon in der ShowProjekte Liste  
    '                    '
    '                    If ShowProjekte.shpListe.ContainsKey(SID) Then

    '                        Dim changeWasValid As Boolean = False
    '                        somethingChanged = False


    '                        hproj = ShowProjekte.getProjectS(SID)

    '                        ' das muss nur visuell korrigiert werden , ansonsten nicht in den Projekt-Definitionen 

    '                        ' nur korrigieren, wenn es nicht ein zusammengesetztes Shape ist 
    '                        If shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle Then
    '                            If shpelement.Height <> boxHeight * 0.8 Then
    '                                shpelement.Height = boxHeight * 0.8
    '                            End If
    '                        End If


    '                        ' just in case - falls jdn rotiert ...
    '                        Try
    '                            shpelement.Rotation = 0.0
    '                        Catch ex As Exception

    '                        End Try


    '                        pname = hproj.name

    '                        With hproj
    '                            If Abs(shpelement.Width - (hproj.dauerInDays / 365) * boxWidth * 12) > 0.02 * ((hproj.dauerInDays / 365) * boxWidth * 12) And _
    '                                Not ProjectBoardDefinitions.My.Settings.drawPhases = True Then

    '                                shpelement.Width = (hproj.dauerInDays / 365) * boxWidth * 12
    '                                somethingChanged = True

    '                            End If

    '                            If zeile <> .tfZeile Or spalte <> .tfspalte Then
    '                                somethingChanged = True
    '                                changeWasValid = False

    '                                ' hier muss geprüft werden, ob das Shape nach oben , also in den NoShow Bereich geschoben wurde 
    '                                If zeile = 0 Then
    '                                    changeWasValid = True
    '                                End If

    '                                If roentgenBlick.isOn Then
    '                                    Call NoshowNeedsofProject(pname)
    '                                End If


    '                                If spalte <> .tfspalte And zeile > 0 Then
    '                                    ' es muss geprüft werden, ob das überhaupt zulässig ist ... 

    '                                    Try
    '                                        .startDate = .startDate.AddMonths(spalte - .tfspalte)
    '                                        changeWasValid = True
    '                                    Catch ex As Exception
    '                                        spalte = .tfspalte
    '                                    End Try


    '                                End If
    '                            End If
    '                        End With



    '                        ' nur wenn sich was verändert hat , muss was getan werden
    '                        If somethingChanged Then

    '                            ' Änderung 12.10.13 wenn das Projekt in Zeile 0 geschoben wird , dann wird es in das NoShow verschoben ...
    '                            If zeile = 0 Then
    '                                ' dann soll das Projekt jetzt in den NoShow Bereich verschoben werden 
    '                                enableOnUpdate = False
    '                                Call awinShowNoShowProject(pname:=pname)
    '                                enableOnUpdate = True


    '                            Else
    '                                zeile = findeMagicBoardPosition(selCollection, pname, zeile, spalte, laengeInMon)

    '                                ' jetzt die Informationen in der Projektliste entsprechend anpassen 

    '                                With hproj

    '                                    .tfZeile = zeile

    '                                End With

    '                                ' Behandlung Multi-Shapes


    '                                If shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMixed Then
    '                                    '
    '                                    ' zusammengesetztes Shape 
    '                                    '

    '                                    ' ab hier: das wird nicht mehr benötigt : die Drawphases wird es so nicht mehr geben ....
    '                                    'Dim phasenName As String
    '                                    'Dim phaseShapeName As String
    '                                    'Dim phasenShpElement As Excel.Shape

    '                                    '' jetzt jedes Shape entsprechend anpassen 
    '                                    'For i = 1 To hproj.CountPhases
    '                                    '    phasenName = hproj.getPhase(i).name
    '                                    '    phaseShapeName = hproj.name & "#" & phasenName & "#" & i.ToString

    '                                    '    Try
    '                                    '        phasenShpElement = shpelement.groupItem(phaseShapeName)
    '                                    '        hproj.CalculateShapeCoord(i, top, left, width, height)

    '                                    '        With phasenShpElement
    '                                    '            .Top = top
    '                                    '            .Left = left
    '                                    '            .Width = width
    '                                    '            .Height = height
    '                                    '            .Rotation = 0.0
    '                                    '        End With

    '                                    '    Catch ex As Exception

    '                                    '    End Try

    '                                    'Next

    '                                    '' jetzt noch das Gesamt Shape, das zusammengesetzte ausrichten 
    '                                    'hproj.CalculateShapeCoord(top, left, width, height)

    '                                    'With shpelement
    '                                    '    .Top = top
    '                                    '    .Left = left
    '                                    '    .Width = width
    '                                    '    .Height = height
    '                                    '    .Rotation = 0.0
    '                                    'End With


    '                                Else
    '                                    hproj.CalculateShapeCoord(top, left, width, height)

    '                                    With shpelement
    '                                        .Top = top
    '                                        .Left = left
    '                                        .Width = width
    '                                        .Height = height
    '                                        .Rotation = 0.0
    '                                        .TextFrame2.TextRange.Text = pname
    '                                    End With

    '                                    '' Änderung 19.8 - falls ein beauftragtes Projekt kopiert wurde, muss es entsprechend Status=<geplant> 
    '                                    '' visualisiert werden 
    '                                    'Call defineShapeAppearance(hproj, shpelement)
    '                                End If


    '                                ' Ende Behandlung 

    '                                If roentgenBlick.isOn Then
    '                                    With roentgenBlick
    '                                        Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
    '                                    End With

    '                                End If
    '                            End If

    '                            ChartsNeedUpdate = ChartsNeedUpdate Or changeWasValid
    '                        Else

    '                            If shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMixed Then
    '                                ' hier dann was machen, wenn auch angeboten wird, die Shapes komplett zu zeichnen 
    '                            Else
    '                                With hproj

    '                                    .tfZeile = zeile
    '                                    '.tfSpalte = spalte
    '                                    .CalculateShapeCoord(top, left, width, height)

    '                                End With

    '                                ' das Shape in das Raster schnappen lassen ....  
    '                                With shpelement
    '                                    .Top = top
    '                                    .Left = left
    '                                    .Width = width
    '                                    .Height = height
    '                                    .Rotation = 0.0

    '                                End With
    '                            End If

    '                        End If

    '                        Try
    '                            If selCollection.Contains(pname) Then
    '                                selCollection.Remove(pname)
    '                            End If
    '                        Catch ex As Exception

    '                        End Try


    '                    Else
    '                        '
    '                        ' das Shape ist neu dazu gekommen, also kopiert worden und muss in die Liste aufgenommen werden 
    '                        updateKennung = 2
    '                        ChartsNeedUpdate = True


    '                        Dim shpName As String = shpelement.Name
    '                        Dim zaehler As Integer = 1

    '                        pname = ""
    '                        ' Änderung 25.3.14
    '                        ' ein kopiertes Projekt sollte jetzt in der nächsten Zeile platziert werden 
    '                        'zeile = findeMagicBoardPosition(selCollection, pname, zeile, spalte, laengeInMon)


    '                        pname = shpelement.Name & " - Kopie " & zaehler

    '                        Dim anzahlZeilen As Integer
    '                        Dim oldproj As clsProjekt
    '                        Try
    '                            oldproj = ShowProjekte.getProject(shpName) ' der shpName ist identisch mit dem Projekt-Namen aus dem kopiert wurde

    '                            ' Änderung 25.3.14 wegen Kopiertes Projekt soll einfach in der nächsten Zeile gezeichnet werden  
    '                            anzahlZeilen = getNeededSpace(oldproj)
    '                            zeile = oldproj.tfZeile + 1
    '                            Call moveShapesDown(zeile, anzahlZeilen)
    '                        Catch ex As Exception
    '                            Throw New ArgumentException("Projekt in OnUpdate nicht gefunden: " & shpName)
    '                            Exit Sub
    '                        End Try

    '                        hproj = New clsProjekt
    '                        oldproj.CopyTo(hproj, False)

    '                        With hproj
    '                            .name = pname
    '                            .shpUID = shpelement.ID.ToString
    '                            '.dauer = laenge
    '                            .tfZeile = zeile + anzahlZeilen - 1
    '                            '.tfSpalte = spalte
    '                            .Status = ProjektStatus(0)
    '                            ' Änderung 8.11 : ein neues Projekt sollte in der Zukunft angelegt werden, wenn oldproj.startdate in der Vergangenheit liegt - 
    '                            ' ansonsten ein Termin ein Monat nach oldproj ..

    '                            If DateDiff(DateInterval.Month, oldproj.startDate, Date.Now) >= 0 Then
    '                                ' das Datum von Oldproj liegt in der Vergangenheit 
    '                                .startDate = Date.Now.AddMonths(1)
    '                            Else
    '                                ' das Ursprungsprojekt leigt auch noch in der Zukunft - deswegen wird es zeitlich in der Nähe positioniert 
    '                                .startDate = oldproj.startDate.AddMonths(1)
    '                            End If

    '                            .variantName = ""
    '                            ' Änderung 19.8 : Bewertungen löschen in dem kopierten Projekt, ausserdem Status auf <geplant> setzen 
    '                            .clearBewertungen()

    '                            ' das Budget wird auf Null gesetzt , ebenso die monatlichen Budget-Werte
    '                            .Erloes = 0
    '                            Dim budgetvalues() As Double
    '                            ReDim budgetvalues(.Dauer - 1)
    '                            .budgetWerte = budgetvalues

    '                        End With

    '                        Dim successful As Boolean = False
    '                        While Not successful
    '                            Try
    '                                key = pname & "#"
    '                                AlleProjekte.Add(key, hproj)
    '                                successful = True
    '                            Catch ex As Exception
    '                                zaehler = zaehler + 1
    '                                pname = shpelement.Name & " - Kopie " & zaehler
    '                                hproj.name = pname
    '                            End Try
    '                        End While

    '                        Try
    '                            If successful Then
    '                                ShowProjekte.Add(hproj)
    '                                selectedProjekte.Add(hproj)
    '                                ChartsNeedUpdate = True
    '                            End If
    '                        Catch ex As Exception

    '                        End Try


    '                        If roentgenBlick.isOn Then
    '                            With roentgenBlick
    '                                Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
    '                            End With
    '                        End If

    '                        If shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMixed Then
    '                            '
    '                            ' zusammengesetztes Shape 
    '                            '
    '                            ' jetzt wird das kopierte Shape gelöscht 
    '                            enableOnUpdate = False
    '                            ' das gibt es noch gar nicht in der Liste, deswegen ist der Befehl an dieser 
    '                            ' stelle nicht nötig 
    '                            ' projectboardShapes.remove(shpelement.Name)
    '                            shpelement.Delete()
    '                            Call ZeichneProjektinPlanTafel(pname, hproj.tfZeile)
    '                            enableOnUpdate = True


    '                        Else
    '                            hproj.CalculateShapeCoord(top, left, width, height)

    '                            With shpelement
    '                                .Name = pname
    '                                .TextFrame2.TextRange.Text = pname
    '                                .Top = top
    '                                .Left = left
    '                                .Width = width
    '                                .Height = height
    '                                .Rotation = 0.0
    '                            End With

    '                            ' Änderung 19.8 - falls ein beauftragtes Projekt kopiert wurde, muss es entsprechend Status=<geplant> 
    '                            ' visualisiert werden 
    '                            Call defineShapeAppearance(hproj, shpelement)
    '                        End If



    '                    End If
    '                ElseIf shapeType = PTshty.phaseE Or shapeType = PTshty.phaseN Then
    '                    ' war vorher: 
    '                    'ElseIf shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeMixed And _
    '                    '       shpelement.Connector = Microsoft.Office.Core.MsoTriState.msoTrue Or _
    '                    '       shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRoundedRectangle And _
    '                    '       shpelement.AlternativeText = "Phase" Then


    '                    If formPhase Is Nothing Then
    '                        formPhase = New frmPhaseInformation
    '                    End If

    '                    If Not formPhase.Visible Then
    '                        If formPhase.IsDisposed Then
    '                            formPhase = New frmPhaseInformation
    '                        End If
    '                    End If


    '                    Call updatePhaseStartDuration(shpelement)
    '                    Call updatePhaseInformation(shpelement)

    '                ElseIf shapeType = PTshty.milestoneN Or shapeType = PTshty.milestoneE Then
    '                    ' war vorher: 
    '                    'ElseIf shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeDiamond Then
    '                    '    'Print("not recognized" & shpelement.autoType & "," & shpelement.Name)

    '                    If formMilestone Is Nothing Then
    '                        formMilestone = New frmMilestoneInformation
    '                    End If

    '                    If Not formMilestone.Visible Then
    '                        If formMilestone.IsDisposed Then
    '                            formMilestone = New frmMilestoneInformation
    '                        End If
    '                    End If

    '                    Call updateMilestoneInformation(shpelement)

    '                ElseIf shapeType = PTshty.status Then
    '                    ' war vorher: 
    '                    'ElseIf shpelement.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval Then
    '                    'Print("not recognized" & shpelement.autoType & "," & shpelement.Name)

    '                    If formStatus Is Nothing Then
    '                        formStatus = New frmStatusInformation
    '                    End If

    '                    If Not formStatus.Visible Then
    '                        If formStatus.IsDisposed Then
    '                            formStatus = New frmStatusInformation
    '                        End If
    '                    End If

    '                    Call updateStatusInformation(shpelement)

    '                End If

    '            Next ' six
    '        End If

    '        ' jetzt muss noch geprüft werden, ob Shapes gelöscht wurden
    '        ' also: existieren in Projekte Einträge, die keine Entsprechung in Shapes Auflistung haben 

    '        For Each kvp As KeyValuePair(Of String, String) In ShowProjekte.shpListe

    '            Try

    '                shpelement = tmpShapes.Item(kvp.Value)

    '            Catch ex As Exception
    '                tmpDelListe.Add(kvp.Value)
    '            End Try



    '        Next

    '        If tmpDelListe.Count > 0 Then

    '            ChartsNeedUpdate = True
    '            updateKennung = 2

    '            For i = 1 To tmpDelListe.Count
    '                pname = tmpDelListe.Item(i)

    '                If roentgenBlick.isOn Then
    '                    Call NoshowNeedsofProject(pname)
    '                    somethingChanged = False
    '                End If


    '                ' Änderung 18.6.2013: notwendig, weil durch Drücken der Del Taste das Shape gelöscht wurde; 
    '                ' anschließendes sofortiges Eintragen eines neuen Projektes hat dann zu einem Fehler geführt 
    '                Try
    '                    hproj = ShowProjekte.getProject(pname)
    '                    key = hproj.name & "#" & hproj.variantName
    '                    Try
    '                        ShowProjekte.Remove(pname)
    '                        AlleProjekte.Remove(key)
    '                        DeletedProjekte.Add(hproj)
    '                    Catch ex1 As Exception

    '                    End Try
    '                Catch ex As Exception

    '                End Try


    '            Next
    '            tmpDelListe.Clear()

    '        End If


    '        ' wenn die Charts geupdated werden müssen ...  

    '        If ChartsNeedUpdate Then
    '            enableOnUpdate = False
    '            Call awinNeuZeichnenDiagramme(updateKennung)
    '            enableOnUpdate = True
    '        End If


    '    End With

    '    appInstance.EnableEvents = formerEE

    'End Sub

End Class
