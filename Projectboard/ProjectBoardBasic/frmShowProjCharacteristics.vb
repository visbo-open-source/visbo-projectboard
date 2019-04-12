Imports ProjectBoardDefinitions
Imports System.Math
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms


Public Class frmShowProjCharacteristics
    Private nrSnapshots As Integer
    Private valueBeauftragung As Integer
    Private minmaxScales(1, 6) As Double
    Private necessary(6) As Boolean
    Private hproj As clsProjekt
    Private showAll As Boolean = False
    Private phaseList As Collection
    Private milestoneList As Collection
    Private typCollection As New Collection
    Private lastAmpel As Integer



    Private Sub timeSlider_Scroll(sender As Object, e As EventArgs) Handles timeSlider.Scroll


        hproj = projekthistorie.ElementAt(nrSnapshots - timeSlider.Value)




        With hproj


            If timeSlider.Value = 0 Then
                snapshotDate.Text = "Aktueller Stand: " & .timeStamp.ToString
            ElseIf timeSlider.Value = valueBeauftragung Then
                snapshotDate.Text = "Beauftragung: " & .timeStamp.ToString
            Else
                snapshotDate.Text = .timeStamp.ToString
            End If


        End With




    End Sub

    Private Sub uebernehmenProjekt_Click(sender As Object, e As EventArgs)
        Call MsgBox(" noch nicht implementiert")
    End Sub

    Private Sub frmShowProjCharacteristics_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.timeMachine, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.timeMachine, PTpinfo.left) = Me.Left

        Dim vglName As String = projekthistorie.Last.name.Trim
        ' Änderung 140717
        'hproj = ShowProjekte.getProject(vglName)

        ' Änderung-Ergänzung 140717 tk: Neuzeichnen des Shapes inkl der gezeigten Phasen/Meilensteine
        hproj = projekthistorie.Last


        Call clearProjektinPlantafel(hproj.name)
        ' jetzt muss das aktuelle Shape rausgenommen werden und mit dem TimeSlider Shape ersetzt werden

        Try
            ShowProjekte.Remove(hproj.name)
            ShowProjekte.Add(hproj)

        Catch ex As Exception

        End Try

        Dim tmpCollection As New Collection
        Call ZeichneProjektinPlanTafel(noCollection:=tmpCollection, pname:=hproj.name, tryzeile:=hproj.tfZeile, _
                                       drawPhaseList:=phaseList, drawMilestoneList:=milestoneList)


        ' Ende Änderung-Ergänzung 140717


        'Call aktualisierePMSForms(hproj)
        Call aktualisiereCharts(hproj, False)
        timeMachineIsOn = False

    End Sub



    Private Sub frmShowProjCharacteristics_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Dim pShape As Excel.Shape

        Me.Top = CInt(frmCoord(PTfrm.timeMachine, PTpinfo.top))
        Me.Left = CInt(frmCoord(PTfrm.timeMachine, PTpinfo.left))


        ' erst mal nur wenig von der Time-Machine anzeigen ...
        With Me
            .compareCurrent.Visible = False
            .compareBeauftragung.Visible = False
            .movetoBeauftragung.Visible = False
            .movetoNext.Visible = False
            .movetoPrevious.Visible = False
            .typSelection.Visible = False
            .Label3.Visible = False
            .showMore.Text = "more ..."
            .Height = 190              ' ur: 28.05.2014
        End With
        showAll = False

        lastAmpel = projekthistorie.Last.ampelStatus

        ' Initialisieren des ComboBox Feldes 
        With typSelection
            .Items.Add(" ")
            .Items.Add("Personalkosten")
            .Items.Add("Sonstige Kosten")
            .Items.Add("Budget")
            .Items.Add("Ergebnis")
            .Items.Add("Strategie und Risiko")
            .Items.Add("Termine")
            .Items.Add("Projekt-Ampel")
            .Items.Add("Meilenstein-Ampeln")
            .Items.Add("Phasen")
            .SelectedIndex = 0
        End With

        timeMachineIsOn = True
        nrSnapshots = projekthistorie.Count - 1
        ' hier müssen einmalig die Max-Scale Werte bestimmt werden und ausserdem die gezeigten Charts gleich entsprechend 
        ' mit dem Max Scale angepasst werden 
        Dim currentvalue As Double
        Dim gesamtKosten As Double, persKosten As Double, sonstKosten As Double, erloes As Double
        Dim i As Integer
        Dim anzDiagrams As Integer
        Dim chtobj As Excel.ChartObject

        Dim vglName As String = projekthistorie.Last.name.Trim

        ' Änderung 140717
        ' jetzt werden die Settings für das Neuzeichnen der Shapes gesetzt 
        'pShape = ShowProjekte.getShape(vglName)
        'typCollection.Add(CInt(PTshty.phaseN).ToString, CInt(PTshty.phaseN).ToString)
        'typCollection.Add(CInt(PTshty.phaseE).ToString, CInt(PTshty.phaseE).ToString)
        'phaseList = projectboardShapes.getAllChildswithType(pShape, typCollection)
        phaseList = projectboardShapes.getPhaseList(vglName)

        'typCollection.Clear()
        'typCollection.Add(CInt(PTshty.milestoneN).ToString, CInt(PTshty.milestoneN).ToString)
        'typCollection.Add(CInt(PTshty.milestoneE).ToString, CInt(PTshty.milestoneE).ToString)
        'milestoneList = projectboardShapes.getAllChildswithType(pShape, typCollection)
        milestoneList = projectboardShapes.getMilestoneList(vglName)

        ' Ende Änderung 140717

        ' hier wird der Index für die Beauftragung bestimmt 


        Dim tmpProj As clsProjekt


        tmpProj = projekthistorie.beauftragung
        If IsNothing(tmpProj) Then
            valueBeauftragung = -1
        Else
            valueBeauftragung = nrSnapshots - projekthistorie.currentIndex
        End If




        hproj = projekthistorie.Last

        For i = 0 To 6
            necessary(i) = False
            minmaxScales(0, i) = 0.0
            minmaxScales(1, i) = 0.0
        Next
        '
        ' hier wird bestimmt, welche Skalierungsfaktoren überhaupt bereücksicht werden müssen 
        '
        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)
            Dim tmpArray() As String
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            If anzDiagrams > 0 Then
                For i = 1 To anzDiagrams
                    chtobj = CType(.ChartObjects(i), Excel.ChartObject)
                    If chtobj.Name <> "" Then

                        tmpArray = chtobj.Name.Split(New Char() {CType("#", Char)}, 5)
                        ' chtoj name ist aufgebaut: pr#PTprdk.kennung#pNAme#Auswahl
                        If tmpArray(0).Trim = "pr" Then

                            If tmpArray(2).Trim = vglName Then
                                If tmpArray(1) = CStr(PTprdk.Phasen) Then
                                    necessary(0) = True
                                ElseIf tmpArray(1) = CStr(PTprdk.PersonalBalken) And tmpArray(3) = "1" Then
                                    ' Personalbedarf
                                    necessary(1) = True
                                ElseIf tmpArray(1) = CStr(PTprdk.PersonalBalken) And tmpArray(3) = "2" Then
                                    ' Personalkosten
                                    necessary(2) = True
                                ElseIf tmpArray(1) = CStr(PTprdk.KostenBalken) And tmpArray(3) = "1" Then
                                    ' Sonstige Kosten
                                    necessary(3) = True
                                ElseIf tmpArray(1) = CStr(PTprdk.KostenBalken) And tmpArray(3) = "2" Then
                                    ' Gesamtkosten
                                    necessary(4) = True
                                ElseIf tmpArray(1) = CStr(PTprdk.StrategieRisiko) Then
                                    necessary(5) = False
                                ElseIf tmpArray(1) = CStr(PTprdk.Ergebnis) Then
                                    necessary(6) = True
                                End If
                            End If


                        End If


                    End If

                Next
            End If
        End With


        '
        ' hier werden die benötigten (necessary) Min- bzw. Max-Werte bestimmt 
        '


        For Each kvp As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste
            ' Phasen Skalierung min, max 
            Dim tmpValues() As Double
            ReDim tmpValues(kvp.Value.anzahlRasterElemente - 1)


            ' Phasen Darstellung

            If necessary(0) Then
                minmaxScales(0, 0) = 0.0
                currentvalue = kvp.Value.dauerInDays
                If currentvalue > minmaxScales(1, 0) Then

                    minmaxScales(1, 0) = currentvalue

                End If
            End If



            ' Personalbedarfe Ressourcen

            If necessary(1) Then
                minmaxScales(0, 1) = 0.0
                tmpValues = kvp.Value.getAlleRessourcen
                currentvalue = tmpValues.Max
                If currentvalue > minmaxScales(1, 1) Then
                    If currentvalue < 80 Then
                        minmaxScales(1, 1) = Round(currentvalue)
                    ElseIf currentvalue < 300 Then
                        minmaxScales(1, 1) = Round(currentvalue)
                    Else
                        minmaxScales(1, 1) = Round(currentvalue)
                    End If
                End If
            End If


            ' Personalkosten
            If necessary(2) Then
                minmaxScales(0, 2) = 0.0
                tmpValues = kvp.Value.getAllPersonalKosten
                persKosten = tmpValues.Sum
                currentvalue = tmpValues.Max
                If currentvalue > minmaxScales(1, 2) Then
                    If currentvalue < 80 Then
                        'minmaxScales(1, 2) = Round(currentvalue / 5 + 0.6) * 5
                        minmaxScales(1, 2) = Round(currentvalue)
                    ElseIf currentvalue < 300 Then
                        'minmaxScales(1, 2) = Round(currentvalue / 10 + 0.6) * 10
                        minmaxScales(1, 2) = Round(currentvalue)
                    Else
                        'minmaxScales(1, 2) = Round(currentvalue / 50 + 0.6) * 50
                        minmaxScales(1, 2) = Round(currentvalue)
                    End If
                End If
            End If


            ' Gesamt Andere Kosten
            If necessary(3) Then
                minmaxScales(0, 3) = 0.0
                tmpValues = kvp.Value.getGesamtAndereKosten
                sonstKosten = tmpValues.Sum
                currentvalue = tmpValues.Max
                If currentvalue > minmaxScales(1, 3) Then
                    If currentvalue < 80 Then
                        'minmaxScales(1, 3) = Round(currentvalue / 5 + 0.6) * 5
                        minmaxScales(1, 3) = Round(currentvalue)
                    ElseIf currentvalue < 300 Then
                        'minmaxScales(1, 3) = Round(currentvalue / 10 + 0.6) * 10
                        minmaxScales(1, 3) = Round(currentvalue)
                    Else
                        'minmaxScales(1, 3) = Round(currentvalue / 50 + 0.6) * 50
                        minmaxScales(1, 3) = Round(currentvalue)
                    End If
                End If
            End If


            ' Gesamt Kosten
            If necessary(4) Then
                minmaxScales(0, 4) = 0.0
                tmpValues = kvp.Value.getGesamtKostenBedarf
                gesamtKosten = tmpValues.Sum
                currentvalue = tmpValues.Max
                If currentvalue > minmaxScales(1, 4) Then
                    If currentvalue < 80 Then
                        'minmaxScales(1, 4) = Round(currentvalue / 5 + 0.6) * 5
                        minmaxScales(1, 4) = Round(currentvalue)
                    ElseIf currentvalue < 300 Then
                        'minmaxScales(1, 4) = Round(currentvalue / 10 + 0.6) * 10
                        minmaxScales(1, 4) = Round(currentvalue)
                    Else
                        'minmaxScales(1, 4) = Round(currentvalue / 50 + 0.6) * 50
                        minmaxScales(1, 4) = Round(currentvalue)
                    End If
                End If
            End If


            ' Strategie Risiko : ist dann zu unterscheiden, wenn die anderen auch unterstützt werden 
            ' wie z.Bsp FitRisikoVol oder ComplexRisk oder Zeit Risk  
            minmaxScales(0, 5) = 0.0
            minmaxScales(1, 5) = 11.0

            ' Ergebnis 
            If necessary(6) Then
                erloes = kvp.Value.Erloes
                tmpValues = kvp.Value.getGesamtKostenBedarf
                gesamtKosten = tmpValues.Sum
                currentvalue = erloes
                If currentvalue > minmaxScales(1, 6) Then
                    If currentvalue < 80 Then
                        'minmaxScales(1, 6) = Round(currentvalue / 5 + 0.6) * 5
                        minmaxScales(1, 6) = Round(currentvalue)
                    ElseIf currentvalue < 300 Then
                        'minmaxScales(1, 6) = Round(currentvalue / 10 + 0.6) * 10
                        minmaxScales(1, 6) = Round(currentvalue)
                    Else
                        'minmaxScales(1, 6) = Round(currentvalue / 50 + 0.6) * 50
                        minmaxScales(1, 6) = Round(currentvalue)
                    End If
                End If

                With kvp.Value
                    currentvalue = erloes - gesamtKosten * (1 + .risikoKostenfaktor)
                    If currentvalue < minmaxScales(0, 6) Then
                        If currentvalue < -300 Then
                            'minmaxScales(0, 6) = Round(currentvalue / 50 - 0.6) * 50
                            minmaxScales(0, 6) = Round(currentvalue)
                        ElseIf currentvalue < -80 Then
                            'minmaxScales(0, 6) = Round(currentvalue / 10 - 0.6) * 10
                            minmaxScales(0, 6) = Round(currentvalue)
                        Else
                            'minmaxScales(0, 6) = Round(currentvalue / 5 - 0.6) * 5
                            minmaxScales(0, 6) = Round(currentvalue)
                        End If
                    End If
                End With
            End If

        Next
        '
        ' jetzt werden wieder alle relevanten Diagramme durchgegangen, um sie auf die entsprechende Skalierung zu setzen ...
        '
        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)
            Dim tmpArray() As String
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            If anzDiagrams > 0 Then
                For i = 1 To anzDiagrams
                    chtobj = CType(.ChartObjects(i), Excel.ChartObject)
                    If chtobj.Name <> "" Then

                        tmpArray = chtobj.Name.Split(New Char() {CType("#", Char)}, 5)
                        ' chtoj name ist aufgebaut: pr#PTprdk.kennung#pNAme#Auswahl

                        If tmpArray(0).Trim = "pr" Then

                            If tmpArray(2).Trim = vglName Then

                                If tmpArray(1) = CStr(PTprdk.Phasen) Then
                                    If necessary(0) Then
                                        With CType(chtobj.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                                            .MinimumScale = minmaxScales(0, 0)
                                            .MaximumScale = CInt(minmaxScales(1, 0) / 365 * 12) + 3
                                        End With
                                    End If

                                ElseIf tmpArray(1) = CStr(PTprdk.PersonalBalken) Then

                                    If necessary(1) And tmpArray(3) = "1" Then

                                        With CType(chtobj.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                                            .MinimumScale = minmaxScales(0, 1)
                                            .MaximumScale = minmaxScales(1, 1)
                                        End With

                                    ElseIf necessary(2) And tmpArray(3) = "2" Then

                                        With CType(chtobj.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                                            .MinimumScale = minmaxScales(0, 2)
                                            .MaximumScale = minmaxScales(1, 2)
                                        End With

                                    End If

                                ElseIf tmpArray(1) = CStr(PTprdk.KostenBalken) Then


                                    If necessary(3) And tmpArray(3) = "1" Then
                                        ' Sonstige Kosten
                                        With CType(chtobj.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                                            .MinimumScale = minmaxScales(0, 3)
                                            .MaximumScale = minmaxScales(1, 3)
                                        End With
                                    ElseIf necessary(4) And tmpArray(3) = "2" Then
                                        ' Gesamtkosten
                                        With CType(chtobj.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                                            .MinimumScale = minmaxScales(0, 4)
                                            .MaximumScale = minmaxScales(1, 4)
                                        End With
                                    End If


                                ElseIf tmpArray(1) = CStr(PTprdk.StrategieRisiko) Then
                                    ' noch zu programmieren 
                                    ' momentan ist nichts zu tun, da nur StrategieRisiko unterstützt wird mit immer fester Skalierung


                                ElseIf tmpArray(1) = CStr(PTprdk.Ergebnis) Then

                                    If necessary(6) Then
                                        With CType(chtobj.Chart.Axes(Excel.XlAxisType.xlValue), Excel.Axis)
                                            .MinimumScale = minmaxScales(0, 6)
                                            .MaximumScale = minmaxScales(1, 6)
                                        End With
                                    End If

                                End If

                            End If

                        End If

                    End If

                Next
            End If
        End With

        'appInstance.ScreenUpdating = True
        'appInstance.ScreenUpdating = formerSU

    End Sub

   
    

    Private Sub compareCurrent_Click(sender As Object, e As EventArgs) Handles compareCurrent.Click
        ' in ProjektHistorie sind die Projekt-Snapshots in aufsteigender Reihenfolge sortiert 

        Dim pname As String = hproj.name
        Dim cproj As clsProjekt

        Try
            cproj = ShowProjekte.getProject(pname)
            Dim top As Double = Me.Top + Me.Height + 5
            Dim left As Double = Me.Left + Me.Width * 0.6


            Call awinCompareProject(hproj, cproj, 4, top, left)
        Catch ex As Exception
            Call MsgBox("Fehler bei Compare " & hproj.name & vbLf & ex.Message)
        End Try
       

    End Sub

    Private Sub compareBeauftragung_Click(sender As Object, e As EventArgs) Handles compareBeauftragung.Click
        ' in ProjektHistorie sind die Projekt-Snapshots in aufsteigender Reihenfolge sortiert 

        Dim cproj As clsProjekt = projekthistorie.First
        Dim top As Double = Me.Top + Me.Height + 5
        Dim left As Double = Me.Left + 20

        If valueBeauftragung >= 0 Then
            cproj = projekthistorie.ElementAt(nrSnapshots - valueBeauftragung)
            Call awinCompareProject(hproj, cproj, 3, top, left)
        Else
            Call MsgBox("es gibt keine Beauftragung")
        End If

    End Sub

   

    Private Sub timeSlider_ValueChanged(sender As Object, e As EventArgs) Handles timeSlider.ValueChanged


        hproj = projekthistorie.ElementAt(nrSnapshots - timeSlider.Value)

        ' Änderung-Ergänzung 140717 tk: Neuzeichnen des Shapes inkl der gezeigten Phasen/Meilensteine

        Call clearProjektinPlantafel(hproj.name)
        ' jetzt muss das aktuelle Shape rausgenommen werden und mit dem TimeSlider Shape ersetzt werden

        Try
            ShowProjekte.Remove(hproj.name)
            ShowProjekte.Add(hproj)

        Catch ex As Exception

        End Try

        Dim tmpCollection As New Collection
        Call ZeichneProjektinPlanTafel(noCollection:=tmpCollection, pname:=hproj.name, tryzeile:=hproj.tfZeile, _
                                       drawPhaseList:=phaseList, drawMilestoneList:=milestoneList)


        ' Ende Änderung-Ergänzung 140717

        With hproj

            If timeSlider.Value = 0 Then
                snapshotDate.Text = "Aktueller Stand"
            ElseIf timeSlider.Value = valueBeauftragung Then
                snapshotDate.Text = .timeStamp.ToString & " (Beauftragung)"
            Else
                snapshotDate.Text = .timeStamp.ToString
            End If

        End With

        'Call aktualisierePMSForms(hproj)
        Call aktualisiereCharts(hproj, False)

    End Sub


    Private Sub typSelection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles typSelection.SelectedIndexChanged

        With typSelection
            Select Case .SelectedIndex
                Case 0
                    ' Personalkosten
                Case 1
                    ' Sonstige Kosten
                Case 2
                    ' Budget
                Case 3
                    ' Ergebnis
                Case 4
                    ' Strategie und Risiko
                Case 5
                    ' Termine
                Case 6
                    ' Projekt-Ampel
                Case 7
                    ' Meilenstein-Ampeln
                Case Else

            End Select
        End With

    End Sub

    Private Sub movetoPrevious_Click(sender As Object, e As EventArgs) Handles movetoPrevious.Click


        If timeSlider.Value = nrSnapshots Then
            My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
        Else
            Try
                With typSelection
                    Select Case .SelectedIndex
                        Case 1
                            ' Personalkosten
                            hproj = projekthistorie.PrevDiff(PThcc.perscost)
                        Case 2
                            ' Sonstige Kosten
                            hproj = projekthistorie.PrevDiff(PThcc.othercost)
                        Case 3
                            ' Budget
                            hproj = projekthistorie.PrevDiff(PThcc.budget)
                        Case 4
                            ' Ergebnis
                            hproj = projekthistorie.PrevDiff(PThcc.ergebnis)
                        Case 5
                            ' Strategie und Risiko
                            hproj = projekthistorie.PrevDiff(PThcc.fitrisk)
                        Case 6
                            ' Termine
                            hproj = projekthistorie.PrevDiff(PThcc.resultdates)
                        Case 7
                            ' Projekt-Ampel
                            hproj = projekthistorie.PrevDiff(PThcc.projektampel)
                        Case 8
                            ' Meilenstein-Ampeln
                            hproj = projekthistorie.PrevDiff(PThcc.resultampel)

                        Case 9
                            ' Phasen
                            hproj = projekthistorie.PrevDiff(PThcc.phasen)
                        Case Else

                    End Select

                    timeSlider.Value = nrSnapshots - projekthistorie.currentIndex
                    If timeSlider.Value = 0 Then
                        snapshotDate.Text = "Aktueller Stand"
                    ElseIf timeSlider.Value = valueBeauftragung Then
                        snapshotDate.Text = hproj.timeStamp.ToString & " (Beauftragung)"
                    Else
                        snapshotDate.Text = hproj.timeStamp.ToString
                    End If

                End With

            Catch ex As Exception
                My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
            End Try
        End If

        

    End Sub

    Private Sub movetoNext_Click(sender As Object, e As EventArgs) Handles movetoNext.Click


        If timeSlider.Value = 0 Then
            My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
        Else
            Try
                With typSelection
                    Select Case .SelectedIndex
                        Case 1
                            ' Personalkosten
                            hproj = projekthistorie.NextDiff(PThcc.perscost)
                        Case 2
                            ' Sonstige Kosten
                            hproj = projekthistorie.NextDiff(PThcc.othercost)
                        Case 3
                            ' Budget
                            hproj = projekthistorie.NextDiff(PThcc.budget)
                        Case 4
                            ' Ergebnis
                            hproj = projekthistorie.NextDiff(PThcc.ergebnis)
                        Case 5
                            ' Strategie und Risiko
                            hproj = projekthistorie.NextDiff(PThcc.fitrisk)
                        Case 6
                            ' Termine
                            hproj = projekthistorie.NextDiff(PThcc.resultdates)
                        Case 7
                            ' Projekt-Ampel
                            hproj = projekthistorie.NextDiff(PThcc.projektampel)
                        Case 8
                            ' Meilenstein-Ampeln
                            hproj = projekthistorie.NextDiff(PThcc.resultampel)

                        Case 9
                            ' Phasen
                            hproj = projekthistorie.NextDiff(PThcc.phasen)
                        Case Else

                    End Select

                    timeSlider.Value = nrSnapshots - projekthistorie.currentIndex
                    If timeSlider.Value = 0 Then
                        snapshotDate.Text = "Aktueller Stand"
                    ElseIf timeSlider.Value = valueBeauftragung Then
                        snapshotDate.Text = hproj.timeStamp.ToString & " (Beauftragung)"
                    Else
                        snapshotDate.Text = hproj.timeStamp.ToString
                    End If

                End With

            Catch ex As Exception
                My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
            End Try
        End If




    End Sub


    Private Sub movetoBeauftragung_Click(sender As Object, e As EventArgs) Handles movetoBeauftragung.Click

        If valueBeauftragung < 0 Then
            Call MsgBox(" das Projekt wurde noch nicht beauftragt")
        Else
            timeSlider.Value = valueBeauftragung
            hproj = projekthistorie.ElementAt(nrSnapshots - timeSlider.Value)
            snapshotDate.Text = hproj.timeStamp.ToString & " (Beauftragung)"
        End If
        

    End Sub

    
   
    Private Sub showMore_Click(sender As Object, e As EventArgs) Handles showMore.Click

        If showAll Then
            ' jetzt soll ausgeblendet werden ....
            With Me
                .compareCurrent.Visible = False
                .compareBeauftragung.Visible = False
                .movetoBeauftragung.Visible = False
                .movetoNext.Visible = False
                .movetoPrevious.Visible = False
                .typSelection.Visible = False
                .Label3.Visible = False
                .showMore.Text = "more ..."
                .Height = 190             ' ur: 28.05.2014
            End With
            showAll = Not showAll
        Else
            ' jetzt soll eingeblendet werden 
            With Me
                .Height = 300           ' ur: 28.05.2014
                .compareCurrent.Visible = True
                .compareBeauftragung.Visible = True
                .movetoBeauftragung.Visible = True
                .movetoNext.Visible = True
                .movetoPrevious.Visible = True
                .typSelection.Visible = True
                .Label3.Visible = True
                .showMore.Text = "less ..."
            End With
            showAll = Not showAll
        End If
    End Sub
End Class