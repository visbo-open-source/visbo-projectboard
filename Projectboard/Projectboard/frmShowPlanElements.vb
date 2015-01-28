Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports ClassLibrary1
Imports Microsoft.Office.Interop.Excel

Public Class frmShowPlanElements

    ' kann von ausserhalb gesetzt werden; gibt an ob das ganze Portfolio angezeigt werden soll
    ' oder nur die selektierten Projekte 
    Friend showModePortfolio As Boolean
    Friend menuOption As Integer
    Friend chkbxShowObjects As Boolean
    Friend chkbxCreateCharts As Boolean


    Private allMilestones As New Collection
    Private allPhases As New Collection
    Private allCosts As New Collection
    Private allRoles As New Collection
    Private allBUs As New Collection
    Private allTyps As New Collection


    Private selectedMilestones As New Collection
    Private selectedPhases As New Collection
    Private selectedCosts As New Collection
    Private selectedRoles As New Collection
    Private selectedBUs As New Collection
    Private selectedTyps As New Collection

    'Private sKeyMilestones As String = ""
    'Private sKeyPhases As String = ""
    'Private sKeyCosts As String = ""
    'Private sKeyRoles As String = ""

    Private backgroundRunning As Boolean = False

    Private Enum PTauswahlTyp
        phase = 0
        meilenstein = 1
        Rolle = 2
        Kostenart = 3
        BusinessUnit = 4
        ProjektTyp = 5
    End Enum


    Private chTyp As String

    

    ''' <summary>
    ''' Koordinaten merken 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmShowPlanElements_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.listselP, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.listselP, PTpinfo.left) = Me.Left

    End Sub

    ''' <summary>
    ''' wird zu Beginn, als "Lade-Routine" für das Formular aufgerufen; besetzt unter anderem die Selection Collections aus dem letzten Filter
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmShowPlanElements_Load(sender As Object, e As EventArgs) Handles Me.Load

        If frmCoord(PTfrm.listselP, PTpinfo.top) > 0 Then
            Me.Top = CInt(frmCoord(PTfrm.listselP, PTpinfo.top))
            Me.Left = CInt(frmCoord(PTfrm.listselP, PTpinfo.left))
        Else
            Me.Top = 60
            Me.Left = 100
        End If

        statusLabel.Text = ""
        statusLabel.Visible = True

        Dim nrShapes As Integer = appearanceDefinitions.Count


        ' jetzt werden anhand des letzten Filters die Collections gesetzt 
        Call retrieveSelections("Last", selectedBUs, selectedTyps, _
                                selectedPhases, selectedMilestones, _
                                selectedRoles, selectedCosts)

        ' jetzt werden die ProjektReport- bzw. PortfolioReport-Vorlagen ausgelesen 
        ' in diesem Fall werden nur die mit Multiprojekt angezeigt 

        If Me.menuOption = PTmenue.multiprojektReport Then
            Dim dateiName As String = ""
            Dim dirname As String = awinPath & RepPortfolioVorOrdner

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
                'Call MsgBox(ex.Message & ": " & dateiName)
            End Try
        End If

        Me.rdbPhases.Checked = True

    End Sub

    ''' <summary>
    ''' Behandlung OK Button drücken
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click


        appInstance.EnableEvents = False
        enableOnUpdate = False

        statusLabel.Text = ""

        ' hier muss jetzt noch der aktuelle rdb ausgelesen werden ..
        If Me.rdbPhases.Checked = True Then

            selectedPhases.Clear()
            For Each element As String In ListBox2.Items
                If Not selectedPhases.Contains(element) Then
                    selectedPhases.Add(element, element)
                End If
            Next


        ElseIf Me.rdbMilestones.Checked = True Then

            selectedMilestones.Clear()
            For Each element As String In ListBox2.Items
                If Not selectedMilestones.Contains(element) Then
                    selectedMilestones.Add(element, element)
                End If
            Next

        ElseIf rdbRoles.Checked = True Then

            selectedRoles.Clear()
            For Each element As String In ListBox2.Items
                If Not selectedRoles.Contains(element) Then
                    selectedRoles.Add(element, element)
                End If
            Next

        ElseIf rdbCosts.Checked = True Then

            selectedCosts.Clear()
            For Each element As String In ListBox2.Items
                If Not selectedCosts.Contains(element) Then
                    selectedCosts.Add(element, element)
                End If
            Next

        ElseIf rdbBU.Checked = True Then

            selectedBUs.Clear()
            For Each element As String In ListBox2.Items
                If Not selectedBUs.Contains(element) Then
                    selectedBUs.Add(element, element)
                End If
            Next

        ElseIf rdbTyp.Checked = True Then

            selectedTyps.Clear()
            For Each element As String In ListBox2.Items
                If Not selectedTyps.Contains(element) Then
                    selectedTyps.Add(element, element)
                End If
            Next
        End If




        ''''
        ''
        ''
        ' jetzt kommt die Fall-Unterscheidung 
        ''
        ''
        ''''

        Dim validOption As Boolean
        If Me.menuOption = PTmenue.visualisieren Then
            validOption = True
        ElseIf showRangeRight - showRangeLeft > 5 Then
            validOption = True
        Else
            validOption = False
        End If


        If Me.menuOption = PTmenue.multiprojektReport Then

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And validOption Then

                Dim vorlagenDateiName As String
                vorlagenDateiName = awinPath & RepPortfolioVorOrdner & _
                                    "\" & repVorlagenDropbox.Text

                Try
                    rdbMilestones.Enabled = False
                    rdbPhases.Enabled = False
                    rdbRoles.Enabled = False
                    rdbCosts.Enabled = False
                    filterBox.Enabled = False
                    ListBox1.Enabled = False
                    OKButton.Enabled = False
                    repVorlagenDropbox.Enabled = False
                    AbbrButton.Cursor = Cursors.Arrow

                    statusLabel.Text = ""
                    statusLabel.Visible = True

                    Me.Cursor = Cursors.WaitCursor

                    ' Alternativ ohne Background Worker

                    BackgroundWorker1.RunWorkerAsync(vorlagenDateiName)

                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try


            Else
                Call MsgBox("bitte mindestens ein Element selektieren bzw. " & vbLf & _
                             "einen Zeitraum angeben ...")
            End If

        ElseIf Me.menuOption = PTmenue.leistbarkeitsAnalyse Then

            Dim myCollection As New Collection

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And validOption Then

                Dim formerSU As Boolean = appInstance.ScreenUpdating
                appInstance.ScreenUpdating = False

                ' Window Position festlegen
                Dim chtop As Double = 50.0 + awinSettings.ChartHoehe1
                Dim chleft As Double = (showRangeRight - 1) * boxWidth + 4

                If selectedPhases.Count > 0 Then
                    chTyp = DiagrammTypen(0)
                    Call zeichneLeistbarkeitsChart(selectedPhases, chTyp, chtop, chleft)
                End If

                If selectedMilestones.Count > 0 Then
                    chTyp = DiagrammTypen(5)
                    Call zeichneLeistbarkeitsChart(selectedMilestones, chTyp, chtop, chleft)
                End If

                If selectedRoles.Count > 0 Then
                    chTyp = DiagrammTypen(1)
                    Call zeichneLeistbarkeitsChart(selectedRoles, chTyp, chtop, chleft)
                End If

                If selectedCosts.Count > 0 Then
                    chTyp = DiagrammTypen(2)
                    Call zeichneLeistbarkeitsChart(selectedCosts, chTyp, chtop, chleft)
                End If

                ' den aktuellen Filter wegschreiben unter dem Namen last 
                
                Call storeFilterAndclearSelections("Last")
                appInstance.ScreenUpdating = formerSU

            Else

            End If

        ElseIf Me.menuOption = PTmenue.visualisieren Then


            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And validOption Then

                If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0) And _
                    (selectedRoles.Count > 0 Or selectedCosts.Count > 0) Then
                    Call MsgBox("es können nur entweder Phasen / Meilensteine oder Rollen oder Kosten angezeigt werden")

                ElseIf selectedPhases.Count > 0 Or selectedMilestones.Count > 0 Then

                    If selectedPhases.Count > 0 Then
                        Call awinZeichnePhasen(selectedPhases, False, True)
                    End If

                    If selectedMilestones.Count > 0 Then
                        ' Phasen anzeigen 
                        Dim farbID As Integer = 4
                        Call awinZeichneMilestones(selectedMilestones, farbID, False, True)

                    End If

                ElseIf selectedRoles.Count > 0 Then
                    Call MsgBox("noch nicht implementiert")

                Else
                    Call MsgBox("noch nicht implementiert")
                End If

                Call storeFilterAndclearSelections("Last")

            Else
                Call MsgBox("bitte mindestens ein Element aus einer der Kategorien selektieren  ")
            End If

        ElseIf menuOption = PTmenue.filterdefinieren Then

            Call storeFilterAndclearSelections("Last")
            Call MsgBox("ok, Filter gespeichert")

        Else

            Call MsgBox("noch nicht unterstützt")

        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

    End Sub

    ''' <summary>
    ''' stellt ggf den vorherigen Zustand an vor-selektierten Items wieder her
    ''' ebenso den Searchkey 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbPhases_CheckedChanged(sender As Object, e As EventArgs) Handles rdbPhases.CheckedChanged

        Dim i As Integer
        statusLabel.Text = ""

        If Me.rdbPhases.Checked Then
            ' clear Listbox1 
            headerLine.Text = "Phasen"
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            filterBox.Text = ""

            chkbxOneChart.Text = "Alles in einem Chart"


            If allPhases.Count = 0 Then
                For i = 1 To PhaseDefinitions.Count
                    allPhases.Add(CStr(PhaseDefinitions.getPhaseDef(i).name))
                Next
            End If


            Call rebuildFormerState(PTauswahlTyp.phase)


        Else
            ' Merken, was ggf. das Filterkriterium war 
            'sKeyPhases = filterBox.Text

            ' Merken welches die selektierten Phasen waren 
            selectedPhases.Clear()
            For Each element As String In ListBox2.Items
                selectedPhases.Add(element, element)
            Next

        End If

    End Sub

    Private Sub rdbMilestones_CheckedChanged(sender As Object, e As EventArgs) Handles rdbMilestones.CheckedChanged

        statusLabel.Text = ""

        If Me.rdbMilestones.Checked Then
            ' clear Listbox1 
            headerLine.Text = "Meilensteine"
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()

            filterBox.Text = ""

            chkbxOneChart.Text = "Alles in einem Chart"

            If allMilestones.Count = 0 Then

                For i As Integer = 1 To MilestoneDefinitions.Count
                    allMilestones.Add(MilestoneDefinitions.elementAt(i - 1).name)
                Next
            End If

            
            Call rebuildFormerState(PTauswahlTyp.meilenstein)



        Else
            
            ' Merken welches die selektierten Phasen waren 
            selectedMilestones.Clear()
            For Each element As String In ListBox2.Items
                selectedMilestones.Add(element, element)
            Next


        End If
    End Sub

    ''' <summary>
    ''' zeigt alle Rollen an, unabhängig davon ob sie in den Projekten vorkommen oder nicht 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbRoles_CheckedChanged(sender As Object, e As EventArgs) Handles rdbRoles.CheckedChanged

        Dim i As Integer

        statusLabel.Text = ""

        If RoleDefinitions.Count = 0 Then
            Call MsgBox("es sind keine Kostenarten definiert !")
        Else
            If Me.rdbRoles.Checked Then
                ' clear Listbox1 
                headerLine.Text = "Rollen"
                ListBox1.Items.Clear()
                ListBox2.Items.Clear()
                filterBox.Text = ""
                chkbxOneChart.Text = "Alles in einem Chart"


                If allRoles.Count = 0 Then
                    For i = 1 To RoleDefinitions.Count
                        allRoles.Add(RoleDefinitions.getRoledef(i).name)
                    Next
                End If


                Call rebuildFormerState(PTauswahlTyp.Rolle)



            Else
                ' Merken, was ggf. das Filterkriterium war 
                'sKeyRoles = filterBox.Text

                ' Merken welches die selektierten Phasen waren 
                selectedRoles.Clear()
                For Each element As String In ListBox2.Items
                    selectedRoles.Add(element, element)
                Next

            End If
        End If

    End Sub

    ''' <summary>
    ''' wenn Radio-Button Kosten gedrückt wird 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbCosts_CheckedChanged(sender As Object, e As EventArgs) Handles rdbCosts.CheckedChanged
        Dim i As Integer

        statusLabel.Text = ""

        If CostDefinitions.Count = 0 Then
            Call MsgBox("es sind keine Kostenarten definiert !")
        Else
            If Me.rdbCosts.Checked Then
                ' clear Listbox1 
                headerLine.Text = "Kostenarten"
                ListBox1.Items.Clear()
                ListBox2.Items.Clear()
                filterBox.Text = ""
                chkbxOneChart.Text = "Alles in einem Chart"

                If allCosts.Count = 0 Then
                    For i = 1 To CostDefinitions.Count
                        allCosts.Add(CostDefinitions.getCostdef(i).name)
                    Next
                End If

                Call rebuildFormerState(PTauswahlTyp.Kostenart)

            Else
                
                ' Merken welches die selektierten Phasen waren 
                selectedCosts.Clear()
                'For Each element As String In ListBox1.SelectedItems
                For Each element As String In ListBox2.Items
                    selectedCosts.Add(element, element)
                Next

            End If
        End If

    End Sub

    ''' <summary>
    ''' Behandlung Radio Button Business Unit drücken 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbBU_CheckedChanged(sender As Object, e As EventArgs) Handles rdbBU.CheckedChanged

        Dim i As Integer

        statusLabel.Text = ""

        If businessUnitDefinitions.Count = 0 Then
            Call MsgBox("es sind keine Business Units definiert !")
        Else
            If Me.rdbBU.Checked Then
                ' clear Listbox1 
                headerLine.Text = "Business Units"
                ListBox1.Items.Clear()
                ListBox2.Items.Clear()
                filterBox.Text = ""

                If allBUs.Count = 0 Then
                    For i = 1 To businessUnitDefinitions.Count
                        allBUs.Add(CStr(businessUnitDefinitions.ElementAt(i - 1).Value.name))
                    Next
                End If

                Call rebuildFormerState(PTauswahlTyp.BusinessUnit)

            Else

                ' Merken welches die selektierten Phasen waren 
                selectedBUs.Clear()

                For Each element As String In ListBox2.Items
                    selectedBUs.Add(element, element)
                Next

            End If
        End If


    End Sub

    Private Sub rdbTyp_CheckedChanged(sender As Object, e As EventArgs) Handles rdbTyp.CheckedChanged

        Dim i As Integer

        statusLabel.Text = ""

        If Projektvorlagen.Count = 0 Then
            Call MsgBox("es sind keine Projektvorlagen definiert !")
        Else
            If Me.rdbTyp.Checked Then
                ' clear Listbox1 
                headerLine.Text = "Projektvorlagen / Generik"
                ListBox1.Items.Clear()
                ListBox2.Items.Clear()

                filterBox.Text = ""
                chkbxOneChart.Text = "Alles in einem Chart"

                If allTyps.Count = 0 Then
                    For i = 1 To Projektvorlagen.Count
                        allTyps.Add(Projektvorlagen.Liste.ElementAt(i - 1).Key)
                    Next
                End If
                

                Call rebuildFormerState(PTauswahlTyp.ProjektTyp)

            Else

                ' Merken welches die selektierten Phasen waren 
                selectedTyps.Clear()

                For Each element As String In ListBox2.Items
                    selectedTyps.Add(element, element)
                Next

            End If
        End If


    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click


        If menuOption = PTmenue.multiprojektReport And backgroundRunning Then

            rdbMilestones.Enabled = True
            rdbPhases.Enabled = True
            rdbRoles.Enabled = True
            rdbCosts.Enabled = True
            filterBox.Enabled = True
            ListBox1.Enabled = True
            OKButton.Enabled = True
            repVorlagenDropbox.Enabled = True
            statusLabel.Text = "Berichterstellung wurde beendet"

            Me.Cursor = Cursors.Arrow
            backgroundRunning = False

            Try
                Me.BackgroundWorker1.CancelAsync()
            Catch ex As Exception

            End Try


        Else
            ListBox1.SelectedItems.Clear()
            ListBox2.Items.Clear()
            filterBox.Text = ""

            If rdbPhases.Checked Then
                selectedPhases.Clear()
            ElseIf rdbMilestones.Checked Then
                selectedMilestones.Clear()
            ElseIf rdbRoles.Checked Then
                selectedRoles.Clear()
            ElseIf rdbCosts.Checked Then
                selectedCosts.Clear()
            ElseIf rdbBU.Checked Then
                selectedBUs.Clear()
            Else
                selectedTyps.Clear()
            End If
        End If


        'MyBase.Close()

    End Sub


    ''' <summary>
    ''' wenn etwas in der Such-Maske eingegeben wird: prüfen, Listbox1 entsprechend ausdünnen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub filterBox_TextChanged(sender As Object, e As EventArgs) Handles filterBox.TextChanged

        Dim suchstr As String = filterBox.Text
        Dim currentNames As New Collection

        If rdbPhases.Checked Then
            currentNames = allPhases
        ElseIf rdbMilestones.Checked Then
            currentNames = allMilestones
        ElseIf rdbRoles.Checked Then
            currentNames = allRoles
        ElseIf rdbCosts.Checked Then
            currentNames = allCosts
        ElseIf rdbBU.Checked Then
            currentNames = allBUs
        ElseIf rdbTyp.Checked Then
            currentNames = allTyps
        End If


        If filterBox.Text = "" Then
            ListBox1.Items.Clear()
            For Each s As String In currentNames
                ListBox1.Items.Add(s)
            Next
        Else
            ListBox1.Items.Clear()
            For Each s As String In currentNames
                If s.Contains(suchstr) Then
                    ListBox1.Items.Add(s)
                End If
            Next
        End If

    End Sub


    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub pictureRoles_Click(sender As Object, e As EventArgs) Handles pictureRoles.Click
        If Me.rdbRoles.Checked = False Then
            rdbRoles.Checked = True
        Else
            rdbRoles.Checked = False
        End If
    End Sub

    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub picturePhasen_Click(sender As Object, e As EventArgs) Handles picturePhasen.Click
        If Me.rdbPhases.Checked = False Then
            rdbPhases.Checked = True
        Else
            rdbPhases.Checked = False
        End If
    End Sub

    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub pictureMilestones_Click(sender As Object, e As EventArgs) Handles pictureMilestones.Click
        If Me.rdbMilestones.Checked = False Then
            rdbMilestones.Checked = True
        Else
            rdbMilestones.Checked = False
        End If
    End Sub

    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub pictureCosts_Click(sender As Object, e As EventArgs) Handles pictureCosts.Click
        If Me.rdbCosts.Checked = False Then
            Me.rdbCosts.Checked = True
        Else
            Me.rdbCosts.Checked = False
        End If
    End Sub

    ''' <summary>
    ''' Klick auf das Bild soll auch den Radiobutton setzen
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub pictureBU_Click(sender As Object, e As EventArgs) Handles pictureBU.Click

        If Me.rdbBU.Checked = False Then
            Me.rdbBU.Checked = True
        Else
            Me.rdbBU.Checked = False
        End If

    End Sub

    ''' <summary>
    ''' stellt den vorherigen Zustand wieder her: welche Werte waren bereits für die betreffende 
    ''' Kategorie ausgewählt
    ''' dabei wird auf die in dieser Klasse definierten Variablen selectedphases, allphases, ... zugegriffen 
    ''' </summary>
    ''' <param name="typ"></param>
    ''' <remarks></remarks>
    Private Sub rebuildFormerState(ByVal typ As Integer)

        'Dim searchkey As String = ""
        Dim tmpCollection As New Collection
        Dim i As Integer
        Dim listOfNames As New Collection

        Select Case typ
            Case PTauswahlTyp.phase
                'searchkey = sKeyPhases
                tmpCollection = selectedPhases
                listOfNames = allPhases

            Case PTauswahlTyp.meilenstein
                'searchkey = sKeyMilestones
                tmpCollection = selectedMilestones
                listOfNames = allMilestones

            Case PTauswahlTyp.Rolle
                'searchkey = sKeyRoles
                tmpCollection = selectedRoles
                listOfNames = allRoles

            Case PTauswahlTyp.Kostenart
                'searchkey = sKeyCosts
                tmpCollection = selectedCosts
                listOfNames = allCosts

            Case PTauswahlTyp.BusinessUnit
                tmpCollection = selectedBUs
                listOfNames = allBUs

            Case PTauswahlTyp.ProjektTyp
                tmpCollection = selectedTyps
                listOfNames = allTyps

        End Select



        ' Filter Box Test setzen 
        For i = 1 To listOfNames.Count
            ListBox1.Items.Add(listOfNames.Item(i))
        Next
        filterBox.Text = ""

        ' jetzt prüfen, ob selectedphases bereits etwas enthält
        ' wenn ja, dann werden diese Items in Listbox2 dargestellt 
        For Each element As String In tmpCollection
            ListBox2.Items.Add(element)
        Next
    End Sub

    Private Sub BackgroundWorker1_Disposed(sender As Object, e As EventArgs) Handles BackgroundWorker1.Disposed



    End Sub


    ''' <summary>
    ''' Hintergrund Prozess - wird nur für die Report Erzeugung benötigt 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim vorlagenDateiName As String = CType(e.Argument, String)

        Try
            With awinSettings

                Call createPPTSlidesFromConstellation(vorlagenDateiName, _
                                                      selectedPhases, selectedMilestones, selectedRoles, selectedCosts, _
                                                      worker, e)

            End With
        Catch ex As Exception
            Call MsgBox("Fehler " & ex.Message)
        End Try
       


    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

    ''' <summary>
    ''' wird durchlaufen, wenn der Hintergrund Prozess mit dem Erstellen der Multiprojektsicht fertig ist 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Me.statusLabel.Text = "...done"
        Me.rdbMilestones.Enabled = True
        Me.rdbPhases.Enabled = True
        Me.rdbRoles.Enabled = True
        Me.rdbCosts.Enabled = True
        Me.filterBox.Enabled = True
        Me.ListBox1.Enabled = True
        Me.OKButton.Enabled = True
        Me.repVorlagenDropbox.Enabled = True
        Me.Cursor = Cursors.Arrow
        Me.statusLabel.Visible = True

        Call storeFilterAndclearSelections("Last")



    End Sub


    ''' <summary>
    ''' ruft das Formular auf, um die Einstellungen für das Multireporting vorzunehmen  
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub einstellungen_Click(sender As Object, e As EventArgs) Handles einstellungen.Click

        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult

        dialogreturn = mppFrm.ShowDialog


    End Sub

    ''' <summary>
    ''' fügt das selektierte Element der Listbox2 hinzu
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub addButton_Click(sender As Object, e As EventArgs) Handles addButton.Click

        Dim i As Integer
        Dim element As Object

        For i = 1 To ListBox1.SelectedItems.Count
            element = ListBox1.SelectedItems.Item(i - 1)
            If ListBox2.Items.Contains(element) Then
                ' nichts tun 
            Else
                ListBox2.Items.Add(element)
            End If
        Next


        ListBox1.SelectedItems.Clear()

    End Sub

    ''' <summary>
    ''' entfernt ein Item aus der Listbox2 - die ausgewöhlten Elemente 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub removeButton_Click(sender As Object, e As EventArgs) Handles removeButton.Click
        Dim i As Integer
        Dim element As Object
        Dim removeCollection As New Collection

        For i = 1 To ListBox2.SelectedItems.Count
            element = ListBox2.SelectedItems.Item(i - 1)
            removeCollection.Add(element)
        Next

        For Each element In removeCollection
            ListBox2.Items.Remove(element)
        Next

    End Sub

    ''' <summary>
    ''' zeichnet das Leistbarkeits-Chart 
    ''' </summary>
    ''' <param name="selCollection">Collection mit den Phasne-, Meilenstein, Rollen- oder Kostenarten</param>
    ''' <param name="chTyp">Typ: es handelt sich um Phasen, rollen, etc. </param>
    ''' <param name="chtop">auf welcher Höhe soll das Chart gezeichnet werden</param>
    ''' <param name="chleft">auf welcher x-Koordinate soll das Chart gezeichnet werden</param>
    ''' <remarks></remarks>
    Private Sub zeichneLeistbarkeitsChart(ByVal selCollection As Collection, ByVal chTyp As String, _
                                              ByRef chtop As Double, ByRef chleft As Double)


        Dim repObj As Excel.ChartObject
        Dim myCollection As Collection

        Dim chWidth As Double
        Dim chHeight As Double

        ' Window Position festlegen 
        chWidth = 265 + (showRangeRight - showRangeLeft - 12 + 1) * boxWidth + (showRangeRight - showRangeLeft) * screen_correct
        chHeight = awinSettings.ChartHoehe1


        If chkbxOneChart.Checked = True Then


            ' alles in einem Chart anzeigen
            myCollection = New Collection
            For Each element As String In selCollection
                myCollection.Add(element, element)
            Next

            repObj = Nothing
            Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                              chWidth, chHeight, False, chTyp, False)

            chtop = chtop + 5
            chleft = chleft + 7
        Else
            ' für jedes ITEM ein eigenes Chart machen
            For Each element As String In selCollection
                ' es muss jedesmal eine neue Collection erzeugt werden - die Collection wird in DiagramList gemerkt
                ' wenn die mit Clear leer gemacht wird, funktioniert der Diagram Update nicht mehr ....
                myCollection = New Collection
                myCollection.Add(element, element)
                repObj = Nothing

                Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                   chWidth, chHeight, False, chTyp, False)

                chtop = chtop + 5
                chleft = chleft + 7
            Next

        End If

    End Sub

    ''' <summary>
    ''' speichert den letzten Filter und setzt die temporären Collections wieder zurück 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub storeFilterAndclearSelections(ByVal fName As String)

        Dim lastFilter As clsFilter


        lastFilter = New clsFilter(fName, selectedBUs, selectedTyps, _
                                  selectedPhases, selectedMilestones, _
                                 selectedRoles, selectedCosts)


        If menuOption = PTmenue.filterdefinieren Then
            filterDefinitions.storeFilter(fName, lastFilter)
        Else
            selFilterDefinitions.storeFilter(fName, lastFilter)
        End If


        Me.selectedPhases.Clear()
        Me.selectedMilestones.Clear()
        Me.selectedRoles.Clear()
        Me.selectedCosts.Clear()
        Me.selectedBUs.Clear()
        Me.selectedTyps.Clear()

        Me.ListBox2.Items.Clear()

    End Sub

    ''' <summary>
    ''' besetzt die Selection Collections mit den Werten des Filters mit Namen fName
    ''' </summary>
    ''' <param name="fName"></param>
    ''' <param name="selectedBUs"></param>
    ''' <param name="selectedTyps"></param>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="selectedRoles"></param>
    ''' <param name="selectedCosts"></param>
    ''' <remarks></remarks>
    Private Sub retrieveSelections(ByVal fName As String, _
                                       ByRef selectedBUs As Collection, ByRef selectedTyps As Collection, _
                                       ByRef selectedPhases As Collection, ByRef selectedMilestones As Collection, _
                                       ByRef selectedRoles As Collection, ByRef selectedCosts As Collection)

        Dim lastFilter As clsFilter

        If menuOption = PTmenue.filterdefinieren Then
            lastFilter = filterDefinitions.retrieveFilter(fName)
        Else
            lastFilter = selFilterDefinitions.retrieveFilter(fName)
            If IsNothing(lastFilter) Then
                lastFilter = filterDefinitions.retrieveFilter(fName)
            End If
        End If


        If Not IsNothing(lastFilter) Then
            selectedBUs = lastFilter.BUs
            selectedTyps = lastFilter.Typs
            selectedPhases = lastFilter.Phases
            selectedMilestones = lastFilter.Milestones
            selectedRoles = lastFilter.Roles
            selectedCosts = lastFilter.Costs

        Else
            selectedBUs = New Collection
            selectedTyps = New Collection
            selectedPhases = New Collection
            selectedMilestones = New Collection
            selectedRoles = New Collection
            selectedCosts = New Collection
        End If

    End Sub

    Private Sub repVorlagenDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles repVorlagenDropbox.SelectedIndexChanged

        If menuOption = PTmenue.filterdefinieren Then
            ' es wurde ein anderer Filter gewählt ... 
        End If

    End Sub
End Class