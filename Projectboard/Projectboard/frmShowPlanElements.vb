Imports ProjectBoardDefinitions
Imports System.ComponentModel
Imports ClassLibrary1
Imports Microsoft.Office.Interop.Excel

Public Class frmShowPlanElements

    ' kann von ausserhalb gesetzt werden; gibt an ob das ganze Portfolio angezeigt werden soll
    ' oder nur die selektierten Projekte 
    Friend showModePortfolio As Boolean
    Friend menuOption As Integer
    Private existingNames As New Collection

    Private selectedMilestones As New Collection
    Private sKeyMilestones As String = ""

    Private selectedPhases As New Collection
    Private sKeyPhases As String = ""

    Private selectedCosts As New Collection
    Private sKeyCosts As String = ""

    Private selectedRoles As New Collection
    Private sKeyRoles As String = ""


    Private backgroundRunning As Boolean = False


    Private Enum PTauswahlTyp
        phase = 0
        meilenstein = 1
        Rolle = 2
        Kostenart = 3
    End Enum



    Private chtop As Double
    Private chleft As Double
    Private chWidth As Double
    Private chHeight As Double
    Private chTyp As String

    

    Private Sub frmShowPlanElements_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.listselP, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.listselP, PTpinfo.left) = Me.Left

    End Sub

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
                    If dateiName.Contains("Multiprojekt") Then
                        repVorlagenDropbox.Items.Add(dateiName)
                    End If

                Next i
            Catch ex As Exception
                'Call MsgBox(ex.Message & ": " & dateiName)
            End Try
        End If



    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim myCollection As Collection
        Dim repObj As Excel.ChartObject

        appInstance.EnableEvents = False
        enableOnUpdate = False

        statusLabel.Text = ""

        ' hier muss jetzt noch der aktuelle rdb ausgelesen werden ..
        If Me.rdbPhases.Checked = True Then

            selectedPhases.Clear()
            For Each element As String In ListBox1.SelectedItems
                If Not selectedPhases.Contains(element) Then
                    selectedPhases.Add(element, element)
                End If
            Next

        ElseIf Me.rdbMilestones.Checked = True Then

            selectedMilestones.Clear()
            For Each element As String In ListBox1.SelectedItems
                If Not selectedMilestones.Contains(element) Then
                    selectedMilestones.Add(element, element)
                End If
            Next

        ElseIf rdbRoles.Checked = True Then

            selectedRoles.Clear()
            For Each element As String In ListBox1.SelectedItems
                If Not selectedRoles.Contains(element) Then
                    selectedRoles.Add(element, element)
                End If
            Next

        ElseIf rdbCosts.Checked = True Then

            selectedCosts.Clear()
            For Each element As String In ListBox1.SelectedItems
                If Not selectedCosts.Contains(element) Then
                    selectedCosts.Add(element, element)
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

        If Me.menuOption = PTmenue.multiprojektReport Then

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And showRangeRight - showRangeLeft > 5 Then

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

        ElseIf Me.menuOption = PTmenue.leistbarkeitsAnalyse Or Me.menuOption = PTmenue.visualisieren Then

            If (selectedPhases.Count > 0 Or selectedMilestones.Count > 0 _
                    Or selectedRoles.Count > 0 Or selectedCosts.Count > 0) _
                    And showRangeRight - showRangeLeft > 5 Then

                If Me.rdbPhases.Checked = True Then

                    If chkbxShowObjects.Checked = True Then

                        ' Phasen anzeigen 

                        Call awinZeichnePhasen(selectedPhases, False, True)

                        If selectedMilestones.Count > 0 Then
                            ' Phasen anzeigen 
                            Dim farbID As Integer = 4
                            Call awinZeichneMilestones(selectedMilestones, farbID, False, True)

                        End If

                        selectedMilestones.Clear()
                        selectedPhases.Clear()

                    End If


                    If chkbxCreateCharts.Checked = True Then

                        Dim formerSU As Boolean = appInstance.ScreenUpdating
                        appInstance.ScreenUpdating = False
                        ' Window Position festlegen 
                        chtop = 50.0 + awinSettings.ChartHoehe1
                        chleft = (showRangeRight - 1) * boxWidth + 4
                        chWidth = 265 + (showRangeRight - showRangeLeft - 12 + 1) * boxWidth + (showRangeRight - showRangeLeft) * screen_correct
                        chHeight = awinSettings.ChartHoehe1
                        chTyp = DiagrammTypen(0)

                        If chkbxOneChart.Checked = True Then


                            ' alles in einem Chart anzeigen 
                            myCollection = New Collection
                            For Each element As String In ListBox1.SelectedItems
                                myCollection.Add(element, element)
                            Next

                            repObj = Nothing
                            Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                              chWidth, chHeight, False, chTyp, False)


                        Else
                            ' für jedes ITEM ein eigenes Chart machen
                            For Each element As String In ListBox1.SelectedItems
                                ' es muss jedesmal eine neue Collection erzeugt werden - die Collection wird in DiagramList gemerkt
                                ' wenn die mit Clear leer gemacht wird, funktioniert der Diagram Update nicht mehr ....
                                myCollection = New Collection
                                myCollection.Add(element, element)
                                repObj = Nothing

                                Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                                   chWidth, chHeight, False, chTyp, False)

                                chtop = chtop + chHeight + 2
                            Next

                        End If

                        appInstance.ScreenUpdating = formerSU

                    End If




                ElseIf Me.rdbMilestones.Checked = True Then
                    ' Milestones anzeigen

                    ' wenn Röntgen Blick an ist: ausschalten und Anzeige löschen
                    ' Alle bisher angezeigten Milestones löschen
                    Dim farbID As Integer = 4

                    If chkbxShowObjects.Checked = True Then

                        If selectedPhases.Count > 0 Then

                            ' Phasen anzeigen 
                            Call awinZeichnePhasen(selectedPhases, False, True)

                        End If



                        ' alles in einem Chart anzeigen 
                        selectedMilestones.Clear()
                        For Each element As String In ListBox1.SelectedItems
                            selectedMilestones.Add(element, element)
                        Next

                        ' Phasen anzeigen 
                        Call awinZeichneMilestones(selectedMilestones, farbID, False, True)

                        selectedMilestones.Clear()
                        selectedPhases.Clear()

                    End If


                    If chkbxCreateCharts.Checked = True Then

                        ' Window Position festlegen 
                        chtop = 50.0 + awinSettings.ChartHoehe1
                        chleft = (showRangeRight - 1) * boxWidth + 4
                        chWidth = 265 + (showRangeRight - showRangeLeft - 12 + 1) * boxWidth + (showRangeRight - showRangeLeft) * screen_correct
                        chHeight = awinSettings.ChartHoehe1
                        chTyp = DiagrammTypen(5)

                        If chkbxOneChart.Checked = True Then

                            ' alles in einem Chart anzeigen 
                            myCollection = New Collection
                            For Each element As String In ListBox1.SelectedItems
                                myCollection.Add(element, element)
                            Next

                            repObj = Nothing
                            Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                              chWidth, chHeight, False, chTyp, False)


                        Else
                            ' für jedes ITEM ein eigenes Chart machen
                            For Each element As String In ListBox1.SelectedItems
                                ' es muss jedesmal eine neue Collection erzeugt werden - die Collection wird in DiagramList gemerkt
                                ' wenn die mit Clear leer gemacht wird, funktioniert der Diagram Update nicht mehr ....
                                myCollection = New Collection
                                myCollection.Add(element, element)
                                repObj = Nothing

                                Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                                   chWidth, chHeight, False, chTyp, False)

                                chtop = chtop + chHeight + 2
                            Next

                        End If

                    End If


                ElseIf Me.rdbRoles.Checked = True Or Me.rdbCosts.Checked = True Then
                    ' Rollen anzeigen 


                    If chkbxShowObjects.Checked = True Then


                    End If

                    If chkbxCreateCharts.Checked = True Then

                        ' Window Position festlegen 
                        chtop = 50.0 + awinSettings.ChartHoehe1
                        chleft = (showRangeRight - 1) * boxWidth + 4
                        chWidth = 265 + (showRangeRight - showRangeLeft - 12 + 1) * boxWidth + (showRangeRight - showRangeLeft) * screen_correct
                        chHeight = awinSettings.ChartHoehe1

                        If Me.rdbRoles.Checked = True Then
                            chTyp = DiagrammTypen(1)
                        Else
                            chTyp = DiagrammTypen(2)
                        End If


                        If chkbxOneChart.Checked = True Then

                            ' alles in einem Chart anzeigen 
                            myCollection = New Collection
                            For Each element As String In ListBox1.SelectedItems
                                myCollection.Add(element, element)
                            Next

                            repObj = Nothing
                            Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                              chWidth, chHeight, False, chTyp, False)


                        Else
                            ' für jedes ITEM ein eigenes Chart machen
                            For Each element As String In ListBox1.SelectedItems
                                ' es muss jedesmal eine neue Collection erzeugt werden - die Collection wird in DiagramList gemerkt
                                ' wenn die mit Clear leer gemacht wird, funktioniert der Diagram Update nicht mehr ....
                                myCollection = New Collection
                                myCollection.Add(element, element)
                                repObj = Nothing

                                Call awinCreateprcCollectionDiagram(myCollection, repObj, chtop, chleft,
                                                                                   chWidth, chHeight, False, chTyp, False)

                                chtop = chtop + chHeight + 2
                            Next

                        End If

                    End If



                    'ElseIf Me.rdbCosts.Checked = True Then
                    ' Kosten anzeigen

                    ' Röntgen-Blick anschalten, wenn nicht eh schon an
                    ' wenn der an war, alle Werte zurücksetzen

                End If

            Else
                Call MsgBox("bitte mindestens ein Element selektieren bzw. " & vbLf & _
                             "einen Zeitraum angeben ...")
            End If

        Else

            Call MsgBox("noch nicht unterstützt")

        End If


        Me.ListBox1.SelectedItems.Clear()
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

        statusLabel.Text = ""

        If Me.rdbPhases.Checked Then
            ' clear Listbox1 
            headerLine.Text = "Phasen"
            ListBox1.Items.Clear()
            existingNames.Clear()
            filterBox.Text = ""
            chkbxShowObjects.Text = "in Projekten anzeigen"
            chkbxCreateCharts.Text = "Summen-Chart"
            chkbxOneChart.Text = "Alles in einem Chart"

            ' showModePortfolio kann nur gesetzt sein, wenn es auch einen selektierten Zeitraum gibt 
            existingNames = ShowProjekte.getPhaseNames
            Call rebuildFormerState(PTauswahlTyp.phase, existingNames)


        Else
            ' Merken, was ggf. das Filterkriterium war 
            sKeyPhases = filterBox.Text

            ' Merken welches die selektierten Phasen waren 
            selectedPhases.Clear()
            For Each element As String In ListBox1.SelectedItems
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
            existingNames.Clear()
            filterBox.Text = ""
            chkbxShowObjects.Text = "in Projekten anzeigen"
            chkbxCreateCharts.Text = "Summen-Chart"
            chkbxOneChart.Text = "Alles in einem Chart"

            existingNames = ShowProjekte.getMilestoneNames
            Call rebuildFormerState(PTauswahlTyp.meilenstein, existingNames)



        Else
            ' Merken, was ggf. das Filterkriterium war 
            sKeyMilestones = filterBox.Text

            ' Merken welches die selektierten Phasen waren 
            selectedMilestones.Clear()
            For Each element As String In ListBox1.SelectedItems
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
                existingNames.Clear()
                filterBox.Text = ""
                chkbxShowObjects.Text = "Werte im Projekt anzeigen"
                chkbxCreateCharts.Text = "Summen-Chart"
                chkbxOneChart.Text = "Alles in einem Chart"

                For i = 1 To RoleDefinitions.Count
                    existingNames.Add(RoleDefinitions.getRoledef(i).name)
                Next

                Call rebuildFormerState(PTauswahlTyp.Rolle, existingNames)



            Else
                ' Merken, was ggf. das Filterkriterium war 
                sKeyRoles = filterBox.Text

                ' Merken welches die selektierten Phasen waren 
                selectedRoles.Clear()
                For Each element As String In ListBox1.SelectedItems
                    selectedRoles.Add(element, element)
                Next

            End If
        End If

    End Sub

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
                existingNames.Clear()
                filterBox.Text = ""
                chkbxShowObjects.Text = "Werte im Projekt anzeigen"
                chkbxCreateCharts.Text = "Summen-Chart"
                chkbxOneChart.Text = "Alles in einem Chart"

                For i = 1 To CostDefinitions.Count
                    existingNames.Add(CostDefinitions.getCostdef(i).name)
                Next

                Call rebuildFormerState(PTauswahlTyp.Kostenart, existingNames)

            Else
                ' Merken, was ggf. das Filterkriterium war 
                sKeyCosts = filterBox.Text

                ' Merken welches die selektierten Phasen waren 
                selectedCosts.Clear()
                For Each element As String In ListBox1.SelectedItems
                    selectedCosts.Add(element, element)
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
            filterBox.Text = ""

            If rdbPhases.Checked Then
                selectedPhases.Clear()
            ElseIf rdbMilestones.Checked Then
                selectedMilestones.Clear()
            ElseIf rdbRoles.Checked Then
                selectedRoles.Clear()
            Else
                selectedCosts.Clear()
            End If
        End If

        
        'MyBase.Close()

    End Sub



    Private Sub chkbxCreateCharts_CheckedChanged(sender As Object, e As EventArgs) Handles chkbxCreateCharts.CheckedChanged

        If chkbxCreateCharts.Checked = True Then
            chkbxOneChart.Visible = True
        Else
            chkbxOneChart.Visible = False
            If chkbxShowObjects.Checked = False Then
                chkbxShowObjects.Checked = True
            End If
        End If


    End Sub

    Private Sub chkbxShowObjects_CheckedChanged(sender As Object, e As EventArgs) Handles chkbxShowObjects.CheckedChanged

        If chkbxShowObjects.Checked = False Then

            If chkbxCreateCharts.Checked = False Then
                chkbxCreateCharts.Checked = True
            End If

        End If

    End Sub

    Private Sub filterBox_TextChanged(sender As Object, e As EventArgs) Handles filterBox.TextChanged

        Dim suchstr As String = filterBox.Text

        If filterBox.Text = "" Then
            ListBox1.Items.Clear()
            For Each s As String In existingNames
                ListBox1.Items.Add(s)
            Next
        Else
            ListBox1.Items.Clear()
            For Each s As String In existingNames
                If s.Contains(suchstr) Then
                    ListBox1.Items.Add(s)
                End If
            Next
        End If

    End Sub


    Private Sub pictureRoles_Click(sender As Object, e As EventArgs) Handles pictureRoles.Click
        If Me.rdbRoles.Checked = False Then
            rdbRoles.Checked = True
        End If
    End Sub

    Private Sub picturePhasen_Click(sender As Object, e As EventArgs) Handles picturePhasen.Click
        If Me.rdbPhases.Checked = False Then
            rdbPhases.Checked = True
        End If
    End Sub

    Private Sub pictureMilestones_Click(sender As Object, e As EventArgs) Handles pictureMilestones.Click
        If Me.rdbMilestones.Checked = False Then
            rdbMilestones.Checked = True
        End If
    End Sub

    Private Sub pictureCosts_Click(sender As Object, e As EventArgs) Handles pictureCosts.Click
        If Me.rdbCosts.Checked = False Then
            Me.rdbCosts.Checked = True
        End If
    End Sub

    Private Sub pictureZoom_Click(sender As Object, e As EventArgs) Handles pictureZoom.Click
        filterBox.Text = ""
    End Sub

    Private Sub rebuildFormerState(ByVal typ As Integer, ByVal listOfNames As Collection)

        Dim searchkey As String = ""
        Dim tmpCollection As New Collection
        Dim i As Integer

        Select Case typ
            Case PTauswahlTyp.phase
                searchkey = sKeyPhases
                tmpCollection = selectedPhases

            Case PTauswahlTyp.meilenstein
                searchkey = sKeyMilestones
                tmpCollection = selectedMilestones

            Case PTauswahlTyp.Rolle
                searchkey = sKeyRoles
                tmpCollection = selectedRoles

            Case PTauswahlTyp.Kostenart
                searchkey = sKeyCosts
                tmpCollection = selectedCosts

        End Select

        If searchkey.Length > 0 Then

            For Each s As String In listOfNames
                If s.Contains(searchkey) Then
                    ListBox1.Items.Add(s)
                End If
            Next

        Else
            For i = 1 To listOfNames.Count
                ListBox1.Items.Add(listOfNames.Item(i))
            Next
        End If

        ' Filter Box Test setzen 

        filterBox.Text = searchkey

        ' jetzt prüfen, ob selectedphases bereits etwas enthält
        ' wenn ja, dann werden diese Items selektiert
        For Each element As String In tmpCollection
            ListBox1.SelectedItem = element
        Next
    End Sub

    Private Sub BackgroundWorker1_Disposed(sender As Object, e As EventArgs) Handles BackgroundWorker1.Disposed

        

    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim vorlagenDateiName As String = CType(e.Argument, String)


        With awinSettings
            Call createPPTSlidesFromConstellation(vorlagenDateiName, _
                                                  selectedPhases, selectedMilestones, selectedRoles, selectedCosts, _
                                                  .mppShowMsName, .mppShowProjectLine, .mppShowAmpel, .mppShowMsDate, .mppStrict, _
                                                  worker, e)
        End With
        

    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusLabel.Text = CType(re.Result, String)

    End Sub

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

    End Sub

    
    Private Sub einstellungen_Click(sender As Object, e As EventArgs) Handles einstellungen.Click

        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult

        dialogreturn = mppFrm.ShowDialog


    End Sub
End Class