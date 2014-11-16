Imports ProjectBoardDefinitions

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


    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim myCollection As Collection
        Dim repObj As Excel.ChartObject

        appInstance.EnableEvents = False
        enableOnUpdate = False

        If ListBox1.SelectedItems.Count > 0 And showRangeRight - showRangeLeft > 5 Then

            If Me.rdbPhases.Checked = True Then

                If chkbxShowObjects.Checked = True Then

                    ' alles in einem Chart anzeigen 
                    selectedPhases.Clear()
                    'For Each element As String In ListBox1.SelectedItems
                    For Each element As String In selectedPhases
                        selectedPhases.Add(element, element)
                    Next

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

End Class