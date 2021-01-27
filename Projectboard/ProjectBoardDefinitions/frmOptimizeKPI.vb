Imports Microsoft.Office.Core
Imports xlNS = Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Windows.Forms

Public Class frmOptimizeKPI

    Friend tmpListe As New SortedList(Of String, String)
    Friend kennung As String
    Public menueOption As Integer
    Friend calledFrom As String = "menu"


    Private Sub frmOptimizeKPI_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim tmpKennung As String
        Dim title As String
        Dim selectedIndex As Integer = 0
        Dim lfdNr As Integer = 0
        Dim tmpstr(3) As String

        ' hier müssen die Buttons sichtbar gesetzt werden
        Me.startButton.Visible = True
        Me.startButton.Top = Me.silverMedal.Top
        Me.startButton.Left = Me.silverMedal.Left

        Me.goldMedal.Visible = False
        Me.silverMedal.Visible = False
        Me.bronceMedal.Visible = False
        Me.abbruchButton.Visible = False
        Me.progressText.Text = ""


        ' dann muss die Dropbox aufgebaut werden
        ' alle Charts überprüfen: sind sie optimierbar, wenn ja, dann Aufnahme in die Liste 
        ' gibt es ein ActiveChart? ja, dann als Default Eintrag nehmen 



        With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.mptPfCharts)), xlNS.Worksheet)


            For Each chtobj As xlNS.ChartObject In CType(.ChartObjects, xlNS.ChartObjects)

                tmpKennung = chtobj.Name
                Try
                    If tmpKennung.Length > 0 Then
                        If DiagramList.contains(tmpKennung) Then
                            tmpstr = chtobj.Chart.ChartTitle.Text.Split(New Char() {CChar("(")}, 3)
                            title = tmpstr(0)
                            tmpListe.Add(tmpKennung, title)
                            Me.auswahlKPI.Items.Add(title)
                            lfdNr = lfdNr + 1
                        End If
                    End If
                Catch ex As Exception

                End Try

            Next

        End With

        ' jetzt soll das Element entsprechend gesetzt werden 
        If Me.auswahlKPI.Items.Count > 0 Then
            Me.auswahlKPI.SelectedIndex = selectedIndex
        Else
            Me.progressText.Text = "es gibt keine optimierbaren Charts ! "
            Me.startButton.Visible = False
            Me.auswahlKPI.Enabled = False
        End If

    End Sub

    Private Sub startButton_Click(sender As Object, e As EventArgs) Handles startButton.Click


        With Me

            auswahlKPI.Enabled = False
            startButton.Visible = False
            abbruchButton.Visible = True
            abbruchButton.Enabled = True
            Me.Cursor = Cursors.WaitCursor

        End With

        If DiagramList.contains(kennung) Then
            BackgroundWorker1.RunWorkerAsync(kennung)
        Else
            Call MsgBox("Fehler: Kennung nicht vorhanden: " & kennung)
        End If



    End Sub

    Private Sub BackgroundWorker1_Disposed(sender As Object, e As EventArgs) Handles BackgroundWorker1.Disposed

        Me.progressText.Text = Me.progressText.Text & " --> durch User abgebrochen"

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim bgKennung As String = CType(e.Argument, String)
        Dim myCollection As Collection
        Dim diagrammTyp As String
        Dim tmpDiagramm As clsDiagramm



        tmpDiagramm = DiagramList.getDiagramm(bgKennung)

        If IsNothing(tmpDiagramm) Then
        Else
            With tmpDiagramm
                myCollection = .gsCollection
                diagrammTyp = .diagrammTyp
            End With

            ' Aufruf der Optimierungs-Schleife ....

            enableOnUpdate = False
            If menueOption = 1 Then
                ' Varianten Optimierung
                Call awinCalcOptimizationVarianten(diagrammTyp, myCollection, worker, e)
            Else
                ' Phasen Freiraum Optimierung 
                Dim OptimierungsErgebnis As New SortedList(Of String, clsOptimizationObject)
                Call awinCalcOptimizationElemFreiheitsgrade(diagrammTyp, myCollection, OptimierungsErgebnis, worker, e)

                If OptimierungsErgebnis.Count > 0 Then


                    For Each kvp In OptimierungsErgebnis
                        Try

                            With kvp.Value

                                Dim pName As String = kvp.Value.projectName
                                Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                                Dim phaseName As String = CStr(myCollection.Item(1))

                                Dim phaseList As Collection = projectboardShapes.getPhaseList(pName)
                                Dim milestoneList As Collection = projectboardShapes.getMilestoneList(pName)

                                Call clearProjektinPlantafel(pName)

                                ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                                ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                                Dim tmpCollection As New Collection
                                Call ZeichneProjektinPlanTafel(tmpCollection, pName, hproj.tfZeile, phaseList, milestoneList)


                            End With
                        Catch ex As Exception
                            Call MsgBox("Projekt: " & kvp.Key & " : Startzeitpunkt liegt in der Vergangenheit ")
                        End Try

                    Next

                    Call visualisiereErgebnis()
                    OptimierungsErgebnis.Clear()


                Else
                    MsgBox("es waren keine Verbesserungen zu erzielen")
                End If



            End If
            ' Änderung tk 29.5: das war vorher auf false !? 
            enableOnUpdate = True
        End If
        




    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged


        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.progressText.Text = CType(re.Result, String)

    End Sub

    Private Sub abbruchButton_Click(sender As Object, e As EventArgs) Handles abbruchButton.Click

        auswahlKPI.Enabled = True
        Me.Cursor = Cursors.Arrow



        With Me

            auswahlKPI.Enabled = True
            startButton.Visible = True
            startButton.Enabled = True
            abbruchButton.Visible = False
            Me.Cursor = Cursors.Arrow

        End With

        Me.progressText.Text = "Optimierung wurde abgebrochen ... "

        Me.BackgroundWorker1.CancelAsync()


    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        Me.Cursor = Cursors.Arrow
        Me.auswahlKPI.Enabled = False

        Me.abbruchButton.Visible = False
        Me.startButton.Visible = False

        If menueOption = 1 Then
            ' Varianten Optimierung 
            If projectConstellations.Contains(autoSzenarioNamen(3)) Then
                Me.bronceMedal.Visible = True
                Me.silverMedal.Visible = True
                Me.goldMedal.Visible = True
            ElseIf projectConstellations.Contains(autoSzenarioNamen(2)) Then
                Me.silverMedal.Visible = True
                Me.goldMedal.Visible = True
            ElseIf projectConstellations.Contains(autoSzenarioNamen(1)) Then
                Me.goldMedal.Visible = True
            Else
                Me.progressText.Text = "es waren keine Verbesserungen zu erzielen !"
            End If
        Else
            Me.progressText.Text = "... Fertig ! "
        End If
        



    End Sub

    Private Sub auswahlKPI_SelectedValueChanged(sender As Object, e As EventArgs) Handles auswahlKPI.SelectedValueChanged

        Dim title As String

        title = CStr(Me.auswahlKPI.SelectedItem)
        kennung = tmpListe.ElementAt(tmpListe.IndexOfValue(title)).Key


    End Sub

    Private Sub goldMedal_Click(sender As Object, e As EventArgs) Handles goldMedal.Click

        appInstance.ScreenUpdating = False
        Call loadSessionConstellation(calcPortfolioKey(autoSzenarioNamen(1), ""), False, True)
        appInstance.ScreenUpdating = True

        Me.progressText.Text = "Geladen: " & autoSzenarioNamen(1)

    End Sub

    Private Sub silverMedal_Click(sender As Object, e As EventArgs) Handles silverMedal.Click

        appInstance.ScreenUpdating = False
        Call loadSessionConstellation(calcPortfolioKey(autoSzenarioNamen(2), ""), False, True)
        appInstance.ScreenUpdating = True

        Me.progressText.Text = "Geladen: " & autoSzenarioNamen(2)

    End Sub

    Private Sub bronceMedal_Click(sender As Object, e As EventArgs) Handles bronceMedal.Click

        appInstance.ScreenUpdating = False
        Call loadSessionConstellation(calcPortfolioKey(autoSzenarioNamen(3), ""), False, True)
        appInstance.ScreenUpdating = True

        Me.progressText.Text = "Geladen: " & autoSzenarioNamen(3)

    End Sub

    
End Class