Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports System.Math
Imports DBAccLayer
'Imports WpfWindow
'Imports WPFPieChart
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Microsoft.Office.Interop


Public Class frmSelectPPTTempl

    Public listOfVorlagen As New Collection
    Public calledfrom As String
    Public awinSelection As Excel.ShapeRange

    Private Sub frmSelectPPTTempl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dateiName As String = ""
        Dim dirname As String = ""
        Dim paramType As Boolean

        ' hier wird  unterschieden, ob Projekt- oder Portfolio Report
        If calledfrom = "Portfolio1" Then
            dirname = awinPath & RepPortfolioVorOrdner
            paramType = False
            Me.einstellungen.Visible = False
        ElseIf calledfrom = "Portfolio2" Then
            dirname = awinPath & RepPortfolioVorOrdner
            paramType = True
            Me.einstellungen.Visible = False
        ElseIf calledfrom = "Projekt" Then
            dirname = awinPath & RepProjectVorOrdner
            Me.einstellungen.Visible = True
        Else
            dirname = awinPath & RepProjectVorOrdner
            Me.einstellungen.Visible = True
        End If

        ' jetzt werden die ProjektReport- bzw. PortfolioReport-Vorlagen ausgelesen 

        Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
        Try
            Dim i As Integer
            For i = 1 To listOfVorlagen.Count
                dateiName = Dir(listOfVorlagen.Item(i - 1))
                If calledfrom = "Projekt" Then
                    If Not dateiName.Contains("Typ II") Then
                        RepVorlagenDropbox.Items.Add(dateiName)
                    End If
                    'RepVorlagenDropbox.Items.Add(dateiName)
                ElseIf calledfrom = "Portfolio1" Or calledfrom = "Portfolio2" Then
                    If Not dateiName.Contains("Typ II") Then
                        RepVorlagenDropbox.Items.Add(dateiName)
                    End If
                Else
                    If dateiName.Contains("Typ II") Then
                        RepVorlagenDropbox.Items.Add(dateiName)
                    End If
                End If

            Next i
        Catch ex As Exception
            'Call MsgBox(ex.Message & ": " & dateiName)
        End Try

    End Sub

    Private Sub createReport_Click(sender As Object, e As EventArgs) Handles createReport.Click

        ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)


        'Dim singleShp As Excel.Shape
        'Dim hproj As clsProjekt
        Dim vglName As String = " "
        'Dim pName As String, variantName As String
        Dim vorlagenDateiName As String = ""
        Dim dirName As String


        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        ' hier muss unterschieden werden, ob Projekt oder Portfolio-Report soll erzeugt werden
        If calledfrom = "Portfolio1" Or calledfrom = "Portfolio2" Then
            dirName = awinPath & RepPortfolioVorOrdner
            vorlagenDateiName = dirName & "\" & RepVorlagenDropbox.Text
            Try
                createReport.Enabled = False
                RepVorlagenDropbox.Enabled = False
                Me.Cursor = Cursors.WaitCursor

                'Call PPTstarten()

                BackgroundWorker1.RunWorkerAsync(vorlagenDateiName)

            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

        Else
            dirName = awinPath & RepProjectVorOrdner
            vorlagenDateiName = dirName & "\" & RepVorlagenDropbox.Text

            'awinSettings.eppExtendedMode = True

            Try
                If selectedProjekte.Count <= 0 Then
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                    For Each tmpname In awinSelection
                        Try
                            selectedProjekte.Add(ShowProjekte.getProject(tmpname), False)
                        Catch ex As Exception

                        End Try
                    Next
                End If

            Catch ex As Exception
                awinSelection = Nothing
                selectedProjekte.Clear(False)
            End Try

            Try
                createReport.Enabled = False
                RepVorlagenDropbox.Enabled = False
                Me.Cursor = Cursors.WaitCursor

                'Call PPTstarten()

                BackgroundWorker2.RunWorkerAsync(vorlagenDateiName)


            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU
        'MyBase.Close()
    End Sub

    Private Sub SelectAbbruch_Click(sender As Object, e As EventArgs) Handles SelectAbbruch.Click

        createReport.Enabled = True
        RepVorlagenDropbox.Enabled = True
        Me.Cursor = Cursors.Arrow

        Me.BackgroundWorker1.CancelAsync()
        Me.BackgroundWorker2.CancelAsync()


        With appInstance
            If Not .EnableEvents Then
                .EnableEvents = True
            End If

            If Not .ScreenUpdating Then
                .ScreenUpdating = True
            End If
        End With

        selectedProjekte.Clear(False)
        'Call MsgBox("Berichterstellung wurde beendet")
        MyBase.Close()

    End Sub


    Private Sub RepVorlagenDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RepVorlagenDropbox.SelectedIndexChanged
        ' hier muss die selektierte Vorlage genommen werden, um damit den dann bei OK-Button Click den Report anzustoßen
        Dim newTemplate As String = RepVorlagenDropbox.Text
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

        Dim vorlagenDateiName As String = CType(e.Argument, String)

        currentReportProfil.name = "Last"

        Dim tmpCollection As New Collection

        currentReportProfil.Phases = copyColltoSortedList(tmpCollection)
        currentReportProfil.Milestones = copyColltoSortedList(tmpCollection)
        currentReportProfil.Roles = copyColltoSortedList(tmpCollection)
        currentReportProfil.Costs = copyColltoSortedList(tmpCollection)
        currentReportProfil.Typs = copyColltoSortedList(tmpCollection)
        currentReportProfil.BUs = copyColltoSortedList(tmpCollection)

        currentReportProfil.CalendarVonDate = StartofCalendar

        Dim vonDate As Date = getDateofColumn(showRangeLeft, False)
        Dim bisDate As Date = getDateofColumn(showRangeRight, True)

        Try
            currentReportProfil.calcRepVonBis(vonDate, bisDate)
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

        Try

            With awinSettings

                If .mppSortiertDauer Then
                    .mppShowAllIfOne = True
                End If

                currentReportProfil.ProjectLine = .mppShowProjectLine
                currentReportProfil.AllIfOne = .mppShowAllIfOne
                currentReportProfil.Ampeln = .mppShowAmpel
                currentReportProfil.UseAbbreviation = .mppUseAbbreviation

                currentReportProfil.PhName = .mppShowPhName
                currentReportProfil.PhDate = .mppShowPhDate
                currentReportProfil.MSName = .mppShowMsName
                currentReportProfil.MSDate = .mppShowMsDate
                currentReportProfil.UseAbbreviation = .mppUseAbbreviation
                currentReportProfil.KwInMilestone = .mppKwInMilestone


                currentReportProfil.VLinien = .mppVertikalesRaster
                currentReportProfil.ShowHorizontals = .mppShowHorizontals
                currentReportProfil.Legend = .mppShowLegend
                currentReportProfil.OnePage = .mppOnePage

                currentReportProfil.SortedDauer = .mppSortiertDauer
                currentReportProfil.ExtendedMode = .mppExtendedMode
                currentReportProfil.FullyContained = .mppFullyContained

                currentReportProfil.projectsWithNoMPmayPass = .mppProjectsWithNoMPmayPass

                ' Dateiname eliminieren, ohne Pfadangaben im ReportProfil speichern
                Dim hstr() As String
                hstr = Split(vorlagenDateiName, "\")
                currentReportProfil.PPTTemplate = hstr(hstr.Length - 1)
                currentReportProfil.isMpp = True

            End With

            Call createPPTSlidesFromConstellation(vorlagenDateiName, _
                                                   tmpCollection, tmpCollection, _
                                                   tmpCollection, tmpCollection, _
                                                   tmpCollection, tmpCollection, True, _
                                                   worker, e)

 
        Catch ex As Exception
            Call MsgBox("Fehler " & ex.Message)
            Call MsgBox(" in BAckground Worker ...")
        End Try





    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        'Me.statusNotification.Text = e.ProgressPercentage.ToString & "% erledigt"

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusNotification.Text = CType(re.Result, String)

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted


        createReport.Enabled = True
        RepVorlagenDropbox.Enabled = True
        Me.Cursor = Cursors.Arrow


    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

        Dim vorlagenDateiName As String = CType(e.Argument, String)
        Dim tmpCollection As New Collection

        currentReportProfil.name = "Last"
        currentReportProfil.Phases = copyColltoSortedList(tmpCollection)
        currentReportProfil.Milestones = copyColltoSortedList(tmpCollection)
        currentReportProfil.Roles = copyColltoSortedList(tmpCollection)
        currentReportProfil.Costs = copyColltoSortedList(tmpCollection)
        currentReportProfil.Typs = copyColltoSortedList(tmpCollection)
        currentReportProfil.BUs = copyColltoSortedList(tmpCollection)

        currentReportProfil.CalendarVonDate = StartofCalendar

        
        Try
            'currentReportProfil.calcRepVonBis(vonDate, bisDate)
            currentReportProfil.calcRepVonBis(StartofCalendar, StartofCalendar)
        Catch ex As Exception
            ' tk keine Exception, weil sonst kein Einzelprojekt Report gemacht werdne kann 
            Throw New ArgumentException(ex.Message)
        End Try
        Try

            With awinSettings

                If .mppSortiertDauer Then
                    .mppShowAllIfOne = True
                End If

                currentReportProfil.ProjectLine = .mppShowProjectLine
                currentReportProfil.AllIfOne = .mppShowAllIfOne
                currentReportProfil.Ampeln = .mppShowAmpel
                currentReportProfil.UseAbbreviation = .mppUseAbbreviation

                currentReportProfil.PhName = .mppShowPhName
                currentReportProfil.PhDate = .mppShowPhDate
                currentReportProfil.MSName = .mppShowMsName
                currentReportProfil.MSDate = .mppShowMsDate
                currentReportProfil.UseAbbreviation = .mppUseAbbreviation
                currentReportProfil.KwInMilestone = .mppKwInMilestone


                currentReportProfil.VLinien = .mppVertikalesRaster
                currentReportProfil.ShowHorizontals = .mppShowHorizontals
                currentReportProfil.Legend = .mppShowLegend
                currentReportProfil.OnePage = .mppOnePage

                currentReportProfil.SortedDauer = .mppSortiertDauer
                currentReportProfil.ExtendedMode = .mppExtendedMode
                currentReportProfil.FullyContained = .mppFullyContained

                currentReportProfil.projectsWithNoMPmayPass = .mppProjectsWithNoMPmayPass

                ' Dateiname eliminieren, ohne Pfadangaben im ReportProfil speichern
                Dim hstr() As String
                hstr = Split(vorlagenDateiName, "\")
                currentReportProfil.PPTTemplate = hstr(hstr.Length - 1)
                currentReportProfil.isMpp = False
            End With


            Call createPPTReportFromProjects(vorlagenDateiName, _
                                             tmpCollection, tmpCollection, _
                                             tmpCollection, tmpCollection, _
                                             tmpCollection, tmpCollection, _
                                             worker, e)

        Catch ex As Exception

        End Try



    End Sub

    Private Sub BackgroundWorker2_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker2.ProgressChanged
        'Me.statusNotification.Text = e.ProgressPercentage.ToString & "% erledigt"

        Dim re As System.ComponentModel.DoWorkEventArgs = CType(e.UserState, System.ComponentModel.DoWorkEventArgs)
        Me.statusNotification.Text = CType(re.Result, String)

    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted


        createReport.Enabled = True
        RepVorlagenDropbox.Enabled = True
        Me.Cursor = Cursors.Arrow


    End Sub

    Private Sub statusNotification_TextChanged(sender As Object, e As EventArgs) Handles statusNotification.TextChanged

    End Sub

    ''' <summary>
    ''' ruft das Formular auf, um die Einstellungen für das ProjektReporting vorzunehmen  
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub einstellungen_Click(sender As Object, e As EventArgs) Handles einstellungen.Click

        Dim mppFrm As New frmMppSettings
        Dim dialogreturn As DialogResult

        If calledfrom = "MS-Project" Then
            mppFrm.calledfrom = calledfrom
        Else
            mppFrm.calledfrom = "frmSelectPPTTempl"
        End If

        dialogreturn = mppFrm.ShowDialog


    End Sub
End Class