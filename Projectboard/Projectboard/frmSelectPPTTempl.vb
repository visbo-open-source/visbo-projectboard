Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports System.Math
Imports MongoDbAccess
'Imports WpfWindow
'Imports WPFPieChart
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel


Public Class frmSelectPPTTempl

    Public listOfVorlagen As New Collection
    Public calledfrom As String
    Public awinSelection As Excel.ShapeRange

    Private Sub frmSelectPPTTempl_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dateiName As String = ""
        Dim dirname As String = ""

        ' hier wird  unterschieden, ob Projekt- oder Portfolio Report
        If calledfrom = "Portfolio" Then
            dirname = awinPath & RepPortfolioVorOrdner
        ElseIf calledfrom = "Projekt" Then
            dirname = awinPath & RepProjectVorOrdner
        End If

        ' jetzt werden die ProjektReport- bzw. PortfolioReport-Vorlagen ausgelesen 

        Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
        Try
            Dim i As Integer
            For i = 1 To listOfVorlagen.Count
                dateiName = Dir(listOfVorlagen.Item(i - 1))
                RepVorlagenDropbox.Items.Add(dateiName)
            Next i
        Catch ex As Exception
            'Call MsgBox(ex.Message & ": " & dateiName)
        End Try

    End Sub

    Private Sub createReport_Click(sender As Object, e As EventArgs) Handles createReport.Click

        Dim request As New Request(awinSettings.databaseName)
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
        If calledfrom = "Portfolio" Then
            dirName = awinPath & RepPortfolioVorOrdner
            vorlagenDateiName = dirName & "\" & RepVorlagenDropbox.Text
            Try
                createReport.Enabled = False
                RepVorlagenDropbox.Enabled = False
                Me.Cursor = Cursors.WaitCursor

                BackgroundWorker1.RunWorkerAsync(vorlagenDateiName)

            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

        ElseIf calledfrom = "Projekt" Then
            dirName = awinPath & RepProjectVorOrdner
            vorlagenDateiName = dirName & "\" & RepVorlagenDropbox.Text

            Try
                'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
                awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
            Catch ex As Exception
                awinSelection = Nothing
            End Try

            Try
                createReport.Enabled = False
                RepVorlagenDropbox.Enabled = False
                Me.Cursor = Cursors.WaitCursor

                BackgroundWorker2.RunWorkerAsync(vorlagenDateiName)
                'Call createPPTSlidesFromConstellation(vorlagenDateiName)

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
       

        Call MsgBox("Berichterstellung wurde beendet")
        MyBase.Close()

    End Sub


    Private Sub RepVorlagenDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RepVorlagenDropbox.SelectedIndexChanged
        ' hier muss die selektierte Vorlage genommen werden, um damit den dann bei OK-Button Click den Report anzustoßen
        Dim newTemplate As String = RepVorlagenDropbox.Text
    End Sub


    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

        Dim vorlagenDateiName As String = CType(e.Argument, String)

        'Call createPPTSlidesFromConstellation(vorlagenDateiName, worker, e)
        Dim tmpCollection As New Collection

        With awinSettings
            Call createPPTSlidesFromConstellation(vorlagenDateiName, _
                                                  tmpCollection, tmpCollection, tmpCollection, tmpCollection, _
                                                  .mppShowMsName, .mppShowProjectLine, .mppShowAmpel, .mppShowMsDate, .mppStrict, _
                                                  worker, e)
        End With
        


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

        Call createPPTReportFromProjects(vorlagenDateiName, worker, e)


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
End Class