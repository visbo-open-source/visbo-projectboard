Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports MongoDbAccess






Public Class ThisWorkbook
    ' Copyright Philipp Koytek et al. 
    ' 2012 ff
    ' Nicht authorisierte Verwendung nicht gestattet 

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

    Private Sub ThisWorkbook_ActivateEvent() Handles Me.ActivateEvent

        Application.DisplayFormulaBar = False

    End Sub

    ''' <summary>
    ''' stellt sicher, daß die Excel Settings in anderen Workbooks wieder gelten
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ThisWorkbook_Deactivate() Handles Me.Deactivate

        Application.DisplayFormulaBar = True

        With Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 0
            .DisplayWorkbookTabs = True
            .GridlineColor = RGB(220, 220, 220)
            .DisplayHeadings = True

            Try
                .FreezePanes = False
            Catch ex As Exception

            End Try


        End With

    End Sub

    Private Sub ThisWorkbook_Startup() Handles Me.Startup

        'Dim cbar As CommandBar


        appInstance = Application

        myProjektTafel = appInstance.ActiveWorkbook.Name

        Dim path As String = CType(appInstance.ActiveWorkbook, Excel.Workbook).Path
       
        ' die Short Cut Menues aus Excel werden hier nicht mehr de-aktiviert 
        ' das wird jetzt nur in Tabelle1, also der Projekt-Tafel gemacht ...
        ' in anderen Excel Sheets ist das weiterhin aktiv 
        'For Each cbar In appInstance.CommandBars

        '    If cbar.Type = MsoBarType.msoBarTypePopup Then
        '        cbar.Enabled = False
        '    End If
        'Next

        magicBoardCmdBar.cmdbars = appInstance.CommandBars



        Try


            appInstance.ScreenUpdating = False

            ' hier werden die Settings aus der Datei ProjectboardConfig.xml ausgelesen.
            ' falls die nicht funktioniert, so werden die My.Settings ausgelesen und verwendet.

            If Not readawinSettings(path) Then

                awinSettings.databaseURL = My.Settings.mongoDBURL
                awinSettings.databaseName = My.Settings.mongoDBname
                awinSettings.globalPath = My.Settings.globalPath
                awinSettings.awinPath = My.Settings.awinPath
                awinSettings.visboTaskClass = My.Settings.TaskClass
                awinSettings.visboAbbreviation = My.Settings.VISBOAbbreviation
                awinSettings.visboAmpel = My.Settings.VISBOAmpel
                awinSettings.visboAmpelText = My.Settings.VISBOAmpelText
                awinSettings.visboresponsible = My.Settings.VISBOresponsible
                awinSettings.visbodeliverables = My.Settings.VISBOdeliverables
                awinSettings.visbopercentDone = My.Settings.VISBOpercentDone
                awinSettings.visboMapping = My.Settings.VISBOMapping
                awinSettings.visboDebug = My.Settings.VISBODebug

            End If

            Call awinsetTypen("ProjectBoard")

        Catch ex As Exception

            Call MsgBox(ex.Message)
            appInstance.Quit()
        Finally
            appInstance.ScreenUpdating = True
            appInstance.ShowChartTipNames = True
            appInstance.ShowChartTipValues = True
        End Try

        anzahlCalls = 0


        'Call awinRightClickinPortfolioAendern()
        ' Änderung 19.1.15 Right Click in den Charts de-aktivieren für BMW 
        'Call awinRightClickinPRCCharts()

    End Sub

    


    Private Sub ThisWorkbook_Open() Handles Me.Open



        If Application.EnableEvents Then
        Else
            Application.EnableEvents = True
        End If

        CType(Application.Workbooks(myProjektTafel), Excel.Workbook).Activate()

        CType(Application.Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet).Activate()

        projectboardWindows(PTwindows.mpt) = Application.ActiveWindow

        With projectboardWindows(PTwindows.mpt)
            .DisplayHeadings = False
            '.Caption = windowNames(PTwindows.mpt)
            .Caption = bestimmeWindowCaption(PTwindows.mpt)
            .DisplayWorkbookTabs = False
            '.ScrollRow = 1
            '.ScrollColumn = 1
            .Visible = True
            .Zoom = 100
            .WindowState = XlWindowState.xlMaximized
        End With


        'If appInstance.Windows.Count < 2 Then
        '    Try
        '        With appInstance
        '            .Windows.Arrange(Excel.XlArrangeStyle.xlArrangeStyleTiled)
        '            .Windows(1).WindowState = XlWindowState.xlMaximized
        '        End With
        '    Catch ex As Exception
        '        ' 
        '    End Try

        'End If


        ' hier wird die Projekt Tafel so dargestellt, daß Zeitraum zu sehen ist ... und ein späteres Diagramm 
        ' Änderung 29.06.14 hier nicht mehr notwendig 
        ' Call awinScrollintoView()


    End Sub

    Private Sub ThisWorkbook_BeforeSave(SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Me.BeforeSave

        Cancel = True


    End Sub

    Private Sub ThisWorkbook_BeforeClose(ByRef Cancel As Boolean) Handles Me.BeforeClose


        Dim projektespeichern As New frmProjekteSpeichern
        Dim returnValue As DialogResult
        Dim cancelAbbruch As Boolean = False

        If loginErfolgreich Then

            ' tk: nur Fragen , wenn die Datenbank überhaupt läuft 
            Try

                If Not noDB Then
                    'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

                    If CType(mongoDBAcc, Request).pingMongoDb() And AlleProjekte.Count > 0 Then
                        returnValue = projektespeichern.ShowDialog

                        If returnValue = DialogResult.Yes Then

                            Call StoreAllProjectsinDB(True)

                        ElseIf returnValue = DialogResult.Cancel Then

                            cancelAbbruch = True
                        End If

                    Else
                        If awinSettings.englishLanguage Then
                            Call MsgBox("no projects to store ...")
                        Else
                            Call MsgBox("keine Projekte zu speichern ...")
                        End If


                    End If

                    If Not cancelAbbruch Then
                        ' die temporären Schutz
                        If CType(mongoDBAcc, Request).cancelWriteProtections(dbUsername) Then
                            If awinSettings.visboDebug Then
                                Call MsgBox("Ihre vorübergehenden Schreibsperren wurden aufgehoben")
                            End If
                        End If
                    End If
                    

                End If


            Catch ex As Exception

            End Try

            

            If cancelAbbruch Then
                Cancel = True
            Else
                appInstance.ScreenUpdating = False
                ' hier sollen jetzt noch die Phasen weggeschrieben werden 
                Try
                    'Call awinWritePhaseDefinitions()
                    Call awinWritePhaseMilestoneDefinitions()
                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Error when writing Projectboard Customization File")
                    Else
                        Call MsgBox("Fehler bei Schreiben Projectboard Customization File")
                    End If

                End Try
                appInstance.ScreenUpdating = True
            End If

        End If

        If Not cancelAbbruch Then
            Call awinKontextReset()
            ' hier wird festgelegt, dass Projectboard.xlsx beim Schließen nicht gespeichert wird, und auch nicht nachgefragt wird.
            'appInstance.EnableEvents = False

            Dim WB As Workbook
            For Each WB In Application.Workbooks
                If WB.Name = myProjektTafel Then
                    Try
                        WB.Saved = True
                    Catch ex As Exception

                    End Try

                End If

            Next


            Application.DisplayAlerts = False
            'Application.Quit()


        End If



    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown

        Try

            Dim cbar As CommandBar

            'die Short Cut Menues aus Excel alle wieder aktivieren ...

            For Each cbar In appInstance.CommandBars

                If cbar.Type = MsoBarType.msoBarTypePopup Then
                    cbar.Enabled = True
                End If
            Next

            'Call MsgBox(" in shutdown")


            Try
                Application.DisplayFormulaBar = True
            Catch ex As Exception

            End Try


            With Application.ActiveWindow
                Try
                    .SplitColumn = 0
                    .SplitRow = 0
                Catch ex As Exception

                End Try

                Try
                    .DisplayWorkbookTabs = True
                Catch ex As Exception

                End Try

                Try
                    .GridlineColor = RGB(220, 220, 220)
                Catch ex As Exception

                End Try

                Try
                    .FreezePanes = False
                Catch ex As Exception

                End Try

                Try
                    .DisplayHeadings = True
                Catch ex As Exception

                End Try

            End With

            appInstance.ShowChartTipNames = True
            appInstance.ShowChartTipValues = True

            'Dim anzWindows As Integer = appInstance.Windows.Count





            appInstance.ScreenUpdating = True

            If Application.Workbooks.Count <= 1 Then
                Dim a As Integer = Application.Workbooks.Count
                'Dim name asstring = Application.Workbooks(1).name
            End If



        Catch ex As Exception


        End Try

    End Sub


    ' ''' <summary>
    ' ''' definiert die Windows und Views, die benötigt werden 
    ' ''' es ist die Tabelle1=mpt aktiviert 
    ' ''' </summary>
    ' ''' <remarks></remarks>
    'Private Sub defineVisboWindowViews()

    '    Dim formerEE As Boolean = appInstance.EnableEvents
    '    Dim formerSU As Boolean = appInstance.ScreenUpdating
    '    Dim formereOU As Boolean = enableOnUpdate

    '    If enableOnUpdate Then
    '        enableOnUpdate = False
    '    End If

    '    If appInstance.EnableEvents Then
    '        appInstance.EnableEvents = False
    '    End If

    '    If appInstance.ScreenUpdating Then
    '        appInstance.ScreenUpdating = False
    '    End If

    '    ' jetzt werden die Windows aufgebaut ...

    '    ' dann werden alle auf invisible gesetzt , bis auf projectboardWindows(mpt)


    '    Dim visboWorkbook As Excel.Workbook = appInstance.Workbooks.Item(myProjektTafel)


    '    'projectboardWindows(PTwindows.mpt) = appInstance.ActiveWindow.NewWindow
    '    projectboardWindows(PTwindows.mpt) = appInstance.ActiveWindow


    '    ' Aus dem aktuellen Window ein benanntes Window machen 

    '    projectboardWindows(PTwindows.mptpr) = appInstance.ActiveWindow.NewWindow

    '    ' jetzt auf das Worksheet positionieren ...
    '    CType(visboWorkbook.Worksheets(arrWsNames(ptTables.mptPrCharts)), Excel.Worksheet).Activate()

    '    With projectboardWindows(PTwindows.mptpr)
    '        .WindowState = Excel.XlWindowState.xlNormal
    '        .EnableResize = True
    '        .DisplayHorizontalScrollBar = True
    '        .DisplayVerticalScrollBar = True
    '        .DisplayFormulas = False
    '        .DisplayHeadings = False
    '        .DisplayGridlines = False
    '        .GridlineColor = RGB(255, 255, 255)
    '        .DisplayWorkbookTabs = False
    '        .Caption = bestimmeWindowCaption(PTwindows.mptpr)
    '        .Visible = False
    '    End With

    '    ' Aufbau des Windows windowNames(4): Charts
    '    projectboardWindows(PTwindows.mptpf) = appInstance.ActiveWindow.NewWindow

    '    ' jetzt das Worksheet aktivieren ...
    '    visboWorkbook.Worksheets.Item(arrWsNames(ptTables.mptPfCharts)).activate()

    '    With projectboardWindows(PTwindows.mptpf)
    '        .WindowState = Excel.XlWindowState.xlNormal
    '        .EnableResize = True
    '        .DisplayHorizontalScrollBar = True
    '        .DisplayVerticalScrollBar = True
    '        .DisplayGridlines = False
    '        .DisplayHeadings = False
    '        .DisplayRuler = False
    '        .DisplayOutline = False
    '        .DisplayWorkbookTabs = False
    '        .Caption = bestimmeWindowCaption(PTwindows.mptpf)
    '        .Visible = False
    '    End With


    '    ' jetzt das Sheet Multiprojekt-Tafel aktivieren
    '    visboWorkbook.Worksheets.Item(arrWsNames(ptTables.MPT)).activate()

    '    'jetzt das MPT Sheet wieder holen 
    '    With projectboardWindows(PTwindows.mpt)
    '        .WindowState = XlWindowState.xlMaximized
    '        .Activate()
    '    End With

    '    ' wieder auf den Ausgangszustand setzen ... 
    '    With appInstance
    '        If .EnableEvents <> formerEE Then
    '            .EnableEvents = formerEE
    '        End If

    '        If .ScreenUpdating <> formerSU Then
    '            .ScreenUpdating = formerSU
    '        End If

    '        If enableOnUpdate <> formereOU Then
    '            enableOnUpdate = formereOU
    '        End If
    '    End With


    'End Sub
  

  


    Private bIShrankTheRibbon As Boolean
    Private Sub ThisWorkbook_WindowActivate(Wn As Microsoft.Office.Interop.Excel.Window) Handles Me.WindowActivate

        If appInstance.Version <> "14.0" Then
        
            If CStr(Wn.Caption).Contains("Chart") Then
                bIShrankTheRibbon = False
                appInstance.ExecuteExcel4Macro("SHOW.TOOLBAR(" & Chr(34) & "Ribbon" & Chr(34) & ",False)")
                bIShrankTheRibbon = True
            End If
        End If

    End Sub


    Private Sub ThisWorkbook_WindowDeactivate(Wn As Microsoft.Office.Interop.Excel.Window) Handles Me.WindowDeactivate
        If appInstance.Version <> "14.0" Then

            If CStr(Wn.Caption).Contains("Chart") Then
                If bIShrankTheRibbon Then
                    'appInstance.ExecuteExcel4Macro("SHOW.TOOLBAR(" & Chr(34) & "Ribbon" & Chr(34) & ",True)")
                    appInstance.ActiveWindow.WindowState = XlWindowState.xlNormal
                End If
            End If
        End If
    End Sub

    Private Sub ThisWorkbook_SheetDeactivate(Sh As Object) Handles Me.SheetDeactivate
        Dim a As Integer = -1
    End Sub
End Class
