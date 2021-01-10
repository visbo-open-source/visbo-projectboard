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
        Try
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
        Catch ex As Exception

        End Try


    End Sub

    Private Sub ThisWorkbook_Startup() Handles Me.Startup

        'Dim cbar As CommandBar


        appInstance = Application

        ' nicht visible setzen
        'appInstance.Visible = False

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
                awinSettings.DBWithSSL = My.Settings.mongoDBWithSSL
                awinSettings.proxyURL = My.Settings.proxyServerURL
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
                awinSettings.visboServer = My.Settings.VISBOServer
                awinSettings.userNamePWD = My.Settings.userNamePWD
                awinSettings.rememberUserPwd = My.Settings.rememberUserPWD

            End If



            ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
            awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
            If My.Settings.rememberUserPWD Then
                awinSettings.userNamePWD = My.Settings.userNamePWD
            Else
                awinSettings.userNamePWD = ""
            End If

            ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
            awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
            If My.Settings.rememberUserPWD Then
                awinSettings.userNamePWD = My.Settings.userNamePWD
            Else
                awinSettings.userNamePWD = ""
            End If

            ' Refresh von Projekte im Cache  in Minuten
            cacheUpdateDelay = 30

            'appInstance.EnableEvents = False
            Call awinsetTypen("ProjectBoard")
            'appInstance.EnableEvents = True

            'appInstance.Visible = True

        Catch ex As Exception

            appInstance.EnableEvents = True

            '   Call MsgBox(ex.Message)
            appInstance.Quit()
        Finally
            appInstance.ScreenUpdating = True
            appInstance.ShowChartTipNames = True
            appInstance.ShowChartTipValues = True
        End Try

        anzahlCalls = 0


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
            .WindowState = XlWindowState.xlMaximized
            .DisplayHeadings = False
            '.Caption = windowNames(PTwindows.mpt)
            .Caption = bestimmeWindowCaption(PTwindows.mpt)
            .DisplayWorkbookTabs = False
            '.ScrollRow = 1
            '.ScrollColumn = 1
            .Visible = True
            If .Width < 1100 Then
                .Zoom = 80
            ElseIf .Width < 1400 Then
                .Zoom = 90
            Else
                .Zoom = 100
            End If


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
        Dim err As New clsErrorCodeMsg

        If loginErfolgreich Then


            ' tk: nur Fragen , wenn die Datenbank überhaupt läuft 
            Try
                My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
                If awinSettings.rememberUserPwd Then
                    My.Settings.userNamePWD = awinSettings.userNamePWD
                    ' um die Settings abzuspeichern
                Else
                    My.Settings.userNamePWD = ""
                End If
                My.Settings.Save()

                My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
                If awinSettings.rememberUserPwd Then
                    My.Settings.userNamePWD = awinSettings.userNamePWD
                    ' um die Settings abzuspeichern
                Else
                    My.Settings.userNamePWD = ""
                End If
                My.Settings.Save()


                If Not noDB Then

                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                        ' nicht fragen - das führt nur zu sehr unangenehmen Überraschungen 

                    Else
                        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() And AlleProjekte.Count > 0 Then
                            returnValue = projektespeichern.ShowDialog

                            If returnValue = DialogResult.Yes Then

                                Call StoreAllProjectsinDB()

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
                    End If


                    ' ur:19.06.2019
                    If Not cancelAbbruch Then
                        ' die temporären Schutz
                        If CType(databaseAcc, DBAccLayer.Request).cancelWriteProtections(dbUsername, err) Then
                            If awinSettings.visboDebug Then
                                Call MsgBox("Ihre vorübergehenden Schreibsperren wurden aufgehoben")
                            End If
                        End If
                    End If


                End If


            Catch ex As Exception

            End Try

            Dim testresult As New Object

            Dim appResult As New SortedList(Of String, clsAppearance)
            Dim custfieldsResult As New clsCustomFieldDefinitions
            Dim customizeResult As New clsCustomization
            Dim customrolesResult As New clsCustomUserRoles
            Dim organisationResult As New clsOrganisation

            If cancelAbbruch Then
                Cancel = True
            Else
                ''Try

                ''    testresult = CType(databaseAcc, DBAccLayer.Request).retrieveAllVCSettingFromDB(err,
                ''                                                                           appResult,
                ''                                                                        custfieldsResult,
                ''                                                                           customizeResult,
                ''                                                                           customrolesResult,
                ''                                                                           organisationResult)
                ''Catch ex As Exception

                ''End Try

                ' dann wird das ProjectboardCustomization File wieder weggespeichert ... 
                If awinSettings.readWriteMissingDefinitions Then

                    appInstance.ScreenUpdating = False

                    ' hier sollen jetzt noch die Phasen und Meilensteine, die hinzugefügt wurden, weggeschrieben werden 
                    Try

                        Dim msgResult As New MsgBoxResult

                        If MilestoneDefsAndPhaseDefsAdded And
                         myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Then
                            If awinSettings.englishLanguage Then
                                msgResult = MsgBox("You want to save the added phases and milestone in the DB ?", vbYesNo)

                            Else
                                msgResult = MsgBox("Sollen die hinzugefügten Phasen und Meilensteine in der DB gespeichert werden?", vbYesNo)

                            End If

                            If msgResult = MsgBoxResult.Yes Then


                                ' jetzt wird geprüft, ob die missingPhaseDefinitions in PhaseDefinitions übertragen werden 
                                ' jetzt wird geprüft, ob die missingMilestoneDefinitions in MilestoneDefinitions übertragen werden 
                                If awinSettings.addMissingPhaseMilestoneDef Then

                                    Call addMissingDefs2Defs()

                                End If
                                'ur: 2019-09-02: nicht mehr in Customization file zurückschreiben, sondern in DB
                                Call awinWritePhaseMilestoneDefinitions()

                                Dim customizations As clsCustomization = get_customSettings()
                                Dim result As Boolean = False
                                result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(customizations,
                                                                                    CStr(settingTypes(ptSettingTypes.customization)),
                                                                                    CStr(settingTypes(ptSettingTypes.customization)),
                                                                                    Nothing,
                                                                                    err)
                                If result = False Then
                                    If awinSettings.englishLanguage Then
                                        Call MsgBox("Error when writing Customizations to DB ")
                                    Else
                                        Call MsgBox("Fehler bei Speichern der Customizations in die DB ")
                                    End If
                                End If
                            End If

                        End If

                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            Call MsgBox("Error when writing Customizations to DB ")
                        Else
                            Call MsgBox("Fehler bei Speichern der Customizations in die DB ")
                        End If

                    End Try
                    appInstance.ScreenUpdating = True
                End If

            End If

        End If

        If Not cancelAbbruch Then
            'Call awinKontextReset()
            ' hier wird festgelegt, dass Projectboard.xlsx beim Schließen nicht gespeichert wird, und auch nicht nachgefragt wird.
            'appInstance.EnableEvents = False

            ' ur:2020-11-23: hier sollte Logfile geschlossen werden.
            ' ''Call logfileSchliessen()

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
            ' Application.Quit()

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

            Dim err As New clsErrorCodeMsg

            Dim logoutErfolgreich As Boolean = CType(databaseAcc, DBAccLayer.Request).logout(err)

            If logoutErfolgreich Then
                If awinSettings.visboDebug Then
                    If awinSettings.englishLanguage Then
                        Call MsgBox(err.errorMsg & vbCrLf & "User don't have access to a VisboCenter any longer!")
                    Else
                        Call MsgBox(err.errorMsg & vbCrLf & "User hat keinen Zugriff mehr zu einem VisboCenter!")
                    End If
                End If

            End If

            appInstance.ScreenUpdating = True

            If Application.Workbooks.Count <= 1 Then
                Dim a As Integer = Application.Workbooks.Count
                'Dim name asstring = Application.Workbooks(1).name
            End If



        Catch ex As Exception

        End Try
        Application.Quit()

    End Sub



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

    Private Sub ThisWorkbook_WindowResize(Wn As Window) Handles Me.WindowResize
        ' ein Vergrößern sollte immer das Chart größer, das heisst breiter werden lassen
        ' das Mitte Window 
        ' beim Verkleinern sollte gar nix passieren , auss

    End Sub
End Class
