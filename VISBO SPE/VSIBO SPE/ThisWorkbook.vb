
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ProjectboardReports
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
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

    Private Sub ThisWorkbook_Startup() Handles Me.Startup
        'Dim cbar As CommandBar
        Dim hstr() As String
        Dim xxx As Long = GetCommandLine
        Dim cline As String = CmdToSTr(xxx)
        Call MsgBox(cline)

        hstr = cline.Split("""")
        'Call MsgBox(hstr(0), hstr(1), hstr(2))

        'visboClient = "VISBO Simple Project Edit / "
        visboClient = "VISBO SPE / "

        ' currentProjektTafelModus auf beginnend mit massEditTermine setzend
        Select Case My.Settings.startModus
            Case "Time"
                currentProjektTafelModus = ptModus.massEditTermine
            Case "Resources"
                currentProjektTafelModus = ptModus.massEditRessSkills
            Case "Cost"
                currentProjektTafelModus = ptModus.massEditCosts
            Case Else
                currentProjektTafelModus = ptModus.massEditTermine
        End Select



        ' Refresh von Projekte im Cache  in Minuten
        cacheUpdateDelay = 10

        appInstance = Application


        logfileNamePath = createLogfileName()

        ' nicht visible setzen
        appInstance.Visible = False

        myProjektTafel = appInstance.ActiveWorkbook.Name

        Dim path As String = CType(appInstance.ActiveWorkbook, Excel.Workbook).Path

        ' die Short Cut Menues aus Excel werden hier nicht mehr de-aktiviert 
        ' das wird jetzt nur in Tabelle1, also der Projekt-Tafel gemacht ...
        ' in anderen Excel Sheets ist das weiterhin aktiv 
        'For Each cbar In appInstance.CommandBars

        '    If cbar.Type = MsoBarType.msoBarTypePopup Then
        '        cbar.Enabled = False
        '    End If
        ''Next

        'ur:220523: Test if esc is no longer necessary
        'magicBoardCmdBar.cmdbars = appInstance.CommandBars



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

            appInstance.EnableEvents = False
            Call speSetTypen()
            appInstance.EnableEvents = True

            appInstance.Visible = True

            speSetTypen_Performed = True

        Catch ex As Exception

            appInstance.EnableEvents = True

            Call MsgBox(ex.Message)
            appInstance.Quit()
        Finally
            appInstance.ScreenUpdating = True
            appInstance.ShowChartTipNames = True
            appInstance.ShowChartTipValues = True
        End Try

        anzahlCalls = 0
    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown
        Call logger(ptErrLevel.logInfo, "VisboSPE", "Add-In was finished!")

        If loginErfolgreich Then

            My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
            If My.Settings.rememberUserPWD Then
                My.Settings.userNamePWD = awinSettings.userNamePWD
            Else
                'Cancel User Login-Data
                My.Settings.userNamePWD = ""
            End If

            ' speichern 
            My.Settings.Save()
        Else
            ' wenn kein erfolgreicher Login stattgefunden hat
            My.Settings.Reset()
        End If


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


            If cancelAbbruch Then
                Cancel = True
            Else

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


    Private Sub ThisWorkbook_BeforeSave(SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Me.BeforeSave

        Cancel = True
    End Sub

    Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
    Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (MyDest As Byte(), MySource As Long, ByVal MySize As Long)
    Function CmdToSTr(Cmd As Long) As String
        Dim Buffer() As Byte
        Dim StrLen As Long

        If Cmd Then
            StrLen = lstrlenW(Cmd) * 2

            If StrLen Then
                ReDim Buffer(StrLen - 1)
                CopyMemory(Buffer, Cmd, StrLen)
                CmdToSTr = UnicodeBytesToString(Buffer)
            Else
                CmdToSTr = ""
            End If
        Else
            CmdToSTr = ""
        End If
    End Function
    Private Function UnicodeBytesToString(
    ByVal bytes() As Byte) As String

        Return System.Text.Encoding.Unicode.GetString(bytes)
    End Function
End Class
