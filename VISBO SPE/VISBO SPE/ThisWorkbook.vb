
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ProjectboardReports
Imports Microsoft.Office.Core
Imports System.Environment
Imports Microsoft.Office.Interop.Excel
Public Class ThisWorkbook
    ' Copyright Philipp Koytek et al. 
    ' 2012 ff
    ' Nicht authorisierte Verwendung nicht gestattet 

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

    Private Sub ThisWorkbook_ActivateEvent() Handles Me.ActivateEvent

        Try
            Application.DisplayFormulaBar = False
            Application.ActiveWindow.DisplayWorkbookTabs = False
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Activiate Workbook SPE", ex.Message)
        End Try


    End Sub

    Private Sub ThisWorkbook_Startup() Handles Me.Startup

        Try
            Dim vpid As String = ""
            Dim vpvid As String = ""
            Dim oneTimeToken As String = ""
            Dim rest As String = ""
            Dim del As String = ""
            Dim parameterString As String = ""
            'Dim cbar As CommandBar

            'Call Auto_open()

            ' Name of the called Client
            visboClient = divClients(client.VisboSPE)


            logfileNamePath = createLogfileName()
            'Call MsgBox(logfileNamePath)

            Dim CmdLine As String 'command-line string

            CmdLine = GetCommandLine() 'get the cmd-line string
            CmdLine = Left$(CmdLine, InStr(CmdLine & vbNullChar, vbNullChar) - 1)

            '----nur zum Test

            Dim hhstr() As String = CmdLine.Split("/")
            Dim hparameter As String = ""
            If hhstr.Length = 2 Then
                Dim curUserAppDir As String = GetFolderPath(SpecialFolder.ApplicationData)
                CmdLine = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE " & curUserAppDir & "\VISBO\VISBO Project Edit\VISBO Project Edit.xlsx/"
            End If

            '----nur zum Test
            'Call logger(ptErrLevel.logInfo, "Startup", "cmdline: " & CmdLine)

            Dim hstr() As String = CmdLine.Split("/")
            'Dim parameter As String = ""
            'If hstr.Length > 2 Then
            '    'Call logger(ptErrLevel.logInfo, "Startup", "parameter1: " & hstr(2))
            'End If
            'If hstr.Length > 3 Then
            '    Call logger(ptErrLevel.logInfo, "Startup", "parameter2: " & hstr(3))
            'End If
            'If hstr.Length > 4 Then
            '    Call logger(ptErrLevel.logInfo, "Startup", "parameter3: " & hstr(4))
            'End If

            If hstr.Length > 2 Then
                For i = 2 To hstr.Length - 2
                    Dim elem As String = hstr(i)
                    Dim bezeichner As String = (elem.Split(":"))(0)
                    Select Case bezeichner
                        Case "vpid"
                            spe_vpid = (elem.Split(":"))(1)
                            'Call logger(ptErrLevel.logInfo, "Startup", "vpid = " & spe_vpid)
                        Case "vpvid"
                            spe_vpvid = (elem.Split(":"))(1)
                            'Call logger(ptErrLevel.logInfo, "Startup", "vpvid = " & spe_vpvid)
                        Case "ott"
                            spe_ott = (elem.Split(":"))(1)
                            'Call logger(ptErrLevel.logInfo, "Startup", "oneTimeToken = " & spe_ott)
                        Case Else
                            rest = (elem.Split(":"))(1)
                            'Call logger(ptErrLevel.logInfo, "Startup", "rest = " & rest)
                    End Select
                Next
            End If


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

                    ' the following settings were defined in customization-configuration
                    'awinSettings.autoAjustChilds = My.Settings.autoAjustChilds
                    'awinSettings.noNewCalculation = My.Settings.noNewCalculation
                    'awinSettings.propAnpassRess = My.Settings.propAnpassRess

                End If

                ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
                awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
                'awinSettings.rememberUserPwd = False

                If My.Settings.rememberUserPWD Then
                    awinSettings.userNamePWD = My.Settings.userNamePWD
                Else
                    awinSettings.userNamePWD = ""
                End If


                appInstance.EnableEvents = False
                Call speSetTypen(spe_ott)

                'Call logger(ptErrLevel.logInfo, "Startup- nach speSetTypen", "English Language =  " & awinSettings.englishLanguage.ToString)

                appInstance.EnableEvents = True

                appInstance.Visible = True

                speSetTypen_Performed = True

            Catch ex As Exception

                appInstance.EnableEvents = True

                appInstance.Quit()
            Finally
                appInstance.ScreenUpdating = True
                appInstance.ShowChartTipNames = True
                appInstance.ShowChartTipValues = True
            End Try

            '' Laden des übergebenen Projektes

            Call loadGivenProject()



            anzahlCalls = 0
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Startup- nach speSetTypen", ex.Message)
        End Try

    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown

        Try
            'Call logger(ptErrLevel.logInfo, "VisboSPE", "Add-In was finished!")

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
        Catch ex As Exception
            Call logger(ptErrLevel.logError, "Shutdown ", ex.Message)
        End Try

    End Sub

    Private Sub ThisWorkbook_BeforeClose(ByRef Cancel As Boolean) Handles Me.BeforeClose


        Dim projektespeichern As New frmProjekteSpeichern
        Dim returnValue As DialogResult
        Dim cancelAbbruch As Boolean = False
        Dim err As New clsErrorCodeMsg
        Dim msg As String = ""

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
                                msg = "no projects to store ..."
                            Else
                                msg = "keine Projekte zu speichern ..."
                            End If
                            If awinSettings.visboDebug Then
                                Call MsgBox(msg)
                            End If

                            'Call logger(ptErrLevel.logInfo, "Worksbook.Before Close", msg)

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

    Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
    'Declare Function lstrlenW Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
    'Declare Function CopyMemory Lib "kernel32" Alias "lstrcpynA" (ByVal MyDest As String, ByVal MySource As Long, ByVal MySize As Long) As Long
    'Function CmdToSTr(Cmd As Long) As String
    '    Dim Buffer As String
    '    Dim StrLen As Long

    '    If Cmd Then
    '        StrLen = lstrlenW(Cmd) * 2

    '        If StrLen > 0 Then
    '            Buffer = Space(StrLen + 1)
    '            Call MsgBox("Vor Buffer: ")
    '            Cmd = CopyMemory(Buffer, Cmd, StrLen + 1)
    '            Call MsgBox("Nach Buffer: ")
    '            'If Cmd > 0 Then
    '            '    Dim PosZero As Long
    '            '    PosZero = InStr(Buffer, Chr(0))
    '            '    If PosZero > 0 Then Buffer = Left(Buffer, PosZero - 1)
    '            'End If
    '            CmdToSTr = Buffer
    '        Else
    '            CmdToSTr = ""
    '        End If
    '    Else
    '        CmdToSTr = ""
    '    End If
    'End Function

    'Sub Auto_open()

    '    Dim CmdLine As String 'command-line string

    '    CmdLine = GetCommandLine() 'get the cmd-line string
    '    CmdLine = Left$(CmdLine, InStr(CmdLine & vbNullChar, vbNullChar) - 1)
    '    Call MsgBox("cmdline1: " & CmdLine)

    '    Dim hstr() As String = CmdLine.Split("/")
    '    Dim parameter As String = ""
    '    If hstr.Length > 2 Then
    '        Call MsgBox("parameter1: " & hstr(2))
    '    End If
    '    If hstr.Length > 3 Then
    '        Call MsgBox("parameter2: " & hstr(3))
    '    End If
    '    If hstr.Length > 4 Then
    '        Call MsgBox("parameter3: " & hstr(4))
    '    End If

    'End Sub
    'Private Function UnicodeBytesToString(
    'ByVal bytes() As Byte) As String
    '    Dim enc As Encoding = New UnicodeEncoding(False, True, True)
    '    Dim value As String = ""
    '    Try
    '        value = enc.GetString(bytes)

    '    Catch e As DecoderFallbackException
    '        Call logger(ptErrLevel.logError, "UnicodeBytesToString", "Unable to decode {0} at index {1}" & " : " & e.Index)

    '    End Try

    '    Return value
    'End Function
End Class
