
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

        Dim vpid As String = ""
        Dim vpvid As String = ""
        Dim oneTimeToken As String = ""
        Dim rest As String = ""
        Dim del As String = ""

        'Dim cbar As CommandBar
        Dim hstr() As String
        Dim cmdLine As String = GetCommandLine()
        Dim cmdID As Int64 = CType(GetCommandLine, Int64)
        Dim cline As String = CmdToSTr(cmdID)
        'Dim cline As String = "C:\Users\UteRittinghaus-Koyte\Dokumente\VISBO-NativeClients\visbo-projectboard\VISBO SPE\VSIBO SPE\bin\Debug\VISBO SPE.xlsx" / """vpid:627a4a80c0bdb36bb7f65062&vpvid:627a5a1fc0bdb36bb7f65a22&ott:eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI2MThhNzViNzIyMDI2NDIyODkwY2NhOWEiLCJlbWFpbCI6InVsaS5wcm9ic3RAdmlzYm8uZGUiLCJzZXNzaW9uIjp7ImlwIjoiOTEuMTAuMTk3LjE4MiIsInRpbWVzdGFtcCI6IjIwMjItMDUtMjNUMTg6Mzk6MDMuODg0WiJ9LCJpYXQiOjE2NTMzMzExNDMsImV4cCI6MTY1MzMzMTI2M30.0V1vu5kDApZqnZs6P7pW_ds7qUwdwT0NcSCbVy9sO70"
        Call MsgBox(cline)

        hstr = cline.Split("/")
        Dim parameter As String = (hstr(hstr.Length - 1))
        Call MsgBox(parameter)
        hstr = parameter.Split("""")
        If hstr.Length > 1 Then
            Call MsgBox(hstr(1))
        End If


        Dim parameterString As String = hstr(1)
        'parameterString = "vpid:627a4a80c0bdb36bb7f65062&vpvid:627a5a1fc0bdb36bb7f65a22&ott:eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI2MThhNzViNzIyMDI2NDIyODkwY2NhOWEiLCJlbWFpbCI6InVsaS5wcm9ic3RAdmlzYm8uZGUiLCJzZXNzaW9uIjp7ImlwIjoiOTEuMTAuMTk3LjE4MiIsInRpbWVzdGFtcCI6IjIwMjItMDUtMjNUMTg6Mzk6MDMuODg0WiJ9LCJpYXQiOjE2NTMzMzExNDMsImV4cCI6MTY1MzMzMTI2M30.0V1vu5kDApZqnZs6P7pW_ds7qUwdwT0NcSCbVy9sO70"

        'parameterString = "vpid:624dcb87e89109508af0ef8b&vpvid:624dcb88e89109508af0efde&ott:eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJfaWQiOiI1YWY4NGU3ZGUxMTdjNGM3ZmI3MGQ1MjAiLCJlbWFpbCI6InV0ZS5yaXR0aW5naGF1cy1rb3l0ZWtAdmlzYm8uZGUiLCJzZXNzaW9uIjp7ImlwIjoiODQuMTYwLjc1LjQzIiwidGltZXN0YW1wIjoiMjAyMi0wNS0yNVQwOTo0NDozOS4zNjNaIn0sImlhdCI6MTY1MzQ3MTg3OSwiZXhwIjoxNjUzNDcxOTk5fQ.y8pZh7WOj5L1ZK50WIPCGah7t1OF10h0EN6TCpGtRm0"

        'Call MsgBox(parameterString)
        hstr = parameterString.Split("&")

        For i = 0 To hstr.Length - 1
            Dim elem As String = hstr(i)
            Dim bezeichner As String = (elem.Split(":"))(0)
            Call MsgBox("bezeichner = " & bezeichner)
            Select Case bezeichner
                Case "vpid"
                    spe_vpid = (elem.Split(":"))(1)
                    Call MsgBox("vpid = " & spe_vpid)
                Case "vpvid"
                    spe_vpvid = (elem.Split(":"))(1)
                    Call MsgBox("vpvid = " & spe_vpvid)
                Case "ott"
                    spe_ott = (elem.Split(":"))(1)
                    Call MsgBox("oneTimeToken = " & spe_ott)
                Case Else
                    rest = (elem.Split(":"))(1)
                    Call MsgBox("rest = " & rest)
            End Select
        Next


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
            Call speSetTypen(spe_ott)
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

        '' Laden des übergebenen Projektes

        Call loadGivenProject()

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

    Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As Long
    Declare Function lstrlenW Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
    Declare Function CopyMemory Lib "kernel32" Alias "lstrcpynA" (ByVal MyDest As String, ByVal MySource As Long, ByVal MySize As Long) As Long
    Function CmdToSTr(Cmd As Long) As String
        Dim Buffer As String
        Dim StrLen As Long

        If Cmd Then
            StrLen = lstrlenW(Cmd) * 2

            If StrLen > 0 Then
                Buffer = Space(StrLen - 1)
                Cmd = CopyMemory(Buffer, Cmd, StrLen + 1)
                If Cmd > 0 Then
                    Dim PosZero As Long
                    PosZero = InStr(Buffer, Chr(0))
                    If PosZero > 0 Then Buffer = Left(Buffer, PosZero - 1)
                End If
                CmdToSTr = Buffer
            Else
                CmdToSTr = ""
            End If
        Else
            CmdToSTr = ""
        End If
    End Function
    Private Function UnicodeBytesToString(
    ByVal bytes() As Byte) As String
        Dim enc As Encoding = New UnicodeEncoding(False, True, True)
        Dim value As String = ""
        Try
            value = enc.GetString(bytes)

        Catch e As DecoderFallbackException
            Call logger(ptErrLevel.logError, "UnicodeBytesToString", "Unable to decode {0} at index {1}" & " : " & e.Index)

        End Try

        Return value
    End Function
End Class
