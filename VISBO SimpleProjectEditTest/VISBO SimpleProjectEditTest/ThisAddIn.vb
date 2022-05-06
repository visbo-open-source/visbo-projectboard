Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ProjectboardReports
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel


Public Class ThisAddIn


    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        'Dim cbar As CommandBar


        'visboClient = "VISBO Simple Project Edit / "
        visboClient = "VISBO SPE"

        ' Refresh von Projekte im Cache  in Minuten
        cacheUpdateDelay = 10


        ' nicht visible setzen
        'appInstance.Visible = False

        'myProjektTafel = appInstance.ActiveWorkbook.Name

        Dim path As String = My.Settings.awinPath

        '' die Short Cut Menues aus Excel werden hier nicht mehr de-aktiviert 
        '' das wird jetzt nur in Tabelle1, also der Projekt-Tafel gemacht ...
        '' in anderen Excel Sheets ist das weiterhin aktiv 
        ''For Each cbar In appInstance.CommandBars

        ''    If cbar.Type = MsoBarType.msoBarTypePopup Then
        ''        cbar.Enabled = False
        ''    End If
        ''Next

        'magicBoardCmdBar.cmdbars = appInstance.CommandBars

        appInstance = Application


        logfileNamePath = createLogfileName()

        'Try


        '    appInstance.ScreenUpdating = False

        '    ' hier werden die Settings aus der Datei ProjectboardConfig.xml ausgelesen.
        '    ' falls die nicht funktioniert, so werden die My.Settings ausgelesen und verwendet.

        '    If Not readawinSettings(path) Then

        '        awinSettings.databaseURL = My.Settings.mongoDBURL
        '        awinSettings.databaseName = My.Settings.mongoDBname
        '        awinSettings.DBWithSSL = My.Settings.mongoDBWithSSL
        '        awinSettings.proxyURL = My.Settings.proxyServerURL
        '        awinSettings.globalPath = My.Settings.globalPath
        '        awinSettings.awinPath = My.Settings.awinPath
        '        awinSettings.visboTaskClass = My.Settings.TaskClass
        '        awinSettings.visboAbbreviation = My.Settings.VISBOAbbreviation
        '        awinSettings.visboAmpel = My.Settings.VISBOAmpel
        '        awinSettings.visboAmpelText = My.Settings.VISBOAmpelText
        '        awinSettings.visboresponsible = My.Settings.VISBOresponsible
        '        awinSettings.visbodeliverables = My.Settings.VISBOdeliverables
        '        awinSettings.visbopercentDone = My.Settings.VISBOpercentDone
        '        awinSettings.visboMapping = My.Settings.VISBOMapping
        '        awinSettings.visboDebug = My.Settings.VISBODebug
        '        awinSettings.visboServer = My.Settings.VISBOServer
        '        awinSettings.userNamePWD = My.Settings.userNamePWD
        '        awinSettings.rememberUserPwd = My.Settings.rememberUserPWD

        '    End If



        '    ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
        '    awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        '    If My.Settings.rememberUserPWD Then
        '        awinSettings.userNamePWD = My.Settings.userNamePWD
        '    Else
        '        awinSettings.userNamePWD = ""
        '    End If

        '    ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
        '    awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        '    If My.Settings.rememberUserPWD Then
        '        awinSettings.userNamePWD = My.Settings.userNamePWD
        '    Else
        '        awinSettings.userNamePWD = ""
        '    End If

        '    ' Refresh von Projekte im Cache  in Minuten
        '    cacheUpdateDelay = 30

        '    'appInstance.EnableEvents = False
        '    Call speSetTypen()
        '    'appInstance.EnableEvents = True

        '    'appInstance.Visible = True

        'Catch ex As Exception

        '    appInstance.EnableEvents = True

        '    '   Call MsgBox(ex.Message)
        '    appInstance.Quit()
        'Finally
        '    appInstance.ScreenUpdating = True
        '    appInstance.ShowChartTipNames = True
        '    appInstance.ShowChartTipValues = True
        'End Try

        anzahlCalls = 0
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
