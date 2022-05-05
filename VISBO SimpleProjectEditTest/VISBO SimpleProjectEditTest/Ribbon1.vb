Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ProjectboardReports
Imports DBAccLayer
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Security.Principal
Imports System.Diagnostics
Imports System.Drawing
'Imports System.Windows
Imports System.Net
Imports System
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Web



'TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

'1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
'   zu behandeln, zum Beispiel das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem
'   Menüband-Designer exportiert haben, verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und
'   ändern Sie den Code für die Verwendung mit dem Programmiermodell für die Menübanderweiterung (RibbonX).

'3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.

'Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("VISBO_SimpleProjectEditTest.Ribbon1.xml")
    End Function

#Region "Menübandrückrufe"
    'Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        Me.ribbon.Invalidate()
    End Sub


    ''' <summary>
    ''' lädt die gewählten Projekte und gewählten Varianten in die Session
    ''' </summary>
    ''' <param name="Control"></param>
    ''' <remarks></remarks>
    Public Sub PTProjectLoad(Control As Office.IRibbonControl)

        Try

            If Not speSetTypen_Performed Then

                'appInstance.ScreenUpdating = False

                ' hier werden die Settings aus der Datei ProjectboardConfig.xml ausgelesen.
                ' falls die nicht funktioniert, so werden die My.Settings ausgelesen und verwendet.

                If Not readawinSettings(awinPath) Then

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
                Call speSetTypen()
                'appInstance.EnableEvents = True

                'appInstance.Visible = True

            End If
        Catch ex As Exception

            appInstance.EnableEvents = True

            '   Call MsgBox(ex.Message)
            appInstance.Quit()
        Finally
            appInstance.ScreenUpdating = True
            appInstance.ShowChartTipNames = True
            appInstance.ShowChartTipValues = True
        End Try


        Dim boardWasEmpty As Boolean = ShowProjekte.Count = 0
        Call PBBDatenbankLoadProjekte(Control)

        ' Window so positionieren, dass die Projekte sichtbar sind ...  
        If ShowProjekte.Count > 0 Then
            Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
            If boardWasEmpty Then
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
        End If


    End Sub



    Public Sub PTProjectSave(control As Office.IRibbonControl)
        Call MsgBox("Save")
    End Sub


    Public Sub PTProjectDelete(control As Office.IRibbonControl)
        Call MsgBox("Delete")
    End Sub


    Public Sub PTProjectCost(control As Office.IRibbonControl)
        Call MsgBox("Cost")
    End Sub

    Public Sub PTProjectTime(control As Office.IRibbonControl)
        Call MsgBox("Time")
    End Sub

    Public Sub PTProjectResources(control As Office.IRibbonControl)
        Call MsgBox("Resources")
    End Sub

#End Region

#Region "Hilfsprogramme"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
