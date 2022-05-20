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
        Return GetResourceText("VSIBO_SPE.Ribbon1.xml")
    End Function

#Region "Menübandrückrufe"
    'Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
        Me.ribbon.Invalidate()
    End Sub
    Public Function imageSuper_GetImage(control As IRibbonControl) As Bitmap

        imageSuper_GetImage = My.Resources.noun_money_100x100
        Select Case control.Id
            Case "Pt6G6B3"
                imageSuper_GetImage = My.Resources.noun_money_100x100
            Case "Pt6G6B4"
                imageSuper_GetImage = My.Resources.noun_stop_watch_100x100
            Case "Pt6G6B5"
                imageSuper_GetImage = My.Resources.noun_bottleneck_100x100
            Case "Pt6G6B6"
                imageSuper_GetImage = My.Resources.visbo_icon_transparent_Bild
            Case "Pt6G6B7"
                imageSuper_GetImage = My.Resources.noun_settings_100x100
        End Select
    End Function

    ''' <summary>
    ''' lädt die gewählten Projekte und gewählten Varianten in die Session
    ''' </summary>
    ''' <param name="Control"></param>
    ''' <remarks></remarks>
    Public Sub PTProjectLoad(Control As Office.IRibbonControl)

        Try
            Dim path As String = "C:\Users\UteRittinghaus-Koyte\Dokumente\VISBO-NativeClients\visbo-projectboard\VISBO SimpleProjectEditTest\VISBO SimpleProjectEditTest\bin\Debug"

            If Not speSetTypen_Performed Then

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

                appInstance.EnableEvents = False
                Call speSetTypen()
                appInstance.EnableEvents = True

                appInstance.Visible = True

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
        Call PBBDatenbankLoadProjekte(Control, False)

        If AlleProjekte.Count > 0 Then
            ' Termine edit aufschalten
            'all MsgBox(currentProjektTafelModus)
            Call massEditRcTeAt(currentProjektTafelModus)
        End If

    End Sub



    Public Sub PTProjectSave(control As Office.IRibbonControl)
        Call MsgBox("Save")
        If AlleProjekte.Count > 0 Then
            ' Mouse auf Wartemodus setzen
            appInstance.Cursor = Excel.XlMousePointer.xlWait
            'Projekte speichern
            Call StoreAllProjectsinDB()
            ' delete all projects from cache
            'AlleProjekte.Clear()
            'Try
            '    Dim currentws As Excel.Worksheet = appInstance.ActiveSheet

            '    Select Case currentProjektTafelModus
            '        Case ptModus.massEditTermine
            '            Call massEditRcTeAt(ptModus.massEditTermine)
            '        Case ptModus.massEditRessSkills
            '            Call massEditRcTeAt(ptModus.massEditRessSkills)
            '        Case ptModus.massEditCosts
            '            Call massEditRcTeAt(ptModus.massEditCosts)

            '    End Select

            'Catch ex As Exception

            'End Try

            ' Mouse wieder auf Normalmodus setzen
            appInstance.Cursor = Excel.XlMousePointer.xlDefault
        End If
    End Sub


    Public Sub PTProjectDelete(control As Office.IRibbonControl)
        Call MsgBox("Delete")
    End Sub


    Public Sub PTProjectCost(control As Office.IRibbonControl)
        currentProjektTafelModus = ptModus.massEditCosts
        ' Call MsgBox(ptModus.massEditCosts.ToString)

        Call massEditRcTeAt(ptModus.massEditCosts)
    End Sub

    Public Sub PTProjectTime(control As Office.IRibbonControl)
        currentProjektTafelModus = ptModus.massEditTermine
        'Call MsgBox(ptModus.massEditTermine.ToString)

        Call massEditRcTeAt(ptModus.massEditTermine)
    End Sub

    Public Sub PTProjectResources(control As Office.IRibbonControl)
        currentProjektTafelModus = ptModus.massEditRessSkills
        'Call MsgBox(ptModus.massEditRessSkills.ToString)

        Call massEditRcTeAt(ptModus.massEditRessSkills)
    End Sub


    Public Sub PTProjectSettings(control As Office.IRibbonControl)
        Call MsgBox("Settings")
    End Sub

    Public Sub PTProjectGoToWebUI(control As Office.IRibbonControl)

        Dim pname As String = ""
        Dim vname As String = ""
        Dim view As String = "Capacity"

        pname = visboZustaende.currentProject.name
        vname = visboZustaende.currentProject.variantName

        Select Case currentProjektTafelModus
            Case ptModus.massEditCosts
                view = "Cost"
            Case ptModus.massEditRessSkills
                view = "Capacity"
            Case ptModus.massEditTermine
                view = "Deadline"
        End Select

        Call FollowHyperlinkToWebsite(visboZustaende.currentProject, view)


        'Call MsgBox("GoToWebUI for " & pname & ":" & vname)
    End Sub




    'Public Sub ImportWorksheet()
    '    ' This macro will import a file into this workbook 
    '    Dim ControlFile As String = appInstance.ActiveWorkbook.Name

    '    Dim currentws As Excel.Worksheet = appInstance.ActiveSheet

    '    Dim wb As Excel.Workbook = appInstance.Workbooks.Open(Filename:="C:\Users\UteRittinghaus-Koyte\Dokumente\VISBO-NativeClients\visbo-projectboard\Projectboard\Projectboard\bin\Debug\" & "Projectboard.xlsx")


    '    ' Private Sub Application_WorkbookBeforeSave(
    '    'ByVal Wb As Microsoft.Office.Interop.Excel.Workbook
    '    'ByVal SaveAsUI As Boolean
    '    'ByRef Cancel As Boolean) Handles Application.WorkbookBeforeSave

    '    If Globals.Factory.HasVstoObject(wb) = True Then
    '        For Each interopSheet As Excel.Worksheet In wb.Worksheets
    '            If Globals.Factory.HasVstoObject(interopSheet) = True Then
    '                Dim vstoSheet As Worksheet = Globals.Factory.GetVstoObject(interopSheet)
    '                If vstoSheet.Controls.Count > 0 Then
    '                    System.Windows.Forms.MessageBox.Show(
    '                        "The VSTO controls are not persisted when you" _
    '                        + " save and close this workbook.",
    '                        "Controls Persistence",
    '                        System.Windows.Forms.MessageBoxButtons.OK,
    '                        System.Windows.Forms.MessageBoxIcon.Warning)
    '                    Exit For
    '                End If
    '            End If
    '        Next
    '    End If


    '    myProjektTafel = wb.Name

    '    Dim newWS As Excel.Worksheet = wb.Worksheets.Item("meRC")
    '    Dim newWSName As String = newWS.Name
    '    appInstance.ActiveWorkbook.Worksheets.Item("meRC").Copy(After:=currentws)

    '    'appInstance.Windows("Projectboard").Activate()
    '    appInstance.ActiveWorkbook.Close(SaveChanges:=False)
    '    'appInstance.Windows(ControlFile).Activate()
    'End Sub
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
