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

    Public Function imageSuper_GetImage(control As IRibbonControl) As Bitmap

        imageSuper_GetImage = My.Resources.noun_money
        Select Case control.Id
            Case "Pt6G6B3"
                imageSuper_GetImage = My.Resources.noun_money
            Case "Pt6G6B4"
                imageSuper_GetImage = My.Resources.noun_stop_watch
            Case "Pt6G6B5"
                imageSuper_GetImage = My.Resources.noun_bottleneck
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

                'appInstance.ScreenUpdating = False

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
                Call speSetTypen()
                'appInstance.EnableEvents = True

                'appInstance.Visible = True

            End If
        Catch ex As Exception

            'appInstance.EnableEvents = True

            '   Call MsgBox(ex.Message)
            'appInstance.Quit()
        Finally
            '    appInstance.ScreenUpdating = True
            '    appInstance.ShowChartTipNames = True
            '    appInstance.ShowChartTipValues = True
        End Try


        Dim boardWasEmpty As Boolean = ShowProjekte.Count = 0
        Call PBBDatenbankLoadProjekte(Control, False)

        If AlleProjekte.Count > 0 Then
            ' Ressourcen edit aufschalten

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

        Call massEditRcTeAt(ptModus.massEditRessSkills)
    End Sub


    ''' <summary>
    ''' aktiviert , je nach Modus die entsprechenden Ribbon Controls 
    ''' </summary>
    ''' <param name="modus"></param>
    ''' <remarks></remarks>
    Public Sub enableControls(ByVal modus As ptModus)

        If modus = ptModus.graficboard Then
            visboZustaende.projectBoardMode = modus


        Else
            visboZustaende.projectBoardMode = modus

        End If

        Me.ribbon.Invalidate()

    End Sub


    ''' <summary>
    ''' es werden nur Projekte an MassEdit übergeben ... sollten Summary Projekte in der Selection sein, werden die erst durch ihre Projekte, die im Show sind, ersetzt 
    ''' </summary>
    ''' <param name="meModus"></param>
    Private Sub massEditRcTeAt(ByVal meModus As ptModus)
        Dim todoListe As New Collection
        Dim projektTodoliste As New Collection
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""


        '' now set visbozustaende
        '' necessary to know whether roles or cost need to be shown in building the forms to select roles , skills and costs 
        'visboZustaende.projectBoardMode = meModus

        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        ' die DB Cache Projekte werden hier weder zurückgesetzt, noch geholt ... das kostet nur Antwortzeit auf Vorhalt
        ' sie werden ggf im MassenEdit geholt, wenn es notwendig ist .... 

        ''ur: 220506:  Call projektTafelInit()

        ''ur: 220506: enableOnUpdate = False
        ' jetzt auf alle Fälle wieder das MPT Window aktivieren ...
        ''ur: 220506: projectboardWindows(PTwindows.mpt).Activate()

        If ShowProjekte.Count > 0 Then

            ' neue Methode 
            todoListe = getProjectSelectionList(True)

            ' check, ob wirklich alle Projekte editiert werden sollen ... 
            If todoListe.Count = ShowProjekte.Count And todoListe.Count > 30 Then
                Dim yesNo As Integer
                yesNo = MsgBox("Wollen Sie wirklich alle Projekte editieren?", MsgBoxStyle.YesNo)
                If yesNo = MsgBoxResult.No Then
                    enableOnUpdate = True
                    Exit Sub
                End If
            End If



            If todoListe.Count > 0 Then

                ' jetzt muss ggf noch showrangeLeft und showrangeRight gesetzt werden  

                Call enableControls(meModus)

                ' hier sollen jetzt die Projekte der todoListe in den Backup Speicher kopiert werden , um 
                ' darauf zugreifen zu können, wenn beim Massen-Edit die Option alle Änderungen verwerfen gewählt wird. 
                'Call saveProjectsToBackup(todoListe)

                ' hier wird die aktuelle Zusammenstellung an Windows gespeichert ...
                'projectboardViews(PTview.mpt) = CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).CustomViews, Excel.CustomViews).Add("View" & CStr(PTview.mpt))

                ' jetzt soll ScreenUpdating auf False gesetzt werden, weil jetzt Windows erzeugt und gewechselt werden 
                'appInstance.ScreenUpdating = False

                Try
                    enableOnUpdate = False

                    If (meModus = ptModus.massEditRessSkills Or meModus = ptModus.massEditCosts) Then

                        If showRangeLeft = 0 Then
                            showRangeLeft = ShowProjekte.getMinMonthColumn(todoListe)
                            showRangeRight = ShowProjekte.getMaxMonthColumn(todoListe)

                            '' ur:220506:  Call awinShowtimezone(showRangeLeft, showRangeRight, True)
                        Else
                            ' beim alten ShowRangeLeft lassen, wenn es Überlappungen gibt ..
                            Dim newLeft As Integer = ShowProjekte.getMinMonthColumn(todoListe)
                            Dim newRight As Integer = ShowProjekte.getMaxMonthColumn(todoListe)

                            If newLeft >= showRangeRight Or newRight <= showRangeLeft Then
                                ' neu bestimmen 
                                '' ur:220506:  Call awinShowtimezone(showRangeLeft, showRangeRight, False)

                                showRangeLeft = ShowProjekte.getMinMonthColumn(todoListe)
                                showRangeRight = ShowProjekte.getMaxMonthColumn(todoListe)

                                '' ur:220506:  Call awinShowtimezone(showRangeLeft, showRangeRight, True)

                            End If
                        End If

                        ' tk 15.2.19 Portfolio Manager darf Summary-Projekte bearbeiten , um sie dann als Vorgaben speichern zu können 
                        ' das wird in der Funktion substituteListeByPVnameIDs geregelt .. 
                        projektTodoliste = substituteListeByPVNameIDs(todoListe)

                        ' jetzt aufbauen der dbCacheProjekte, names are pvnames
                        Call buildCacheProjekte(projektTodoliste, namesArePvNames:=True)

                        Call writeOnlineMassEditRessCost(projektTodoliste, showRangeLeft, showRangeRight, meModus)


                    ElseIf meModus = ptModus.massEditTermine Then
                        ' tk 15.2.19 Portfolio Manager darf Summary-Projekte bearbeiten , um sie dann als Vorgaben speichern zu können 
                        ' das wird in der Funktion substituteListeByPVnameIDs geregelt .. 
                        projektTodoliste = substituteListeByPVNameIDs(todoListe)

                        ' jetzt aufbauen der dbCacheProjekte, names are pvnames
                        Call buildCacheProjekte(projektTodoliste, namesArePvNames:=True)

                        Call writeOnlineMassEditTermine(projektTodoliste)

                    ElseIf meModus = ptModus.massEditAttribute Then
                        ' tk 15.2.19 Portfolio Manager darf Summary-Projekte bearbeiten , um sie dann als Vorgaben speichern zu können 
                        ' das wird in der Funktion substituteListeByPVnameIDs geregelt .. 
                        projektTodoliste = substituteListeByPVNameIDs(todoListe)

                        ' jetzt aufbauen der dbCacheProjekte, names are pNames
                        Call buildCacheProjekte(todoListe, namesArePvNames:=False)

                        Call writeOnlineMassEditAttribute(projektTodoliste)
                    Else
                        Exit Sub
                    End If

                    appInstance.EnableEvents = True



                    Try

                        If Not IsNothing(projectboardWindows(PTwindows.mpt)) Then
                            projectboardWindows(PTwindows.massEdit) = projectboardWindows(PTwindows.mpt).NewWindow
                        Else
                            projectboardWindows(PTwindows.massEdit) = appInstance.ActiveWindow.NewWindow
                        End If

                    Catch ex As Exception
                        projectboardWindows(PTwindows.massEdit) = appInstance.ActiveWindow.NewWindow
                    End Try

                    ' jetzt das Massen-Edit Sheet Ressourcen / Kosten aktivieren 
                    Dim tableTyp As Integer = ptTables.meRC

                    If (meModus = ptModus.massEditRessSkills Or meModus = ptModus.massEditCosts) Then
                        tableTyp = ptTables.meRC
                    ElseIf meModus = ptModus.massEditTermine Then
                        tableTyp = ptTables.meTE
                    ElseIf meModus = ptModus.massEditAttribute Then
                        tableTyp = ptTables.meAT
                    Else
                        tableTyp = ptTables.meRC
                    End If

                    With CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Worksheets(arrWsNames(tableTyp)), Excel.Worksheet)
                        .Activate()
                    End With


                    With projectboardWindows(PTwindows.massEdit)
                        'With appInstance.ActiveWindow

                        Try
                            .FreezePanes = False
                            .Split = False

                            If (meModus = ptModus.massEditRessSkills Or meModus = ptModus.massEditCosts) Then

                                If awinSettings.meExtendedColumnsView = True Then
                                    .SplitRow = 1
                                    .SplitColumn = 7
                                    .FreezePanes = True
                                Else
                                    .SplitRow = 1
                                    .SplitColumn = 6
                                    .FreezePanes = True
                                End If
                                .DisplayHeadings = False

                            ElseIf meModus = ptModus.massEditTermine Then
                                .SplitRow = 1
                                .SplitColumn = 6
                                .FreezePanes = True
                                .DisplayHeadings = True

                            ElseIf meModus = ptModus.massEditAttribute Then
                                .SplitRow = 1
                                .SplitColumn = 5
                                .FreezePanes = True
                                .DisplayHeadings = True

                            Else
                                Exit Sub
                            End If

                            .DisplayFormulas = False
                            .DisplayGridlines = True
                            '.GridlineColor = RGB(220, 220, 220)
                            .GridlineColor = Excel.XlRgbColor.rgbBlack
                            .DisplayWorkbookTabs = False
                            .Caption = bestimmeWindowCaption(PTwindows.massEdit, tableTyp:=tableTyp)
                            .WindowState = Excel.XlWindowState.xlMaximized
                            .Activate()
                        Catch ex As Exception
                            Call MsgBox("Fehler in massEditRcTeAt")
                        End Try


                    End With


                    ' tk 4.3.19 
                    ' jetzt das Multiprojekt Window ausblenden ...
                    projectboardWindows(PTwindows.mpt).Visible = False

                    ' jetzt auch alle anderen ggf offenen pr und pf Windows unsichtbar machen ... 
                    Try
                        If Not IsNothing(projectboardWindows(PTwindows.mptpf)) Then
                            projectboardWindows(PTwindows.mptpf).Visible = False
                        End If
                    Catch ex As Exception

                    End Try

                    Try
                        If Not IsNothing(projectboardWindows(PTwindows.mptpr)) Then
                            projectboardWindows(PTwindows.mptpr).Visible = False
                        End If
                    Catch ex As Exception

                    End Try

                    ' Ende Ausblenden 






                Catch ex As Exception
                    Call MsgBox("Fehler: " & ex.Message)
                    If appInstance.EnableEvents = False Then
                        appInstance.EnableEvents = True
                    End If

                End Try

            Else
                enableOnUpdate = True
                If appInstance.EnableEvents = False Then
                    appInstance.EnableEvents = True
                End If
                If awinSettings.englishLanguage Then
                    Call MsgBox("no projects apply to criterias ...")
                Else
                    Call MsgBox("Es gibt keine Projekte, die zu der Auswahl passen ...")
                End If
            End If


        Else
            enableOnUpdate = True
            If appInstance.EnableEvents = False Then
                appInstance.EnableEvents = True
            End If

            If awinSettings.englishLanguage Then
                Call MsgBox("no active projects ...")
            Else
                Call MsgBox("Es gibt keine aktiven Projekte ...")
            End If

        End If


        'appInstance.ScreenUpdating = True
        'If appInstance.ScreenUpdating = False Then
        '    appInstance.ScreenUpdating = True
        'End If


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
