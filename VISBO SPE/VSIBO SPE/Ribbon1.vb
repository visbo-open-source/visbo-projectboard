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
            Case "Pt6G6B9"
                imageSuper_GetImage = My.Resources.noun_chart_100x100
            Case "Pt6G6B8"
                imageSuper_GetImage = My.Resources.noun_gantt_chart_100x100
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

        Dim projektespeichern As New frmProjekteSpeichern
        Dim returnValue As DialogResult
        Dim cancelAbbruch As Boolean = False
        Dim err As New clsErrorCodeMsg


        Try
            'Dim path As String = "C:\Users\UteRittinghaus-Koyte\Dokumente\VISBO-NativeClients\visbo-projectboard\VISBO SPE\VSIBO SPE\bin\Debug"
            Dim path As String = ""

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

                Try
                    Dim clearOK As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()
                Catch ex As Exception
                    Call logger(ptErrLevel.logError, "PTProjectLoad", "Warning: no Cache clearing " & ex.Message)
                End Try

                ' Refresh von Projekte im Cache  in Minuten
                cacheUpdateDelay = 30

                appInstance.EnableEvents = False
                Call speSetTypen("")
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


        Dim boardWasEmpty As Boolean = editProjekteInSPE.Count = 0

        If Not boardWasEmpty Then
            If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() And AlleProjekte.Count > 0 Then
                returnValue = projektespeichern.ShowDialog

                If returnValue = DialogResult.Yes Then

                    Call StoreAllProjectsinDB()

                End If
            End If
            AlleProjekte.Clear()
            ShowProjekte.Clear()
            editProjekteInSPE.Clear()
            Call clearTable(currentProjektTafelModus)
        End If

        If spe_vpid <> "" And spe_vpvid <> "" Then

            'holen des Projekte mit vpid... und vpvid...
            Dim hproj As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectVersionfromDB(spe_vpid, spe_vpvid, err)
            If Not IsNothing(hproj) Then
                ShowProjekte.Add(hproj, False)
                editProjekteInSPE.Add(hproj, False)
                AlleProjekte.Add(hproj, False)
            Else
                Call PBBDatenbankLoadProjekte(Control, False)
            End If
            spe_vpid = ""
            spe_vpvid = ""
        Else
            Call PBBDatenbankLoadProjekte(Control, False)
        End If

        appInstance.EnableEvents = True

        If AlleProjekte.Count > 0 Then
            ' Termine edit aufschalten
            visboZustaende.currentProject = AlleProjekte.getProject(0)
            Call massEditRcTeAt(currentProjektTafelModus)
        End If

    End Sub



    Public Sub PTProjectSave(control As Office.IRibbonControl)
        'Call MsgBox("Save")
        If AlleProjekte.Count > 0 Then
            ' Mouse auf Wartemodus setzen
            appInstance.Cursor = Excel.XlMousePointer.xlWait
            'Projekte speichern
            Call StoreAllProjectsinDB()

            ' Mouse wieder auf Normalmodus setzen
            appInstance.Cursor = Excel.XlMousePointer.xlDefault
        End If
    End Sub


    Public Sub PTProjectDelete(control As Office.IRibbonControl)

        'delete all projects from cache
        AlleProjekte.Clear()
        ShowProjekte.Clear()
        editProjekteInSPE.Clear()

        Try
            Dim currentws As Excel.Worksheet = appInstance.ActiveSheet

            Select Case currentProjektTafelModus
                Case ptModus.massEditTermine
                    Call massEditRcTeAt(ptModus.massEditTermine)
                Case ptModus.massEditRessSkills
                    Call massEditRcTeAt(ptModus.massEditRessSkills)
                Case ptModus.massEditCosts
                    Call massEditRcTeAt(ptModus.massEditCosts)

            End Select

        Catch ex As Exception

        End Try

    End Sub


    Public Sub PTProjectCost(control As Office.IRibbonControl)

        If editProjekteInSPE.Count > 0 Then
            currentProjektTafelModus = ptModus.massEditCosts
            ' Call MsgBox(ptModus.massEditCosts.ToString)

            Call massEditRcTeAt(ptModus.massEditCosts)
        End If

    End Sub

    Public Sub PTProjectTime(control As Office.IRibbonControl)

        If editProjekteInSPE.Count > 0 Then
            currentProjektTafelModus = ptModus.massEditTermine
            'Call MsgBox(ptModus.massEditTermine.ToString)

            Call massEditRcTeAt(ptModus.massEditTermine)
        End If

    End Sub

    Public Sub PTProjectResources(control As Office.IRibbonControl)

        If editProjekteInSPE.Count > 0 Then
            currentProjektTafelModus = ptModus.massEditRessSkills
            'Call MsgBox(ptModus.massEditRessSkills.ToString)

            Call massEditRcTeAt(ptModus.massEditRessSkills)
        End If

    End Sub


    Public Sub PTProjectEditSettings(control As Office.IRibbonControl)
        Dim settingsEdit As New frmProjectEditSettings
        settingsEdit.ShowDialog()
    End Sub

    Public Sub PTProjectGoToWebUI(control As Office.IRibbonControl)

        Dim pname As String = ""
        Dim vname As String = ""
        Dim view As String = "Capacity"

        If editProjekteInSPE.Count > 0 Then
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
        End If

        'Call MsgBox("GoToWebUI for " & pname & ":" & vname)
    End Sub

    Public Sub PTLoadPiC_ConstellationFromDB(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

        ' Timer
        Dim sw As clsStopWatch
        sw = New clsStopWatch
        sw.StartTimer()

        Dim loadConstellationFrm As New frmLoadConstellation
        Dim storedAtOrBefore As Date = Date.Now.Date.AddHours(23).AddMinutes(59)

        Dim timeStampsCollection As New Collection

        Dim dbPortfolioNames As New SortedList(Of String, String)
        Dim cTimestamp As Date
        Dim initMessage As String = "Es sind dabei folgende Probleme aufgetreten" & vbLf & vbLf


        Dim outPutCollection As New Collection
        Dim outputLine As String = ""

        Dim successMessage As String = initMessage
        Dim returnValue As DialogResult

        Dim addToContext As Boolean = False

        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        Call projektTafelInit()


        ' Wenn das Laden eines Portfolios aus dem Menu Datenbank aufgerufen wird, so werden erneut alle Portfolios aus der Datenbank geholt


        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

            dbPortfolioNames = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, err)

            If dbPortfolioNames.Count > 0 Then

                Try
                    enableOnUpdate = False

                    loadConstellationFrm.addToSession.Checked = False
                    loadConstellationFrm.addToSession.Visible = False

                    loadConstellationFrm.loadAsSummary.Visible = False
                    loadConstellationFrm.loadAsSummary.Checked = False

                    loadConstellationFrm.constellationsToShow = dbPortfolioNames
                    loadConstellationFrm.retrieveFromDB = True

                    returnValue = loadConstellationFrm.ShowDialog

                    sw.StartTimer()

                    If returnValue = DialogResult.OK Then

                        ' it is only possible to load one portfolio as Context 
                        ' now just load the Context in AlleProjekte resp. ShowProjekte 

                        ' if AlleProjekte already contains the Name#variantName then don't do replace it 
                        ' if ShowProjekte already contains the name, then don't replace it. 

                        ' should not be possible to load an Portfolio of earlier dates into Context 
                        storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)


                        Dim constellationsToDo As New clsConstellations

                        ' Liste der ausgewählten Portfolio/Variante Paaren (pro Portfolio nur eine Variante)
                        Dim constellationsChecked As New SortedList(Of String, String)

                        ' WaitCursor einschalten ...
                        Cursor.Current = Cursors.WaitCursor

                        ' liste, welche Portfolios und Portfolio-Varianten geladen werden sollen, wird erstellt
                        constellationsChecked = New SortedList(Of String, String)

                        For Each tNode As TreeNode In loadConstellationFrm.TreeViewPortfolios.Nodes
                            If tNode.Checked Then
                                Dim checkedVariants As Integer = 0          ' enthält die Anzahl ausgwählter Varianten des pName
                                For Each vNode As TreeNode In tNode.Nodes
                                    If vNode.Checked Then
                                        If Not constellationsChecked.ContainsKey(tNode.Text) Then
                                            Dim vname As String = deleteBrackets(vNode.Text)
                                            constellationsChecked.Add(tNode.Text, vname)
                                        Else
                                            Call MsgBox("Portfolio '" & tNode.Text & "' mehrfach ausgewählt!")
                                        End If
                                        checkedVariants = checkedVariants + 1
                                    End If
                                Next
                                If tNode.Nodes.Count = 0 Or checkedVariants = 0 Then
                                    If Not constellationsChecked.ContainsKey(tNode.Text) Then
                                        constellationsChecked.Add(tNode.Text, "")
                                    End If

                                ElseIf tNode.Nodes.Count > 0 And checkedVariants = 1 Then
                                    ' alles schon getan
                                Else
                                    Call MsgBox("Error in Portfolio-Selection")
                                End If
                            End If
                        Next

                        If constellationsChecked.Count = 1 Then
                            ' tk , 14.11.22 AlleProjekte, ShowProjekte must not be changed 
                            ' it contains one or more projects which were loaded to be edited 
                            projectConstellations.clearLoadedPortfolios()

                            ' hole Portfolio (pName,vName) aus den db
                            Dim portfolioName As String = constellationsChecked.First.Key
                            Dim portfolioVariantName As String = constellationsChecked.First.Value
                            Dim PiCPortfolio As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(portfolioName,
                                                                                                                               dbPortfolioNames(portfolioName),
                                                                                                                               cTimestamp, err,
                                                                                                                               variantName:=portfolioVariantName,
                                                                                                                               storedAtOrBefore:=storedAtOrBefore)
                            projectConstellations.Add(PiCPortfolio)

                            sw.StartTimer()

                            ' now load the single projects of the Portfolio 
                            If Not IsNothing(PiCPortfolio) Then
                                Dim msgKey As String = calcProjektKey(portfolioName, portfolioVariantName)
                                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & portfolioName, " start of Operation ... ")

                                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In PiCPortfolio.Liste

                                    Dim pName As String = getPnameFromKey(kvp.Key)
                                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                                    If kvp.Value.show = True Then
                                        Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, storedAtOrBefore)

                                        If Not IsNothing(hproj) Then

                                            If Not AlleProjekte.Containskey(calcProjektKey(pName, vName)) Then
                                                AlleProjekte.Add(hproj)
                                            End If

                                            If Not ShowProjekte.contains(pName) Then
                                                ' removes hproj from ShowProjekte, if already in there
                                                ShowProjekte.Add(hproj)
                                            End If

                                        Else
                                            Call logger(ptErrLevel.logWarning, "Loading " & kvp.Key & " failed ..", " Operation continued ...")
                                        End If
                                    End If

                                Next

                                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & portfolioName, " End of Operation ... ")

                            Else
                                Dim msgTxt As String = "Load Portfolio " & portfolioName & " failed .."
                                Call logger(ptErrLevel.logError, "Load Portfolio " & portfolioName, " failed ..")
                                Throw New ArgumentException(msgTxt)
                            End If

                            sw.EndTimer()

                        End If

                    End If


                    Cursor.Current = Cursors.Default


                    ' Timer
                    If awinSettings.visboDebug Then
                        Call MsgBox("PTLadenKonstellation 3rd Part took: " & sw.EndTimer & "milliseconds")
                    End If

                    enableOnUpdate = True
                Catch ex As Exception

                End Try

            Else
                If awinSettings.englishLanguage Then
                    Call MsgBox("there is no portfolio ...")
                Else
                    Call MsgBox("kein Portfolio vorhanden ...")
                End If

            End If


        Else
            Call MsgBox("Datebase Connection lost ...")
        End If



    End Sub

    ''' <summary>
    ''' zeigt zwei Windows an, bestehend aus der Massen-Edit Ressourcen bz. Kosten  Tabelle und der meCharts Tabelle   
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTSPEshowCharts(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg
        Dim former_showRangeLeft As Integer = showRangeLeft
        Dim former_showRangeRight As Integer = showRangeRight

        ' ur:2022.03.29: change the timezone because of TSO orga
        Dim warningFrm As New frmChangedTimeZone
        If Not notAgain Then
            warningFrm.ShowDialog()
        End If

        ' tk 16.11.22 switch off Budget Chart 
        awinSettings.fullProtocol = False

        Dim timeZoneWasOff As Boolean = setTimeZoneIfTimeZonewasOff(True)

        ' whether or there need to be three or four charts
        Dim withSkills As Boolean = RoleDefinitions.getAllSkillIDs.Count > 0

        ' das Ganze nur machen, wenn das Chart nicht ohnehin schon gezeigt wird ... 
        Try
            If Not IsNothing(projectboardWindows(PTwindows.meChart)) Then
                Exit Sub
            End If
        Catch ex As Exception
        End Try

        Try
            If IsNothing(projectboardWindows(PTwindows.massEdit)) Then
                projectboardWindows(PTwindows.massEdit) = appInstance.ActiveWindow
            End If
        Catch ex As Exception

        End Try

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Dim meWS As Excel.Worksheet = CType(appInstance.ActiveSheet, Excel.Worksheet)

        Dim currentRow As Integer
        Dim currentColumn As Integer
        Dim prcTyp As String
        Dim pName As String = ""
        Dim hproj As clsProjekt = Nothing


        ' das Markieren der selektierten Projekte einschalten ..
        awinSettings.showValuesOfSelected = True


        Try
            currentRow = appInstance.ActiveCell.Row
            currentColumn = appInstance.ActiveCell.Column
        Catch ex As Exception
            currentRow = 2
            currentColumn = visboZustaende.meColRC
        End Try

        Dim rcName As String = CStr(meWS.Cells(currentRow, visboZustaende.meColRC).value)
        Dim rcNameID As String = getRCNameIDfromExcelRange(CType(meWS.Range(meWS.Cells(currentRow, visboZustaende.meColRC), meWS.Cells(currentRow, visboZustaende.meColRC + 1)), Excel.Range))

        Dim rcNameTeamID As Integer = -1
        Dim rcID As Integer = RoleDefinitions.parseRoleNameID(rcNameID, rcNameTeamID)



        If IsNothing(rcName) Then
            rcName = ""
        End If

        If IsNothing(rcNameID) Then
            rcNameID = ""
        End If

        ' jetzt ist entweder was gefunden oder es ist komplett ohne Werte 
        If rcName = "" Then
            currentRow = 2
            Try
                prcTyp = DiagrammTypen(1)
                rcName = RoleDefinitions.getDefaultTopNodeName
            Catch ex As Exception
                prcTyp = DiagrammTypen(1)
                rcName = ""
            End Try

        Else
            If RoleDefinitions.containsNameOrID(rcNameID) Then
                prcTyp = DiagrammTypen(1)
            ElseIf CostDefinitions.containsName(rcName) Then
                prcTyp = DiagrammTypen(2)
            Else
                prcTyp = DiagrammTypen(1)
                rcName = RoleDefinitions.getDefaultTopNodeName
                rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, "")
            End If

        End If

        pName = CStr(CType(appInstance.ActiveSheet, Excel.Worksheet).Cells(currentRow, visboZustaende.meColpName).value)

        Dim visboWorkbook As Excel.Workbook = appInstance.Workbooks.Item(myProjektTafel)


        ' now get Windows Parameters such as maxScreenWidth and maxScreenHeight 

        Call setWindowParameters()

        With projectboardWindows(PTwindows.massEdit)
            .WindowState = Excel.XlWindowState.xlNormal
            .EnableResize = True
        End With


        projectboardWindows(PTwindows.meChart) = appInstance.ActiveWindow.NewWindow

        ' tk 14.11. take Table 
        'visboWorkbook.Worksheets.Item(arrWsNames(ptTables.meCharts)).activate()
        visboWorkbook.Worksheets.Item(arrWsNames(ptTables.meAT)).activate()
        With projectboardWindows(PTwindows.meChart)
            .WindowState = Excel.XlWindowState.xlNormal
            .EnableResize = True
            .DisplayHorizontalScrollBar = False
            .DisplayVerticalScrollBar = False
            .DisplayGridlines = False
            .DisplayHeadings = False
            .DisplayRuler = False
            .DisplayOutline = False
            .DisplayWorkbookTabs = False
            .Caption = bestimmeWindowCaption(PTwindows.meChart)
        End With


        visboWorkbook.Windows.Arrange(Excel.XlArrangeStyle.xlArrangeStyleHorizontal)

        ' in Abhängigkeit von der Resolution soll jetzt mehr oder weniger prozentualer Platz spendiert werden 
        'Dim teilungsfaktor As Double = 0.7
        Dim teilungsfaktor As Double = 0.67
        If maxScreenHeight < 520 Then
            teilungsfaktor = 0.6
        End If

        ' jetzt die Größen anpassen 
        With projectboardWindows(PTwindows.massEdit)
            .Top = 0
            .Left = 1.0 + frmCoord(PTfrm.basis, PTpinfo.left)
            '.Height = 3 / 4 * maxScreenHeight
            .Height = teilungsfaktor * maxScreenHeight
            .Width = maxScreenWidth - 7.0        ' -7.0, damit der Scrollbar angeklickt werden kann
        End With

        ' jetzt die Größen anpassen 
        With projectboardWindows(PTwindows.meChart)
            .Top = teilungsfaktor * maxScreenHeight + 1
            .Left = 1.0 + frmCoord(PTfrm.basis, PTpinfo.left)
            .Height = (1 - teilungsfaktor) * maxScreenHeight - 1
            .Width = maxScreenWidth - 7.0        ' -7.0, damit der Scrollbar angeklickt werden kann
        End With


        ' Check: was ist das aktuelle Sheet 
        'Dim checkSheet As Object = projectboardWindows(1).ActiveSheet

        ' jetzt das Mass-Edit Window aktivieren 

        projectboardWindows(PTwindows.massEdit).Activate()
        'With CType(projectboardWindows(1).ActiveSheet, Excel.Worksheet)
        '    CType(.Cells(currentRow, currentColumn), Excel.Range).Activate()
        'End With

        Dim anz As Integer = appInstance.ActiveWorkbook.Windows.Count

        ' jetzt werden die Charts ggf erzeugt ...  
        If CType(CType(projectboardWindows(PTwindows.meChart).ActiveSheet, Excel.Worksheet).ChartObjects, Excel.ChartObjects).Count = 0 Then
            ' sie müssen erzeugt werden

            ' jetzt das Projekt Ergebnis Chart anzeigen
            Dim dummyObj As Excel.ChartObject = Nothing
            Dim chLeft As Double = 2

            ' show 4 Windows, if there are SKills

            ''ur:2022.03.29: in future only 2/3 charts
            'Dim stdBreite As Double = (projectboardWindows(PTwindows.meChart).UsableWidth - 12) / 3
            ' tk 7.4 show Budget as well
            Dim stdBreite As Double = (projectboardWindows(PTwindows.meChart).UsableWidth - 12)

            If awinSettings.fullProtocol Then
                stdBreite = (projectboardWindows(PTwindows.meChart).UsableWidth - 12) / 2
            End If

            Dim showFourDiagrams As Boolean = (withSkills And visboZustaende.projectBoardMode = ptModus.massEditRessSkills)
            If showFourDiagrams Then
                ''ur:2022.03.29: in future only 2/3 charts
                'stdBreite = (projectboardWindows(PTwindows.meChart).UsableWidth - 12) / 4
                If awinSettings.fullProtocol Then
                    stdBreite = (projectboardWindows(PTwindows.meChart).UsableWidth - 12) / 3
                Else
                    stdBreite = (projectboardWindows(PTwindows.meChart).UsableWidth - 12) / 2
                End If

            End If

            Dim chWidth As Double = stdBreite
            Dim chHeight As Double = projectboardWindows(PTwindows.meChart).UsableHeight - 2

            ' tk 21.12.22 0.7 was formerly 0.8
            If chHeight < 0.7 * projectboardWindows(PTwindows.meChart).Height Then
                chHeight = 0.7 * projectboardWindows(PTwindows.meChart).Height
            End If

            'Dim chTop As Double = 5
            Dim chTop As Double = 0

            ''ur: 2022.03.29: no longern shown because of new TSO-orga
            '' show the project Profit/Lost Diagram
            If editProjekteInSPE.contains(pName) Then
                hproj = editProjekteInSPE.getProject(pName)

                selectedProjekte.Clear(False)
                selectedProjekte.Add(hproj, False)

                If awinSettings.fullProtocol Then
                    Call createProjektErgebnisCharakteristik2(hproj, dummyObj, PThis.current, chTop, chLeft, chWidth, chHeight, False, True)
                End If

            End If


            ' now show Utilization Chart
            ' das Auslastungs-Chart Orga-Einheit
            Dim repObj As Excel.ChartObject = Nothing
            If awinSettings.fullProtocol Then
                chLeft = chLeft + chWidth + 2
            End If

            chWidth = stdBreite



            Dim myCollection As New Collection
            If rcName <> "" Then
                myCollection.Add(rcName)
                Call awinCreateprcCollectionDiagram(myCollection, repObj, chTop, chLeft,
                                                                       chWidth, chHeight, False, prcTyp, True, CDbl(awinSettings.fontsizeTitle))

                ' show only when skill are relevant 
                If showFourDiagrams Then
                    ' das Auslastungs-Chart Skill
                    repObj = Nothing
                    chLeft = chLeft + chWidth + 2
                    chWidth = stdBreite

                    myCollection.Clear()
                    If rcNameID = "" Then
                        rcNameID = RoleDefinitions.bestimmeRoleNameID(rcName, "")
                    End If
                    myCollection.Add(rcNameID)
                    Call awinCreateprcCollectionDiagram(myCollection, repObj, chTop, chLeft,
                                                                           chWidth, chHeight, False, prcTyp, True, CDbl(awinSettings.fontsizeTitle),
                                                                           isMESkillChart:=True)
                End If
            End If



            ' now Show Soll-Ist Vergleich mit Plan vs Last_Plan oder Plan vs Beauftragung
            ' tk do not show soll-ist
            'Dim obj As Excel.ChartObject = Nothing
            'chLeft = chLeft + chWidth + 2
            'chWidth = stdBreite

            '' hier muss jetzt das lproj bestimmt werden 
            'Dim lproj As clsProjekt = Nothing


            'Dim comparisonTyp As Integer
            'Dim qualifier2 As String = ""
            'Dim teamID As Integer = -1

            'Dim scInfo As New clsSmartPPTChartInfo
            'With scInfo
            '    .hproj = hproj
            '    .vergleichsArt = PTVergleichsArt.beauftragung
            '    .einheit = PTEinheiten.euro
            '    .prPF = ptPRPFType.project
            '    If visboZustaende.projectBoardMode = ptModus.massEditRessSkills Then
            '        .elementTyp = ptElementTypen.roles
            '        scInfo.einheit = PTEinheiten.personentage
            '    Else
            '        .elementTyp = ptElementTypen.costs
            '        scInfo.einheit = PTEinheiten.euro
            '    End If
            '    .chartTyp = PTChartTypen.Balken
            '    .detailID = PTprdk.KostenBalken2
            'End With

            'Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString

            'If awinSettings.meCompareVsLastPlan Then
            '    Dim vpID As String = ""
            '    lproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, vpID, awinSettings.meDateForLastPlan, err)
            '    comparisonTyp = PTprdk.KostenBalken2

            '    scInfo.vergleichsArt = PTVergleichsArt.planungsstand

            '    scInfo.vergleichsDatum = awinSettings.meDateForLastPlan
            '    scInfo.vglProj = lproj
            '    scInfo.vergleichsTyp = PTVergleichsTyp.standVom
            '    scInfo.q2 = rcName
            '    scInfo.detailID = PTprdk.KostenBalken2
            'Else

            '    If awinSettings.meCompareWithLastVersion Then

            '        lproj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, Date.Now, err)
            '        comparisonTyp = PTprdk.KostenBalken2

            '        scInfo.vglProj = lproj
            '        scInfo.vergleichsTyp = PTVergleichsTyp.letzter
            '        scInfo.q2 = ""
            '        scInfo.detailID = PTprdk.KostenBalken2

            '        If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
            '            If myCustomUserRole.specifics.Length > 0 Then
            '                If RoleDefinitions.containsNameOrID(myCustomUserRole.specifics) Then

            '                    comparisonTyp = PTprdk.PersonalBalken2
            '                    scInfo.q2 = RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, teamID).name

            '                End If
            '            End If

            '        ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Then

            '            ' wenn es ein Team-Member ist , soll nachgesehen werden, ob es für das Team Vorgaben gibt 
            '            ' wenn nein, dann soll die Kostenstelle der Person genommen werden, sofern sie 
            '            If rcName <> "" Then
            '                Dim potentialParents() As Integer = RoleDefinitions.getIDArray(myCustomUserRole.specifics)

            '                If Not IsNothing(potentialParents) And Not IsNothing(lproj) Then

            '                    Dim tmpParentName As String = ""

            '                    If rcNameTeamID = -1 Then
            '                        tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
            '                    Else
            '                        Dim tmpTeamName As String = RoleDefinitions.getRoleDefByID(rcNameTeamID).name
            '                        tmpParentName = RoleDefinitions.chooseParentFromList(tmpTeamName, potentialParents)
            '                        If tmpParentName = "" Then
            '                            tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
            '                        Else
            '                            Dim tmpParentNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpParentName, "")
            '                            If lproj.containsRoleNameID(tmpParentNameID) Then
            '                                ' passt bereits 
            '                            Else
            '                                tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
            '                            End If

            '                        End If
            '                    End If

            '                    If tmpParentName <> "" Then
            '                        scInfo.q2 = tmpParentName
            '                    End If
            '                End If


            '            End If

            '        End If



            '    Else
            '        lproj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)
            '        comparisonTyp = PTprdk.KostenBalken

            '        scInfo.vglProj = lproj
            '        scInfo.vergleichsTyp = PTVergleichsTyp.erster
            '        scInfo.q2 = ""
            '        scInfo.detailID = PTprdk.KostenBalken

            '        If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
            '            If myCustomUserRole.specifics.Length > 0 Then
            '                If RoleDefinitions.containsNameOrID(myCustomUserRole.specifics) Then

            '                    comparisonTyp = PTprdk.PersonalBalken
            '                    scInfo.q2 = RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, teamID).name

            '                End If
            '            End If

            '        ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Then

            '            If rcName <> "" Then
            '                Dim potentialParents() As Integer = RoleDefinitions.getIDArray(myCustomUserRole.specifics)

            '                If Not IsNothing(potentialParents) And Not IsNothing(lproj) Then

            '                    Dim tmpParentName As String = ""

            '                    If rcNameTeamID = -1 Then
            '                        tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
            '                    Else
            '                        Dim tmpTeamName As String = RoleDefinitions.getRoleDefByID(rcNameTeamID).name
            '                        tmpParentName = RoleDefinitions.chooseParentFromList(tmpTeamName, potentialParents)
            '                        If tmpParentName = "" Then
            '                            tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
            '                        Else
            '                            Dim tmpParentNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpParentName, "")
            '                            If lproj.containsRoleNameID(tmpParentNameID) Then
            '                                ' passt bereits 
            '                            Else
            '                                tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
            '                            End If

            '                        End If
            '                    End If

            '                    If tmpParentName <> "" Then
            '                        scInfo.q2 = tmpParentName
            '                    End If

            '                End If


            '            End If
            '        End If

            '    End If

            'End If

            ''Dim vglBaseline As Boolean = Not IsNothing(lproj)
            'Dim reportObj As Excel.ChartObject = Nothing


            'Try

            '    Call createRessBalkenOfProject(scInfo, 2, reportObj, chTop, chLeft, chHeight, chWidth, True,
            '                                       calledFromMassEdit:=True)

            '    ' alt, am 20.2. durch obiges ersetzt 
            '    'If scInfo.q2 = "" Then
            '    '    Call createCostBalkenOfProject(hproj, lproj, reportObj, 2, chTop, chLeft, chHeight, chWidth, False, comparisonTyp)
            '    'Else
            '    '    Call createRessBalkenOfProject(scInfo, 2, reportObj, chTop, chLeft, chHeight, chWidth, True,
            '    '                                   calledFromMassEdit:=True)
            '    'End If


            'Catch ex As Exception

            'End Try

            ' tk end of do not show soll-Ist


        Else
            ' sie sind schon da 

        End If

        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True

        ' jetzt das Chart-Window aktivieren (sonst bleibt Ribbon stehen)
        projectboardWindows(PTwindows.meChart).Activate()


        showRangeLeft = former_showRangeLeft
        showRangeRight = former_showRangeRight


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
