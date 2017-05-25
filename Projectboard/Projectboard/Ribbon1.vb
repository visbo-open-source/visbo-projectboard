Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ClassLibrary1
Imports MongoDbAccess
Imports WPFPieChart
Imports Microsoft.Office.Core
'Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
'Imports MSProject = Microsoft.Office.Interop.MSProject
Imports System.Security.Principal
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows



'TODO: Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

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
    Implements Microsoft.Office.Core.IRibbonExtensibility

    Private ribbon As Microsoft.Office.Core.IRibbonUI

    Private tempSkipChanges As Boolean = False

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Microsoft.Office.Core.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("ExcelWorkbook1.Ribbon1.xml")
    End Function

#Region "Menübandrückrufe"
    'Erstellen Sie hier Rückrufmethoden. Weitere Informationen über das Hinzufügen von Rückrufmethoden erhalten Sie, indem Sie das Menüband-XML-Element im Projektmappen-Explorer markieren und dann F1 drücken.
    Public Sub Ribbon_Load(ByVal ribbonUI As Microsoft.Office.Core.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Sub PTNeueKonstellation(control As IRibbonControl)


        Dim storeConstellationFrm As New frmStoreConstellation
        Dim returnValue As DialogResult
        Dim constellationName As String

        Dim returnRequest As Boolean = False
        Dim controlID As String = control.Id
        Dim jetzt As Date = Date.Now

        Call projektTafelInit()


        '
        If AlleProjekte.Count > 0 Then
            returnValue = storeConstellationFrm.ShowDialog  ' Aufruf des Formulars zur Eingabe des Portfolios

            If returnValue = DialogResult.OK Then
                constellationName = storeConstellationFrm.ComboBox1.Text


                Call storeSessionConstellation(constellationName)

                ' setzen der public variable, welche Konstellation denn jetzt gesetzt ist
                currentConstellationName = constellationName
            End If
        Else
            Call MsgBox("Es sind keine Projekte in der Projekt-Tafel geladen!")
        End If
        ' 
        ' Ende alte Version; vor dem 26.10.14




        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' speichert die ausgewählten SessionConstellations in die Datenbank 
    ''' dabei wird sichergestellt, dass alle Projekte, die 
    ''' noch gar nicht in der DB existieren oder die sich im Vergleich zur DB-Versaion geändert haben, 
    ''' auch gespeichert werden 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTStoreKonstellationsToDB(control As IRibbonControl)

        Dim storeConstellationFrm As New frmLoadConstellation
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim DBtimeStamp As Date = Date.Now

        Dim outPutCollection As New Collection


        With storeConstellationFrm
            If awinSettings.englishLanguage Then
                .Text = "store Scenario(s) in Datenbase"
            Else
                .Text = "Szenario(s) in Datenbank speichern"
            End If

            .constellationsToShow = projectConstellations
            .retrieveFromDB = False
            .lblStandvom.Visible = False
            .requiredDate.Visible = False
            .addToSession.Visible = False
        End With

        Dim returnValue As DialogResult = storeConstellationFrm.ShowDialog

        If returnValue = DialogResult.OK Then

            For i As Integer = 1 To storeConstellationFrm.ListBox1.SelectedItems.Count

                Dim constellationName As String = CStr(storeConstellationFrm.ListBox1.SelectedItems.Item(i - 1))
                Dim currentConstellation As clsConstellation = projectConstellations.getConstellation(constellationName)

                Call storeSingleConstellationToDB(outPutCollection, currentConstellation)

            Next

        End If

    End Sub

    Sub PTLadenKonstellation(control As IRibbonControl)

        Dim loadFromDatenbank As String = "PT5G1B1"
        Dim loadConstellationFrm As New frmLoadConstellation
        Dim storedAtOrBefore As Date = Date.Now
        Dim ControlID As String = control.Id
        Dim timeStampsCollection As New Collection
        Dim dbConstellations As New clsConstellations

        Dim initMessage As String = "Es sind dabei folgende Probleme aufgetreten" & vbLf & vbLf

        Dim successMessage As String = initMessage
        Dim returnValue As DialogResult

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        Call projektTafelInit()

        ' Wenn das Laden eines Portfolios aus dem Menu Datenbank aufgerufen wird, so werden erneut alle Portfolios aus der Datenbank geholt

        If ControlID = loadFromDatenbank And Not noDB Then

            If request.pingMongoDb() Then

                dbConstellations = request.retrieveConstellationsFromDB()

                Try
                    timeStampsCollection = request.retrieveZeitstempelFromDB()
                    'Dim heute As String = Date.Now.ToString
                    If timeStampsCollection.Count > 0 Then
                        With loadConstellationFrm
                            .constellationsToShow = dbConstellations
                            .retrieveFromDB = True
                            If timeStampsCollection.Count > 0 Then
                                '.earliestDate = CDate(timeStampsCollection.Item(1))
                                .earliestDate = CDate(timeStampsCollection.Item(timeStampsCollection.Count)).Date.AddHours(23).AddMinutes(59)
                            Else
                                .earliestDate = Date.Now.Date.AddHours(23).AddMinutes(59)
                            End If

                            '.listOfTimeStamps = timeStampsCollection
                        End With
                    End If

                Catch ex As Exception

                End Try

            Else
                Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
            End If
        Else
            With loadConstellationFrm
                .constellationsToShow = projectConstellations
                .retrieveFromDB = False
            End With
        End If

        enableOnUpdate = False

        If AlleProjekte.Count > 0 Then
            loadConstellationFrm.addToSession.Checked = False
        Else
            loadConstellationFrm.addToSession.Checked = False
            loadConstellationFrm.addToSession.Visible = False
        End If



        returnValue = loadConstellationFrm.ShowDialog

        If returnValue = DialogResult.OK Then

            appInstance.ScreenUpdating = False

            If Not IsNothing(loadConstellationFrm.requiredDate.Value) Then
                storedAtOrBefore = CDate(loadConstellationFrm.requiredDate.Value)
            Else
                storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)
            End If

            Dim constellationsToDo As New clsConstellations

            For Each tmpName As String In loadConstellationFrm.ListBox1.SelectedItems

                Dim constellation As clsConstellation = projectConstellations.getConstellation(tmpName)
                If Not IsNothing(constellation) Then
                    If Not constellationsToDo.Contains(constellation.constellationName) Then
                        constellationsToDo.Add(constellation)
                    End If
                Else
                    constellation = dbConstellations.getConstellation(tmpName)
                    If Not IsNothing(constellation) Then
                        If Not constellationsToDo.Contains(constellation.constellationName) Then
                            constellationsToDo.Add(constellation)
                        End If
                        projectConstellations.Add(constellation)
                    End If

                End If

            Next


            Dim clearBoard As Boolean = Not loadConstellationFrm.addToSession.Checked
            'Dim clearSession As Boolean = ((ControlID = loadFromDatenbank) And clearBoard)
            Dim clearSession As Boolean = False
            If constellationsToDo.Count > 0 Then
                Call showConstellations(constellationsToDo, clearBoard, clearSession, storedAtOrBefore)
            End If

            ' jetzt muss die Info zu den Schreibberechtigungen geholt werden 
            If Not noDB Then
                writeProtections.adjustListe = request.retrieveWriteProtectionsFromDB(AlleProjekte)
            End If

            appInstance.ScreenUpdating = True

        End If

        enableOnUpdate = True

    End Sub

    Sub PTAendernKonstellation(control As IRibbonControl)

        Call PBBChangeCurrentPortfolio()

    End Sub
    Sub PTRemoveKonstellation(control As IRibbonControl)

        Dim ControlID As String = control.Id

        Dim removeConstFilterFrm As New frmRemoveConstellation
        Dim constFilterName As String

        Dim returnValue As DialogResult

        Call projektTafelInit()


        Dim deleteDatenbank As String = "Pt5G3B1"
        Dim deleteFromSession As String = "PT2G3M1B3"
        Dim deleteFilter As String = "Pt6G3B5"

        Dim removeFromDB As Boolean

        If ControlID = deleteDatenbank And Not noDB Then
            removeConstFilterFrm.frmOption = "ProjConstellation"
            removeFromDB = True

            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            If request.pingMongoDb() Then
                projectConstellations = request.retrieveConstellationsFromDB()
            Else
                Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
                removeFromDB = False
            End If

        ElseIf ControlID = deleteFromSession Then
            removeConstFilterFrm.frmOption = "ProjConstellation"
            removeFromDB = False

        ElseIf ControlID = deleteFilter And Not noDB Then
            removeConstFilterFrm.frmOption = "DBFilter"
            removeFromDB = True

            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            If request.pingMongoDb() Then
                filterDefinitions.filterListe = request.retrieveAllFilterFromDB(False)
            Else
                Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
                removeFromDB = False
            End If

        Else
            removeFromDB = False
        End If

        enableOnUpdate = False

        returnValue = removeConstFilterFrm.ShowDialog

        If returnValue = DialogResult.OK Then
            If ControlID = deleteDatenbank Or _
                ControlID = deleteFromSession Then

                constFilterName = removeConstFilterFrm.ListBox1.Text

                Call awinRemoveConstellation(constFilterName, removeFromDB)
                Call MsgBox(constFilterName & " wurde gelöscht ...")

                If constFilterName = currentConstellationName Then

                    ' aktuelle Konstellation unter dem Namen 'Last' speichern
                    'Call storeSessionConstellation("Last")
                    'currentConstellationName = "Last"
                Else
                    ' aktuelle Konstellation bleibt unverändert
                End If


            End If
            If ControlID = deleteFilter Then

                Dim removeOK As Boolean = False
                Dim filter As clsFilter = Nothing

                constFilterName = removeConstFilterFrm.ListBox1.Text

                filter = filterDefinitions.retrieveFilter(constFilterName)

                Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                If request.pingMongoDb() Then

                    ' Filter muss aus der Datenbank gelöscht werden.

                    removeOK = request.removeFilterFromDB(filter)
                    If removeOK = False Then
                        Call MsgBox("Fehler bei Löschen des Filters: " & constFilterName)
                    Else
                        ' DBFilter ist nun aus der DB gelöscht
                        ' hier: wird der Filter nun noch aus der Filterliste gelöscht
                        Call filterDefinitions.filterListe.Remove(constFilterName)
                        Call MsgBox(constFilterName & " wurde gelöscht ...")
                    End If
                Else
                    Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!" & vbLf & "DB Filter '" & filter.name & "'konnte in der Datenbank nicht gelöscht werden")
                    removeOK = False
                End If

            End If
        End If
        enableOnUpdate = True

    End Sub


    Sub PT5StoreProjects(control As IRibbonControl)

        Dim storedProj As Integer = 0
        Dim msgresult As Integer
        Dim awinSelection As Excel.ShapeRange
        Dim emptySelection As Boolean = True

        Call projektTafelInit()

        Try
            If AlleProjekte.Count > 0 Then

                Try
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                If IsNothing(awinSelection) Then
                    emptySelection = True
                ElseIf awinSelection.Count > 0 Then
                    emptySelection = False
                    storedProj = StoreSelectedProjectsinDB()    ' es werden die selektierten Projekte einschl. der geladenen Varianten 
                Else
                    emptySelection = True
                End If

                ' in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

                If emptySelection Then
                    msgresult = MsgBox("Es wurde kein Projekt selektiert. " & vbLf & "Alle Projekte und Varianten speichern?", MsgBoxStyle.OkCancel)

                    If msgresult = MsgBoxResult.Ok Then
                        ' Mouse auf Wartemodus setzen
                        appInstance.Cursor = Excel.XlMousePointer.xlWait
                        'Projekte speichern
                        Call StoreAllProjectsinDB()
                        ' Mouse wieder auf Normalmodus setzen
                        appInstance.Cursor = Excel.XlMousePointer.xlDefault
                    End If
                Else
                    'Call MsgBox("Es wurden " & storedProj & " Projekte gespeichert!")
                End If

            Else
                Call MsgBox("keine Projekte zu speichern ...")
            End If
        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try

        Call awinDeSelect()

        ' jetzt werden die Protection-Darstellung der Projekt-Namen auf der Multiprojekt-Tafel wieder aktualisiert 


    End Sub

    ''' <summary>
    ''' alles, d.h Projekte, Szenarien und Abhängigkeiten speichern
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT5StoreEverything(control As IRibbonControl)

        Dim storedProj As Integer = 0

        Call projektTafelInit()

        Try
            If AlleProjekte.Count > 0 Then

                appInstance.Cursor = Excel.XlMousePointer.xlWait

                'Projekte, Szenarien und Abhängigkeiten  speichern
                Call StoreAllProjectsinDB(True)

                ' Mouse wieder auf Normalmodus setzen
                appInstance.Cursor = Excel.XlMousePointer.xlDefault

            Else
                Call MsgBox("es gibt nichts zu speichern ...")
            End If
        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try

        Call awinDeSelect()

    End Sub
    ''' <summary>
    ''' löscht die ausgewählten Projekte aus der Datenbank 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT5DeleteProjectsInDB(control As IRibbonControl)

        Call PBBDeleteProjectsInDB(control)

    End Sub

    Sub PT5DeleteProjectsInDBExceptF1(control As IRibbonControl)
        Call PBBDeleteProjectsInDB(control)
    End Sub


    ''' <summary>
    ''' ruft den Portfolio Browser auf, um Projekte, die in der Datenbank liegen zu schützen bzw. die Projekte, die aktuell in der Session geladen sind  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT5SetWriteProtection(control As IRibbonControl)
        Call PBBWriteProtections(control, True)
    End Sub


    ''' <summary>
    ''' löscht alles, was aktuell in der Session ist 
    ''' Projekte, Charts, Shapes ... 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT6G3ClearSession(control As IRibbonControl)

        Call projektTafelInit()

        ' Bestätigungs-Fenster aufrufen 
        Dim bestaetigeLoeschen As New frmconfirmDeletePrj
        Dim returnValue As DialogResult

        If awinSettings.englishLanguage Then
            bestaetigeLoeschen.botschaft = "Please confirm the reset of the complete session"
        Else
            bestaetigeLoeschen.botschaft = "Bitte bestätigen Sie das Löschen der kompletten Session"
        End If

        returnValue = bestaetigeLoeschen.ShowDialog

        If returnValue = DialogResult.Cancel Then
            ' nichts tun
        Else

            Call clearCompleteSession()
            
        End If

    End Sub

    Sub PTXHealingCustomFieldsOfVariants(control As IRibbonControl)
        

    End Sub

    ''' <summary>
    ''' löscht alle Beschriftungen in der PRojekt-Tafel 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT6DeleteBeschriftung(control As IRibbonControl)

        Dim todoList As New Collection
        
        Call projektTafelInit()


        Call deleteBeschriftungen()



    End Sub



    Sub PT6DeleteCharts(control As IRibbonControl)

        Call projektTafelInit()

        Dim currentWsName As String
        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            currentWsName = arrWsNames(ptTables.MPT)
        Else
            currentWsName = arrWsNames(ptTables.meRC)
        End If

        Call deleteChartsInSheet(currentWsName)


    End Sub
    Sub PT0SaveCockpit(control As IRibbonControl)


        Dim i As Integer = 1
        Dim storeCockpitFrm As New frmStoreCockpit
        Dim returnValue As DialogResult
        Dim cockpitName As String
        Try


            Call projektTafelInit()

            appInstance.ScreenUpdating = False
            enableOnUpdate = False

            Call awinDeSelect()

            Dim anzDiagrams As Integer = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)).ChartObjects, Excel.ChartObjects).Count


            If anzDiagrams > 0 Then


                ' hier muss die Auswahl des Names für das Cockpit erfolgen

                returnValue = storeCockpitFrm.ShowDialog  ' Aufruf des Formulars zur Eingabe des Cockpitnamens

                If returnValue = DialogResult.OK Then

                    cockpitName = storeCockpitFrm.ComboBox1.Text

                    Call awinStoreCockpit(cockpitName)

                Else

                    appInstance.ScreenUpdating = True
                    enableOnUpdate = True

                End If


                ' hier muss eventuell ein Neuzeichnen erfolgen
            Else
                Call MsgBox("Es ist kein Chart angezeigt")
                appInstance.ScreenUpdating = True
                enableOnUpdate = True
            End If

        Catch ex As Exception
            Throw New ArgumentException("PT0SaveCockpit: Fehler:  ", ex.Message)
        End Try

    End Sub

    Sub PT0ShowCockpit(control As IRibbonControl)


        Dim i As Integer = 1
        Dim loadCockpitFrm As New frmLoadCockpit
        Dim returnValue As DialogResult
        Dim cockpitName As String

        Call projektTafelInit()

        '' ''If showRangeRight - showRangeLeft >= minColumns - 1 Then
        '' ''Else
        '' ''    Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
        '' ''End If

        If ShowProjekte.Count < 1 Then

            Dim msgtxt As String = ""
            If awinSettings.englishLanguage Then
                msgtxt = "Please first load at least one project"
            Else
                msgtxt = "Bitte laden Sie zunächst mindestens ein Projekt"
            End If
            Call MsgBox(msgtxt)

        Else

            If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                ' alles ok , bereits gesetzt 

            Else
                If selectedProjekte.Count > 0 Then
                    showRangeLeft = selectedProjekte.getMinMonthColumn
                    showRangeRight = selectedProjekte.getMaxMonthColumn
                Else
                    showRangeLeft = ShowProjekte.getMinMonthColumn
                    showRangeRight = ShowProjekte.getMaxMonthColumn
                End If
                Call awinShowtimezone(showRangeLeft, showRangeRight, True)
            End If

            Dim awinSelection As Excel.ShapeRange

            Try
                'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
                awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
            Catch ex As Exception
                awinSelection = Nothing
            End Try


            appInstance.EnableEvents = False
            enableOnUpdate = False

            ' hier muss die Auswahl des Names für das Cockpit erfolgen

            returnValue = loadCockpitFrm.ShowDialog  ' Aufruf des Formulars zur Eingabe des Cockpitnamens

            If returnValue = DialogResult.OK Then

                cockpitName = loadCockpitFrm.ListBox1.Text

                appInstance.ScreenUpdating = False

                If loadCockpitFrm.deleteOtherCharts.Checked Then
                    ' erst alle anderen Charts löschen ... 
                    Dim currentWsName As String
                    If visboZustaende.projectBoardMode = ptModus.graficboard Then
                        currentWsName = arrWsNames(ptTables.MPT)
                    Else
                        currentWsName = arrWsNames(ptTables.meRC)
                    End If

                    Call deleteChartsInSheet(currentWsName)
                End If

                Try
                    Call awinLoadCockpit(cockpitName)

                    Dim hproj As clsProjekt = Nothing

                    ' nur wenn ein Projekt selektiert wurde, werden die Projekt-Charts aktualisiert
                    If Not awinSelection Is Nothing Then


                        If awinSelection.Count = 1 Then
                            Dim singleShp As Excel.Shape

                            ' jetzt die Aktion durchführen ...
                            singleShp = awinSelection.Item(1)

                            Try
                                hproj = ShowProjekte.getProject(singleShp.Name, True)
                            Catch ex As Exception
                                Call MsgBox("Projekt nicht gefunden ..." & singleShp.Name)
                                Exit Sub
                            End Try

                            Call aktualisiereCharts(hproj, True)

                            Call awinDeSelect()
                        End If
                    Else
                        Try
                            hproj = ShowProjekte.getProject(1)
                        Catch ex As Exception
                            Call MsgBox("Projekt nicht gefunden ..." & hproj.name)
                            Exit Sub
                        End Try

                        Call aktualisiereCharts(hproj, True)

                        Call awinDeSelect()

                    End If

                    Call awinNeuZeichnenDiagramme(9)

                Catch ex As Exception
                    appInstance.ScreenUpdating = True
                    Call MsgBox("Fehler beim Laden ..")
                End Try


                appInstance.ScreenUpdating = True

            Else
                appInstance.ScreenUpdating = True

            End If

        End If

        ' hier muss eventuell ein Neuzeichnen erfolgen
        enableOnUpdate = True
        appInstance.EnableEvents = True
    End Sub


    ''' <summary>
    ''' wird aktuell verwendet , um eine Stelle für Testen bestimmter Funktionalitäten zu haben
    ''' ohne dass eine neue Ribbon Erweiterung gemacht werden muss
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub awinTestNewFunctions(control As IRibbonControl)
        'Call MsgBox("Anzahl Aufrufe: " & anzahlCalls)

        Dim awinSelection As Excel.ShapeRange
        Dim i As Integer
        Dim hproj As clsProjekt
        Dim singleShp As Excel.Shape
        Dim ausgabeString As String = ""
        Dim vglWert As Integer
        Dim curCoord() As Double
        Dim key As String

        Call projektTafelInit()


        enableOnUpdate = False



        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)

        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' Es muss mindestens 1 Projekt selektiert sein
            For i = 1 To awinSelection.Count

                singleShp = awinSelection.Item(i)
                key = singleShp.Name
                hproj = ShowProjekte.getProject(singleShp.Name, True)
                vglWert = calcYCoordToZeile(singleShp.Top)
                curCoord = projectboardShapes.getCoord(singleShp.Name)

                ausgabeString = ausgabeString & hproj.name & ": " & hproj.tfZeile.ToString & _
                                 " - " & vglWert.ToString & "; " & _
                                 calcXCoordToDate(singleShp.Left).ToShortDateString & " vs. " & hproj.startDate.ToShortDateString & _
                                 " vs. " & calcXCoordToDate(curCoord(1)).ToShortDateString & singleShp.Left.ToString & vbLf


            Next i


        End If

        Call awinDeSelect()
        Call MsgBox(ausgabeString)

        enableOnUpdate = True

        

    End Sub

    Sub PT0ShowProjektInfo1(control As IRibbonControl)

        With visboZustaende
            If IsNothing(formProjectInfo1) And .projectBoardMode = ptModus.massEditRessCost Then

                formProjectInfo1 = New frmProjectInfo1
                Call updateProjectInfo1(visboZustaende.lastProject, visboZustaende.lastProjectDB)

                formProjectInfo1.Show()
            End If

        End With

    End Sub


    ''' <summary>
    ''' Rename Funktion für ein Projekt
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Rename(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange


        Dim phaseList As New Collection
        Dim milestoneList As New Collection
        Dim neuerVariantenName As String = ""
        Dim ok As Boolean = True
        Dim nameCollection As New Collection
        Dim abbruch As Boolean = False


        Try
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

            Call projektTafelInit()

            enableOnUpdate = False

            Try
                awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
            Catch ex As Exception
                awinSelection = Nothing
            End Try

            If Not awinSelection Is Nothing Then

                If awinSelection.Count = 1 Then


                    ' jetzt die Aktion durchführen ...
                    Dim pName As String = CStr(awinSelection.Item(1).Name)
                    ' wurde ein shape selektiert, das auch ein Projekt ist ? 

                    If ShowProjekte.contains(pName) Then

                        ' hier jetzt prüfen, ob es sich um ein geschütztes Projekt handelt ... 
                        hproj = ShowProjekte.getProject(pName)

                        Dim isProtectedbyOthers As Boolean = Not tryToprotectProjectforMe(hproj.name, hproj.variantName)
                        ' wenn schon das gewählte Projekt geschützt ist , dann gar nichts weiter machen ... 
                        If Not isProtectedbyOthers Then
                            ' hier das Fomular zur Eingabe des neuen Namens aufrufen ... 
                            Dim renameForm As New frmRenameProject
                            With renameForm
                                .oldName.Text = pName
                                .newName.Text = pName
                            End With

                            Dim result As DialogResult = renameForm.ShowDialog

                            If result = DialogResult.OK Then

                                Dim newName As String = renameForm.newName.Text

                                ' jetzt wird in der Datenbank umbenannt 
                                Try
                                    If request.projectNameAlreadyExists(pName, "", Date.Now) Or _
                                        request.projectNameAlreadyExists(pName, hproj.variantName, Date.Now) Then

                                        ok = request.renameProjectsInDB(pName, newName, dbUsername)
                                        If Not ok Then
                                            If awinSettings.englishLanguage Then
                                                Call MsgBox("rename cancelled: there is at least one write-protected variant for Project " & pName)
                                            Else
                                                Call MsgBox("Rename nicht durchgeführt: es gibt mindestens eine schreibgeschützte Variante im Projekt " & pName)
                                            End If
                                        Else
                                            writeProtections.adjustListe = request.retrieveWriteProtectionsFromDB(AlleProjekte)
                                        End If
                                    End If
                                Catch ex As Exception
                                    Call MsgBox("Fehlende Berechtigung?" & vbLf & ex.Message)
                                End Try


                                ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                                ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                                Try
                                    If ok Then

                                        Dim variantNamesCollection As Collection = AlleProjekte.getVariantNames(pName, False)
                                        hproj = ShowProjekte.getProject(pName)

                                        ' jetzt werden alle Vorkommen in den Session Constellations umbenannt 
                                        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste
                                            Dim anzahl As Integer = kvp.Value.renameProject(pName, newName)
                                        Next

                                        ' jetzt werden alle Vorkommen in Dependencies umbenannt 
                                        '  ....

                                        ' merken , welche Phasen, Meilensteine aktuell gezeigt werden 
                                        phaseList = projectboardShapes.getPhaseList(pName)
                                        milestoneList = projectboardShapes.getMilestoneList(pName)
                                        Dim key As String = calcProjektKey(hproj)
                                        ShowProjekte.Remove(pName)
                                        Call clearProjektinPlantafel(pName)


                                        ' jetzt müssen auch alle in der Session / AlleProjekte vorhandenen Varianten umbenannt werden 
                                        For Each vName As String In variantNamesCollection
                                            key = calcProjektKey(pName, vName)
                                            Dim tmpProj As clsProjekt = AlleProjekte.getProject(key)
                                            If Not IsNothing(tmpProj) Then
                                                AlleProjekte.Remove(key)
                                                tmpProj.name = newName
                                                key = calcProjektKey(newName, vName)
                                                AlleProjekte.Add(tmpProj)
                                            End If
                                        Next

                                        ' gibt es die Standard-Variante ? 
                                        key = calcProjektKey(pName, "")
                                        If AlleProjekte.Containskey(key) Then
                                            Dim tmpProj As clsProjekt = AlleProjekte.getProject(key)
                                            AlleProjekte.Remove(key)
                                            tmpProj.name = newName
                                            key = calcProjektKey(newName, "")
                                            AlleProjekte.Add(tmpProj)
                                        End If


                                        hproj.name = newName
                                        ShowProjekte.Add(hproj)

                                        Dim tmpCollection As New Collection

                                        Call ZeichneProjektinPlanTafel(tmpCollection, newName, hproj.tfZeile, phaseList, milestoneList)

                                    End If

                                Catch ex As Exception

                                    If awinSettings.englishLanguage Then
                                        Call MsgBox("Error when renaming project: " & ex.Message)
                                    Else
                                        Call MsgBox("Fehler bei Rename Projekt: " & ex.Message)
                                    End If


                                End Try
                            End If

                            ' jetzt kann der Schutz wieder freigegeben werden ...

                        Else
                            If awinSettings.englishLanguage Then
                                Call MsgBox("Project " & pName & " is write-protected and cannot be renamed!")
                            Else
                                Call MsgBox("Projekt " & pName & " ist schreibgeschützt und kann nicht umbenannt werden!")
                            End If
                        End If


                    Else
                        If awinSettings.englishLanguage Then
                            Call MsgBox("please select a project first ...")
                        Else
                            Call MsgBox("bitte Projekt selektieren ...")
                        End If

                    End If

                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox("please select just 1 project only ...")
                    Else
                        Call MsgBox("bitte nur ein Projekt selektieren ...")
                    End If

                End If


            Else
                If awinSettings.englishLanguage Then
                    Call MsgBox("please select a project first ...")
                Else
                    Call MsgBox("bitte Projekt selektieren ...")
                End If
            End If

            enableOnUpdate = True
        Catch ex As Exception

        End Try


    End Sub

    Sub PT2ProjektNeu(control As IRibbonControl)

        Dim ProjektEingabe As New frmProjektEingabe1
        Dim returnValue As DialogResult
        Dim zeile As Integer = 0
        ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        Call projektTafelInit()

        enableOnUpdate = False


        returnValue = ProjektEingabe.ShowDialog

        If returnValue = DialogResult.OK Then
            With ProjektEingabe
                Dim buName As String = CStr(.businessUnitDropBox.SelectedItem)

                If Not noDB Then

                    Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                    If request.pingMongoDb() Then

                        If Not request.projectNameAlreadyExists(projectname:=.projectName.Text, variantname:="", storedAtorBefore:=Date.Now) Then

                            ' Projekt existiert noch nicht in der DB, kann also eingetragen werden


                            Call TrageivProjektein(.projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart), _
                                               CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile, _
                                               CType(.sFit.Text, Double), CType(.risiko.Text, Double), CDbl(.volume.Text), _
                                               CStr(""), buName)
                        Else
                            Call MsgBox(" Projekt '" & .projectName.Text & "' existiert bereits in der Datenbank!")
                        End If
                    Else

                        Call MsgBox("Datenbank- Verbindung ist unterbrochen !")
                        appInstance.ScreenUpdating = True

                        ' Projekt soll trotzdem angezeigt werden
                        Call TrageivProjektein(.projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart), _
                                               CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile, _
                                               CType(.sFit.Text, Double), CType(.risiko.Text, Double), CDbl(.volume.Text), _
                                               CStr(""), buName)

                    End If

                Else

                    appInstance.ScreenUpdating = True

                    ' Projekt soll trotzdem angezeigt werden
                    Call TrageivProjektein(.projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart), _
                                           CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile, _
                                           CType(.sFit.Text, Double), CType(.risiko.Text, Double), CDbl(.volume.Text), _
                                           CStr(""), buName)

                End If

            End With
        End If

        ''If Not currentConstellationName.EndsWith("(*)") Then
        ''    currentConstellationName = currentConstellationName & " (*)"
        ''End If

        If currentConstellationName <> calcLastSessionScenarioName() Then
            currentConstellationName = calcLastSessionScenarioName()
        End If

        'Call storeSessionConstellation("Last")

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' eine neue Variante anlegen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2VarianteNeu(control As IRibbonControl)

        Call PBBVarianteNeu(control)

    End Sub

    ''' <summary>
    ''' beschriftet die ausgewählten Projekte 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTBeschriften(control As IRibbonControl)

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try


        ' wenn nichts selektiert ist, sollen alle beschriftet werden 

        Dim annotateFrm As New frmAnnotateProject
        annotateFrm.Show()


        'If Not awinSelection Is Nothing Then

        '    If awinSelection.Count >= 1 Then

        '        Dim annotateFrm As New frmAnnotateProject
        '        annotateFrm.Show()

        '    Else
        '        Call MsgBox("bitte mindestens ein Projekt selektieren ...")
        '    End If
        'Else
        '    Call MsgBox("bitte mindestens ein Projekt selektieren ...")
        'End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' stellt die selektierten Projekte im Extended View dar;
    ''' Proj.extendedView wird dabei auf true gesetzt 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTExtendedView(control As IRibbonControl)

        Dim awinSelection As Excel.ShapeRange
        Dim hproj As clsProjekt
        Dim singleShp As Excel.Shape
        Dim i As Integer

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count >= 1 Then

                ' Es muss mindestens 1 Projekt selektiert sein
                For i = 1 To awinSelection.Count


                    singleShp = awinSelection.Item(i)

                    Try
                        hproj = ShowProjekte.getProject(singleShp.Name, True)
                        hproj.extendedView = True

                    Catch ex As Exception
                        Call MsgBox(" Fehler in extended Darstellung " & singleShp.Name & " , Modul: PTExtendedView")
                        enableOnUpdate = True
                        Exit Sub
                    End Try

                Next i

                appInstance.ScreenUpdating = True

                Call awinClearPlanTafel()
                Call awinZeichnePlanTafel(True)

                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte mindestens ein Projekt selektieren ...")
            End If
        Else
            Call MsgBox("bitte mindestens ein Projekt selektieren ...")
        End If

        awinDeSelect()
        enableOnUpdate = True

    End Sub
    ''' <summary>
    ''' stellt die selektierten Projekte im normalen Modus dar;
    ''' Proj.extendedView wird dabei auf false gesetzt 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTLineView(control As IRibbonControl)

        Dim awinSelection As Excel.ShapeRange
        Dim hproj As clsProjekt
        Dim singleShp As Excel.Shape
        Dim i As Integer

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count >= 1 Then

                ' Es muss mindestens 1 Projekt selektiert sein
                For i = 1 To awinSelection.Count


                    singleShp = awinSelection.Item(i)

                    Try
                        hproj = ShowProjekte.getProject(singleShp.Name, True)
                        hproj.extendedView = False

                    Catch ex As Exception
                        Call MsgBox(" Fehler in einzeiliger Darstellung " & singleShp.Name & " , Modul: PTLineView")
                        enableOnUpdate = True
                        Exit Sub
                    End Try

                Next i

                appInstance.ScreenUpdating = False

                Call awinClearPlanTafel()
                Call awinZeichnePlanTafel(True)

                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte mindestens ein Projekt selektieren ...")
            End If
        Else
            Call MsgBox("bitte mindestens ein Projekt selektieren ...")
        End If

        awinDeSelect()
        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' aktiviert die selektierte Variante 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2VarianteAktiv(control As IRibbonControl)

        Call PBBVarianteAktiv(control)


    End Sub

    ''' <summary>
    ''' die Variante, die übernommen werden soll, muss bereits in der Showprojekte sein und selektiert sein
    ''' Das Projekt wird zur Standard-Variante 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2VarianteUebernehmen(control As IRibbonControl)

        Dim awinSelection As Excel.ShapeRange
        Dim i As Integer
        Dim hproj As clsProjekt
        Dim singleShp As Excel.Shape
        'Dim ausgabeString As String = ""
        'Dim vglWert As Integer
        'Dim curCoord() As Double
        Dim key As String

        Call projektTafelInit()


        enableOnUpdate = False



        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)

        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' Es muss mindestens 1 Projekt selektiert sein
            For i = 1 To awinSelection.Count

                singleShp = awinSelection.Item(i)
                hproj = ShowProjekte.getProject(singleShp.Name, True)

                ' jetzt prüfen : die Variante kann nur dann zur Standard-Variante gemacht werden, 
                ' wenn die Standard-Variante nicht geschützt ist ..


                If tryToprotectProjectforMe(hproj.name, "") Then
                    ' ist erlaubt ...
                    ' das Projekt zur Standard Variante machen 


                    Dim oldvName As String = hproj.variantName
                    Dim newvName As String = ""

                    ' die aktuelle Variante aus der AlleProjekte rausnehmen 
                    key = calcProjektKey(hproj)
                    AlleProjekte.Remove(key)

                    ' das bisherige Standard Projekt aus der AlleProjekte rausnehmen 
                    key = calcProjektKey(hproj.name, "")
                    AlleProjekte.Remove(key)

                    'jetzt die aktuelle Variante zur Standard Variante machen 
                    hproj.variantName = ""
                    hproj.timeStamp = Date.Now
                    If hproj.Status = ProjektStatus(0) Then
                        hproj.Status = ProjektStatus(1)
                    End If

                    ' die "neue" Standard Variante in AlleProjekte aufnehmen 
                    AlleProjekte.Add(hproj)

                    ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                    ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                    Dim tmpCollection As New Collection
                    Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, hproj.tfZeile, tmpCollection, tmpCollection)

                    ' jetzt müssen noch alle Projekt-Constellationen aktualisiert werden 
                    Call projectConstellations.updateVariantName(hproj.name, oldvName, newvName)

                Else
                    If hproj.variantName = "" Then
                        If awinSettings.englishLanguage Then
                            Call MsgBox("The project " & hproj.name & " is already the base-variant")

                        Else
                            Call MsgBox("Projekt " & hproj.name & " ist bereits die Standard-Variante")
                        End If
                    Else
                        ' ist nicht erlaubt ... 
                        If awinSettings.englishLanguage Then
                            Call MsgBox("The base variant of project " & hproj.name & " is protected" & vbLf & _
                                        "and cannot be replaced by another variant")

                        Else
                            Call MsgBox("Projekt " & hproj.name & " ist in der Standard-Variante geschützt" & vbLf & _
                                        "und kann daher nicht von einer anderen Variante überschrieben werden")
                        End If
                    End If

                End If

            Next i

            If currentConstellationName <> calcLastSessionScenarioName() Then
                currentConstellationName = calcLastSessionScenarioName()
            End If

        End If

    End Sub

    ''' <summary>
    ''' Es werden Projekte, die Varianten haben angezeigt in einem TreeView
    ''' Hier können Varianten ausgewählt werden, die gelöscht werden sollen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2VarianteLoeschen(control As IRibbonControl)

        Call PBBVarianteLoeschen(control)


    End Sub

    ''' <summary>
    ''' ein Formular wird aufgeschaltet zum Hinzufügen von Abbildungs-Regeln unbekannte Begriffe zu bekannten Begriffen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT5editDictionary(control As IRibbonControl)

        Dim editDictionary As New frmEditWoerterbuch


        Call projektTafelInit()

        editDictionary.Show()


    End Sub

    Sub PT5changeTimeSpan(control As IRibbonControl)

        Dim mvTimeSpan As New frmMoveTimeSpan
        'Dim returnValue As DialogResult

        Call projektTafelInit()

        appInstance.EnableEvents = False

        'returnValue = mvTimeSpan.Showdialog
        ' in dieser auskommentierten Variante ist es sehr langsam ... deshalb als modales Fenster
        If showRangeRight <> showRangeLeft Then
            mvTimeSpan.Show()
        Else
            Call MsgBox("bitte zuerst eine Zeitspanne definieren")
        End If


        appInstance.EnableEvents = True


    End Sub

    Sub PTDefineDependencies(control As IRibbonControl)

        Dim defineDependencies As New frmDependencies
        Dim result As DialogResult
        Dim awinSelection As Excel.ShapeRange



        Call projektTafelInit()

        enableOnUpdate = False

        If ShowProjekte.Count > 0 Then

            Try
                awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
            Catch ex As Exception
                awinSelection = Nothing
            End Try

            If Not awinSelection Is Nothing Then

                If awinSelection.Count > 1 Then

                    result = defineDependencies.ShowDialog()
                Else

                    Call MsgBox("Bitte zunächst  mindestens zwei Projekte selektieren!")
                End If
            Else
                Call MsgBox("Bitte zunächst mindestens zwei Projekte selektieren!")
            End If

        Else
            Call MsgBox("Es sind keine Projekte geladen!")
        End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' Ressourcen und Kosten eines Projektes bearbeiten 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Resources(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim pname As String
        Dim awinSelection As Excel.ShapeRange
        Dim hproj As clsProjekt

        Call projektTafelInit()



        ' es wird vbeim Betreten der Tabelle2 nochmal auf False gesetzt ... und insbesondere bei Activate Tabelle1 (!) auf true gesetzt, nicht vorher wieder
        enableOnUpdate = False


        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                    pname = hproj.name
                Catch ex As Exception
                    Call MsgBox(" Fehler in EditProject " & singleShp.Name & " , Modul: Tom2G1Resources")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                If hproj.Status = ProjektStatus(0) Then
                    ' jetzt werden die Daten aus hproj in Edit Ressourcen worksheet geschrieben ... 
                    appInstance.ScreenUpdating = False
                    Call awinStoreProjForEditRess(hproj)
                    Dim oldShpID As Integer = CInt(hproj.shpUID)

                    ' hier wird das non-modale Dialog Fenster aufgerufen 
                    Dim confirmEdit As New frmConfirmEditRess

                    confirmEdit.selectedProject = hproj.name
                    confirmEdit.Show()

                    With CType(appInstance.Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
                        .Activate()
                    End With
                    appInstance.ScreenUpdating = True
                Else
                    Call MsgBox("bitte erst eine Variante anlegen ...")
                End If




            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If


        ' das muss hier de-aktiviert werden, weil durch non-modalen Aufruf des Formulars enableonupdate wieder auf true gesetzt wird 
        ' enableOnUpdate = True



    End Sub
    ''' <summary>
    ''' aktiviert , je nach Modus die entsprechenden Ribbon Controls 
    ''' </summary>
    ''' <param name="modus"></param>
    ''' <remarks></remarks>
    Private Sub enableControls(ByVal modus As Integer)

        If modus = ptModus.graficboard Then
            visboZustaende.projectBoardMode = modus
            Call visboZustaende.clearAuslastungsArray()


        ElseIf modus = ptModus.massEditRessCost Then
            visboZustaende.projectBoardMode = modus

        End If

        Me.ribbon.Invalidate()

    End Sub

    ''' <summary>
    ''' setzt die Menu-Labels im Ribbon, je nachdem auf Englisch oder deutsch
    ''' </summary>
    ''' <param name="control"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function bestimmeLabel(control As IRibbonControl) As String
        Dim tmpLabel As String = "?"
        Select Case control.Id
            Case "PTproj" ' Project
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt"
                Else
                    tmpLabel = "Project"
                End If
            Case "PTMEC" ' Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Charts u. Info"
                Else
                    tmpLabel = "Charts and Info"
                End If

            Case "PTMEC1" ' Rollen und Kosten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Rollen u. Kosten"
                Else
                    tmpLabel = "Roles and Cost"
                End If

            Case "PTMEC2" ' Budget/Kosten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Budget/Kosten"
                Else
                    tmpLabel = "Budget/Cost"
                End If

            Case "PTMEC3" ' Formular Forecast Gegenüberstellung 
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Gewinn/Verlust"
                Else
                    tmpLabel = "Profit/Loss"
                End If
            Case "PTX" ' Multiprojekt-Info
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Multiprojekt-Info"
                Else
                    tmpLabel = "Multiproject-Info"
                End If

            Case "PTXG1" ' Analysen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Analysen"
                Else
                    tmpLabel = "Analyzes"
                End If

            Case "PT3G1B1" 'Ampeln
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ampeln"
                Else
                    tmpLabel = "Traffic Lights"
                End If

            Case "PTXG1B2" ' Meilenstein-Ampeln
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Meilenstein-Ampeln"
                Else
                    tmpLabel = "Milestone Trafficlights"
                End If

            Case "PT3G1M1" ' Planelemente visualisieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Phasen/Meilensteine"
                Else
                    tmpLabel = "Phases/Milestones"
                End If

            Case "PTXG1B4" ' Auswahl über Namen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Namen"
                Else
                    tmpLabel = "Select by Names"
                End If

            Case "PTXG1B5" ' Auswahl über Projekt-Struktur
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Projekt-Struktur"
                Else
                    tmpLabel = "Select by Structure"
                End If

            Case "PTOPTB1" ' Optimieren hier wir onAction gemacht
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Optimieren"
                Else
                    tmpLabel = "Optimization"
                End If

            Case "PTPf" ' Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Charts"
                Else
                    tmpLabel = "Charts"
                End If

            Case "PTXG1M2" ' Engpass Analyse
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Rollen/Kosten/Meilensteine/Phasen"
                Else
                    tmpLabel = "Ressources/Costs/Milestones/Phases"
                End If

            Case "PTXG1B6" ' Auswahl über Namen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Namen"
                Else
                    tmpLabel = "Select by Names"
                End If

            Case "PTXG1B7" ' Auswahl über Hierarchie
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Projekt-Struktur"
                Else
                    tmpLabel = "Select by Structure"
                End If

            Case "PTXG1B10" ' größter Engpass
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "größter Engpass"
                Else
                    tmpLabel = "Worst bottleneck"
                End If

            Case "PTXG1B3" ' Auslastung
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auslastung"
                Else
                    tmpLabel = "Capacity Utilization"
                End If

            Case "PTXG1B8" ' Strategie/Risiko/Marge
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Strategie / Risiko"
                Else
                    tmpLabel = "Strategy / Risk"
                End If

            Case "PTFKB1" ' Budget/Kosten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Gewinn / Verlust"
                Else
                    tmpLabel = "Profit / Loss"
                End If

            Case "PTFKB2" ' Ergebnis/Auslastung
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ergebnis/Auslastung"
                Else
                    tmpLabel = "Profit/Utilization"
                End If

            Case "PT0" ' Einzelprojekt-Info
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Einzelprojekt-Info"
                Else
                    tmpLabel = "Project-Info"
                End If

            Case "PT0G1" ' Analysen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Analysen"
                Else
                    tmpLabel = "Analyzes"
                End If

            Case "PT0G1B0" ' Projekt-Ampel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt-Ampel"
                Else
                    tmpLabel = "Trafficlight"
                End If

            Case "PT0G1B2" ' Meilenstein-Ampel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Meilenstein-Ampel"
                Else
                    tmpLabel = "Milestone Trafficlights"
                End If

            Case "PT0G1M0" ' Planelemente visualisieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Phasen/Meilensteine visualisieren"
                Else
                    tmpLabel = "Visualize Phases/Milestones"
                End If

            Case "PT0G1B8" ' Auswahl über Namen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt Filter"
                Else
                    tmpLabel = "Project Filter"
                End If

            Case "PT0G1B9" ' Auswahl über Projekt-Struktur
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Projekt-Struktur"
                Else
                    tmpLabel = "Select by Structure"
                End If

            Case "PT3G1B5" ' Zeit-Maschine
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Zeit-Maschine"
                Else
                    tmpLabel = "Time Machine"
                End If

            Case "PT7G1" ' Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Charts"
                Else
                    tmpLabel = "Charts"
                End If

            Case "PT0G1M0B2" ' Phasen Übersicht
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Phasen Übersicht"
                Else
                    tmpLabel = "Overview Phases"
                End If

            Case "PT0G1M2B7" ' Meilenstein Trend Analyse
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Meilenstein Trend Analyse"
                Else
                    tmpLabel = "Milestone Trend Analysis"
                End If

            Case "PT0G1M1B1" ' Personal Bedarfe
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Personal Bedarfe"
                Else
                    tmpLabel = "Ressource Needs"
                End If

            Case "PT0G1M1B2" ' Kosten Bedarfe
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Kosten Bedarfe"
                Else
                    tmpLabel = "Cost-Type Needs"
                End If

            Case "PT0G1M1B3" ' Formular Forecast Gegenüberstellung 
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Gewinn/Verlust"
                Else
                    tmpLabel = "Profit/Loss"
                End If
            Case "PT0G1B3" ' Strategie/Risiko/Marge
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Strategie/Risiko"
                Else
                    tmpLabel = "Strategy/Risk"
                End If

            Case "PT0G1B4" ' Strategie/Risiko/Abhängigkeiten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Strategie/Risiko/Abhängigkeiten"
                Else
                    tmpLabel = "Strategy/Risk/Dependencies"
                End If
            Case "PT7G1M0" ' Add Portfolio Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio-Charts"
                Else
                    tmpLabel = "Add Portfolio Charts"
                End If
            Case "PT7G1M1" ' Add Project Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt-Charts"
                Else
                    tmpLabel = "Add Project Charts"
                End If
            Case "PT7G1M2" ' Plan/Aktuell
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Plan versus Aktuell"
                Else
                    tmpLabel = "Plan versus Actual"
                End If

            Case "PT7G1M2B1" ' Personalkosten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Personalkosten"
                Else
                    tmpLabel = "Personnel Cost"
                End If

            Case "PT7G1M2B2" ' Sonstige Kosten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Sonstige Kosten"
                Else
                    tmpLabel = "Other Cost"
                End If

            Case "PT7G1M2B3" ' Gesamt Kosten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Gesamt Kosten"
                Else
                    tmpLabel = "Total Cost"
                End If

            Case "PT0G1B7" ' Ergebnis
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ergebnis"
                Else
                    tmpLabel = "Profit"
                End If

            Case "PT7" ' Cockpit
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Cockpit"
                Else
                    tmpLabel = "Cockpit"
                End If

            Case "PT0G1M3B1" ' Cockpit visualisieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Cockpit laden"
                Else
                    tmpLabel = "Load Cockpit"
                End If

            Case "PT0G1M3B2" ' Cockpit speichern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Chart-Set sichern als Cockpit"
                Else
                    tmpLabel = "Save current Chart-Set as Cockpit"
                End If

            Case "PT0G1M3B3" ' Cockpit-Charts löschen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Charts löschen"
                Else
                    tmpLabel = "Delete Charts"
                End If

            Case "PT1" ' Reports
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Reports"
                Else
                    tmpLabel = "Reports"
                End If

            Case "PT1G1" ' Powerpoint Report Profile
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Powerpoint Report Profile"
                Else
                    tmpLabel = "Powerpoint Report Profiles"
                End If

            Case "PT1G1M0" ' Report-Profil definieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Report definieren"
                Else
                    tmpLabel = "Define a Report"
                End If

            Case "PT1G1M01" ' Einzelprojekt-Berichte
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt Report"
                Else
                    tmpLabel = "Project Report"
                End If

            Case "PT1G1M01B0" ' Typ I
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "ohne Element-Auswahl"
                Else
                    tmpLabel = "without element-selection"
                End If

            Case "PT1G1M1" ' Typ II
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "mit Element-Auswahl"
                Else
                    tmpLabel = "with element-selection"
                End If

            Case "PT1G1M1B1" ' Auswahl über Namen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Namen"
                Else
                    tmpLabel = "Select by Names"
                End If

            Case "PT1G1M1B2" ' Auswahl über Hierarchie
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Projekt-Struktur"
                Else
                    tmpLabel = "Select by Structure"
                End If

            Case "PT1G1M02" ' Multiprojekt-Berichte
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio Report"
                Else
                    tmpLabel = "Portfolio Report"
                End If

            Case "PT1G1B2" ' Typ I
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "ohne Element-Auswahl"
                Else
                    tmpLabel = "without element-selection"
                End If

            Case "PT1G1M2" ' Typ II
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "mit Element-Auswahl"
                Else
                    tmpLabel = "with element-selection"
                End If

            Case "PT1G1M2B1" ' Auswahl über Namen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Namen"
                Else
                    tmpLabel = "Select by Names"
                End If

            Case "PT1G1M2B2" ' Auswahl über Hierarchie
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Projekt-Struktur"
                Else
                    tmpLabel = "Select by Structure"
                End If

            Case "PT1G1B4" ' letztes Report-Profil speichern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "letzte Report Definition als Vorlage speichern"
                Else
                    tmpLabel = "Store last Report definition as pre-defined"
                End If

            Case "PT1G1B5" ' Report-Profil ausführen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "vordefinierten Report erstellen"
                Else
                    tmpLabel = "Select pre-defined Report"
                End If

            Case "PT1G1B1" ' Report Sprache
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Report Sprache"
                Else
                    tmpLabel = "Report Language"
                End If

            Case "PT1G1B6" ' Report Generator Template erstellen Sprache
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Batch-File für Report Erstellung erzeugen"
                Else
                    tmpLabel = "Create Batch-File for mass Report Creation"
                End If

            Case "PT2" ' Bearbeiten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Bearbeiten"
                Else
                    tmpLabel = "Edit"
                End If

            Case "PT2G1" ' Einzelprojekte
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Einzelprojekte"
                Else
                    tmpLabel = "Singleprojects"
                End If

            Case "PT2G1M0" ' Neues Projekt anlegen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Neues Projekt"
                Else
                    tmpLabel = "New Project"
                End If

            Case "PT2G1M0B0" ' Neu auf Basis Vorlage
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Neu auf Basis Vorlage"
                Else
                    tmpLabel = "Based on template"
                End If

            Case "PT2G1M0B1" ' Neu auf Basis Projekt
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Neu auf Basis Projekt"
                Else
                    tmpLabel = "Based on project"
                End If

            Case "PT2G1M1" ' Variante
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Varianten"
                Else
                    tmpLabel = "Variants"
                End If

            Case "PT2G1M1B0" ' neue Variante anlegen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Neue Variante"
                Else
                    tmpLabel = "New Variant"
                End If

            Case "PT2G1M1B1" ' Variante aktivieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Variante aktivieren"
                Else
                    tmpLabel = "Activate Variant"
                End If

            Case "PT2G1M1B2" ' Variante löschen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Variante löschen"
                Else
                    tmpLabel = "Delete Variant"
                End If

            Case "PT2G1M1B3" ' Variante zum Standard machen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Variante zum Standard machen"
                Else
                    tmpLabel = "Set Variant as base-variant"
                End If

            Case "PT2G1M2B4" ' Ressource/Kostenart hinzufügen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Rolle hinzufügen"
                Else
                    tmpLabel = "Add Resource"
                End If
            Case "PT2G1M2B7" ' Ressource/Kostenart hinzufügen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Kostenart hinzufügen"
                Else
                    tmpLabel = "Add Cost"
                End If

            Case "PT2G1M2B5" ' Ressource/Kostenart löschen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Rolle/Kostenart löschen"
                Else
                    tmpLabel = "Delete resource / cost"
                End If

            Case "PTmassEdit" 'Editieren im MassEdit
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Edit"
                Else
                    tmpLabel = "Edit"
                End If

            Case "PT2G1M2B1" ' Ressourcen und Kosten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ändern von Ressourcen und Kosten"
                Else
                    tmpLabel = "Modify monthly Resource and Cost Needs"
                End If

            Case "PT2G1M2B2" ' Strategie/Risiko/Budget
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Strategie/Risiko/Budget"
                Else
                    tmpLabel = "Strategy/Risk/Budget"
                End If

            Case "PT2G1M2B3" ' Zeitspanne f. Projektstart
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Zeitspanne f. Projektstart"
                Else
                    tmpLabel = "Timespan for projectstart"
                End If

            Case "PTMECsettings" ' Einstellungen beim Editieren Ressourcen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Einstellungen"
                Else
                    tmpLabel = "Settings"
                End If

            Case "PT6G2B3" ' prozentuale Auslastungs-Werte anzeigen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Prozentuale Auslastungs-Werte anzeigen"
                Else
                    tmpLabel = "Show percentil values"
                End If

            Case "PT6G2B4" ' Platzhalter Rollen automatisch reduzieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Platzhalter Rollen automatisch reduzieren"
                Else
                    tmpLabel = "Automatically reduce placeholder values"
                End If

            Case "PT6G2B5" ' Sortierung ermöglichen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Sortierung ermöglichen"
                Else
                    tmpLabel = "Enable sorting"
                End If

            Case "PTfreezeB1" ' Fixieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Fixieren"
                Else
                    tmpLabel = "Freeze"
                End If

            Case "PTfreezeB2" ' Fixierung aufheben
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Fixierung zum Bewegen aufheben"
                Else
                    tmpLabel = "De-Freeze for moving"
                End If

            Case "PTzurück" ' zurück zur Multiprojekt-Tafel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Zurück zu Projekt Tafel"
                Else
                    tmpLabel = "Back to Project Board"
                End If

            Case "PTback" ' OnAction zur Multiprojekt-Tafel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Zurück"
                Else
                    tmpLabel = "Back"
                End If

            Case "PT2G1M2B6" ' Änderungen verwerfen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Änderungen verwerfen"
                Else
                    tmpLabel = "Skip changes"
                End If

            Case "PT3G1M2" ' Beschriften
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Beschriftungen"
                Else
                    tmpLabel = "Annotations"
                End If

            Case "PT2G1B4" ' Beschriften ON
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ein"
                Else
                    tmpLabel = "ON"
                End If

            Case "PT2G1B5" ' Beschriftungen löschen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Aus"
                Else
                    tmpLabel = "OFF"
                End If

            Case "PT2G1B6" ' Extended View
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Extended Ansicht"
                Else
                    tmpLabel = "Expanded View"
                End If

            Case "PT2G1B7" ' Rollup View
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Linien Ansicht"
                Else
                    tmpLabel = "Collapsed View"
                End If

            Case "PT2G1B1" ' Umbenennen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt umbenennen"
                Else
                    tmpLabel = "Rename Project"
                End If

            Case "PT2G1B3" ' Umbenennen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Variante umbenennen"
                Else
                    tmpLabel = "Rename Variant"
                End If

            Case "PT2G2" 'Projekte/Varianten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekte/Varianten"
                Else
                    tmpLabel = "Projects/Variants"
                End If

            Case "PT2G2B2" ' Portfolio/s anzeigen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio/s anzeigen"
                Else
                    tmpLabel = "Show Portfolio/s"
                End If


            Case "PT2G2B4" ' Editieren Portfolio
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio"
                Else
                    tmpLabel = "Portfolio"
                End If


            Case "PT2G3" ' Session
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Session"
                Else
                    tmpLabel = "Session"
                End If

            Case "PT2G3M1" ' aus der Session löschen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Löschen aus Session"
                Else
                    tmpLabel = "Delete in Session"
                End If

            Case "PT2G3M1B1" ' alles
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Alles"
                Else
                    tmpLabel = "Whole Session"
                End If

            Case "PT2G3M1B2" ' einzelne Projekte und Varianten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt/e"
                Else
                    tmpLabel = "Project/s"
                End If

            Case "PT2G3M1B3" ' Szenario
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio/s"
                Else
                    tmpLabel = "Portfolio/s"
                End If

            Case "PT2G3M2" ' Zeichen-Elemente löschen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Zeichen-Elemente löschen"
                Else
                    tmpLabel = "Delete Shapes"
                End If

            Case "PT2G3M4B1" ' Meilensteine, Phasen, Stati
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Meilensteine, Phasen, Stati"
                Else
                    tmpLabel = "Milestones, Phases, Stati"
                End If

            Case "PT2G3M4B3" ' Beschriftungen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Beschriftungen"
                Else
                    tmpLabel = "Annotations"
                End If

            Case "PT2G3M4B2" ' Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Charts"
                Else
                    tmpLabel = "Charts"
                End If
            Case "PT0G1BX" ' Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Charts löschen"
                Else
                    tmpLabel = "Delete Charts"
                End If

            Case "PT4" ' Datenmanagement
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Datenmanagement"
                Else
                    tmpLabel = "Data management"
                End If

            Case "PTExit" 'Beenden
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Beenden"
                Else
                    tmpLabel = "Exit"
                End If

            Case "PT4G1" ' IMPORT
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Import"
                Else
                    tmpLabel = "Import from Filesystem"
                End If

            Case "PT4G1B6" ' VISBO Open XML
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "VISBO Open XML"
                Else
                    tmpLabel = "VISBO Open XML"
                End If

            Case "PT4G1B1" ' Import VISBO-Steckbriefe
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "VISBO-Steckbriefe"
                Else
                    tmpLabel = "VISBO project briefs"
                End If

            Case "PT4G1B2" ' Import Excel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Excel"
                Else
                    tmpLabel = "Excel"
                End If

            Case "PT4G1B4" ' Import MS Project
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "MS Project"
                Else
                    tmpLabel = "MS Project"
                End If

            Case "PT4G1B3" ' Import RPLAN RXF
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "RPLAN RXF"
                Else
                    tmpLabel = "RPLAN RXF"
                End If

            Case "PT4G1B7" ' Import Projekte (Batch)
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Batch Projekt-Erzeugung"
                Else
                    tmpLabel = "Batch Project Creation"
                End If

            Case "PT4G1B5" ' Import Scenario Definition
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio Liste"
                Else
                    tmpLabel = "Portfolio List"
                End If

            Case "PT4G2" ' EXPORT
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Export"
                Else
                    tmpLabel = "Export to Filesystem"
                End If

            Case "PT4G1M1B3" ' Export VISBO-Steckbriefe
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "VISBO-Steckbriefe"
                Else
                    tmpLabel = "VISBO project briefs"
                End If

            Case "PT4G1M1B2" 'Export Visbo OPen Xml
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "VISBO Open XML"
                Else
                    tmpLabel = "VISBO Open XML"
                End If
            Case "PT4G1M1B1" 'Export Excel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Excel"
                Else
                    tmpLabel = "Excel"
                End If

            Case "PT4G1M0B2" ' exportieren von Meilensteine und Phasen nach Auswahl
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl von Meilensteine und Phasen"
                Else
                    tmpLabel = "Selection of Milestones and Phases"
                End If

            Case "PT4G2M3" ' exportieren von Portfolio Liste
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio Liste"
                Else
                    tmpLabel = "Portfolio List"
                End If


            Case "PT4G1B7" ' Export FC-52
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Export FC-52"
                Else
                    tmpLabel = "Export FC-52"
                End If

            Case "PT4G1M3" ' Export Excel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Excel"
                Else
                    tmpLabel = "Excel"
                End If

                'Case "PT4G2M3" ' Datenbank
                '    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '        tmpLabel = "Datenbank"
                '    Else
                '        tmpLabel = "Database"
                '    End If

            Case "PT5G1" ' Load from Database

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Laden von DB "
                Else
                    tmpLabel = "Load from Database"
                End If

            Case "PT5G1B1" ' Portfolio/s
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio/s"
                Else
                    tmpLabel = "Portfolio/s"
                End If

            Case "PT5G1B3" ' Project/s
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt/e"
                Else
                    tmpLabel = "Project/s"
                End If

            Case "PT5G2" ' Speichern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Speichern in DB"
                Else
                    tmpLabel = "Store to Database"
                End If

            Case "Pt5G2B1" ' Portfolio/s
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio/s"
                Else
                    tmpLabel = "Portfolio/s"
                End If

            Case "Pt5G2B3" ' Projekt/e
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt/e"
                Else
                    tmpLabel = "Project/s"
                End If

            Case "Pt5G2B4" ' Alles Speichern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Alles Speichern"
                Else
                    tmpLabel = "Store everything"
                End If

            Case "PT5G3" ' Löschen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Löschen aus DB"
                Else
                    tmpLabel = "Delete in Database"
                End If

            Case "Pt5G3B1" ' Multiprojekt-Szenario
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio/s"
                Else
                    tmpLabel = "Portfolio/s"
                End If

            Case "PT5G3M2" ' Projekte/Varianten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt/e"
                Else
                    tmpLabel = "Project/s"
                End If

            Case "Pt5G3B3" ' Projekte/Varianten/TimeStamps auswählen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekte/Varianten/TimeStamps auswählen"
                Else
                    tmpLabel = "Select Projects/Variants/TimeStamps"
                End If

            Case "Pt5G3B4" ' X Versionen behalten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "X Versionen behalten"
                Else
                    tmpLabel = "Keep X Versions"
                End If


            Case "PT2G2B5" ' Sperre setzen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Schreibschutz setzen/aufheben"
                Else
                    tmpLabel = "Set/Unset Write-Protection"
                End If

            Case "PTedit"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Edit"
                Else
                    tmpLabel = "Edit"
                End If
            Case "PTview"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ansicht"
                Else
                    tmpLabel = "View"
                End If
            Case "PTfilter"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Filter"
                Else
                    tmpLabel = "Filter"
                End If
            Case "PTsort"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Sortieren"
                Else
                    tmpLabel = "Sort"
                End If
            Case "PTcharts"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Charts"
                Else
                    tmpLabel = "Charts"
                End If
            Case "PTreport"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Reports"
                Else
                    tmpLabel = "Reports"
                End If
            Case "PTeinst"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Einstellungen"
                Else
                    tmpLabel = "Settings"
                End If
            Case "PThelp"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Help"
                Else
                    tmpLabel = "Help"
                End If
            Case "PTlizenz"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Lizenzen"
                Else
                    tmpLabel = "Licenses"
                End If
            Case "PT6" ' Einstellungen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Einstellungen"
                Else
                    tmpLabel = "Settings"
                End If

            Case "PT6G1" ' Visualisierung
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Visualisierung"
                Else
                    tmpLabel = "Visualization"
                End If

            Case "PT6G1B4" ' Anzeigen selekt. Objekte im Summenchart
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Anzeigen selekt. Objekte im Summenchart"
                Else
                    tmpLabel = "Show selected objects in charts"
                End If

            Case "PT6G1B5" ' Ampeln anzeigen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ampeln anzeigen"
                Else
                    tmpLabel = "Show trafficlights"
                End If

            Case "PT6G2" ' Berechnung
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Berechnung"
                Else
                    tmpLabel = "Calculation"
                End If

            Case "PT6G2B1" ' Bei Dehnen/Stauchen proportional anpassen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Bei Dehnen/Stauchen proportional anpassen"
                Else
                    tmpLabel = "Adjust resource/cost when expanding/shortening"
                End If

            Case "PT6G2B2" ' Phasenhäufigkeit anteilig berechnen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Phasenhäufigkeit anteilig berechnen"
                Else
                    tmpLabel = "Calculate phase-frequency proportionally"
                End If

            Case "PT6G3" ' Diverses
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Diverses"
                Else
                    tmpLabel = "Miscellaneous"
                End If

            Case "Pt6G3B3" ' Zeitraum verschieben
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Zeitraum verschieben"
                Else
                    tmpLabel = "Move Timespan"
                End If

            Case "Pt6G3B4" ' Wörterbuch editieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Wörterbuch editieren"
                Else
                    tmpLabel = "Edit Synonyms"
                End If

            Case Else
                tmpLabel = "undefined"
        End Select

        bestimmeLabel = tmpLabel

    End Function
    ''' <summary>
    ''' Setzt die Sichtbarkeit der Menu-Buttons je nach visboZustaende.projectBoardMode
    ''' </summary>
    ''' <param name="control">Control des Menubuttons, von dem diese Routine über "getVisible" im Ribbon aufgerufen wird</param>
    ''' <returns>true: wenn der entsprechende Menubutton sichtbar sein soll
    '''          false: wenn der entsprechende Menubutton unsichtbar sein soll </returns>
    ''' <remarks></remarks>
    Function chckVisibility(control As IRibbonControl) As Boolean
        If visboZustaende.projectBoardMode = ptModus.graficboard Then
            Select Case control.Id
                Case "PTMEC" ' Massen-Edit Charts
                    chckVisibility = False
                Case "PTmassEdit" ' Mass-Edit bearbeiten
                    chckVisibility = False
                Case "PT2G1M2B4" ' Bearbeiten - Zeile (Rolle) einfügen
                    chckVisibility = False
                Case "PT2G1M2B5" ' Bearbeiten - Zeile löschen
                    chckVisibility = False
                Case "PT2G1M2B6" ' Bearbeiten - Änderungen verwerfen
                    chckVisibility = False
                Case "PT2G1M2B7" ' Bearbeiten - Zeile (Kostenart) einfügen
                    chckVisibility = False
                Case "PTzurück" ' Zurück
                    chckVisibility = False
                Case "PTMECsettings" ' Massen-Edit Einstellungen/Settings
                    chckVisibility = False
                Case "PT6G2B3" ' Einstellungen - Berechnung - prozentuale Auslastungs-Werte anzeigen
                    chckVisibility = False
                Case "PT6G2B4" ' Platzhalter Rollen automatisch reduzieren
                    chckVisibility = False
                Case "PT6G2B5" ' Sortierung ermöglichen
                    chckVisibility = False
                Case Else
                    ' alle anderen werden sichtbar gemacht
                    chckVisibility = True
            End Select
        Else
            Select Case control.Id

                Case "PTproj"
                    chckVisibility = False
                Case "PTedit"
                    chckVisibility = False
                Case "PTview"
                    chckVisibility = False
                Case "PTfilter"
                    chckVisibility = False
                Case "PTsort"
                    chckVisibility = False
                Case "PToptimize"
                    chckVisibility = False
                Case "PTcharts"
                    chckVisibility = False
                Case "PTreport"
                    chckVisibility = False
                Case "PTeinst"
                    chckVisibility = False
                Case "PThelp"
                    chckVisibility = False
                Case "PTlizenz"
                    chckVisibility = False

                Case "PT2G1M2B6" ' Mass-Edit Änderungen verwerfen
                    chckVisibility = False


                    '' ''Case "PTX" ' Multiprojekt-Info
                    '' ''    chckVisibility = False
                    '' ''Case "PT0" ' Einzelprojekt-Info
                    '' ''    chckVisibility = False
                    '' ''Case "PT7" ' Cockpit
                    '' ''    chckVisibility = False
                    '' ''Case "PT1" ' Reports
                    '' ''    chckVisibility = False
                    '' ''Case "PT6" ' Einstellungen
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1M0" ' neues Projekt anlegen
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1M1" ' Variante
                    '' ''    chckVisibility = False
                    '' ''    'Case "PT2G1M1B0" ' neue Variante anlegen
                    '' ''    '    chckVisibility = False
                    '' ''    'Case "PT2G1M1B1" '  Variante aktivieren
                    '' ''    '    chckVisibility = False
                    '' ''    'Case "PT2G1M1B2" ' Variante löschen    
                    '' ''    '    chckVisibility = False
                    '' ''    'Case "PT2G1M1B3" ' Variante übernehmen    
                    '' ''    '    chckVisibility = False
                    '' ''Case "PT2G1M2" ' Editieren   
                    '' ''    chckVisibility = False
                    '' ''    'Case "PT2G1M2B2" ' Strategie/Risiko/Budget   
                    '' ''    '    chckVisibility = False
                    '' ''    'Case "PT2G1M2B3" ' Zeitspanne f. Projektstart   
                    '' ''    '    chckVisibility = False
                    '' ''Case "PT2G1B2" ' Fixieren
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1B3" ' Fixierung aufheben
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1M2B6" ' Änderungen verwerfen
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1B4" ' Beschriften
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1B5" ' alle Beschriftungen löschen
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1B6" ' Extended View
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1B7" ' Extended View aufheben
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G2" ' Bearbeiten - Multiprojekt-Szenario
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G2B5" ' Schutz setzen / aufheben 
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G2s3" ' Separator vor Schutz 
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1s4" ' Separator vor Extended View
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G3" ' Bearbeiten - Session
                    '' ''    chckVisibility = False
                    '' ''Case "PT4" ' Datenmanagement
                    '' ''    chckVisibility = False
                    '' ''Case "PT6G1" ' Einstellungen - Visualisierung
                    '' ''    chckVisibility = False
                    '' ''Case "PT6G2B1" ' Einstellungen - Berechnung - Dehnen/Stauchen
                    '' ''    chckVisibility = False
                    '' ''Case "PT6G2B2" ' Phasenhäufigkeit anteilig berechnen
                    '' ''    chckVisibility = False
                    '' ''Case "PT6G3" ' Lade- und Import-Vorgänge
                    '' ''    chckVisibility = False
                    '' ''Case "PT2G1B8" ' umbenennen 
                    '' ''    chckVisibility = False
                    '' ''Case "PT7G1M1" ' Projekt-Charts
                    '' ''    chckVisibility = False
                    '' ''Case "PT7G1M0" ' Portfolio-Charts
                    '' ''    chckVisibility = False
                Case Else
                    chckVisibility = True
            End Select

        End If
    End Function
    Sub Tom2G2MassEdit(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange
        Dim todoListe As New Collection
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        ' zurücksetzen der dbCache-Projekte
        dbCacheProjekte.Clear(False)

        Call projektTafelInit()

        enableOnUpdate = False

        If ShowProjekte.Count > 0 Then

            If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                ' alles ok , bereits gesetzt 

            Else
                If selectedProjekte.Count > 0 Then
                    showRangeLeft = selectedProjekte.getMinMonthColumn
                    showRangeRight = selectedProjekte.getMaxMonthColumn
                Else
                    showRangeLeft = ShowProjekte.getMinMonthColumn
                    showRangeRight = ShowProjekte.getMaxMonthColumn
                End If
                Call awinShowtimezone(showRangeLeft, showRangeRight, True)
            End If

            Try
                awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
            Catch ex As Exception
                awinSelection = Nothing
            End Try

            If IsNothing(awinSelection) Then

                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    todoListe.Add(kvp.Key, kvp.Key)

                    If Not noDB Then

                        ' prüfen, ob es überhaupt schon in der Datenbank existiert ...
                        If request.projectNameAlreadyExists(kvp.Value.name, kvp.Value.variantName, Date.Now) Then
                            Dim dbProj As clsProjekt = request.retrieveOneProjectfromDB(kvp.Value.name, kvp.Value.variantName, Date.Now)
                            dbCacheProjekte.upsert(dbProj)
                        End If

                    End If

                Next

            Else

                For i As Integer = 1 To awinSelection.Count
                    singleShp = awinSelection.Item(i)
                    Dim hproj As clsProjekt
                    Try
                        hproj = ShowProjekte.getProject(singleShp.Name, True)
                        todoListe.Add(hproj.name, hproj.name)

                        If Not noDB Then
                            ' wenn es in der DB existiert, dann im Cache aufbauen 
                            If request.projectNameAlreadyExists(hproj.name, hproj.variantName, Date.Now) Then
                                ' für den Datenbank Cache aufbauen 
                                Dim dbProj As clsProjekt = request.retrieveOneProjectfromDB(hproj.name, hproj.variantName, Date.Now)
                                dbCacheProjekte.upsert(dbProj)
                            End If
                        End If

                    Catch ex As Exception

                    End Try
                Next
            End If

            ' wenn es jetzt etwas zu tun gibt ... 
            If todoListe.Count > 0 Then
                ' alles ok ...
                ' wenn die Charts da bleiben, kann das zu Fehlern führen ... 
                'appInstance.ScreenUpdating = False
                'Call awinStoreCockpit("_Last")
                Call deleteChartsInSheet(arrWsNames(ptTables.MPT))

                Call enableControls(ptModus.massEditRessCost)

                ' hier sollen jetzt die Projekte der todoListe in den Backup Speicher kopiert werden , um 
                ' darauf zugreifen zu können, wenn beim Massen-Edit die Option alle Änderungen verwerfen gewählt wird. 
                'Call saveProjectsToBackup(todoListe)

                Try
                    enableOnUpdate = False
                    Call writeOnlineMassEditRessCost(todoListe, showRangeLeft, showRangeRight)
                    appInstance.EnableEvents = True

                    With CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
                        .Activate()
                    End With

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
            End If



        Else
            enableOnUpdate = True
            If appInstance.EnableEvents = False Then
                appInstance.EnableEvents = True
            End If

            If awinSettings.englishLanguage Then
                Call MsgBox("no projects in session!")
            Else
                Call MsgBox("Es sind keine Projekte geladen!")
            End If

        End If


        'Call MsgBox("ok, zurück ...")

        ' das läuft neben dem Activate Befehl, deshalb soll das hier auskommentiert werden ... 
        'enableOnUpdate = True
        'appInstance.EnableEvents = True

    End Sub

    Sub PTbackToProjectBoard(control As IRibbonControl)


        ' jetzt muss gecheckt werden, welche dbCache Projekte immer noch identisch zum ShowProjekte Pendant sind
        ' deren temp Schutz muss dann wieder aufgehoben werden ... 
        For Each kvp As KeyValuePair(Of String, clsProjekt) In dbCacheProjekte.liste

            If ShowProjekte.contains(kvp.Value.name) Then
                Dim hproj As clsProjekt = ShowProjekte.getProject(kvp.Value.name)
                Dim pvName As String = calcProjektKey(hproj.name, hproj.variantName)

                If hproj.isIdenticalTo(kvp.Value) Then
                    ' temp Schutz aufheben 
                    If writeProtections.isProtected(pvName, dbUsername) Then
                        ' nichts tun , es ist von jdn anderem geschützt 
                        '
                    ElseIf writeProtections.isPermanentProtected(pvName) Then
                        ' nichts tun, es ist permanent protected 
                        '
                    Else
                        ' den temporären Schutz von mir zurücknehmen 
                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, _
                                                   dbUsername, dbPasswort)
                        Dim wpItem As New clsWriteProtectionItem(pvName, ptWriteProtectionType.project, _
                                                                  dbUsername, False, False)
                        If request.setWriteProtection(wpItem) Then
                            ' erfolgreich
                            writeProtections.upsert(wpItem)
                        Else
                            ' nicht erfolgreich
                            wpItem = request.getWriteProtection(hproj.name, hproj.variantName)
                            writeProtections.upsert(wpItem)
                        End If
                    End If
                Else
                    ' temporär geschützt lassen ...
                End If
            End If
        Next

        ' zurücksetzen 
        dbCacheProjekte.Clear(False)

        Call projektTafelInit()

        If tempSkipChanges Then
            'Call restoreProjectsFromBackup()
            Call MsgBox("restored ...")
            tempSkipChanges = False
        End If

        Call enableControls(ptModus.graficboard)

        appInstance.EnableEvents = False
        ' wird ohnehin zu Beginn des MassenEdits ausgeschaltet  
        'enableOnUpdate = False

        appInstance.ScreenUpdating = False
        Call deleteChartsInSheet(arrWsNames(ptTables.meRC))

        ' das eigentliche, ursprüngliche Windows wird wieder angezeigt ...
        With projectboardWindows(0)
            .Activate()
            .Visible = True
            .WindowState = Excel.XlWindowState.xlMaximized
        End With

        ' jetzt werden die Windows gelöscht, falls sie überhaupt existieren  ...
        If Not IsNothing(projectboardWindows(1)) Then
            projectboardWindows(1).Close()
            projectboardWindows(1) = Nothing
        End If

        If Not IsNothing(projectboardWindows(2)) Then
            projectboardWindows(2).Close()
            projectboardWindows(2) = Nothing
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

        ' tk , das Dalassen der Charts kann zu Fehlern führen ... 

        'Call awinLoadCockpit("_Last")
        'appInstance.ScreenUpdating = True
        ' der ScreenUpdating wird im Tabelle1.Activate gesetzt, falls auf False
        With CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)
            .Activate()
        End With

        appInstance.ScreenUpdating = True

    End Sub

    ''' <summary>
    ''' fügt im MassenEdit Sheet eine Zeile ein, macht aber sonst noch nichts, es werden also noch keinerlei Änderungen am 
    ''' betroffenen Projekt vorgenommen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTzeileEinfuegen(control As IRibbonControl)

        Dim currentCell As Excel.Range
        appInstance.EnableEvents = False

        Try

            ' jetzt werden die Validation-Strings für alles, alleRollen, alleKosten und die einzelnen SammelRollen aufgebaut 
            Dim validationStrings As SortedList(Of String, String) = createMassEditRcValidations()

            currentCell = CType(appInstance.ActiveCell, Excel.Range)

            'Dim columnEndData As Integer = CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Range("EndData"), Excel.Range).Column

            Dim columnEndData As Integer = visboZustaende.meColED
            Dim columnStartData As Integer = visboZustaende.meColSD

            Dim columnRC As Integer = visboZustaende.meColRC

            Dim hoehe As Double = CDbl(currentCell.Height)
            currentCell.EntireRow.Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown)
            Dim zeile As Integer = currentCell.Row

            ' Blattschutz aufheben ... 
            If Not awinSettings.meEnableSorting Then
                ' es muss der Blattschutz aufgehoben werden, nachher wieder aktiviert werden ...
                With CType(appInstance.ActiveSheet, Excel.Worksheet)
                    .Unprotect(Password:="x")
                End With
            End If



            With CType(appInstance.ActiveSheet, Excel.Worksheet)

                If Not awinSettings.meExtendedColumnsView Then
                    appInstance.ScreenUpdating = False
                    ' einblenden ... 
                    .Range("MahleInfo").EntireColumn.Hidden = False
                End If


                Dim copySource As Excel.Range = CType(.Range(.Cells(zeile, 1), .Cells(zeile, 1).offset(0, columnEndData - 1)), Excel.Range)
                Dim copyDestination As Excel.Range = CType(.Range(.Cells(zeile - 1, 1), .Cells(zeile - 1, 1).offset(0, columnEndData - 1)), Excel.Range)
                copySource.Copy(Destination:=copyDestination)

                CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Rows(zeile - 1), Excel.Range).RowHeight = hoehe

                For c As Integer = columnStartData - 3 To columnEndData + 1
                    CType(.Cells(zeile - 1, c), Excel.Range).Value = Nothing
                Next

                ' jetzt wieder ausblenden ... 
                If Not awinSettings.meExtendedColumnsView Then
                    ' ausblenden ... 
                    .Range("MahleInfo").EntireColumn.Hidden = True
                    appInstance.ScreenUpdating = True
                End If
            End With

            ' jetzt wird auf die Ressourcen-/Kosten-Spalte positioniert 
            CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Cells(zeile - 1, columnRC), Excel.Range).Select()

            With CType(CType(appInstance.ActiveSheet, Excel.Worksheet).Cells(zeile - 1, columnRC), Excel.Range)

                ' jetzt für die Zelle die Validation neu bestimmen, der Blattschutz muss aufgehoben sein ...  

                Try
                    If Not IsNothing(.Validation) Then
                        .Validation.Delete()
                    End If
                    ' jetzt wird die ValidationList aufgebaut
                    ' ist es eine Rolle ? 
                    If control.Id = "PT2G1M2B4" Then
                        ' Rollen
                        .Validation.Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, _
                                               Formula1:=validationStrings.Item("alleRollen"))
                    ElseIf control.Id = "PT2G1M2B7" Then
                        ' Kosten
                        .Validation.Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, _
                                                                       Formula1:=validationStrings.Item("alleKosten"))
                    Else
                        ' undefiniert, darf eigentlich nie vorkommen, aber just in case ...
                        .Validation.Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, _
                                               Formula1:=validationStrings.Item("alles"))
                    End If

                Catch ex As Exception

                End Try

            End With

            ' jetzt wird der Old-Value gesetzt 
            With visboZustaende
                If CStr(CType(appInstance.ActiveCell, Excel.Range).Value) <> "" Then
                    Call MsgBox("Fehler 099 in PTzeileEinfügen")
                End If
                .oldValue = ""
                .meMaxZeile = CType(CType(appInstance.ActiveSheet, Excel.Worksheet).UsedRange, Excel.Range).Rows.Count
            End With


            ' jetzt den Blattschutz wiederherstellen ... 
            If Not awinSettings.meEnableSorting Then
                ' es muss der Blattschutz wieder aktiviert werden ... 
                With CType(appInstance.ActiveSheet, Excel.Worksheet)
                    .Protect(Password:="x", UserInterfaceOnly:=True, _
                             AllowFormattingCells:=True, _
                             AllowInsertingColumns:=False,
                             AllowInsertingRows:=True, _
                             AllowDeletingColumns:=False, _
                             AllowDeletingRows:=True, _
                             AllowSorting:=True, _
                             AllowFiltering:=True)
                    .EnableSelection = Excel.XlEnableSelection.xlUnlockedCells
                    .EnableAutoFilter = True
                End With
            End If

        Catch ex As Exception
            Call MsgBox("Fehler beim Kopieren einer Zeile ...")
        End Try

        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' löscht im MassenEdit Sheet eine Zeile, das heisst, die Rolle bzw. Kostenart wird rausgenommen 
    ''' es bleibt aber pro Projekt-/Phase eine leere Zeile als Platzhalter stehen  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTzeileLoeschen(control As IRibbonControl)

        Dim currentCell As Excel.Range
        Dim meWS As Excel.Worksheet = CType(appInstance.Worksheets(arrWsNames(ptTables.meRC)), Excel.Worksheet)
        appInstance.EnableEvents = False

        Dim ok As Boolean = True

        Try

            currentCell = CType(appInstance.ActiveCell, Excel.Range)
            Dim zeile As Integer = currentCell.Row

            If zeile >= 2 And zeile <= visboZustaende.meMaxZeile Then
                Dim columnEndData As Integer = visboZustaende.meColED
                Dim columnStartData As Integer = visboZustaende.meColSD
                Dim columnRC As Integer = visboZustaende.meColRC


                Dim pName As String = CStr(meWS.Cells(zeile, 2).value)
                Dim vName As String = CStr(meWS.Cells(zeile, 3).value)
                Dim phaseName As String = CStr(meWS.Cells(zeile, 4).value)
                Dim phaseNameID As String = calcHryElemKey(phaseName, False)
                Dim curComment As Excel.Comment = CType(meWS.Cells(zeile, 4), Excel.Range).Comment
                If Not IsNothing(curComment) Then
                    phaseNameID = curComment.Text
                End If

                Dim rcName As String = CStr(meWS.Cells(zeile, columnRC).value)

                ' hier wird die Rolle- bzw. Kostenart aus der Projekt-Phase gelöscht 
                Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                Dim cphase As clsPhase = hproj.getPhaseByID(phaseNameID)

                If IsNothing(rcName) Then
                    ' nichts tun
                ElseIf rcName.Trim.Length = 0 Then
                    ' nichts tun ... 
                ElseIf RoleDefinitions.containsName(rcName) Then
                    ' es handelt sich um eine Rolle
                    ' das darf aber nur gelöscht werden, wenn die Phase komplett im showrangeleft / showrangeright liegt 
                    If phaseWithinTimeFrame(hproj.Start, cphase.relStart, cphase.relEnde, _
                                             showRangeLeft, showRangeRight, True) Then
                        cphase.removeRoleByName(rcName)
                    Else
                        Call MsgBox("die Phase wird nicht vollständig angezeigt - deshalb kann die Rolle " & rcName & vbLf & _
                                    " nicht gelöscht werden ...")
                        ok = False
                    End If

                ElseIf CostDefinitions.containsName(rcName) Then
                    ' es handelt sih um eine Kostenart 
                    If phaseWithinTimeFrame(hproj.Start, cphase.relStart, cphase.relEnde, _
                                             showRangeLeft, showRangeRight, True) Then
                        cphase.removeCostByName(rcName)
                    Else
                        Call MsgBox("die Phase wird nicht vollständig angezeigt - deshalb kann die Kostenart " & rcName & vbLf & _
                                    " nicht gelöscht werden ...")
                        ok = False
                    End If


                End If


                If ok Then
                    ' jetzt wird die Zeile gelöscht, wenn sie nicht die letzte ihrer Art ist
                    ' denn es sollte für weitere Eingaben immer wenigstens ein Projekt-/Phasen-Repräsentant da sein 
                    If noDuplicatesInSheet(pName, phaseNameID, Nothing, zeile) Then
                        ' diese Zeile nicht löschen, soll weiter als Platzhalter für diese Projekt-Phase dienen können 
                        ' aber die Werte müssen alle gelöscht werden 
                        For ix As Integer = columnRC To columnEndData + 1
                            CType(meWS.Cells(zeile, ix), Excel.Range).Value = ""
                        Next
                    Else
                        CType(meWS.Rows(zeile), Excel.Range).Delete()
                    End If

                    ' jetzt wird auf die Ressourcen-/Kosten-Spalte positioniert 
                    CType(meWS.Cells(zeile, columnRC), Excel.Range).Select()

                    ' jetzt wird der Old-Value gesetzt 
                    With visboZustaende
                        .oldValue = CStr(CType(meWS.Cells(zeile, columnRC), Excel.Range).Value)
                        .meMaxZeile = CType(meWS.UsedRange, Excel.Range).Rows.Count
                    End With

                Else
                    ' nichts tun 
                End If


            Else
                Call MsgBox(" es können nur Zeilen aus dem Datenbereich gelöscht werden ...")
            End If

        Catch ex As Exception
            Call MsgBox("Fehler beim Löschen einer Zeile ..." & vbLf & ex.Message)
        End Try

        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' Attribute eines Projektes bearbeiten 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Attribute(control As IRibbonControl)

        Dim ProjektAendern As New frmProjektAendern
        Dim returnValue As DialogResult

        Dim singleShp As Excel.Shape

        Dim awinSelection As Excel.ShapeRange
        Dim hproj As clsProjekt
        Dim databaseName As String = awinSettings.databaseName

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)

                    ' hier prüfen, ob das Projekt bereits in der DB existiert ...
                    ' und ob es von anderen geschützt ist 
                    ' wenn ja, dann soll es temporär geschützt werden 

                    If tryToprotectProjectforMe(hproj.name, hproj.variantName) Then

                        With ProjektAendern
                            .projectName.Text = hproj.name
                            .vorlagenName.Text = hproj.VorlagenName
                            For Each kvp As KeyValuePair(Of Integer, clsBusinessUnit) In businessUnitDefinitions
                                .businessUnit.Items.Add(kvp.Value.name)
                            Next
                            ' jetzt werden die Werte im Fenster vorbesetzt ...
                            .businessUnit.Text = hproj.businessUnit
                            .Erloes.Text = hproj.Erloes.ToString
                            .risiko.Text = hproj.Risiko.ToString("0.0")
                            .sFit.Text = hproj.StrategicFit.ToString("0.0")
                        End With
                        ' Aufruf Dialog Fenster 
                        returnValue = ProjektAendern.ShowDialog

                        If returnValue = DialogResult.OK Then
                            With hproj
                                .timeStamp = Date.Now

                                If .Erloes <> CType(ProjektAendern.Erloes.Text, Double) Then
                                    If .Erloes = 0 Then
                                        .Erloes = CType(ProjektAendern.Erloes.Text, Double)

                                        ' Workaround: 

                                        ' tk, Änderung 19.1.17 nicht mehr notwendig ..
                                        ' Call awinCreateBudgetWerte(hproj)
                                    Else
                                        Try
                                            ' tk 19.1.17, nicht mehr notwendig, gibt jetzt Methode budgetWerte 
                                            'Call awinUpdateBudgetWerte(hproj, CType(ProjektAendern.Erloes.Text, Double))
                                            .Erloes = CType(ProjektAendern.Erloes.Text, Double)
                                        Catch ex As Exception
                                            .Erloes = CType(ProjektAendern.Erloes.Text, Double)
                                            ' Workaround: 
                                            ' Änderung tk, wird nicht mehr benötigt , gibt jetzt Methode budgetWerte 
                                            ' Call awinCreateBudgetWerte(hproj)
                                        End Try

                                    End If
                                End If

                                Dim tmpValue As Integer = hproj.dauerInDays

                                .StrategicFit = CType(ProjektAendern.sFit.Text, Double)
                                .Risiko = CType(ProjektAendern.risiko.Text, Double)
                                .businessUnit = CType(ProjektAendern.businessUnit.Text, String)

                            End With

                            Call awinNeuZeichnenDiagramme(5)
                        End If
                    Else
                        ' das Projekt darf vom Nutzer nicht verändert werden , weil von anderem Nutzer geschützt 
                        If awinSettings.englishLanguage Then
                            Call MsgBox(hproj.name & ", " & hproj.variantName & " is protected " & vbLf & _
                                        "and cannot be modified. You could instead create a variant.")
                        Else
                            Call MsgBox(hproj.name & ", " & hproj.variantName & " ist geschützt " & vbLf & _
                                        "und kann nicht verändert werden. Sie können jedoch eine Variante anlegen.")
                        End If
                    End If

                    
                Catch ex As Exception
                    Call MsgBox(" Fehler in EditProject " & singleShp.Name & " , Modul: Tom2G1Attribute")
                    Exit Sub
                End Try


            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' earliest und latest Start eines Projektes ändern 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1EarliestLatestStart(control As IRibbonControl)

        Dim setStartEnd As New frmEarliestLatestStart

        Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim i As Integer
        Dim hproj As clsProjekt
        Dim singleShp As Excel.Shape
        Dim pname As String
        Dim todoListe As New Collection
        Dim errMessage As String = ""
        Dim initMsg As String = "bitte erst eine Variante anlegen"

        Call projektTafelInit()

        ' es wird vbeim Betreten der Tabelle2 nochmal auf False gesetzt ... und insbesondere bei Activate Tabelle1 (!) auf true gesetzt, nicht vorher wieder
        enableOnUpdate = False

        ' Änderung 2.7.14 tk : Vorbedingung sicherstellen: nur Projekte, die noch nicht beauftragt sind, können noch verschoben und 
        ' werden
        '
        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' Es muss mindestens 1 Projekt selektiert sein
            For i = 1 To awinSelection.Count

                singleShp = awinSelection.Item(i)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                    pname = hproj.name
                Catch ex As Exception
                    Call MsgBox(" Fehler! Projekt " & singleShp.Name & " nicht im Hauptspeicher")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                If hproj.Status = ProjektStatus(0) Then
                    ' nur dann macht das Setzen von earliest / latest Sinn ...

                    todoListe.Add(hproj.name)

                    If i = 1 Then

                        ' jetzt die Aktion durchführen ...

                        With setStartEnd

                            .EarliestStart.Value = hproj.earliestStart
                            .LatestStart.Value = hproj.latestStart

                        End With


                    Else

                        With setStartEnd

                            If .EarliestStart.Value <> hproj.earliestStart Or .LatestStart.Value <> hproj.latestStart Then

                                .EarliestStart.Value = 0
                                .LatestStart.Value = 0

                            End If

                        End With


                    End If
                Else
                    errMessage = errMessage & vbLf & hproj.name
                End If

            Next i

            If todoListe.Count > 0 Then

                returnValue = setStartEnd.ShowDialog

                If returnValue = DialogResult.OK Then

                    For i = 1 To todoListe.Count

                        pname = CStr(todoListe.Item(i))

                        ' jetzt die Aktion durchführen ...
                        Try
                            hproj = ShowProjekte.getProject(pname)
                            With setStartEnd

                                hproj.earliestStart = .EarliestStart.Value
                                hproj.latestStart = .LatestStart.Value
                                hproj.earliestStartDate = hproj.startDate.AddMonths(.EarliestStart.Value)
                                hproj.latestStartDate = hproj.startDate.AddMonths(.LatestStart.Value)

                            End With
                        Catch ex As Exception
                            Call MsgBox(" Fehler! Projekt " & pname & " earliest/latest kann nicht gesetzt werden")
                            enableOnUpdate = True
                            Exit Sub
                        End Try

                    Next i

                    Call MsgBox("ok, frühester und spätester Start gesetzt")

                ElseIf returnValue = DialogResult.Cancel Then
                    'Call MsgBox("Default soll gelten")

                End If

            End If

            If errMessage.Length > 0 Then
                Call MsgBox(initMsg & vbLf & errMessage)
            End If

        Else

            Call MsgBox("Es muss mindestens ein Projekt selektiert sein")

        End If

        Call awinDeSelect()

        'appInstance.ScreenUpdating = True
        enableOnUpdate = True


    End Sub

    ' am 21.3.17 von tk rausgenommen 
    ''' <summary>
    ''' Projekt ins Noshow stellen  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    ''' 
    ''Sub Tom2G1NoShow(control As IRibbonControl)

    ''    Dim singleShp As Excel.Shape
    ''    'Dim SID As String

    ''    Dim awinSelection As Excel.ShapeRange

    ''    Call projektTafelInit()

    ''    Dim formerEE As Boolean = appInstance.EnableEvents
    ''    appInstance.EnableEvents = False

    ''    enableOnUpdate = False

    ''    Try
    ''        'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
    ''        awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
    ''    Catch ex As Exception
    ''        awinSelection = Nothing
    ''    End Try

    ''    If Not awinSelection Is Nothing Then

    ''        ' jetzt die Aktion durchführen ...

    ''        For Each singleShp In awinSelection

    ''            Dim shapeArt As Integer
    ''            shapeArt = kindOfShape(singleShp)

    ''            With singleShp
    ''                If isProjectType(shapeArt) Then

    ''                    Call awinShowNoShowProject(pname:=.Name)

    ''                End If
    ''            End With
    ''        Next

    ''    Else
    ''        Call MsgBox("vorher Projekt selektieren ...")
    ''    End If

    ''    enableOnUpdate = True
    ''    appInstance.EnableEvents = formerEE
    ''End Sub

    ''' <summary>
    ''' neues Formular zur Auswahl Phasen/Meilensteine/Rollen/Kosten anzeigen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub NameHierarchySelAction(control As IRibbonControl)

        Dim formerES As Boolean = awinSettings.meEnableSorting

        Call PBBNameHierarchySelAction(control.Id)


        If control.Id = "PTMEC1" And awinSettings.meEnableSorting <> formerES Then
            Me.ribbon.Invalidate()
        End If
        

    End Sub


    Sub AnalyseLeistbarkeit001(ByVal control As IRibbonControl)


        Call PBBAnalyseLeistbarkeit001(control.Id)



    End Sub

    ' am 21.3.17 rausgenommen 
    ''' <summary>
    ''' Projekt ins Show zurückholen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    ''Sub Tom2G1Show(control As IRibbonControl)

    ''    Dim getBackToShow As New frmGetProjectbackFromNoshow

    ''    Dim returnValue As DialogResult

    ''    Call projektTafelInit()

    ''    enableOnUpdate = False
    ''    appInstance.ScreenUpdating = False

    ''    If AlleProjekte.Count > 0 And ShowProjekte.Count <> AlleProjekte.Count Then

    ''        returnValue = getBackToShow.ShowDialog
    ''    Else
    ''        If AlleProjekte.Count = 0 Then
    ''            Call MsgBox("Es sind keine Projekte geladen!  ")
    ''        Else
    ''            Call MsgBox("Es gibt keine Projekte in der Warteschlange !")
    ''        End If
    ''    End If



    ''    appInstance.ScreenUpdating = True
    ''    enableOnUpdate = True
    ''End Sub
    ''' <summary>
    ''' Änderungen akzeptieren 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Accept(control As IRibbonControl)

        Dim singleShp As Excel.Shape


        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then
                        Call awinBeauftragung(pname:=.Name, type:=0)
                    End If
                End With
            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

    End Sub



    ''' <summary>
    ''' Projekt beauftragen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Beauftragen(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        If ShowProjekte.contains(.Name) Then
                            Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)

                            If tryToprotectProjectforMe(hproj.name, hproj.variantName) Then
                                Call awinBeauftragung(pname:=hproj.name, type:=1)

                            Else
                                If awinSettings.englishLanguage Then
                                    Call MsgBox(hproj.name & ", " & hproj.variantName & " is protected " & vbLf & _
                                                "and cannot be modified. You could instead create a variant.")
                                Else
                                    Call MsgBox(hproj.name & ", " & hproj.variantName & " ist geschützt " & vbLf & _
                                                "und kann nicht verändert werden. Sie können jedoch eine Variante anlegen.")
                                End If
                            End If
                        End If

                    End If
                End With
            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

    End Sub

    ''' <summary>
    ''' Beauftragung zurücknehmen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2GXBeauftragen(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        If ShowProjekte.contains(.Name) Then
                            Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)

                            ' darf nur gemacht werden, wenn Varianten-NAme <> ""
                            If hproj.variantName <> "" Then
                                If tryToprotectProjectforMe(hproj.name, hproj.variantName) Then
                                    Call awinCancelBeauftragung(hproj.name)
                                Else
                                    If awinSettings.englishLanguage Then
                                        Call MsgBox(hproj.name & ", " & hproj.variantName & " is protected " & vbLf & _
                                                    "and cannot be modified. You could instead create a variant.")
                                    Else
                                        Call MsgBox(hproj.name & ", " & hproj.variantName & " ist geschützt " & vbLf & _
                                                    "und kann nicht verändert werden. Sie können jedoch eine Variante anlegen.")
                                    End If
                                End If
                            Else
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Base-Variant must not be de-freezed ..." & vbLf & _
                                                "please create a variant first ...")
                                Else
                                    Call MsgBox("die Fixierung der Standard Variante kann nicht aufgehoben werden ..." & vbLf & _
                                                "bitte erstellen Sie zu diesem Zweck eine Variante ...")
                                End If
                            End If
                            
                        End If

                    End If
                End With
            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

    End Sub



    ''' <summary>
    ''' Projekt löschen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Loeschen(control As IRibbonControl)

        Call PBBLoeschen(control)

    End Sub
    ' ur: 31.03.2017 trägt nur zur Verwirrung bei
    ' ReportProfil kann nun bei Report-Erstellung bearbeitet

    ' '' ''' <summary>
    ' '' ''' EinzelProjekt Report mit selektierter Vorlage erstellen
    ' '' ''' </summary>
    ' '' ''' <param name="control"></param>
    ' '' ''' <remarks></remarks>
    ' ''Sub awinBHTCReport(control As IRibbonControl)

    ' ''    ' Hierarchie auswählen, Einzelprojekt Berichte 
    ' ''    appInstance.ScreenUpdating = False

    ' ''    Call PBBBHTCHierarchySelAction(control.Id, Nothing)

    ' ''    appInstance.ScreenUpdating = True

    ' ''End Sub



    ''' <summary>
    ''' EinzelProjekt Report mit selektierter Vorlage erstellen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G1Doku(control As IRibbonControl)

        Dim awinSelection As Excel.ShapeRange
        Dim returnValue As DialogResult
        Dim getReportVorlage As New frmSelectPPTTempl

        Call projektTafelInit()

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If awinSelection Is Nothing Then
            Call MsgBox("vorher Projekt/e selektieren ...")
        Else
            enableOnUpdate = False
            appInstance.ScreenUpdating = False
            appInstance.EnableEvents = False

            getReportVorlage.calledfrom = "Projekt"


            ' sichern der awinSettings.mpp... Einstellungen

            ' Änderung tk 23.2.2016: das sollte nicht mehr explizit gesetzt werden - andernfalls kann man in Einzelprojekt-Reports 
            ' keine Zeiraumbetrachtungen mehr machen , ausserdem würde es sich anbieten, in den Swimlane Reports Einzel-Zeilen zu zeichnen 
            'Dim sav_mppShowAllIfOne As Boolean = awinSettings.mppShowAllIfOne
            'awinSettings.mppShowAllIfOne = True
            'Dim sav_mppExtendedMode As Boolean = awinSettings.mppExtendedMode
            'awinSettings.mppExtendedMode = True
            ' Settings für Einzelprojekt-Reports
            'awinSettings.eppExtendedMode = True


            ' Formular zum Auswählen der Report-Vorlage wird aufgerufen

            returnValue = getReportVorlage.ShowDialog

            'awinSettings.eppExtendedMode = False

            ' Zurücksetzen der gesicherten und veränderten Einstellungen

            ' Änderung tk 23.2.2016 
            'awinSettings.mppExtendedMode = sav_mppExtendedMode
            'awinSettings.mppShowAllIfOne = sav_mppShowAllIfOne

            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True
            enableOnUpdate = True
        End If

    End Sub

    Public Sub awinImportMassenEdit(control As IRibbonControl)

        ' Übernahme 

        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult
        Dim getMassenEditImport As New frmSelectImportFiles
        Dim wasNotEmpty As Boolean = False

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        'dateiName = awinPath & projektInventurFile

        getMassenEditImport.menueAswhl = PTImpExp.massenEdit
        returnValue = getMassenEditImport.ShowDialog

        If returnValue = DialogResult.OK Then
            dateiName = getMassenEditImport.selectedDateiName

            Try

                If My.Computer.FileSystem.FileExists(dateiName) Then


                    If ShowProjekte.Count > 0 Then
                        wasNotEmpty = True
                        'Call storeSessionConstellation("Last")
                        ' hier sollte jetzt auch ein ClearPlan-Tafel gemacht werden ...
                        Call awinClearPlanTafel()
                    End If

                    appInstance.Workbooks.Open(dateiName)
                    'Dim scenarioName As String = appInstance.ActiveWorkbook.Name
                    'Dim positionIX As Integer = scenarioName.IndexOf(".xls") - 1
                    'Dim tmpName As String = ""
                    'For ih As Integer = 0 To positionIX
                    '    tmpName = tmpName & scenarioName.Chars(ih)
                    'Next

                    Dim scenarioName As String = "ME"

                    ' alle Import Projekte erstmal löschen
                    ImportProjekte.Clear(False)


                    Call importiereMassenEdit()
                    appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                    Dim sessionConstellation As clsConstellation = verarbeiteImportProjekte(scenarioName, True)

                    ' ''If wasNotEmpty Then
                    ' ''    Call awinClearPlanTafel()
                    ' ''End If

                    '' ''Call awinZeichnePlanTafel(True)
                    ' ''Call awinZeichnePlanTafel(False)
                    ' ''Call awinNeuZeichnenDiagramme(2)

                    If sessionConstellation.count > 0 Then

                        If projectConstellations.Contains(scenarioName) Then
                            projectConstellations.Remove(scenarioName)
                        End If

                        projectConstellations.Add(sessionConstellation)
                        Call loadSessionConstellation(scenarioName, False, False, True)
                    Else
                        Call MsgBox("keine Projekte importiert ...")
                    End If

                    'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
                    'Call importProjekteEintragen(importDate, ProjektStatus(1))

                    If ImportProjekte.Count > 0 Then
                        ImportProjekte.Clear(False)
                    End If
                Else

                    Call MsgBox("bitte Datei auswählen ...")
                End If


            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
            End Try
        Else
            Call MsgBox(" Import Scenario wurde abgebrochen")
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True



    End Sub

    Public Sub Tom2G4B1InventurImport(control As IRibbonControl)
        ' Übernahme 

        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult

        Dim getInventurImport As New frmSelectImportFiles
        Dim wasNotEmpty As Boolean = False

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' wenn noch etwas in der session ist , warnen ! 
        If AlleProjekte.Count > 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("this function is only available with an empty session" & vbLf & _
                            "please store and clear your session first")
            Else
                Call MsgBox("diese Funktionalität ist nur möglich mit einer leeren Session" & vbLf & _
                            "bitte speichern Sie ggf. ihre Projekte und setzen die Session zurück.")
            End If
        Else
            ' Aktion durchführen ...
            getInventurImport.menueAswhl = PTImpExp.simpleScen
            returnValue = getInventurImport.ShowDialog

            If returnValue = DialogResult.OK Then
                dateiName = getInventurImport.selectedDateiName

                Try

                    If My.Computer.FileSystem.FileExists(dateiName) Then

                        If ShowProjekte.Count > 0 Then
                            wasNotEmpty = True
                            'Call storeSessionConstellation("Last")
                            ' hier sollte jetzt auch ein ClearPlan-Tafel gemacht werden ...
                            Call awinClearPlanTafel()
                        End If

                        appInstance.Workbooks.Open(dateiName)
                        Dim scenarioName As String = appInstance.ActiveWorkbook.Name
                        Dim positionIX As Integer = scenarioName.IndexOf(".xls") - 1
                        Dim tmpName As String = ""
                        For ih As Integer = 0 To positionIX
                            tmpName = tmpName & scenarioName.Chars(ih)
                        Next
                        scenarioName = tmpName.Trim

                        ' alle Import Projekte erstmal löschen
                        ImportProjekte.Clear(False)


                        Call awinImportProjektInventur()
                        appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                        Dim sessionConstellation As clsConstellation = verarbeiteImportProjekte(scenarioName)

                        ' ''If wasNotEmpty Then
                        ' ''    Call awinClearPlanTafel()
                        ' ''End If

                        '' ''Call awinZeichnePlanTafel(True)
                        ' ''Call awinZeichnePlanTafel(False)
                        ' ''Call awinNeuZeichnenDiagramme(2)

                        If sessionConstellation.count > 0 Then

                            If projectConstellations.Contains(scenarioName) Then
                                projectConstellations.Remove(scenarioName)
                            End If

                            projectConstellations.Add(sessionConstellation)
                            Call loadSessionConstellation(scenarioName, False, False, True)
                        Else
                            Call MsgBox("keine PRojekte importiert ...")
                        End If

                        'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
                        'Call importProjekteEintragen(importDate, ProjektStatus(1))

                        If ImportProjekte.Count > 0 Then
                            ImportProjekte.Clear(False)
                        End If
                    Else

                        Call MsgBox("bitte Datei auswählen ...")
                    End If


                Catch ex As Exception
                    appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                    Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
                End Try
            Else
                'Call MsgBox(" Import Scenario wurde abgebrochen")
            End If

        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True



    End Sub

    Public Sub Tom2G4B1ScenarioImport(control As IRibbonControl)
        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult

        Dim getScenarioImport As New frmSelectImportFiles
        Dim wasNotEmpty As Boolean = False

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False


        
            ' Aktion durchführen ...
        getScenarioImport.menueAswhl = PTImpExp.scenariodefs
        returnValue = getScenarioImport.ShowDialog

        If returnValue = DialogResult.OK Then
            dateiName = getScenarioImport.selectedDateiName

            Try

                If My.Computer.FileSystem.FileExists(dateiName) Then

                    If ShowProjekte.Count > 0 Then
                        wasNotEmpty = True
                        'Call storeSessionConstellation("Last")
                        ' hier sollte jetzt auch ein ClearPlan-Tafel gemacht werden ...
                        Call awinClearPlanTafel()
                    End If

                    appInstance.Workbooks.Open(dateiName)
                    Dim scenarioName As String = appInstance.ActiveWorkbook.Name
                    Dim positionIX As Integer = scenarioName.IndexOf(".xls") - 1
                    Dim tmpName As String = ""
                    For ih As Integer = 0 To positionIX
                        tmpName = tmpName & scenarioName.Chars(ih)
                    Next
                    scenarioName = tmpName.Trim


                    Dim newConstellation As clsConstellation = importScenarioDefinition(scenarioName)
                    appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                    ' ''If wasNotEmpty Then
                    ' ''    Call awinClearPlanTafel()
                    ' ''End If

                    '' ''Call awinZeichnePlanTafel(True)
                    ' ''Call awinZeichnePlanTafel(False)
                    ' ''Call awinNeuZeichnenDiagramme(2)

                    If newConstellation.count > 0 Then

                        If projectConstellations.Contains(scenarioName) Then
                            projectConstellations.Remove(scenarioName)
                        End If

                        If projectConstellations.Contains(scenarioName) Then
                            projectConstellations.Remove(scenarioName)
                        End If

                        projectConstellations.Add(newConstellation)
                        'Call loadSessionConstellation(scenarioName, False, False, True)
                    Else
                        Call MsgBox("keine Projekte für Szenario erkannt ...")
                    End If

                    'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
                    'Call importProjekteEintragen(importDate, ProjektStatus(1))

                    If ImportProjekte.Count > 0 Then
                        ImportProjekte.Clear(False)
                    End If
                Else

                    Call MsgBox("bitte Datei auswählen ...")
                End If


            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
            End Try
        Else
            'Call MsgBox(" Import Scenario wurde abgebrochen")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    ''' <summary>
    ''' importiert die Modul Batch Datei und legt entsprechende PRojekte an 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub Tom2G4B1ModulImport(control As IRibbonControl)

        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult
        Dim getModuleImport As New frmSelectImportFiles

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        'dateiName = awinPath & projektInventurFile

        getModuleImport.menueAswhl = PTImpExp.modulScen
        returnValue = getModuleImport.ShowDialog

        If returnValue = DialogResult.OK Then
            dateiName = getModuleImport.selectedDateiName

            Try
                appInstance.Workbooks.Open(dateiName)

                ' alle Import Projekte erstmal löschen
                ImportProjekte.Clear(False)
                Call awinImportModule(myCollection)
                appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
                Call importProjekteEintragen(importDate, ProjektStatus(1))

            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
            End Try
        Else
            Call MsgBox(" Import Scenario wurde abgebrochen")
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


    End Sub

    ''' <summary>
    ''' importiert die Modul Batch Datei und fügt den Projekten die entsprechenden Elemente regelbasiert an  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub PT4G1B10AddModularImport(control As IRibbonControl)

        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult
        Dim getModuleImport As New frmSelectImportFiles

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False


        getModuleImport.menueAswhl = PTImpExp.addElements
        returnValue = getModuleImport.ShowDialog

        If returnValue = DialogResult.OK Then
            dateiName = getModuleImport.selectedDateiName
            Dim ruleSet As New clsAddElements
            Dim ok As Boolean = True
            Try
                appInstance.Workbooks.Open(dateiName)

                ' jetzt werden die Regeln ausgelesen ...
                Call awinReadAddOnRules(ruleSet)
                appInstance.ActiveWorkbook.Close(SaveChanges:=True)

            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox("Fehler bei Lesen " & vbLf & dateiName & vbLf & ex.Message)
                ok = False
            End Try


            Dim awinSelection As Excel.ShapeRange
            Dim hproj As clsProjekt

            Try
                awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
            Catch ex As Exception
                awinSelection = Nothing
            End Try


            If ok Then

                Dim allOK As Boolean = True

                If Not awinSelection Is Nothing Then

                    ' jetzt die Aktion durchführen ...


                    For Each singleShp As Excel.Shape In awinSelection
                        Dim shapeArt As Integer
                        shapeArt = kindOfShape(singleShp)


                        With singleShp
                            If isProjectType(shapeArt) Then
                                hproj = ShowProjekte.getProject(singleShp.Name, True)

                                Try
                                    Call awinApplyAddOnRules(hproj, ruleSet)
                                Catch ex As Exception
                                    Call MsgBox(hproj.name & ":" & vbLf & ex.Message)
                                    allOK = False
                                End Try

                            End If
                        End With
                    Next

                    If allOK Then
                        Call MsgBox("ok, alle Projekte wurden um " & ruleSet.name & " ergänzt")
                    Else
                        Call MsgBox("ok, ergänzt, was möglich war ...")
                    End If


                Else

                    For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                        Try
                            Call awinApplyAddOnRules(kvp.Value, ruleSet)
                        Catch ex As Exception
                            Call MsgBox(kvp.Value.name & ":" & vbLf & ex.Message)
                            allOK = False
                        End Try


                    Next

                    If allOK Then
                        Call MsgBox("ok, alle Projekte wurden um " & ruleSet.name & " ergänzt")
                    Else
                        Call MsgBox("ok, ergänzt, was möglich war ...")
                    End If

                End If

            End If

        Else
            Call MsgBox(" Ergänzungs-Vorgang wurde abgebrochen")
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


    End Sub

    ''' <summary>
    ''' importiert eine Excel Datei mit Phasen und Meilensteinen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub Tom2G4B1RPLANImport(control As IRibbonControl)


        Dim dateiName As String = ""
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult
        Dim getRPLANImport As New frmSelectImportFiles
        Dim listofVorlagen As New Collection
        'Dim xlsRplanImport As Excel.Workbook

        Call projektTafelInit()


        'dateiName = awinPath & projektInventurFile

        getRPLANImport.menueAswhl = PTImpExp.rplan
        returnValue = getRPLANImport.ShowDialog

        If returnValue = DialogResult.OK Then

            listofVorlagen = getRPLANImport.selImportFiles

            Dim i As Integer
            For i = 1 To listofVorlagen.Count


                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                enableOnUpdate = False

                dateiName = listofVorlagen.Item(i).ToString

                Try
                    appInstance.Workbooks.Open(dateiName)

                    '' '' alle Import Projekte erstmal löschen
                    ImportProjekte.Clear(False)
                    myCollection.Clear()
                    'Call bmwImportProjektInventur(myCollection)
                    Call planExcelImport(myCollection, False, dateiName)
                    'Call bmwImportProjekteITO15(myCollection, False)

                    appInstance.ActiveWorkbook.Close(SaveChanges:=True)
                    ' xlsRplanImport.Close(SaveChanges:=True)


                    appInstance.ScreenUpdating = True
                    Call importProjekteEintragen(importDate, ProjektStatus(1))

                    'Call awinWritePhaseDefinitions()
                    'Call awinWritePhaseMilestoneDefinitions()

                Catch ex As Exception
                    appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                    Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
                End Try

            Next i

            ' ''appInstance.ScreenUpdating = True
            ' ''Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
        Else
            'Call MsgBox(" Import RPLAN-Projekte wurde abgebrochen")
            'Call logfileSchreiben(" Import RPLAN-Projekte wurde abgebrochen", dateiName, -1)
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub
    ' ''Public Sub Tom2G4B1RPLANImport(control As IRibbonControl)


    ' ''    Dim dateiName As String
    ' ''    Dim myCollection As New Collection
    ' ''    Dim importDate As Date = Date.Now
    ' ''    Dim returnValue As DialogResult
    ' ''    Dim getRPLANImport As New frmSelectRPlanImport

    ' ''    Call projektTafelInit()

    ' ''    appInstance.EnableEvents = False
    ' ''    appInstance.ScreenUpdating = False
    ' ''    enableOnUpdate = False

    ' ''    'dateiName = awinPath & projektInventurFile

    ' ''    getRPLANImport.menueAswhl = PTImpExp.rplan
    ' ''    returnValue = getRPLANImport.ShowDialog

    ' ''    If returnValue = DialogResult.OK Then
    ' ''        dateiName = getRPLANImport.selectedDateiName

    ' ''        Try
    ' ''            appInstance.Workbooks.Open(dateiName)

    ' ''            ' alle Import Projekte erstmal löschen
    ' ''            ImportProjekte.Clear()
    ' ''            'Call bmwImportProjektInventur(myCollection)
    ' ''            Call rplanExcelImport(myCollection, False)
    ' ''            'Call bmwImportProjekteITO15(myCollection, False)
    ' ''            appInstance.ActiveWorkbook.Close(SaveChanges:=True)

    ' ''            appInstance.ScreenUpdating = True
    ' ''            Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))

    ' ''            'Call awinWritePhaseDefinitions()
    ' ''            'Call awinWritePhaseMilestoneDefinitions()

    ' ''        Catch ex As Exception
    ' ''            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
    ' ''            Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
    ' ''        End Try
    ' ''    Else
    ' ''        Call MsgBox(" Import RPLAN-Projekte wurde abgebrochen")
    ' ''    End If



    ' ''    enableOnUpdate = True
    ' ''    appInstance.EnableEvents = True
    ' ''    appInstance.ScreenUpdating = True

    ' ''End Sub

    Public Sub Tom2G4B3RPLANRxfImport(control As IRibbonControl)


        Dim dateiName As String = ""
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult
        Dim getRPLANImport As New frmSelectImportFiles
        Dim protokoll As New SortedList(Of Integer, clsProtokoll)

        ' öffnen des LogFiles
        Call logfileOpen()

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        'dateiName = awinPath & projektInventurFile

        getRPLANImport.menueAswhl = PTImpExp.rplanrxf
        returnValue = getRPLANImport.ShowDialog

        If returnValue = DialogResult.OK Then
            dateiName = getRPLANImport.selectedDateiName

            Try

                ' alle Import Projekte erstmal löschen
                ImportProjekte.Clear(False)

                Call logfileSchreiben("Beginn RXFImport ", dateiName, -1)

                Call RXFImport(myCollection, dateiName, False, protokoll)

                'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
                Call importProjekteEintragen(importDate, ProjektStatus(1))

                Dim result As Integer = MsgBox("Soll ein Protokoll geschrieben werden?", MsgBoxStyle.YesNo)
                If result = MsgBoxResult.Yes Then

                    appInstance.ScreenUpdating = True

                    ' Tabellenblattname aus dateiname erstellen fürs Protokoll (dateiname ohne ".rxf" Extension)
                    Dim tstr As String() = Split(dateiName, "\", -1)
                    Dim hstr As String = tstr(tstr.Length - 1)
                    tstr = Split(hstr, ".", 2)
                    Dim tabblattname As String = tstr(0)

                    appInstance.ScreenUpdating = False
                    ' Protokoll aus der Liste protokoll in Logfile mit tabellenblatt tabblattname ausleiten
                    Call writeProtokoll(protokoll, tabblattname)
                End If


                ' tk Änderung 26.11.15 das muss doch nach dem Import noch nicht gemacht werden
                ' sondern erst nach Editieren Wörterbuch oder ganz am Schluss beim Beenden 
                'Call awinWritePhaseDefinitions()
                'Call awinWritePhaseMilestoneDefinitions()

            Catch ex As Exception

                Call MsgBox(ex.Message & vbLf & dateiName & vbLf & "Fehler bei RXFImport ")
            End Try

        Else
            'Call MsgBox(" RXF-Import RPLAN-Projekte wurde abgebrochen")
            'Call logfileSchreiben(" RXF-Import RPLAN-Projekte wurde abgebrochen", dateiName, -1)
        End If


        ' Schließen des LogFiles
        Call logfileSchliessen()

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Public Sub Tom2G4M1Import(control As IRibbonControl)

        If Not noDB Then
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        End If
        Dim hproj As New clsProjekt
        Dim cproj As New clsProjekt
        Dim vglName As String = " "
        Dim outputString As String = ""
        'Dim dirName As String
        Dim dateiName As String
        Dim pname As String
        Dim importDate As Date = Date.Now
        'Dim importDate As Date = "31.10.2013"
        Dim listofVorlagen As Collection
        Dim projektInventurFile As String = "ProjektInventur.xlsm"

        Dim getVisboImport As New frmSelectImportFiles
        Dim returnValue As DialogResult

        Call logfileOpen()

        getVisboImport.menueAswhl = PTImpExp.visbo
        returnValue = getVisboImport.ShowDialog

        If returnValue = DialogResult.OK Then

            listofVorlagen = getVisboImport.selImportFiles

            Call projektTafelInit()

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False
            enableOnUpdate = False

            Dim myCollection As New Collection



            '' ''dirName = awinPath & msprojectFilesOrdner
            ' ''dirName = importOrdnerNames(PTImpExp.msproject)
            ' ''listOfVorlagen = My.Computer.FileSystem.GetFiles(dirName, FileIO.SearchOption.SearchTopLevelOnly, "*.mpp")

            ' alle Import Projekte erstmal löschen
            ImportProjekte.Clear(False)


            ' jetzt müssen die Projekte ausgelesen werden, die in dateiListe stehen 
            Dim i As Integer
            For i = 1 To listofVorlagen.Count
                dateiName = listofVorlagen.Item(i).ToString
                ' öffnen des LogFiles


                If dateiName = projektInventurFile Then

                    ' nichts machen 

                Else
                    Dim skip As Boolean = False


                    Try
                        appInstance.Workbooks.Open(dateiName)
                        Call logfileSchreiben("Beginn Import ", dateiName, -1)

                    Catch ex1 As Exception
                        Call logfileSchreiben("Fehler bei Öffnen der Datei ", dateiName, -1)
                        skip = True
                    End Try

                    If Not skip Then
                        pname = ""
                        hproj = New clsProjekt
                        Try
                            Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)

                            Try
                                Dim keyStr As String = calcProjektKey(hproj)
                                ImportProjekte.Add(hproj, False)
                                myCollection.Add(calcProjektKey(hproj))
                            Catch ex2 As Exception
                                Call MsgBox("Projekt kann nicht zweimal importiert werden ...")
                            End Try

                            appInstance.ActiveWorkbook.Close(SaveChanges:=False)

                        Catch ex1 As Exception
                            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                            Call logfileSchreiben(ex1.Message, "", anzFehler)
                            Call MsgBox(ex1.Message)
                            'Call MsgBox("Fehler bei Import von Projekt " & hproj.name & vbCrLf & "Siehe Logfile")
                        End Try



                    End If



                End If


            Next i


            Try
                Call importProjekteEintragen(importDate, ProjektStatus(1))
                'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
            Catch ex As Exception
                Call MsgBox("Fehler bei Import : " & vbLf & ex.Message)
            End Try

        Else

            'Call logfileSchreiben("Import wurde abgebrochen", "", -1)

        End If



        Call logfileSchliessen()

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


    End Sub

    Public Sub Tom2G4M2ImportMSProject(control As IRibbonControl)

        If Not noDB Then
            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        End If
        Dim hproj As New clsProjekt
        Dim cproj As New clsProjekt
        Dim vglName As String = " "
        Dim outputString As String = ""
        Dim dateiName As String
        Dim getMSImport As New frmSelectImportFiles
        Dim returnValue As DialogResult

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False



        getMSImport.menueAswhl = PTImpExp.msproject
        returnValue = getMSImport.ShowDialog

        If returnValue = DialogResult.OK Then


            Dim importDate As Date = Date.Now
            'Dim importDate As Date = "31.10.2013"
            ''Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String)
            Dim listofVorlagen As Collection
            listofVorlagen = getMSImport.selImportFiles

            Call projektTafelInit()

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False
            enableOnUpdate = False

            Dim myCollection As New Collection



            '' ''dirName = awinPath & msprojectFilesOrdner
            ' ''dirName = importOrdnerNames(PTImpExp.msproject)
            ' ''listOfVorlagen = My.Computer.FileSystem.GetFiles(dirName, FileIO.SearchOption.SearchTopLevelOnly, "*.mpp")

            ' alle Import Projekte erstmal löschen
            ImportProjekte.Clear(False)


            ' jetzt müssen die Projekte ausgelesen werden, die in dateiListe stehen 
            Dim i As Integer
            For i = 1 To listofVorlagen.Count
                dateiName = listofVorlagen.Item(i).ToString


                ' '' ''Dim skip As Boolean = False


                ' '' ''Try
                ' '' ''    appInstance.Workbooks.Open(dateiName)
                ' '' ''Catch ex1 As Exception
                ' '' ''    'Call MsgBox("Fehler bei Öffnen der Datei " & dateiName)
                ' '' ''    skip = True
                ' '' ''End Try

                ' '' ''If Not skip Then
                ' '' ''    pname = ""
                hproj = New clsProjekt
                Try
                    Call awinImportMSProject("", dateiName, hproj, importDate)

                    Try
                        Dim keyStr As String = calcProjektKey(hproj)
                        ImportProjekte.Add(hproj, False)
                        myCollection.Add(calcProjektKey(hproj))
                    Catch ex2 As Exception
                        Call MsgBox("Projekt kann nicht zweimal importiert werden ...")
                    End Try

                    ' ''appInstance.ActiveWorkbook.Close(SaveChanges:=False)

                Catch ex1 As Exception
                    ''appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                    Call MsgBox(ex1.Message)
                    Call MsgBox("Fehler bei Import von Projekt " & hproj.name)
                End Try

            Next i


            ' '' ''End If

            Try
                'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
                Call importProjekteEintragen(importDate, ProjektStatus(1))
            Catch ex As Exception

                Call MsgBox("Fehler bei Import : " & vbLf & ex.Message)
            End Try



        End If


        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True




    End Sub

    ''' <summary>
    ''' exportiert selektierte/alle Files in eine Excel Datei, die genauso aufgebaut ist , wie die BMW Import Datei  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub planExcelExport(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim outputString As String = ""
        Dim fileListe As New SortedList(Of String, String)
        Dim exportFileName As String = "Export_" & Date.Now.ToString.Replace(":", ".") & ".xlsx"
        Dim ok As Boolean

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                ' hier muss jetzt die todo Liste aufgebaut werden 

                'Dim shapeArt As Integer
                'shapeArt = kindOfShape(singleShp)

                With singleShp
                    'If isProjectType(shapeArt) Then

                    Try

                        hproj = ShowProjekte.getProject(singleShp.Name, True)
                        fileListe.Add(hproj.name, hproj.name)

                    Catch ex As Exception

                        Call MsgBox(singleShp.Name & ": Fehler bei Aufbau todo Liste für Export ...")

                    End Try

                    'End If
                End With

            Next

        Else
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                fileListe.Add(kvp.Key, kvp.Key)
            Next
        End If

        ' hier muss jetzt das File Projekt Detail aufgemacht werden ...
        Try
            appInstance.Workbooks.Open(awinPath & requirementsOrdner & excelExportVorlage)
            ok = True
        Catch ex As Exception
            ok = False
        End Try

        If ok Then
            Dim zeile As Integer = 2
            For Each kvp As KeyValuePair(Of String, String) In fileListe

                Try
                    hproj = ShowProjekte.getProject(kvp.Key)

                    ' jetzt wird dieses Projekt exportiert ... 
                    Try
                        'Call bmwExportProject(hproj, zeile)
                        Call planExportProject(hproj, zeile)
                        outputString = outputString & hproj.name & " erfolgreich .." & vbLf
                    Catch ex As Exception
                        outputString = outputString & hproj.name & " nicht erfolgreich .." & vbLf & _
                                        ex.Message & vbLf & vbLf
                    End Try



                Catch ex As Exception

                    Call MsgBox(ex.Message)

                End Try

            Next

            Try
                ' Schließen der Export Datei unter neuem Namen, original Zustand bleibt erhalten
                'appInstance.ActiveWorkbook.Close(SaveChanges:=True, Filename:=awinPath & exportFilesOrdner & "\" & _
                '                                 exportFileName)
                appInstance.ActiveWorkbook.Close(SaveChanges:=True, Filename:=exportOrdnerNames(PTImpExp.rplan) & "\" & _
                                                 exportFileName)
                Call MsgBox(outputString & "exportiert !")
            Catch ex As Exception

                Call MsgBox("Fehler beim Speichern der Export Datei")

            End Try

        End If


        Call awinDeSelect()
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True



    End Sub

    Public Sub awinWritePrioList(control As IRibbonControl)
        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Try
            Call writeProjektsForSequencing()
        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try


        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
    End Sub

    ''' <summary>
    ''' exportiert alle angezeigten Projekte in eine Massen-Edit Datei 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub awinWriteProjektBedarfeXLSX(control As IRibbonControl)

        If showRangeLeft <= 0 And Not showRangeRight > showRangeLeft Then
            Call MsgBox("bitte  einen Zeitraum angeben")
            Exit Sub
        End If

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Try
            If control.Id = "PT4G2M3B1" Then
                ' Call writeProjektBedarfeXLSX(showRangeLeft, showRangeRight, 0)
                Call writeProjektPhasenBedarfeXLSX(showRangeLeft, showRangeRight, 0)
            ElseIf control.Id = "PT4G2M3B2" Then
                ' Call writeProjektBedarfeXLSX(showRangeLeft, showRangeRight, 1)
                Call writeProjektPhasenBedarfeXLSX(showRangeLeft, showRangeRight, 1)
            ElseIf control.Id = "PT4G2M3B3" Then
                'Call writeProjektBedarfeXLSX(showRangeLeft, showRangeRight, 2)
                Call writeProjektPhasenBedarfeXLSX(showRangeLeft, showRangeRight, 2)
            End If
        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try


        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    ''' <summary>
    ''' exportiert selektierte / alle Files in eine Excel Datei; 
    ''' verwendet dabei die Vorlage in Requirements bmwFC52Vorlage.xlsx
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub bmwFC52Export(control As IRibbonControl)

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Call awinWriteFC52()

        Call awinDeSelect()
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub


    Public Sub Tom2G4M1Export(control As IRibbonControl)


        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim outputString As String = ""
        Dim outPutCollection As New Collection

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Try
                    ' hier muss jetzt das File Projekt Detail aufgemacht werden ...
                    appInstance.Workbooks.Open(awinPath & projektAustausch)

                    Dim shapeArt As Integer
                    shapeArt = kindOfShape(singleShp)

                    With singleShp
                        If isProjectType(shapeArt) Then

                            Try
                                hproj = ShowProjekte.getProject(singleShp.Name, True)

                                ' jetzt wird dieses Projekt exportiert ... 
                                Try
                                    Call awinExportProjectmitHrchy(hproj)

                                    outputString = hproj.getShapeText & " erfolgreich .."
                                    outPutCollection.Add(outputString)
                                Catch ex As Exception
                                    outputString = hproj.getShapeText & " nicht erfolgreich .."
                                    outPutCollection.Add(outputString)
                                End Try



                            Catch ex As Exception
                                outputString = singleShp.Name & " nicht gefunden ..."
                                outPutCollection.Add(outputString)
                            End Try

                        End If
                    End With
                    Try
                        ' Schließen der Datei ProjektSteckbrief ohne abspeichern der Änderungen, original Zustand bleibt erhalten
                        appInstance.ActiveWorkbook.Close(SaveChanges:=False, Filename:=awinPath & projektAustausch)
                    Catch ex As Exception

                        outputString = "Fehler beim Schließen der Projektaustausch Vorlage"
                        outPutCollection.Add(outputString)

                    End Try
                Catch ex As Exception

                    outputString = "Fehler beim Öffnen der Projektaustausch Vorlage"
                    outPutCollection.Add(outputString)

                End Try


            Next

            If outPutCollection.Count > 0 Then
                Call showOutPut(outPutCollection, _
                                 "Exportieren Steckbriefe", _
                                 "erfolgreich exportierte Dateien liegen in " & vbLf & _
                                 exportOrdnerNames(PTImpExp.visbo))
            End If

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If


        Call awinDeSelect()
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True




    End Sub

    ''' <summary>
    ''' erstellt die Summary Zuordnungs-Datei 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G4M2B1ZuordnungRP(control As IRibbonControl)


        Dim fileName As String
        Dim zeile As Integer = 2
        Dim ok As Boolean

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False


        fileName = "Vorlage Zuordnung.xlsx"

        ' öffnen der Excel Datei 
        Try
            appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & fileName)
        Catch ex As Exception
            Call MsgBox("File " & fileName & " nicht gefunden ... Abbruch")
            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True
            enableOnUpdate = True
            Exit Sub
        End Try




        Call awinExportRessZuordnung(0, " ")


        Try

            appInstance.ActiveWorkbook.SaveAs(awinPath & projektRessOrdner & "\Summary.xlsx", _
                                      ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
            ok = True
            appInstance.ActiveWorkbook.Close()

        Catch ex As Exception
            ok = False
            appInstance.ActiveWorkbook.Close()
        End Try


        If ok Then
            Call MsgBox("ok, Datei erstellt ...")
        Else
            Call MsgBox("Fehler bei Save as ..\summary.xlsx")
        End If


        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' erstellt die Zuordnungs-Datei Ressourcen -> Projekt
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G4M2B2ZuordnungRP(control As IRibbonControl)

        Dim initialeVorlageName As String, kapaFileName As String
        Dim zeile As Integer = 2
        Dim anzRollen As Integer
        Dim i As Integer
        Dim initMessage As String = "bitte die Kapazitäten eintragen zu folgenden Rollen" & vbLf
        Dim infoMessage As String = initMessage
        Dim zuordnungsOrdner As String = projektRessOrdner & "\" & "Projekt Zuordnungen"

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False



        ' für jede Ressource eine eigene Datei machen
        anzRollen = RoleDefinitions.Count

        Dim ok As Boolean = True
        Dim roleName As String

        For i = 1 To anzRollen

            roleName = RoleDefinitions.getRoledef(i).name.Trim
            kapaFileName = roleName & " Kapazität.xlsx"

            ' öffnen der Excel Datei 
            Try

                appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & kapaFileName)
                ok = True

            Catch ex As Exception

                initialeVorlageName = "template Kapazität.xlsx"
                ok = False

                Try
                    appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & initialeVorlageName)
                    Try
                        appInstance.ActiveWorkbook.SaveAs(awinPath & projektRessOrdner & "\" & kapaFileName, _
                                      ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

                        infoMessage = infoMessage & kapaFileName & vbLf
                    Catch ex2 As Exception

                    End Try



                Catch ex1 As Exception
                    Call MsgBox("File " & initialeVorlageName & " nicht gefunden ... Abbruch" & vbLf & vbLf & _
                                "dieses File muss im Ordner " & awinPath & projektRessOrdner & "abgelegt werden")
                    appInstance.EnableEvents = True
                    appInstance.ScreenUpdating = True
                    enableOnUpdate = True
                    Exit Sub
                End Try

            End Try


            If ok Then

                Dim curFilename As String = roleName & " Projekt-Zuordnung" & " " & Date.Now.ToString("MMM yy") & ".xlsx"


                Try
                    Call awinExportRessZuordnung(1, roleName)
                    'appInstance.ActiveWorkbook.Save()

                    appInstance.ActiveWorkbook.SaveAs(Filename:=awinPath & zuordnungsOrdner & "\" & curFilename, _
                                                      ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)


                Catch ex As Exception

                    Call MsgBox("Fehler bei Zuordnung " & roleName)
                End Try

            End If


            appInstance.ActiveWorkbook.Close(SaveChanges:=False)



        Next

        If infoMessage.Length > initMessage.Length Then
            ' in diesem Fall wurden  nur die Kapazität-Zuordnungs-Files erstellt 
            infoMessage = infoMessage & vbLf & vbLf & "es wurden noch keine Zuordnungs-Dateien erstellt!"
            Call MsgBox(infoMessage)
        Else
            Call MsgBox("ok, Dateien erstellt ...")
        End If



        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True


    End Sub

    Sub PTDemoModusHistory(control As IRibbonControl, ByRef pressed As Boolean)

        demoModusHistory = Not demoModusHistory
        pressed = demoModusHistory

    End Sub

    Sub PTTestFunktion5(control As IRibbonControl)

        Dim demoModusDate As New frmdemoModusDate
        Dim returnValue As DialogResult

        Call projektTafelInit()


        demoModusHistory = True

        returnValue = demoModusDate.ShowDialog

        If returnValue = DialogResult.OK Then

            If demoModusHistory Then
                Call MsgBox("Demo Modus History: Ein" & vbLf & "neues Datum: " & historicDate)
            Else
                Call MsgBox("Demo Modus History: Aus")
            End If

        Else
            If demoModusHistory Then
                Call MsgBox("Demo Modus History: Ein" & vbLf & "altes Datum: " & historicDate)
            Else
                Call MsgBox("Demo Modus History: Aus")
            End If

        End If






    End Sub


    Public Sub PT5phasenZeichnenInit(control As IRibbonControl, ByRef pressed As Boolean)

        pressed = awinSettings.drawphases

    End Sub

    Public Sub PT5phasenZeichnen(control As IRibbonControl, ByRef pressed As Boolean)

        Dim i As Integer
        Dim hproj As clsProjekt

        Call projektTafelInit()

        Cursor.Current = Cursors.WaitCursor

        ' erstmal alle Beschriftungen löschen 
        Call deleteBeschriftungen()

        If pressed Then
            ' jetzt werden die Projekt-Symbole inkl Phasen Darstellung gezeichnet
            awinSettings.drawphases = True
            Call awinClearPlanTafel()
            Call awinZeichnePlanTafel(True)
        Else
            ' extendedView der einzelnen Projekte, sofern gesetzt, entfernen
            For i = 1 To ShowProjekte.Count
                hproj = ShowProjekte.getProject(i)
                hproj.extendedView = False
            Next
            ' jetzt werden die Projekt-Symbole ohne Phasen Darstellung gezeichnet 
            awinSettings.drawphases = False
            'Call awinLoadConstellation("Last")
            Call awinClearPlanTafel()
            Call awinZeichnePlanTafel(True)
        End If

        Cursor.Current = Cursors.Default

    End Sub

    Public Function PTShowSelectedObjects(control As IRibbonControl) As Boolean

        PTShowSelectedObjects = awinSettings.showValuesOfSelected

    End Function

    Sub awinSetShowSelObj(control As IRibbonControl, ByRef pressed As Boolean)

        awinSettings.showValuesOfSelected = pressed
        
    End Sub


    Public Function PTPropAnpassen(control As IRibbonControl) As Boolean

        PTPropAnpassen = awinSettings.propAnpassRess

    End Function

    Sub awinSetPropAnpass(control As IRibbonControl, ByRef pressed As Boolean)


        awinSettings.propAnpassRess = pressed
        

    End Sub

    Public Function PTPhaseAnteilig(control As IRibbonControl) As Boolean
        PTPhaseAnteilig = awinSettings.phasesProzentual
    End Function

    Sub awinSetPhaseAnteilig(control As IRibbonControl, ByRef pressed As Boolean)
        awinSettings.phasesProzentual = pressed
    End Sub

    Public Function PTProzAuslastung(control As IRibbonControl) As Boolean
        PTProzAuslastung = awinSettings.mePrzAuslastung
    End Function

    Sub awinPTProzAuslastung(control As IRibbonControl, ByRef pressed As Boolean)
        awinSettings.mePrzAuslastung = pressed

        ' jetzt muss der Auslastungs-Array neu aufgebaut werden 
        visboZustaende.clearAuslastungsArray()
        If awinSettings.meExtendedColumnsView Then
            Call updateMassEditAuslastungsValues(showRangeLeft, showRangeRight, Nothing)
        End If


    End Sub

    Public Function PTSkipChanges(control As IRibbonControl) As Boolean
        PTSkipChanges = tempSkipChanges
    End Function

    Sub awinPTSkipChanges(control As IRibbonControl, ByRef pressed As Boolean)
        tempSkipChanges = pressed
    End Sub

    Public Function PTenableSorting(control As IRibbonControl) As Boolean
        PTenableSorting = awinSettings.meEnableSorting
    End Function

    Sub awinPTenableSorting(control As IRibbonControl, ByRef pressed As Boolean)
        awinSettings.meEnableSorting = pressed

        If awinSettings.meEnableSorting Then
            With CType(appInstance.ActiveSheet, Excel.Worksheet)
                .Unprotect("x")
                .EnableSelection = Excel.XlEnableSelection.xlNoRestrictions
            End With
        Else
            With CType(appInstance.ActiveSheet, Excel.Worksheet)
                .Protect(Password:="x", UserInterfaceOnly:=True, _
                         AllowFormattingCells:=True, _
                         AllowInsertingColumns:=False,
                         AllowInsertingRows:=True, _
                         AllowDeletingColumns:=False, _
                         AllowDeletingRows:=True, _
                         AllowSorting:=True, _
                         AllowFiltering:=True)
                .EnableSelection = Excel.XlEnableSelection.xlUnlockedCells
                .EnableAutoFilter = True
            End With
        End If
    End Sub

    Public Function PTautomaticReduce(control As IRibbonControl) As Boolean
        PTautomaticReduce = awinSettings.meAutoReduce
    End Function

    Sub awinPTautomaticReduce(control As IRibbonControl, ByRef pressed As Boolean)
        awinSettings.meAutoReduce = pressed
    End Sub
    'Public Sub PT6StriktPressed(control As IRibbonControl, ByRef pressed As Boolean)

    '    pressed = awinSettings.mppStrict

    'End Sub

    'Public Sub PT6SetStrict(control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppStrict = True
    '    Else
    '        awinSettings.mppStrict = False
    '    End If

    'End Sub

    'Public Sub PT6fullyContainedPressed(control As IRibbonControl, ByRef pressed As Boolean)

    '    pressed = awinSettings.mppFullyContained

    'End Sub

    'Public Sub PT6SetfullyContained(control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppFullyContained = True
    '    Else
    '        awinSettings.mppFullyContained = False
    '    End If

    'End Sub


    'Public Sub PT6DateTextPressed(control As IRibbonControl, ByRef pressed As Boolean)
    '    pressed = awinSettings.mppShowMsDate
    'End Sub


    'Public Sub PT6SetShowDate(Control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppShowMsDate = True
    '    Else
    '        awinSettings.mppShowMsDate = False
    '    End If

    'End Sub


    'Public Sub PT6NameTextPressed(control As IRibbonControl, ByRef pressed As Boolean)
    '    pressed = awinSettings.mppShowMsName
    'End Sub


    'Public Sub PT6SetShowName(Control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppShowMsName = True
    '    Else
    '        awinSettings.mppShowMsName = False
    '    End If

    'End Sub

    'Public Sub PT6ProjectLinePressed(control As IRibbonControl, ByRef pressed As Boolean)
    '    pressed = awinSettings.mppShowProjectLine
    'End Sub


    'Public Sub PT6SetShowProjectLine(Control As IRibbonControl, ByRef pressed As Boolean)

    '    If pressed Then
    '        awinSettings.mppShowProjectLine = True
    '    Else
    '        awinSettings.mppShowProjectLine = False
    '    End If

    'End Sub

    Public Function PT6AmpelnPressed(control As IRibbonControl) As Boolean
        PT6AmpelnPressed = awinSettings.mppShowAmpel
    End Function


    Public Sub PT6SetShowAmpeln(Control As IRibbonControl, ByRef pressed As Boolean)


        awinSettings.mppShowAmpel = pressed
        

    End Sub



    ''' <summary>
    ''' lädt die gewählten Projekte und gewählten Varianten in die Session
    ''' </summary>
    ''' <param name="Control"></param>
    ''' <remarks></remarks>
    Public Sub PT5DatenbankLoadProjekte(Control As IRibbonControl)

        Call PBBDatenbankLoadProjekte(Control)


    End Sub


    Public Function PT5loadprojectsInit(control As IRibbonControl) As Boolean

        PT5loadprojectsInit = awinSettings.applyFilter


    End Function

    Public Sub PT5loadProjectsOnChange(control As IRibbonControl, ByRef pressed As Boolean)

        Call projektTafelInit()

        If pressed Then
            ' jetzt sollen die Projekte gemäß Time Span geladen werden - auch bei Veränderung des TimeSpan 
            awinSettings.applyFilter = True
            ' noch zu tun 
            ' Call awinloadProjectsFromDB()
        Else

            ' jetzt werden bei TimeSpan Änderung keine Projekte nachgeladen 
            awinSettings.applyFilter = False


        End If


    End Sub
    ''' <summary>
    ''' Charakteristik Phasen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B1Phasen(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim top As Double, left As Double, width As Double, height As Double
        Dim hproj As clsProjekt
        Dim scale As Double
        'Dim SID As String

        Call projektTafelInit()

        enableOnUpdate = False

        Dim awinSelection As Excel.ShapeRange

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                Catch ex As Exception
                    Call MsgBox("Projekt nicht gefunden ..." & singleShp.Name)
                    enableOnUpdate = True
                    Exit Sub
                End Try

                top = singleShp.Top + boxHeight + 2
                left = singleShp.Left - 5
                If left <= 0 Then
                    left = 5
                End If

                height = 380
                width = hproj.dauerInDays / 365 * 12 * boxWidth + 7
                scale = hproj.dauerInDays


                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False

                repObj = Nothing
                Dim noColorCollection As New Collection
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True
            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' für BMW Akquise erzeugt 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B1Phasen2(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim pname As String

        Call projektTafelInit()

        enableOnUpdate = False

        Dim awinSelection As Excel.ShapeRange

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                pname = singleShp.Name

                Try
                    hproj = ShowProjekte.getProject(pname, True)
                    pname = hproj.name
                Catch ex As Exception
                    Call MsgBox("Projekt nicht gefunden ..." & pname)
                    enableOnUpdate = True
                    Exit Sub
                End Try

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False


                Dim tmpCollection As New Collection
                ' bestimme die Anzahl Zeilen, die benötigt wird 
                Dim anzahlZeilen As Integer = hproj.calcNeededLines(tmpCollection, tmpCollection, awinSettings.drawphases, False)
                Call moveShapesDown(tmpCollection, hproj.tfZeile + 1, anzahlZeilen, 0)
                'Call ZeichneProjektinPlanTafel2(pname, hproj.tfZeile)
                Dim listCollection As New Collection
                Call ZeichneProjektinPlanTafel(tmpCollection, pname, hproj.tfZeile, listCollection, listCollection)


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True
            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True

    End Sub

    ''' <summary>
    ''' Charakteristik Personal Bedarfe
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B2Resources(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 1
        Dim top As Double, left As Double, width As Double, height As Double

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                Dim ok As Boolean = True
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 180

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                Catch ex As Exception
                    ok = False
                    hproj = Nothing
                End Try

                If ok Then

                    Dim repObj As Excel.ChartObject
                    appInstance.EnableEvents = False
                    appInstance.ScreenUpdating = False

                    repObj = Nothing

                    width = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 10, 6 * boxWidth + 10)

                    Try
                        Call createRessBalkenOfProject(hproj, repObj, auswahl, top, left, height, width)

                        ' jetzt wird das Pie-Diagramm gezeichnet 
                        left = left + width + 10
                        width = boxWidth * 14
                        height = boxHeight * 10
                        repObj = Nothing
                        Call createRessPieOfProject(hproj, repObj, auswahl, top, left, height, width)
                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                End If


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True
            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True



    End Sub

    ''' <summary>
    ''' Charakteristik Personalkosten
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B3PKosten(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 2 ' steuert die Auswahl als Personalkosten
        Dim top As Double, left As Double, width As Double, height As Double

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 180

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try



                width = hproj.anzahlRasterElemente * boxWidth + 10

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Excel.ChartObject = Nothing

                Try
                    Call createRessBalkenOfProject(hproj, repObj, auswahl, top, left, height, width)

                    ' jetzt wird das Pie-Diagramm gezeichnet 
                    left = left + width + 10
                    width = boxWidth * 14
                    height = boxHeight * 10
                    repObj = Nothing
                    Call createRessPieOfProject(hproj, repObj, auswahl, top, left, height, width)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True



    End Sub

    ''' <summary>
    ''' Charakteristik Andere Kosten
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B4AKosten(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 1
        Dim top As Double, left As Double, width As Double, height As Double

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 180

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                width = hproj.anzahlRasterElemente * boxWidth + 10
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Excel.ChartObject = Nothing

                Call createCostBalkenOfProject(hproj, repObj, auswahl, top, left, height, width)

                ' jetzt wird das Pie-Diagramm gezeichnet 
                left = left + width + 10
                width = boxWidth * 14
                height = boxHeight * 10
                repObj = Nothing

                Try
                    Call createCostPieOfProject(hproj, repObj, auswahl, top, left, height, width)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' Charakteristik Gesamtkosten
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B5GKosten(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 2 ' das steuert , dass die Gesamtkosten angezeigt werden 
        Dim top As Double, left As Double, width As Double, height As Double

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 180

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                width = hproj.anzahlRasterElemente * boxWidth + 10

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Excel.ChartObject = Nothing


                Try
                    Call createCostBalkenOfProject(hproj, repObj, auswahl, top, left, height, width)
                    ' jetzt wird das Pie-Diagramm gezeichnet 
                    left = left + width + 10
                    width = boxWidth * 14
                    height = boxHeight * 10
                    repObj = Nothing
                    Call createCostPieOfProject(hproj, repObj, auswahl, top, left, height, width)
                Catch ex As Exception
                    Call MsgBox(ex.Message)
                End Try


                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub



    ''' <summary>
    ''' Charakteristik Strategie / Risiko 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B6SFIT(control As IRibbonControl)


        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next
            Dim obj As Excel.ChartObject = Nothing
            Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.FitRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    ''' <summary>
    ''' Charakteristik Strategie / Risiko / Abhängigkeiten
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B6SFITDEP(control As IRibbonControl)


        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next
            Dim obj As Excel.ChartObject = Nothing
            Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.FitRisikoDependency, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub


    Sub Tom2G2M1B6SFITVOl(control As IRibbonControl)

        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next
            Dim obj As Excel.ChartObject = Nothing

            Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.FitRisikoVol, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
            'Call awinCreateStratRiskVolumeDiagramm(myCollection, obj, True, False, True, True, top, left, width, height)
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

    End Sub

    Sub Tom2G2M1B6Abhaengigkeit(control As IRibbonControl)


        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim deleteList As New Collection
        Dim hproj As clsProjekt
        Dim pname As String

        Dim activeNumber As Integer             ' Kennzahl: auf wieviele Projekte strahlt es aus ?
        Dim passiveNumber As Integer            ' Kennzahl: von wievielen Projekten abhängig 
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name, .Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 12 * boxWidth
                        height = 8 * boxHeight

                    End If
                End With
            Next

            Dim i As Integer
            For i = 1 To myCollection.Count
                pname = CStr(myCollection.Item(i))
                Try
                    hproj = ShowProjekte.getProject(pname)
                    activeNumber = allDependencies.activeNumber(pname, PTdpndncyType.inhalt)
                    passiveNumber = allDependencies.passiveNumber(pname, PTdpndncyType.inhalt)
                    If activeNumber = 0 And passiveNumber = 0 Then
                        deleteList.Add(pname)
                    End If
                Catch ex As Exception

                End Try
            Next

            ' jetzt müssen die Projekte rausgenommen werden, die keine Abhängigkeiten haben 
            For i = 1 To deleteList.Count
                pname = CStr(deleteList.Item(i))
                Try
                    myCollection.Remove(pname)
                Catch ex As Exception

                End Try
            Next

            If myCollection.Count > 0 Then
                Dim obj As Excel.ChartObject = Nothing
                Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.Dependencies, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
            Else
                Call MsgBox("diese Projekte haben keine Abhängigkeiten")
            End If



        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU


    End Sub

    Sub Tom2G2M1B6CRisk(control As IRibbonControl)

        Dim top As Double, left As Double, width As Double, height As Double
        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim sichtbarerBereich As Excel.Range
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        myCollection.Add(.Name)
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                        width = 300
                        height = 280

                    End If
                End With
            Next

            If myCollection.Count > 1 Then

                With appInstance.ActiveWindow
                    sichtbarerBereich = .VisibleRange
                    left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 500) / 2
                    top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                End With

                width = 500
                height = 450
            End If

            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.ComplexRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
        appInstance.ScreenUpdating = formerSU

        ' die Projekte sollen hier doch nicht deselektiert werden, weil dadurch die awinNeuZeichnenDiagramm aufgerufen wird und damit auch die awinUpdatePortfolioDiagrams
        ' was dazu führt, dass alle Projekt in der Projektliste wieder in das Diagramm eingezeichnet werden.
        'Call awinDeSelect()



    End Sub


    ''' <summary>
    ''' zeigt den Soll-Ist Vergleich für das gewählte Projekt an 
    ''' Beauftragung / letzter Plan-Stand / aktueller Plan-Stand
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M2M1B2SollIstPKosten(control As IRibbonControl)

        Call projektTafelInit()
        ' auswahl steuert , dass die Personal-Kosten angezeigt werden 
        Dim auswahl As Integer = 1

        Dim vglBaseline As Boolean = True

        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)

    End Sub

    ''' <summary>
    ''' zeigt den Soll-Ist Vergleich für das gewählte Projekt an 
    ''' Beauftragung / letzter Plan-Stand / aktueller Plan-Stand
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M2M2B2SollIstAKosten(control As IRibbonControl)
        ' auswahl steuert , welche Kosten angezeigt werden
        Dim auswahl As Integer = 2
        Dim vglBaseline As Boolean = True
        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "

        Call projektTafelInit()

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)

    End Sub


    Sub Tom2G2M2M3B2SollIstGKosten(control As IRibbonControl)

        ' auswahl steuert , welche Kosten angezeigt werden
        Dim auswahl As Integer = 3
        Dim vglBaseline As Boolean = True

        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "

        Call projektTafelInit()

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)

    End Sub

    ''' <summary>
    ''' Fortschritts-Chart im Vergleich zur Beauftragung
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M4Fortschritt1(control As IRibbonControl)

        Call projektTafelInit()

        Call awinStatusAnzeige(1, 1, " ")

    End Sub

    ''' <summary>
    ''' Fortschritts-Chart im Vergleich zur letzten Planungs-Freigabe
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M4Fortschritt2(control As IRibbonControl)

        Call projektTafelInit()
        Call awinStatusAnzeige(2, 1, " ")

    End Sub



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="auswahl"></param>
    ''' <param name="typ"></param>
    ''' <remarks></remarks>
    Private Sub awinSollIstVergleich(ByVal auswahl As Integer, ByVal typ As String, ByVal vglBaseline As Boolean)
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, width As Double, height As Double
        Dim reportobj As Excel.ChartObject
        Dim heute As Date = Date.Now
        Dim vglName As String = " "
        Dim pName As String = ";"
        Dim variantName As String = ""

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                With singleShp
                    top = .Top + boxHeight + 5
                    left = .Left - 5
                End With
                height = 300
                width = 400

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)

                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.getShapeText
                    End If
                Else
                    projekthistorie = New clsProjektHistorie
                End If

                With hproj
                    pName = .name
                    variantName = .variantName
                End With

                If vglName <> hproj.getShapeText Then
                    If request.pingMongoDb() Then
                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, hproj)
                    Else
                        Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
                        projekthistorie.clear()
                    End If

                Else
                    ' der aktuelle Stand hproj muss hinzugefügt werden 
                    Dim lastElem As Integer = projekthistorie.Count - 1
                    projekthistorie.RemoveAt(lastElem)
                    projekthistorie.Add(Date.Now, hproj)
                End If

                Dim nrSnapshots As Integer = projekthistorie.Count

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                reportobj = Nothing

                Dim qualifier As String = " "

                Try
                    If typ = "Curve" Then
                        Call createSollIstCurveOfProject(hproj, reportobj, heute, auswahl, qualifier, vglBaseline, top, left, height, width)
                    Else
                        Call createSollIstOfProject(hproj, reportobj, heute, auswahl, qualifier, vglBaseline, top, left, height, width)
                    End If
                Catch ex As Exception

                End Try

                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True


    End Sub


    ''' <summary>
    ''' zeigt die Fortschrittsanzeige an 
    ''' </summary>
    ''' <param name="compareTyp">
    ''' 0=erste Eintrag in Projekt-Historie 
    ''' 1=Beauftragung
    ''' 2=letzte Freigabe
    ''' 3=letzter DB-Eintrag in Projekthistorie
    ''' </param>
    ''' <param name="auswahl">
    ''' 1=Personalkosten
    ''' 2=Sonstige Kosten
    ''' 3=Gesamtkosten
    ''' 3=rolle + Qualifier
    ''' 5=kostenart + qualifier
    ''' </param>
    ''' <param name="qualifier">
    ''' gibt an, um welche Rolle / Kostenart es sich handelt - falls auswahl = 4 oder 5
    ''' </param>
    ''' <remarks></remarks>
    Private Sub awinStatusAnzeige(ByVal compareTyp As Integer, ByVal auswahl As Integer, ByVal qualifier As String)
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, width As Double, height As Double
        Dim reportObj As Excel.ChartObject
        Dim heute As Date = Date.Now
        Dim vglName As String = " "
        Dim pName As String = ";"
        Dim variantName As String = ""
        Dim projektliste As New Collection
        Dim first As Boolean = True

        Call projektTafelInit()

        enableOnUpdate = False



        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try


        ' jetzt die Aktion durchführen für alle selektierten 

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection
                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        Try
                            hproj = ShowProjekte.getProject(.Name, True)
                            pName = hproj.name
                            If istLaufendesProjekt(hproj) Then

                                Try
                                    projektliste.Add(pName, pName)
                                Catch ex1 As Exception

                                End Try

                            End If

                        Catch ex As Exception

                        End Try

                        If first Then
                            top = .Top + boxHeight + 5
                            left = .Left - 5
                            first = False
                        End If

                    End If
                End With
            Next

            height = 300
            width = 400

            If projektliste.Count > 0 Then
                ' Diagramm erstellen 

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                reportObj = Nothing


                Dim tmpObj As Excel.ChartObject = Nothing
                Call awinCreateStatusDiagram1(projektliste, tmpObj, compareTyp, auswahl, qualifier, True, True, _
                                               top, left, width, height)

                reportObj = CType(tmpObj, Excel.ChartObject)
                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True


            End If
        Else
            Call MsgBox("es wurden keine laufenden Projekte selektert ...")
        End If


        ' Ende 


        enableOnUpdate = True


    End Sub



    Sub Tom2G2M5M2B5ShowMilestones(control As IRibbonControl)


        Dim farbTyp As Integer = 4
        Dim numberIt As Boolean = False
        Dim namelist As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Call awinZeichneMilestones(namelist, farbTyp, numberIt, False)

        enableOnUpdate = True
        appInstance.EnableEvents = True
        'appInstance.ScreenUpdating = formerSU


    End Sub
    '' '' '' '' '' '' ''' ur: 13.09.2016: wurde durch  ersetzt
    ' '' '' '' '' '' ''' <summary>
    ' '' '' '' '' '' ''' zeigt bei den ausgewählten Projekten die gewählten  erst eine Liste, aus der man die Namen auswählen kann 
    ' '' '' '' '' '' ''' zeigt dann alle Meilensteine, die zu dieser Liste gehören 
    ' '' '' '' '' '' ''' wenn Projekte selektiert sind: zeige nur die Meilensteine dieser Projekte an 
    ' '' '' '' '' '' ''' wenn nichts selektiert ist: Fehler MEldung 
    ' '' '' '' '' '' ''' </summary>
    ' '' '' '' '' '' ''' <param name="control"></param>
    ' '' '' '' '' '' ''' <remarks></remarks>
    ' '' '' '' '' ''Sub PTShowMilestonesByName(control As IRibbonControl)



    ' '' '' '' '' ''    Dim listOfItems As New Collection
    ' '' '' '' '' ''    Dim nameList As New Collection
    ' '' '' '' '' ''    Dim title As String = "Meilensteine visualisieren"

    ' '' '' '' '' ''    Dim repObj As Object = Nothing

    ' '' '' '' '' ''    Dim singleShp As Excel.Shape
    ' '' '' '' '' ''    Dim myCollection As New Collection
    ' '' '' '' '' ''    Dim hproj As clsProjekt
    ' '' '' '' '' ''    Dim awinSelection As Excel.ShapeRange
    ' '' '' '' '' ''    Dim selektierteProjekte As New clsProjekte

    ' '' '' '' '' ''    Call projektTafelInit()

    ' '' '' '' '' ''    Try
    ' '' '' '' '' ''        awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
    ' '' '' '' '' ''    Catch ex As Exception
    ' '' '' '' '' ''        awinSelection = Nothing
    ' '' '' '' '' ''    End Try

    ' '' '' '' '' ''    If Not awinSelection Is Nothing Then

    ' '' '' '' '' ''        ' jetzt die Aktion durchführen ...

    ' '' '' '' '' ''        For Each singleShp In awinSelection

    ' '' '' '' '' ''            Try
    ' '' '' '' '' ''                hproj = ShowProjekte.getProject(singleShp.Name, True)
    ' '' '' '' '' ''                selektierteProjekte.Add(hproj)
    ' '' '' '' '' ''            Catch ex As Exception
    ' '' '' '' '' ''                Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
    ' '' '' '' '' ''            End Try

    ' '' '' '' '' ''        Next

    ' '' '' '' '' ''        nameList = selektierteProjekte.getMilestoneNames

    ' '' '' '' '' ''        If nameList.Count > 0 Then

    ' '' '' '' '' ''            For Each tmpName As String In nameList
    ' '' '' '' '' ''                listOfItems.Add(tmpName)
    ' '' '' '' '' ''            Next

    ' '' '' '' '' ''            ' jetzt stehen in der listOfItems die Namen der Meilensteine - alphabetisch sortiert 
    ' '' '' '' '' ''            Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "andere löschen")


    ' '' '' '' '' ''            With auswahlFenster

    ' '' '' '' '' ''                .chTyp = DiagrammTypen(5)

    ' '' '' '' '' ''            End With
    ' '' '' '' '' ''            auswahlFenster.Show()

    ' '' '' '' '' ''        Else
    ' '' '' '' '' ''            Call MsgBox("keine Meilensteine in den selektierten Projekten vorhanden ..")
    ' '' '' '' '' ''        End If


    ' '' '' '' '' ''    Else
    ' '' '' '' '' ''        Call MsgBox("Bitte mindestens ein Projekt selektieren ... ")
    ' '' '' '' '' ''        Exit Sub
    ' '' '' '' '' ''    End If







    ' '' '' '' '' ''End Sub
    '' '' '' '' '' '' ''' ur: 13.09.2016: wurde durch ersetzt
    ' '' '' '' '' '' ''' <summary>
    ' '' '' '' '' '' ''' zeigt bei den ausgewählten Projekten die gewählten  erst eine Liste, aus der man die Namen auswählen kann 
    ' '' '' '' '' '' ''' zeigt dann alle Meilensteine, die zu dieser Liste gehören 
    ' '' '' '' '' '' ''' wenn Projekte selektiert sind: zeige nur die Meilensteine dieser Projekte an 
    ' '' '' '' '' '' ''' wenn nichts selektiert ist: zeige die Namen der Meilensteine aus allen Projekten  
    ' '' '' '' '' '' ''' </summary>
    ' '' '' '' '' '' ''' <param name="control"></param>
    ' '' '' '' '' '' ''' <remarks></remarks>
    ' '' '' '' '' ''Public Sub PTShowAllMilestonesByName(Control As IRibbonControl)

    ' '' '' '' '' ''    Dim listOfItems As New Collection
    ' '' '' '' '' ''    Dim nameList As New Collection
    ' '' '' '' '' ''    Dim title As String = "Meilensteine visualisieren"

    ' '' '' '' '' ''    Dim repObj As Object = Nothing

    ' '' '' '' '' ''    Call projektTafelInit()
    ' '' '' '' '' ''    Call awinDeSelect()

    ' '' '' '' '' ''    If ShowProjekte.Count > 0 Then
    ' '' '' '' '' ''        If showRangeRight - showRangeLeft > 5 Then

    ' '' '' '' '' ''            nameList = ShowProjekte.getMilestoneNames

    ' '' '' '' '' ''            If nameList.Count > 0 Then

    ' '' '' '' '' ''                For Each tmpName As String In nameList
    ' '' '' '' '' ''                    listOfItems.Add(tmpName)
    ' '' '' '' '' ''                Next

    ' '' '' '' '' ''                ' jetzt stehen in der listOfItems die Namen der Meilensteine - alphabetisch sortiert 
    ' '' '' '' '' ''                Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "andere löschen")


    ' '' '' '' '' ''                With auswahlFenster

    ' '' '' '' '' ''                    .chTyp = DiagrammTypen(5)

    ' '' '' '' '' ''                End With
    ' '' '' '' '' ''                auswahlFenster.Show()

    ' '' '' '' '' ''            Else
    ' '' '' '' '' ''                Call MsgBox("keine Meilensteine in den selektierten Projekten vorhanden ..")
    ' '' '' '' '' ''            End If
    ' '' '' '' '' ''        Else
    ' '' '' '' '' ''            Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
    ' '' '' '' '' ''        End If
    ' '' '' '' '' ''    Else
    ' '' '' '' '' ''        Call MsgBox("Es sind keine Projekte geladen!")
    ' '' '' '' '' ''    End If



    ' '' '' '' '' ''End Sub

    '' '' '' '' '' '' ''' ur: 10.7.2015: wurde durch awinShowMilestoneTrend ersetzt
    '' '' '' '' '' '' ''' <summary>
    '' '' '' '' '' '' ''' zeigt zu dem ausgewählten Projekt die Meilenstein Trendanalyse an 
    '' '' '' '' '' '' ''' dazu wird erst ein Fenster aufgeschaltet, aus dem der oder die Namen des betreffenden Meilensteins ausgewählt werden können 
    '' '' '' '' '' '' ''' </summary>
    '' '' '' '' '' '' ''' <param name="control"></param>
    '' '' '' '' '' '' ''' <remarks></remarks>
    '' '' '' '' '' ''Sub PTShowMilestoneTrend(control As IRibbonControl)

    '' '' '' '' '' ''    Dim request As New Request(awinSettings.databaseName, dbUsername, dbPasswort)
    '' '' '' '' '' ''    Dim singleShp As Excel.Shape
    '' '' '' '' '' ''    Dim listOfItems As New Collection
    '' '' '' '' '' ''    Dim listOfMSNames As New Collection
    '' '' '' '' '' ''    Dim nameList As New SortedList(Of Date, String)
    '' '' '' '' '' ''    Dim title As String = "Meilensteine auswählen"
    '' '' '' '' '' ''    Dim hproj As clsProjekt
    '' '' '' '' '' ''    Dim awinSelection As Excel.ShapeRange
    '' '' '' '' '' ''    Dim selektierteProjekte As New clsProjekte
    '' '' '' '' '' ''    Dim top As Double, left As Double, height As Double, width As Double
    '' '' '' '' '' ''    Dim repObj As Excel.ChartObject = Nothing

    '' '' '' '' '' ''    Dim pName As String, vglName As String = " "
    '' '' '' '' '' ''    Dim variantName As String

    '' '' '' '' '' ''    Call projektTafelInit()

    '' '' '' '' '' ''    Try
    '' '' '' '' '' ''        awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
    '' '' '' '' '' ''    Catch ex As Exception
    '' '' '' '' '' ''        awinSelection = Nothing
    '' '' '' '' '' ''    End Try
    '' '' '' '' '' ''    If request.pingMongoDb() Then

    '' '' '' '' '' ''        If Not awinSelection Is Nothing Then

    '' '' '' '' '' ''            ' eingangs-prüfung, ob auch nur ein Element selektiert wurde ...
    '' '' '' '' '' ''            If awinSelection.Count = 1 Then

    '' '' '' '' '' ''                ' Aktion durchführen ...

    '' '' '' '' '' ''                singleShp = awinSelection.Item(1)

    '' '' '' '' '' ''                Try
    '' '' '' '' '' ''                    hproj = ShowProjekte.getProject(singleShp.Name)
    '' '' '' '' '' ''                    nameList = hproj.getMilestones

    '' '' '' '' '' ''                    ' jetzt muss die ProjektHistorie aufgebaut werden 
    '' '' '' '' '' ''                    With hproj
    '' '' '' '' '' ''                        pName = .name
    '' '' '' '' '' ''                        variantName = .variantName
    '' '' '' '' '' ''                    End With

    '' '' '' '' '' ''                    If Not projekthistorie Is Nothing Then
    '' '' '' '' '' ''                        If projekthistorie.Count > 0 Then
    '' '' '' '' '' ''                            vglName = projekthistorie.First.getShapeText
    '' '' '' '' '' ''                        End If
    '' '' '' '' '' ''                    Else
    '' '' '' '' '' ''                        projekthistorie = New clsProjektHistorie
    '' '' '' '' '' ''                    End If

    '' '' '' '' '' ''                    If vglName <> hproj.getShapeText Then

    '' '' '' '' '' ''                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
    '' '' '' '' '' ''                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
    '' '' '' '' '' ''                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
    '' '' '' '' '' ''                        projekthistorie.Add(Date.Now, hproj)


    '' '' '' '' '' ''                    Else
    '' '' '' '' '' ''                        ' der aktuelle Stand hproj muss hinzugefügt werden 
    '' '' '' '' '' ''                        Dim lastElem As Integer = projekthistorie.Count - 1
    '' '' '' '' '' ''                        projekthistorie.RemoveAt(lastElem)
    '' '' '' '' '' ''                        projekthistorie.Add(Date.Now, hproj)
    '' '' '' '' '' ''                    End If

    '' '' '' '' '' ''                    If nameList.Count > 0 Then


    '' '' '' '' '' ''                        appInstance.EnableEvents = False
    '' '' '' '' '' ''                        enableOnUpdate = False

    '' '' '' '' '' ''                        repObj = Nothing



    '' '' '' '' '' ''                        For Each kvp As KeyValuePair(Of Date, String) In nameList

    '' '' '' '' '' ''                            Dim msname As String = ""
    '' '' '' '' '' ''                            msname = elemNameOfElemID(kvp.Value)
    '' '' '' '' '' ''                            listOfMSNames.Add(msname)
    '' '' '' '' '' ''                            listOfItems.Add(kvp.Value)
    '' '' '' '' '' ''                        Next

    '' '' '' '' '' ''                        With singleShp
    '' '' '' '' '' ''                            top = .Top + boxHeight + 5
    '' '' '' '' '' ''                            left = .Left - 5
    '' '' '' '' '' ''                        End With

    '' '' '' '' '' ''                        height = 2 * ((nameList.Count - 1) * 20 + 110)
    '' '' '' '' '' ''                        width = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 10, 24 * boxWidth + 10)

    '' '' '' '' '' ''                        'Try

    '' '' '' '' '' ''                        '    Call createMsTrendAnalysisOfProject(hproj, repObj, listOfItems, top, left, height, width)

    '' '' '' '' '' ''                        'Catch ex As Exception

    '' '' '' '' '' ''                        '    Call MsgBox(ex.Message)

    '' '' '' '' '' ''                        'End Try



    '' '' '' '' '' ''                        ' jetzt stehen in der listOfItems die Namen der Meilensteine - alphabetisch sortiert 
    '' '' '' '' '' ''                        Dim auswahlFenster As New ListSelectionWindow(listOfMSNames, title)


    '' '' '' '' '' ''                        With auswahlFenster

    '' '' '' '' '' ''                            .kennung = " "
    '' '' '' '' '' ''                            .chTyp = DiagrammTypen(6)
    '' '' '' '' '' ''                            .chTop = top
    '' '' '' '' '' ''                            .chLeft = left
    '' '' '' '' '' ''                            .chWidth = width
    '' '' '' '' '' ''                            .chHeight = height

    '' '' '' '' '' ''                        End With
    '' '' '' '' '' ''                        auswahlFenster.Show()

    '' '' '' '' '' ''                    Else
    '' '' '' '' '' ''                        Call MsgBox("keine Meilensteine in den selektierten Projekten vorhanden ..")
    '' '' '' '' '' ''                    End If

    '' '' '' '' '' ''                Catch ex As Exception
    '' '' '' '' '' ''                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
    '' '' '' '' '' ''                End Try

    '' '' '' '' '' ''            Else
    '' '' '' '' '' ''                Call MsgBox("bitte nur ein Projekt selektieren ...")
    '' '' '' '' '' ''            End If
    '' '' '' '' '' ''        Else
    '' '' '' '' '' ''            Call MsgBox("vorher ein Projekt selektieren ...")
    '' '' '' '' '' ''        End If

    '' '' '' '' '' ''    Else
    '' '' '' '' '' ''        Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
    '' '' '' '' '' ''        'projekthistorie.clear()
    '' '' '' '' '' ''    End If
    '' '' '' '' '' ''    enableOnUpdate = True
    '' '' '' '' '' ''    appInstance.EnableEvents = True





    '' '' '' '' '' ''End Sub

    Sub PT0ShowProjektStatus(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            For Each singleShp In awinSelection

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                    Call zeichneStatusSymbolInPlantafel(hproj, 0)
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                End Try



            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True





    End Sub

    ''' <summary>
    ''' zeigt die Abhängigkeiten der ausgewählten Projekte an ...
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT0ShowDependencies(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim myCollection As New Collection
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim atleastOne As Boolean = False

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            ' jetzt die Aktion durchführen ...

            ' erst noch alle Connectoren löschen ... 

            Call awinDeleteProjectChildShapes(4)

            For Each singleShp In awinSelection

                Try

                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                    Call zeichneDependenciesOfProject(hproj, PTdpndncyType.inhalt, 0)
                    atleastOne = True

                Catch ex As Exception
                    'Call MsgBox("Projekt " & singleShp.Name & " hat keine Abhängigkeiten")
                End Try



            Next

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True



    End Sub

    Sub Tom2G2M5B3NoShowSymbols(control As IRibbonControl)

        Call projektTafelInit()
        Call awinDeleteProjectChildShapes(0)
        Call deleteBeschriftungen()

        If visboZustaende.showTimeZoneBalken And showRangeLeft > 0 And showRangeRight > 0 Then
            Call awinShowtimezone(showRangeLeft, showRangeRight, True)
        End If
    End Sub


    ''' <summary>
    ''' löscht alle angezeigten Milestones
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M5B3NoShowMilestones(control As IRibbonControl)

        Call projektTafelInit()
        Call awinDeleteProjectChildShapes(1)

    End Sub

    ' '' ''ur: 13.09.2016: eliminiert, da nicht mehr benutzt

    ' '' ''Sub PT0VisualizePhases(control As IRibbonControl)

    ' '' ''    Dim i As Integer
    ' '' ''    Dim von As Integer, bis As Integer

    ' '' ''    Dim listOfItems As New Collection
    ' '' ''    Dim existingNames As New Collection

    ' '' ''    Dim title As String = "Phasen visualisieren"
    ' '' ''    Dim phaseName As String
    ' '' ''    Dim hproj As clsProjekt


    ' '' ''    Dim awinSelection As Excel.ShapeRange
    ' '' ''    Dim selektierteProjekte As New clsProjekte
    ' '' ''    Dim singleshp As Excel.Shape

    ' '' ''    Call projektTafelInit()

    ' '' ''    appInstance.EnableEvents = False
    ' '' ''    enableOnUpdate = False

    ' '' ''    Try
    ' '' ''        awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
    ' '' ''    Catch ex As Exception
    ' '' ''        awinSelection = Nothing
    ' '' ''    End Try


    ' '' ''    Dim anzElem As Integer = selektierteProjekte.Count

    ' '' ''    If Not awinSelection Is Nothing Then

    ' '' ''        ' jetzt die Aktion durchführen ...

    ' '' ''        For Each singleshp In awinSelection

    ' '' ''            Try
    ' '' ''                hproj = ShowProjekte.getProject(singleshp.Name, True)
    ' '' ''                selektierteProjekte.Add(hproj)
    ' '' ''            Catch ex As Exception
    ' '' ''                Call MsgBox("Projekt " & singleshp.Name & " nicht gefunden ...")
    ' '' ''            End Try

    ' '' ''        Next


    ' '' ''        existingNames = selektierteProjekte.getPhaseNames

    ' '' ''        If existingNames.Count > 0 Then

    ' '' ''            ' jetzt werden die Namen in der Reihenfolge, wie sie in der Phasen-Definition stehen in der listofItems eingetragen ..

    ' '' ''            For i = 1 To PhaseDefinitions.Count
    ' '' ''                phaseName = PhaseDefinitions.getPhaseDef(i).name

    ' '' ''                If existingNames.Contains(phaseName) Then
    ' '' ''                    listOfItems.Add(PhaseDefinitions.getPhaseDef(i).name)
    ' '' ''                End If

    ' '' ''            Next

    ' '' ''            ' jetzt stehen in der listOfItems die Namen der Phasen 
    ' '' ''            Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "andere löschen")

    ' '' ''            von = showRangeLeft
    ' '' ''            bis = showRangeRight
    ' '' ''            With auswahlFenster
    ' '' ''                .chTop = 50.0
    ' '' ''                .chLeft = (showRangeRight - 1) * boxWidth + 4
    ' '' ''                .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
    ' '' ''                .chHeight = awinSettings.ChartHoehe1
    ' '' ''                .chTyp = DiagrammTypen(0)

    ' '' ''            End With
    ' '' ''            auswahlFenster.Show()

    ' '' ''        Else
    ' '' ''            Call MsgBox("keine Phasen vorhanden ...")

    ' '' ''        End If



    ' '' ''    Else

    ' '' ''        Call MsgBox("bitte mindestens ein Projekt selektieren ...")

    ' '' ''    End If

    ' '' ''    enableOnUpdate = True
    ' '' ''    appInstance.EnableEvents = True



    ' '' ''End Sub

    ' '' ''ur: 13.09.2016: eliminiert, da nicht mehr benutzt


    ' '' ''Sub PT0VisualizePhasesAll(control As IRibbonControl)

    ' '' ''    Dim i As Integer
    ' '' ''    Dim von As Integer, bis As Integer

    ' '' ''    Dim listOfItems As New Collection
    ' '' ''    Dim existingNames As New Collection

    ' '' ''    Dim title As String = "Phasen visualisieren"
    ' '' ''    Dim phaseName As String

    ' '' ''    Call projektTafelInit()
    ' '' ''    Call awinDeSelect()

    ' '' ''    If ShowProjekte.Count > 0 Then

    ' '' ''        If showRangeRight - showRangeLeft > 5 Then


    ' '' ''            existingNames = ShowProjekte.getPhaseNames

    ' '' ''            ' jetzt werden die Namen in der Reihenfolge, wie sie in der Phasen-Definition stehen in der listofItems eingetragen ..

    ' '' ''            For i = 1 To PhaseDefinitions.Count
    ' '' ''                phaseName = PhaseDefinitions.getPhaseDef(i).name

    ' '' ''                If existingNames.Contains(phaseName) Then
    ' '' ''                    listOfItems.Add(PhaseDefinitions.getPhaseDef(i).name)
    ' '' ''                End If

    ' '' ''            Next

    ' '' ''            ' jetzt stehen in der listOfItems die Namen der Phasen 
    ' '' ''            Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "andere löschen")

    ' '' ''            von = showRangeLeft
    ' '' ''            bis = showRangeRight
    ' '' ''            With auswahlFenster
    ' '' ''                .chTop = 50.0
    ' '' ''                .chLeft = (showRangeRight - 1) * boxWidth + 4
    ' '' ''                .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
    ' '' ''                .chHeight = awinSettings.ChartHoehe1
    ' '' ''                .chTyp = DiagrammTypen(0)

    ' '' ''            End With
    ' '' ''            auswahlFenster.Show()
    ' '' ''        Else
    ' '' ''            Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
    ' '' ''        End If
    ' '' ''    Else
    ' '' ''        Call MsgBox("Es sind keine Projekte geladen!")
    ' '' ''    End If
    ' '' ''End Sub

    ' '' ''ur: 13.09.2016: eliminiert, da nicht mehr benutzt

    ' '' ''Sub PT0ShowPortfolioPhasen(control As IRibbonControl)

    ' '' ''    Dim i As Integer
    ' '' ''    Dim von As Integer, bis As Integer
    ' '' ''    'Dim myCollection As Collection
    ' '' ''    Dim listOfItems As New Collection
    ' '' ''    'Dim left As Double, top As Double, height As Double, width As Double

    ' '' ''    Dim phaseName As String

    ' '' ''    Call projektTafelInit()

    ' '' ''    If ShowProjekte.Count > 0 Then

    ' '' ''        If showRangeRight - showRangeLeft > 5 Then

    ' '' ''            For i = 1 To PhaseDefinitions.Count
    ' '' ''                phaseName = PhaseDefinitions.getPhaseDef(i).name
    ' '' ''                Try
    ' '' ''                    listOfItems.Add(phaseName, phaseName)
    ' '' ''                Catch ex As Exception

    ' '' ''                End Try

    ' '' ''            Next

    ' '' ''            ' jetzt stehen in der listOfItems die Namen der Rollen 
    ' '' ''            Dim auswahlFenster As New ListSelectionWindow(listOfItems, "Phasen auswählen", "pro Item ein Chart")

    ' '' ''            von = showRangeLeft
    ' '' ''            bis = showRangeRight
    ' '' ''            With auswahlFenster
    ' '' ''                .chTop = 50.0 + awinSettings.ChartHoehe1
    ' '' ''                .chLeft = (showRangeRight - 1) * boxWidth + 4
    ' '' ''                .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
    ' '' ''                .chHeight = awinSettings.ChartHoehe1
    ' '' ''                .chTyp = DiagrammTypen(0)
    ' '' ''            End With
    ' '' ''            auswahlFenster.Show()
    ' '' ''        Else
    ' '' ''            Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
    ' '' ''        End If
    ' '' ''    Else
    ' '' ''        Call MsgBox("Es sind noch keine Projekte geladen!")
    ' '' ''    End If
    ' '' ''End Sub

    ' '' ''ur: 13.09.2016: eliminiert, da nicht mehr benutzt


    ' '' ''Sub PTShowMilestoneSummen(control As IRibbonControl)

    ' '' ''    Dim von As Integer, bis As Integer

    ' '' ''    Dim listOfItems As New Collection

    ' '' ''    Dim nameList As New Collection

    ' '' ''    Call projektTafelInit()

    ' '' ''    If ShowProjekte.Count > 0 Then
    ' '' ''        If showRangeRight - showRangeLeft > 5 Then

    ' '' ''            nameList = ShowProjekte.getMilestoneNames

    ' '' ''            If nameList.Count > 0 Then

    ' '' ''                For Each tmpName As String In nameList
    ' '' ''                    listOfItems.Add(tmpName)
    ' '' ''                Next

    ' '' ''                ' jetzt stehen in der listOfItems die Namen der Rollen 
    ' '' ''                Dim auswahlFenster As New ListSelectionWindow(listOfItems, "Meilensteine auswählen", "pro Item ein Chart")

    ' '' ''                von = showRangeLeft
    ' '' ''                bis = showRangeRight
    ' '' ''                With auswahlFenster
    ' '' ''                    .kennung = "sum"
    ' '' ''                    .chTop = 50.0 + awinSettings.ChartHoehe1
    ' '' ''                    .chLeft = (showRangeRight - 1) * boxWidth + 4
    ' '' ''                    .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
    ' '' ''                    .chHeight = awinSettings.ChartHoehe1
    ' '' ''                    .chTyp = DiagrammTypen(5)
    ' '' ''                End With
    ' '' ''                auswahlFenster.Show()


    ' '' ''            Else
    ' '' ''                Call MsgBox("keine Meilensteine in den selektierten Projekten vorhanden ..")
    ' '' ''            End If
    ' '' ''        Else
    ' '' ''            Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
    ' '' ''        End If
    ' '' ''    Else
    ' '' ''        Call MsgBox("Es sind keine Projekte geladen!")
    ' '' ''    End If

    ' '' ''End Sub


    Sub PT0ShowAuslastung(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' Keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim obj As Excel.ChartObject = Nothing
        Dim myCollection As New Collection

        Call projektTafelInit()

        If showRangeLeft > 0 And (showRangeRight - showRangeLeft >= 1) Then
            appInstance.ScreenUpdating = False
            appInstance.EnableEvents = False
            enableOnUpdate = False


            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

            If myCollection.Count > 0 Then

                top = 180
                width = 340
                left = showRangeRight * boxWidth + 4
                If left < 0 Then
                    left = 4
                End If
                height = awinSettings.ChartHoehe2

                Try
                    Call awinCreateAuslastungsDiagramm(obj, top, left, width, height, False)

                    top = top + height + 10
                    Call createAuslastungsDetailPie(obj, 1, top, left, height, width, False)

                    ' jetzt Unterauslastung
                    top = top + height + 10
                    Call createAuslastungsDetailPie(obj, 2, top, left, height, width, False)

                Catch ex As Exception
                    Call MsgBox("keine Information vorhanden")
                End Try

            Else

                If ShowProjekte.Count = 0 Then
                    Call MsgBox("es sind keine Projekte angezeigt")

                Else
                    If showRangeRight - showRangeLeft < minColumns - 1 Then
                        If awinSettings.englishLanguage Then
                            Call MsgBox("please define a timeframe first ...")
                        Else
                            Call MsgBox("bitte wählen Sie zuerst einen Zeitraum aus ...")
                        End If
                    Else
                        Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                                    "gibt es keine Projekte ")
                    End If
                End If

            End If



            appInstance.ScreenUpdating = True
            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please define a timeframe first ...")
            Else
                Call MsgBox("bitte wählen Sie zuerst einen Zeitraum aus ...")
            End If
        End If


        
    End Sub

    Sub PTXShowEngpass(control As IRibbonControl)

        Dim i As Integer
        Dim von As Integer, bis As Integer
        Dim myCollection As New Collection
        Dim listOfItems As New Collection
        Dim left As Double, top As Double, height As Double, width As Double
        Dim roleName As String
        Dim engpass As String = ""
        Dim engpassValue As Double = -100000.0
        Dim curValue As Double

        Dim repObj As Excel.ChartObject = Nothing

        Call projektTafelInit()

        'appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False
        enableOnUpdate = False

        If (showRangeRight - showRangeLeft) >= minColumns - 1 Then

            If ShowProjekte.Count > 0 Then

                For i = 1 To RoleDefinitions.Count
                    roleName = RoleDefinitions.getRoledef(i).name
                    With ShowProjekte
                        curValue = .getAuslastungsValues(roleName, 1).Sum - .getAuslastungsValues(roleName, 2).Sum
                        If curValue > engpassValue Then
                            engpassValue = curValue
                            engpass = roleName
                        End If
                    End With
                Next

                If engpass <> "" Then
                    myCollection.Add(engpass, engpass)
                    von = showRangeLeft
                    bis = showRangeRight

                    height = awinSettings.ChartHoehe1
                    top = 180

                    If von > 1 Then
                        left = showRangeRight * boxWidth + 4
                    Else
                        left = 0
                    End If

                    Dim breite As Integer = System.Math.Max(bis - von, 6)
                    
                    width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct

                    Call awinCreateprcCollectionDiagram(myCollection, repObj, top, left, width, height, False, DiagrammTypen(1), False)

                Else
                    Call MsgBox("kein Engpass gefunden")
                End If
            Else
                Call MsgBox("Es sind keine Projekte geladen! ")
            End If
        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please define a timeframe first ...")
            Else
                Call MsgBox("Bitte wählen Sie zuerst einen Zeitraum aus ...")
            End If

        End If

        'appInstance.ScreenUpdating = True
        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

    ' '' ''ur: 13.09.2016: eliminiert, da nicht mehr benutzt


    ' '' ''Sub PT0ShowPersonalBedarfe(control As IRibbonControl)

    ' '' ''    Dim i As Integer
    ' '' ''    Dim von As Integer, bis As Integer
    ' '' ''    'Dim myCollection As Collection
    ' '' ''    Dim listOfItems As New Collection
    ' '' ''    'Dim left As Double, top As Double, height As Double, width As Double

    ' '' ''    Dim repObj As Object = Nothing
    ' '' ''    Dim title As String = "Rollen auswählen"

    ' '' ''    Call projektTafelInit()

    ' '' ''    'appInstance.ScreenUpdating = False
    ' '' ''    'appInstance.EnableEvents = False
    ' '' ''    'enableOnUpdate = False

    ' '' ''    If ShowProjekte.Count > 0 Then

    ' '' ''        If showRangeRight - showRangeLeft > 5 Then


    ' '' ''            For i = 1 To RoleDefinitions.Count
    ' '' ''                listOfItems.Add(RoleDefinitions.getRoledef(i).name)
    ' '' ''            Next

    ' '' ''            ' jetzt stehen in der listOfItems die Namen der Rollen 
    ' '' ''            Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "pro Item ein Chart")

    ' '' ''            von = showRangeLeft
    ' '' ''            bis = showRangeRight
    ' '' ''            With auswahlFenster
    ' '' ''                .chTop = 100.0 + awinSettings.ChartHoehe1
    ' '' ''                .chLeft = ((von - 1) / 3 - 1) * 3 * boxWidth + 32.8 + von * screen_correct
    ' '' ''                .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
    ' '' ''                .chHeight = awinSettings.ChartHoehe1
    ' '' ''                .chTyp = DiagrammTypen(1)
    ' '' ''            End With

    ' '' ''            auswahlFenster.Show()
    ' '' ''        Else
    ' '' ''            Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
    ' '' ''        End If
    ' '' ''    Else
    ' '' ''        Call MsgBox("Es sind noch keine Projekte geladen!")
    ' '' ''    End If

    ' '' ''    'appInstance.ScreenUpdating = True
    ' '' ''    'appInstance.EnableEvents = True
    ' '' ''    'enableOnUpdate = True

    ' '' ''End Sub

    ' '' ''ur: 13.09.2016: eliminiert, da nicht mehr benutzt


    ' '' ''Sub PT0ShowKostenBedarfe(control As IRibbonControl)

    ' '' ''    Dim i As Integer
    ' '' ''    Dim von As Integer, bis As Integer
    ' '' ''    'Dim myCollection As Collection
    ' '' ''    Dim listOfItems As New Collection
    ' '' ''    'Dim left As Double, top As Double, height As Double, width As Double
    ' '' ''    Dim repObj As Object = Nothing
    ' '' ''    Dim title As String = "Kostenarten auswählen"

    ' '' ''    Call projektTafelInit()

    ' '' ''    'appInstance.EnableEvents = False
    ' '' ''    'enableOnUpdate = False
    ' '' ''    If ShowProjekte.Count > 0 Then

    ' '' ''        If showRangeRight - showRangeLeft > 5 Then

    ' '' ''            For i = 1 To CostDefinitions.Count
    ' '' ''                listOfItems.Add(CostDefinitions.getCostdef(i).name)
    ' '' ''            Next

    ' '' ''            ' jetzt stehen in der listOfItems die Namen der Rollen 
    ' '' ''            'Dim auswahlFenster As New ListSelectionWindow(listOfItems, title)
    ' '' ''            Dim auswahlFenster As New ListSelectionWindow(listOfItems, title, "pro Item ein Chart")


    ' '' ''            von = showRangeLeft
    ' '' ''            bis = showRangeRight
    ' '' ''            With auswahlFenster
    ' '' ''                .chTop = 50.0
    ' '' ''                .chLeft = (showRangeRight - 1) * boxWidth + 4
    ' '' ''                .chWidth = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct
    ' '' ''                .chHeight = awinSettings.ChartHoehe1
    ' '' ''                .chTyp = DiagrammTypen(2)

    ' '' ''            End With

    ' '' ''            auswahlFenster.Show()


    ' '' ''        Else
    ' '' ''            Call MsgBox("Bitte wählen Sie einen Zeitraum von mindestens 6 Monaten aus!")
    ' '' ''        End If
    ' '' ''    Else
    ' '' ''        Call MsgBox("Es sind noch keine Projekte geladen!")
    ' '' ''    End If

    ' '' ''    'appInstance.EnableEvents = True
    ' '' ''    'enableOnUpdate = True

    ' '' ''End Sub

    Sub PT0ShowZieleUebersicht(control As IRibbonControl)

        Dim ControlID As String = control.Id
        Dim relevanteProjekte As clsProjekte
        Dim chtObject As Excel.ChartObject = Nothing
        'Dim top As Double, left As Double, width As Double, height As Double
        Dim future As Integer = 0
        Dim formerAmpelSetting As Boolean = awinSettings.mppShowAmpel
        awinSettings.mppShowAmpel = True


        Dim myCollection As New Collection
        myCollection.Add("Ziele")

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False
        If ControlID = "PT0G1B2" Then
            relevanteProjekte = selectedProjekte
        Else
            Call awinDeSelect() ' evt. vorhandene Selektion entfernen, da über Multiprojekt-Info
            relevanteProjekte = ShowProjekte
        End If

        If relevanteProjekte.Count > 0 Then
            If showRangeRight - showRangeLeft >= minColumns - 1 Then

                ' betrachte sowohl Vergangenheit als auch Gegenwart
                future = 0

                Dim wpfInput As New Dictionary(Of String, clsWPFPieValues)
                Dim valueItem As New clsWPFPieValues

                ' Nicht bewertet 
                With valueItem
                    .value = relevanteProjekte.getColorsInMonth(0, future).Sum
                    .name = "nicht bewertet"
                    .color = CType(awinSettings.AmpelNichtBewertet, UInt32)
                End With
                wpfInput.Add(valueItem.name, valueItem)

                valueItem = New clsWPFPieValues
                ' Grün bewertet
                With valueItem
                    .value = relevanteProjekte.getColorsInMonth(1, future).Sum
                    .name = "OK"
                    .color = CType(awinSettings.AmpelGruen, UInt32)
                End With
                wpfInput.Add(valueItem.name, valueItem)

                valueItem = New clsWPFPieValues
                ' Gelb bewertet
                With valueItem
                    .value = relevanteProjekte.getColorsInMonth(2, future).Sum
                    .name = "nicht vollständig"
                    .color = CType(awinSettings.AmpelGelb, UInt32)
                End With
                wpfInput.Add(valueItem.name, valueItem)

                valueItem = New clsWPFPieValues
                ' Rot bewertet
                With valueItem
                    .value = relevanteProjekte.getColorsInMonth(3, future).Sum
                    .name = "Zielverfehlung"
                    .color = CType(awinSettings.AmpelRot, UInt32)
                End With
                wpfInput.Add(valueItem.name, valueItem)


                Dim pieChartZieleV As New PieChartWindow(wpfInput)

                With pieChartZieleV
                    .Title = "Ziele-Erreichung " & textZeitraum(showRangeLeft, showRangeRight)
                    '.Top = frmCoord(PTfrm.ziele, PTpinfo.top)
                    '.Left = frmCoord(PTfrm.ziele, PTpinfo.left)
                End With

                pieChartZieleV.Show()
            Else
                Call MsgBox("Bitte wählen Sie einen Zeitraum aus!")
            End If

        Else
            If ControlID = "PT0G1B2" Then
                Call MsgBox("Bitte zuerst ein Projekt selektieren! ")
            Else
                Call MsgBox("Es sind keine Projekte geladen!")
            End If

        End If

        'awinSettings.mppShowAmpel = formerAmpelSetting

        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub



    Sub PT0ShowStrategieRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()


        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then

            appInstance.EnableEvents = False
            enableOnUpdate = False

            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

            If myCollection.Count > 0 Then

                With appInstance.ActiveWindow
                    sichtbarerBereich = .VisibleRange
                    left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                    If left < CDbl(sichtbarerBereich.Left) Then
                        left = CDbl(sichtbarerBereich.Left) + 2
                    End If

                    top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                    If top < CDbl(sichtbarerBereich.Top) Then
                        top = CDbl(sichtbarerBereich.Top) + 2
                    End If

                End With

                width = 600
                height = 450

                Dim obj As Excel.ChartObject = Nothing

                Try
                    Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
                Catch ex As Exception

                End Try

            Else

                If ShowProjekte.Count = 0 Then
                    Call MsgBox("es sind keine Projekte angezeigt")

                Else
                    Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                                "gibt es keine Projekte")
                End If


            End If



            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please define a timeframe first ...")
            Else
                Call MsgBox("bitte wählen Sie zuerst einen Zeitraum aus ...")
            End If
        End If
        

        

    End Sub

    ''' <summary>
    ''' zeigt das Portfolio Chart Strategie, Risiko, Abhängigkeiten an 
    ''' die Größe der Kugel entspricht der Anzahl der Abhängigkeiten, also wieviele Projekte sind von diesem Projekt abhängig  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT0ShowSFitRisikoDependency(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoDependency, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If


        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True


    End Sub

    Sub PT0ShowStratRisikoVolume(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoVol, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
                'Call awinCreateStratRiskVolumeDiagramm(myCollection, obj, False, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If


        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True

    End Sub

    ''' <summary>
    ''' zeigt zwei Windows an, bestehend aus der Massen-Edit Ressourcen Tabelle und der meCharts Tabelle   
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTMEShowCharts(control As IRibbonControl)


        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Dim currentRow As Integer
        Dim currentColumn As Integer
        Dim prcTyp As String

        Try
            currentRow = appInstance.ActiveCell.Row
            currentColumn = appInstance.ActiveCell.Column
        Catch ex As Exception
            currentRow = 2
            currentColumn = visboZustaende.meColRC
        End Try

        Dim rcName As String = CStr(CType(appInstance.ActiveSheet, Excel.Worksheet).Cells(currentRow, visboZustaende.meColRC).value)
        If IsNothing(rcName) Then
            rcName = ""
        End If

        Do While rcName = "" And currentRow <= visboZustaende.meMaxZeile
            currentRow = currentRow + 1
            rcName = CStr(CType(appInstance.ActiveSheet, Excel.Worksheet).Cells(currentRow, visboZustaende.meColRC).value)
            If IsNothing(rcName) Then
                rcName = ""
            End If
        Loop

        ' jetzt ist entweder was gefunden oder es ist komplett ohne Werte 
        If rcName = "" Then
            currentRow = 2
            Try
                prcTyp = DiagrammTypen(1)
                rcName = RoleDefinitions.getRoledef(1).name
            Catch ex As Exception
                prcTyp = DiagrammTypen(1)
                rcName = ""
            End Try

        Else
            If RoleDefinitions.containsName(rcName) Then
                prcTyp = DiagrammTypen(1)
            ElseIf CostDefinitions.containsName(rcName) Then
                prcTyp = DiagrammTypen(2)
            Else
                prcTyp = DiagrammTypen(1)
                rcName = RoleDefinitions.getRoledef(1).name
            End If

        End If


        Dim buildcustomView As Boolean = True
        Dim viewName As String = viewNames(1)
        Dim visboWorkbook As Excel.Workbook = appInstance.Workbooks.Item(myProjektTafel)


        projectboardWindows(0) = appInstance.ActiveWindow

        ' Aus dem aktuellen Window ein benanntes Window machen 
        projectboardWindows(1) = appInstance.ActiveWindow.NewWindow
        With projectboardWindows(1)
            .WindowState = Excel.XlWindowState.xlNormal
            .EnableResize = True
            .SplitRow = 1
            .FreezePanes = True
            .DisplayFormulas = False
            .DisplayHeadings = False
            .DisplayGridlines = True
            .GridlineColor = RGB(220, 220, 220)
            .DisplayWorkbookTabs = False
            .Caption = windowNames(1)
        End With

        ' Aufbau des Windows windowNames(4): Charts
        projectboardWindows(2) = appInstance.ActiveWindow.NewWindow
        visboWorkbook.Worksheets.Item(arrWsNames(ptTables.meCharts)).activate()
        With projectboardWindows(2)
            .WindowState = Excel.XlWindowState.xlNormal
            .EnableResize = True
            .DisplayGridlines = False
            .DisplayHeadings = False
            .DisplayRuler = False
            .DisplayVerticalScrollBar = True
            .DisplayHorizontalScrollBar = True
            .DisplayWorkbookTabs = False
            .Caption = windowNames(4)
        End With

        ' jetzt das Ursprungs-Window ausblenden ...
        For Each tmpWindow As Excel.Window In visboWorkbook.Windows
            If (CStr(tmpWindow.Caption) <> windowNames(4)) And (CStr(tmpWindow.Caption) <> windowNames(1)) Then
                tmpWindow.Visible = False
            End If
        Next

        ' jetzt die verbleibenden arrangieren ...
        visboWorkbook.Windows.Arrange(Excel.XlArrangeStyle.xlArrangeStyleHorizontal)

        ' jetzt die Größen anpassen 
        With projectboardWindows(1)
            .Top = 0
            .Height = 3 / 4 * maxScreenHeight
        End With

        ' jetzt die Größen anpassen 
        With projectboardWindows(2)
            .Top = 3 / 4 * maxScreenHeight + 3
            .Height = 1 / 4 * maxScreenHeight - 3
        End With

        ' Check: was ist das aktuelle Sheet 
        'Dim checkSheet As Object = projectboardWindows(1).ActiveSheet

        ' jetzt das Mass-Edit Window aktivieren 
        projectboardWindows(1).Activate()
        With CType(projectboardWindows(1).ActiveSheet, Excel.Worksheet)
            CType(.Cells(currentRow, currentColumn), Excel.Range).Activate()
        End With

        Dim anz As Integer = appInstance.ActiveWorkbook.Windows.Count

        ' jetzt werden die Charts ggf erzeugt ...  
        If CType(CType(projectboardWindows(2).ActiveSheet, Excel.Worksheet).ChartObjects, Excel.ChartObjects).Count = 0 Then
            ' sie müssen erzeugt werden

            ' erst das PRCCollectionChart ...
            Dim repObj As Excel.ChartObject = Nothing
            Dim chWidth As Double = projectboardWindows(2).UsableWidth / 4 - 2
            Dim chHeight As Double = projectboardWindows(2).UsableHeight - 10
            Dim chTop As Double = 5
            Dim chLeft As Double = 2 * chWidth
            Dim myCollection As New Collection

            myCollection.Add(rcName)

            Call awinCreateprcCollectionDiagram(myCollection, repObj, chTop, chLeft,
                                                                   chWidth, chHeight, False, prcTyp, True)

            ' jetzt das Portfolio Chart Budget anzeigen ... 
            'Dim obj As Excel.ChartObject = Nothing
            'chLeft = 3 * chWidth
            'Call awinCreateBudgetErgebnisDiagramm(obj, chTop, chLeft, chWidth, chHeight, False, True)

        Else
            ' sie sind schon da 

        End If



        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True

    End Sub

    Sub PT0ShowAbhaengigkeiten(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range
        Dim deleteList As New Collection
        Dim hproj As clsProjekt
        Dim pname As String

        Dim activeNumber As Integer             ' Kennzahl: auf wieviele Projekte strahlt es aus ?
        Dim passiveNumber As Integer            ' Kennzahl: von wievielen Projekten abhängig 

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then


            Dim i As Integer
            For i = 1 To myCollection.Count
                pname = CStr(myCollection.Item(i))
                Try
                    hproj = ShowProjekte.getProject(pname)
                    activeNumber = allDependencies.activeNumber(pname, PTdpndncyType.inhalt)
                    passiveNumber = allDependencies.passiveNumber(pname, PTdpndncyType.inhalt)
                    If activeNumber = 0 And passiveNumber = 0 Then
                        deleteList.Add(pname)
                    End If
                Catch ex As Exception

                End Try
            Next

            ' jetzt müssen die Projekte rausgenommen werden, die keine Abhängigkeiten haben 
            For i = 1 To deleteList.Count
                pname = CStr(deleteList.Item(i))
                Try
                    myCollection.Remove(pname)
                Catch ex As Exception

                End Try
            Next


            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                If myCollection.Count > 0 Then
                    Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.Dependencies, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
                Else
                    Call MsgBox(" es gibt in diesem Zeitraum keine Projekte mit Abhängigkeiten")
                End If


            Catch ex As Exception

            End Try

        Else
            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If
        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub


    ''' <summary>
    ''' zeigt an , welche Projekte Management Attention verdienen/benötigen, weil sie besser/schlechter als der letzte Stand geplant laufen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT0ShowAttentionL(control As IRibbonControl)

        Dim selectionType As Integer
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range
        Dim deleteList As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' hier muss noch geklärt werden, welche Projekte betrachtet werden; es mcht keinen Sinn, 
        'das nur an den TimeFrame zu koppeln, es geht im wesentlichen um aktuell laufende und vergangene Projekte 
        ' Frage : was ist mit bereits beauftragten Projekten, die noch gar nicht begonnen haben, deren Planung aber bereits schlechter als beauftragt ist ? 

        selectionType = PTpsel.lfundab
        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow

                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2

                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                If myCollection.Count > 0 Then

                    Try
                        Call awinCreateBetterWorsePortfolio(ProjektListe:=myCollection, repChart:=obj, showAbsoluteDiff:=True, isTimeTimeVgl:=False, vglTyp:=1, _
                                                        charttype:=PTpfdk.betterWorseL, bubbleColor:=0, bubbleValueTyp:=PTbubble.strategicFit, showLabels:=True, chartBorderVisible:=True, _
                                                        top:=top, left:=left, width:=width, height:=height)
                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                Else
                    Call MsgBox(" es gibt in diesem Zeitraum keine laufenden / abgeschlossenen Projekte")
                End If


            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub

    Sub PT0ShowAttentionB(control As IRibbonControl)

        Dim selectionType As Integer
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range
        Dim deleteList As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' hier muss noch geklärt werden, welche Projekte betrachtet werden; es mcht keinen Sinn, 
        'das nur an den TimeFrame zu koppeln, es geht im wesentlichen um aktuell laufende und vergangene Projekte 
        ' Frage : was ist mit bereits beauftragten Projekten, die noch gar nicht begonnen haben, deren Planung aber bereits schlechter als beauftragt ist ? 

        selectionType = PTpsel.lfundab
        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450

            Dim obj As Excel.ChartObject = Nothing

            Try
                If myCollection.Count > 0 Then

                    Try
                        Call awinCreateBetterWorsePortfolio(ProjektListe:=myCollection, repChart:=obj, showAbsoluteDiff:=True, isTimeTimeVgl:=False, vglTyp:=1, _
                                                        charttype:=PTpfdk.betterWorseB, bubbleColor:=0, bubbleValueTyp:=PTbubble.strategicFit, showLabels:=True, chartBorderVisible:=True, _
                                                        top:=top, left:=left, width:=width, height:=height)
                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                Else
                    Call MsgBox(" es gibt in diesem Zeitraum keine laufenden bzw. abgeschlossenen Projekte")
                End If


            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If


        appInstance.EnableEvents = True
        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub


    Sub PT0ShowComplexRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)


        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450


            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.ComplexRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If


        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True

        Call awinDeSelect()

    End Sub

    Sub PT0ShowZeitRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        If myCollection.Count > 0 Then

            With appInstance.ActiveWindow
                sichtbarerBereich = .VisibleRange
                left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - 600) / 2
                If left < CDbl(sichtbarerBereich.Left) Then
                    left = CDbl(sichtbarerBereich.Left) + 2
                End If

                top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - 450) / 2
                If top < CDbl(sichtbarerBereich.Top) Then
                    top = CDbl(sichtbarerBereich.Top) + 2
                End If

            End With

            width = 600
            height = 450


            Dim obj As Excel.ChartObject = Nothing

            Try
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.ZeitRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte")
            End If

        End If



        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True

        Call awinDeSelect()

    End Sub

    Sub PTOPTVariantenOptimieren(control As IRibbonControl)


        Dim optimierungsFenster As New frmOptimizeKPI
        Dim returnValue As DialogResult


        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        ' Varianten-Optimierung 
        optimierungsFenster.menueOption = 1

        returnValue = optimierungsFenster.ShowDialog
        'optmierungsFenster.Show()

        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

    Sub PTOPTFreiraumOptimieren(control As IRibbonControl)

        Dim optimierungsFenster As New frmOptimizeKPI
        Dim returnValue As DialogResult


        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        ' Spielraum-Optimierung 
        optimierungsFenster.menueOption = 2

        returnValue = optimierungsFenster.ShowDialog
        'optmierungsFenster.Show()

        appInstance.EnableEvents = True
        enableOnUpdate = True
    End Sub

    Sub PT0ShowPortfolioBudgetCost(control As IRibbonControl)
        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim myCollection As New Collection

        Call projektTafelInit()

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then

            appInstance.EnableEvents = False
            enableOnUpdate = False

            Dim formerES As Boolean = awinSettings.meEnableSorting

            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

            If myCollection.Count > 0 Then

                Dim sichtbarerBereich As Excel.Range

                height = awinSettings.ChartHoehe2
                width = 450

                With appInstance.ActiveWindow

                    sichtbarerBereich = .VisibleRange
                    If visboZustaende.projectBoardMode = ptModus.graficboard Then
                        left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - width) / 2
                        If left < CDbl(sichtbarerBereich.Left) Then
                            left = CDbl(sichtbarerBereich.Left) + 2
                        End If
                    Else
                        left = 5
                    End If



                    top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - height) / 2
                    If top < CDbl(sichtbarerBereich.Top) Then
                        top = CDbl(sichtbarerBereich.Top) + 2
                    End If

                End With

                Dim obj As Excel.ChartObject = Nothing
                Call awinCreateBudgetErgebnisDiagramm(obj, top, left, width, height, False, False)

            Else

                If ShowProjekte.Count = 0 Then
                    If awinSettings.englishLanguage Then
                        Call MsgBox("no projects visualized ...")
                    Else
                        Call MsgBox("es sind keine Projekte angezeigt ...")
                    End If


                Else
                    If showRangeRight - showRangeLeft < minColumns - 1 Then
                        If awinSettings.englishLanguage Then
                            Call MsgBox("please define a timeframe first ...")
                        Else
                            Call MsgBox("bitte wählen Sie zuerst einen Zeitraum aus ...")
                        End If
                    Else
                        If awinSettings.englishLanguage Then
                            Call MsgBox("there are no projects in Timeframe " & textZeitraum(showRangeLeft, showRangeRight))
                        Else
                            Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                                    "gibt es keine Projekte ")
                        End If

                    End If
                End If

            End If

            If control.Id = "PTMEC2" And awinSettings.meEnableSorting <> formerES Then
                Me.ribbon.Invalidate()
            End If

            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please define a timeframe first ...")
            Else
                Call MsgBox("bitte wählen Sie zuerst einen Zeitraum aus ...")
            End If
        End If


    End Sub


    Sub PT0ShowPortfolioErgebnis(control As IRibbonControl)
        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim myCollection As New Collection

        Call projektTafelInit()

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            appInstance.EnableEvents = False
            enableOnUpdate = False

            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

            If myCollection.Count > 0 Then

                Dim sichtbarerBereich As Excel.Range

                height = awinSettings.ChartHoehe2
                width = 450

                With appInstance.ActiveWindow
                    sichtbarerBereich = .VisibleRange
                    left = CDbl(sichtbarerBereich.Left) + (CDbl(sichtbarerBereich.Width) - width) / 2
                    If left < CDbl(sichtbarerBereich.Left) Then
                        left = CDbl(sichtbarerBereich.Left) + 2
                    End If

                    top = CDbl(sichtbarerBereich.Top) + (CDbl(sichtbarerBereich.Height) - height) / 2
                    If top < CDbl(sichtbarerBereich.Top) Then
                        top = CDbl(sichtbarerBereich.Top) + 2
                    End If

                End With



                Dim obj As Excel.ChartObject = Nothing
                Call awinCreateErgebnisDiagramm(obj, top, left, width, height, False, False)


            Else

                If awinSettings.englishLanguage Then
                    Call MsgBox("there are no projects in Timeframe " & textZeitraum(showRangeLeft, showRangeRight))
                Else
                    Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                            "gibt es keine Projekte ")
                End If



            End If

            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("there are no projects in Timeframe " & textZeitraum(showRangeLeft, showRangeRight))
            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf & _
                        "gibt es keine Projekte ")
            End If
        End If


    End Sub



    Sub Tom2G2M5M1B3ShowStatus(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim heute As Integer = getColumnOfDate(Date.Now)
        Dim myCollection As New Collection

        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        myCollection = ShowProjekte.withinTimeFrame(selectionType, heute, heute)

        If myCollection.Count > 0 Then

            Dim nummerieren As Boolean = False
            Call awinZeichneStatus(nummerieren)

        Else

            Call MsgBox("es gibt keine aktuell laufenden Projekte")

        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub


    ''' <summary>
    ''' Charakteristik Projekt-Ergebnis
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B7Ergebnis(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False


        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                Dim formerSU As Boolean = appInstance.ScreenUpdating
                Dim formerEE As Boolean = appInstance.EnableEvents
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                Dim dummyObj As Excel.ChartObject = Nothing
                Dim hproj As clsProjekt

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)

                    Try

                        Call createProjektErgebnisCharakteristik2(hproj, dummyObj, PThis.current)
                    Catch ex1 As Exception
                        Call MsgBox("Fehler bei Diagramm erzeugen: " & ex1.Message)
                    End Try

                Catch ex As Exception

                    Call MsgBox("Name nicht gefunden : " & singleShp.Name)

                End Try



                appInstance.ScreenUpdating = formerSU
                appInstance.EnableEvents = formerEE
            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True



    End Sub

    ''' <summary>
    ''' Vergleichen mit Beauftragung / Freigabe
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M2B1Auftrag(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)
                Dim hproj As clsProjekt

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                    Dim cproj As New clsProjekt
                    Dim top As Double = singleShp.Top + boxHeight + 2
                    Dim left As Double = singleShp.Left - boxWidth
                    If left <= 0 Then
                        left = 1
                    End If
                    Call awinCompareProject(hproj, cproj, 0, top, left)

                Catch ex As Exception
                    Call MsgBox("Fehler bei Beauftragung " & vbLf & ex.Message)
                End Try


                'Call awinCompareProject(pname1:=singleShp.Name, pname2:=" ", compareType:=0)

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' die Phasen zweier Projekte vergleichen  - Darstellung in einem Chart
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G3M1B1PhasenVgl(control As IRibbonControl)

        Dim singleShp1 As Excel.Shape, singleShp2 As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 2 Then
                ' jetzt die Aktion durchführen ...
                singleShp1 = awinSelection.Item(1)
                singleShp2 = awinSelection.Item(2)

                Try
                    hproj = ShowProjekte.getProject(singleShp1.Name, True)
                    cproj = ShowProjekte.getProject(singleShp2.Name, True)
                Catch ex As Exception
                    Call MsgBox("Projekt nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try


                top = singleShp1.Top + boxHeight + 2
                left = singleShp1.Left - 5
                If left <= 0 Then
                    left = 1
                End If

                height = 380

                width = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 7, cproj.anzahlRasterElemente * boxWidth + 7)
                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)
                'width = hproj1.Dauer * boxWidth + 7
                'scale = hproj1.Dauer

                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False

                repObj = Nothing
                Dim htitel As String = hproj.name
                Dim ctitel As String = cproj.name
                Call awinCompareProjectPhases(hproj, htitel, cproj, ctitel, 3, repObj)


                appInstance.ScreenUpdating = True

            Else
                Call MsgBox("bitte zwei Projekte selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' die Phasen zweier Projekte vergleichen  - Darstellung in zwei Charts
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G3M1B2PhasenVgl(control As IRibbonControl)

        Dim singleShp1 As Excel.Shape, singleShp2 As Excel.Shape
        Dim hproj As New clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then


            If awinSelection.Count = 1 Then

                Dim vproj As clsProjektvorlage
                ' jetzt die Aktion durchführen ...
                singleShp1 = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp1.Name, True)
                    vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                    If IsNothing(vproj) Then
                        Call MsgBox("Vorlage" & hproj.VorlagenName & " nicht gefunden ...")
                        enableOnUpdate = True
                        Exit Sub
                    End If
                    cproj = New clsProjekt
                    vproj.copyTo(cproj)
                    cproj.startDate = hproj.startDate

                Catch ex As Exception
                    Call MsgBox("Vorlage" & hproj.VorlagenName & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try


                top = singleShp1.Top + boxHeight + 2
                left = singleShp1.Left - 5
                If left <= 0 Then
                    left = 5
                End If

                height = 380
                width = System.Math.Max(hproj.dauerInDays / 365 * 12 * boxWidth + 7, cproj.dauerInDays / 365 * 12 * boxWidth + 7)
                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)

                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False


                noColorCollection = getPhasenUnterschiede(hproj, cproj)

                repObj = Nothing
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                With repObj
                    top = .Top + .Height + 3
                End With


                repObj = Nothing
                Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, PThis.vorlage)
                appInstance.ScreenUpdating = True

            ElseIf awinSelection.Count = 2 Then
                ' jetzt die Aktion durchführen ...
                singleShp1 = awinSelection.Item(1)
                singleShp2 = awinSelection.Item(2)

                Try
                    hproj = ShowProjekte.getProject(singleShp1.Name, True)
                    cproj = ShowProjekte.getProject(singleShp2.Name, True)
                Catch ex As Exception
                    Call MsgBox("Projekt nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try


                top = singleShp1.Top + boxHeight + 2
                left = singleShp1.Left - 5
                If left <= 0 Then
                    left = 5
                End If

                height = 380
                width = System.Math.Max(hproj.dauerInDays / 365 * 12 * boxWidth + 7, cproj.dauerInDays / 365 * 12 * boxWidth + 7)
                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)

                Dim repObj As Excel.ChartObject
                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False

                noColorCollection = getPhasenUnterschiede(hproj, cproj)

                repObj = Nothing
                Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                With repObj
                    top = .Top + .Height + 3
                End With


                repObj = Nothing
                Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, PThis.current)
                appInstance.ScreenUpdating = True
                'Call awinCompareProjectPhases(name1:=singleShp1.Name, _
                '                              name2:=singleShp2.Name, _
                '                              compareType:=3)
            Else
                Call MsgBox("bitte zwei Projekte selektieren")

            End If
        Else
            Call MsgBox("ein Projekt selektieren, um mit Vorlage zu vergleichen" & vbLf & _
                        " oder zwei Projekte für den Vergleich untereinander")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    Sub PT3G1B2PhasenVgl(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection
        Dim vglName As String = " "
        Dim pName As String = "", variantName As String
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try


        If request.pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                If awinSelection.Count = 1 Then

                    Try
                        Dim lastElem As Integer

                        ' jetzt die Aktion durchführen ...
                        singleShp1 = awinSelection.Item(1)


                        Try
                            hproj = ShowProjekte.getProject(singleShp1.Name, True)
                        Catch ex As Exception
                            Call MsgBox("Projekt nicht gefunden ...")
                            enableOnUpdate = True
                            Exit Sub
                        End Try

                        ' jetzt ggf die Projekt-Historie aufbauen

                        If Not projekthistorie Is Nothing Then
                            If projekthistorie.Count > 0 Then
                                vglName = projekthistorie.First.getShapeText
                            End If
                        Else
                            projekthistorie = New clsProjektHistorie
                        End If

                        With hproj
                            pName = .name
                            variantName = .variantName
                        End With

                        If vglName <> hproj.getShapeText Then

                            ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                            projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                                storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                            projekthistorie.Add(Date.Now, hproj)
                            lastElem = projekthistorie.Count - 1


                        Else
                            ' der aktuelle Stand hproj muss hinzugefügt werden 
                            lastElem = projekthistorie.Count - 1
                            projekthistorie.RemoveAt(lastElem)
                            projekthistorie.Add(Date.Now, hproj)
                        End If


                        If projekthistorie.Count <= 1 Then

                            Call MsgBox(" es gibt zu diesem Projekt noch keine Historie")

                        Else

                            cproj = projekthistorie.ElementAt(lastElem - 1)

                            top = singleShp1.Top + boxHeight + 2
                            left = singleShp1.Left - 5
                            If left <= 0 Then
                                left = 5
                            End If

                            height = 380
                            width = System.Math.Max(hproj.dauerInDays / 365 * 12 * boxWidth + 7, cproj.dauerInDays / 365 * 12 * boxWidth + 7)
                            scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)

                            Dim repObj As Excel.ChartObject
                            appInstance.EnableEvents = False
                            appInstance.ScreenUpdating = False

                            noColorCollection = getPhasenUnterschiede(hproj, cproj)

                            repObj = Nothing
                            Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                            With repObj
                                top = .Top + .Height + 3
                            End With


                            repObj = Nothing
                            Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, PThis.letzterStand)

                            appInstance.ScreenUpdating = True

                        End If
                    Catch ex As Exception

                        Call MsgBox("es gibt keine Historie zu " & pName)

                    End Try


                Else
                    Call MsgBox("bitte nur ein Projekt selektieren")

                End If
            Else
                Call MsgBox("ein Projekt selektieren, um es mit seinem letzten Stand zu vergleichen")
            End If
        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
            projekthistorie.clear()
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    ''' <summary>
    ''' vergleicht die Phasen Termine des aktuellen Projektes mit der Beauftragung
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT3G1B3PhasenVgl(control As IRibbonControl)

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt, cproj As clsProjekt
        Dim top As Double, left As Double, width As Double, height As Double
        Dim scale As Double
        Dim noColorCollection As New Collection
        Dim vglName As String = " "
        Dim pName As String, variantName As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If request.pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                If awinSelection.Count = 1 Then

                    Dim lastElem As Integer

                    ' jetzt die Aktion durchführen ...
                    singleShp1 = awinSelection.Item(1)


                    Try
                        hproj = ShowProjekte.getProject(singleShp1.Name, True)
                    Catch ex As Exception
                        Call MsgBox("Projekt nicht gefunden ...")
                        enableOnUpdate = True
                        Exit Sub
                    End Try

                    ' jetzt ggf die Projekt-Historie aufbauen

                    If Not projekthistorie Is Nothing Then
                        If projekthistorie.Count > 0 Then
                            vglName = projekthistorie.First.getShapeText
                        End If
                    Else
                        projekthistorie = New clsProjektHistorie
                    End If

                    With hproj
                        pName = .name
                        variantName = .variantName
                    End With

                    If vglName <> hproj.getShapeText Then

                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, hproj)
                        lastElem = projekthistorie.Count - 1


                    Else
                        ' der aktuelle Stand hproj muss hinzugefügt werden 
                        lastElem = projekthistorie.Count - 1
                        projekthistorie.RemoveAt(lastElem)
                        projekthistorie.Add(Date.Now, hproj)
                    End If


                    If projekthistorie.Count = 1 Then

                        Call MsgBox(" es gibt zu diesem Projekt noch keine Historie")

                    Else


                        Try
                            cproj = projekthistorie.beauftragung
                            If IsNothing(cproj) Then
                                cproj = projekthistorie.First
                            End If

                            top = singleShp1.Top + boxHeight + 2
                            left = singleShp1.Left - 5
                            If left <= 0 Then
                                left = 5
                            End If

                            height = 380
                            width = System.Math.Max(hproj.dauerInDays / 365 * 12 * boxWidth + 7, cproj.dauerInDays / 365 * 12 * boxWidth + 7)
                            scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)

                            Dim repObj As Excel.ChartObject
                            appInstance.EnableEvents = False
                            appInstance.ScreenUpdating = False

                            noColorCollection = getPhasenUnterschiede(hproj, cproj)

                            repObj = Nothing
                            Call createPhasesBalken(noColorCollection, hproj, repObj, scale, top, left, height, width, PThis.current)

                            With repObj
                                top = .Top + .Height + 3
                            End With


                            repObj = Nothing
                            Call createPhasesBalken(noColorCollection, cproj, repObj, scale, top, left, height, width, PThis.beauftragung)

                        Catch ex As Exception

                            Call MsgBox("es ist kein Beauftragungs-Stand vorhanden")

                        End Try


                    End If

                Else
                    Call MsgBox("bitte nur ein Projekt selektieren")

                End If
            Else
                Call MsgBox("ein Projekt selektieren, um es mit seiner Beauftragung zu vergleichen")
            End If

        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
        End If
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Sub Tom2G3M1B2ResourceVgl(control As IRibbonControl)

        Dim singleShp1 As Excel.Shape, singleShp2 As Excel.Shape
        'Dim SID As String

        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 2 Then
                ' jetzt die Aktion durchführen ...
                singleShp1 = awinSelection.Item(1)
                singleShp2 = awinSelection.Item(2)

                Dim hproj As clsProjekt
                Dim cproj As clsProjekt
                Try
                    hproj = ShowProjekte.getProject(singleShp1.Name, True)
                    cproj = ShowProjekte.getProject(singleShp2.Name, True)
                    Dim top As Double = singleShp1.Top + boxHeight + 2
                    Dim left As Double = singleShp1.Left - boxWidth
                    If left <= 0 Then
                        left = 1
                    End If
                    Call awinCompareProject(hproj, cproj, 3, top, left)
                Catch ex As Exception
                    Call MsgBox("Fehler bei Compare" & vbLf & ex.Message)
                End Try

            Else
                Call MsgBox("bitte zwei Projekte selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If


        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    Sub awinShowTrendSR(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, height As Double, width As Double
        Dim vglName As String = " "

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.ScreenUpdating = False


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If request.pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                If awinSelection.Count = 1 And isProjectType(kindOfShape(awinSelection.Item(1))) Then
                    ' jetzt die Aktion durchführen ...
                    singleShp = awinSelection.Item(1)


                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                    With hproj
                        pName = .name
                        variantName = .variantName
                    End With

                    If Not projekthistorie Is Nothing Then
                        If projekthistorie.Count > 0 Then
                            vglName = projekthistorie.First.getShapeText
                        End If
                    Else
                        projekthistorie = New clsProjektHistorie
                    End If

                    If vglName <> hproj.getShapeText Then

                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, hproj)

                    Else
                        ' der aktuelle Stand hproj muss hinzugefügt werden 
                        Dim lastElem As Integer = projekthistorie.Count - 1
                        projekthistorie.RemoveAt(lastElem)
                        projekthistorie.Add(Date.Now, hproj)
                    End If

                    Dim nrSnapshots As Integer = projekthistorie.Count

                    If nrSnapshots > 0 Then
                        With singleShp
                            top = .Top + boxHeight + 2
                            left = .Left - 3
                        End With
                        width = System.Math.Max(nrSnapshots * boxWidth * 0.65, 300)

                        height = 16 * boxHeight
                        Dim repObj As Excel.ChartObject = Nothing
                        Call createTrendSfit(repObj, top, left, height, width)

                    Else
                        Call MsgBox("es gibt noch keine Projekt-Historie zu " & pName)
                    End If




                Else
                    Call MsgBox("bitte nur ein Projekt selektieren")
                    'For Each singleShp In awinSelection
                    '    With singleShp
                    '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                    '            nrSelPshp = nrSelPshp + 1
                    '            SID = .ID.ToString
                    '        End If
                    '    End With
                    'Next
                End If
            Else
                Call MsgBox("vorher Projekt selektieren ...")
            End If
        Else
            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
        End If

        enableOnUpdate = True
        appInstance.ScreenUpdating = True




    End Sub


    Sub awinShowTrendKPI(control As IRibbonControl)
        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, height As Double, width As Double
        Dim vglName As String = " "


        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.ScreenUpdating = False


        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count = 1 And isProjectType(kindOfShape(awinSelection.Item(1))) Then
                ' jetzt die Aktion durchführen ...
                singleShp = awinSelection.Item(1)


                hproj = ShowProjekte.getProject(singleShp.Name, True)
                With hproj
                    pName = .name
                    variantName = .variantName
                End With

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.getShapeText
                    End If
                Else
                    projekthistorie = New clsProjektHistorie
                End If

                If vglName <> hproj.getShapeText Then
                    If request.pingMongoDb() Then
                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, hproj)
                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
                        projekthistorie.clear()
                    End If

                Else
                    ' der aktuelle Stand hproj muss hinzugefügt werden 
                    Dim lastElem As Integer = projekthistorie.Count - 1
                    projekthistorie.RemoveAt(lastElem)
                    projekthistorie.Add(Date.Now, hproj)
                End If

                Dim nrSnapshots As Integer = projekthistorie.Count

                If nrSnapshots > 0 Then
                    With singleShp
                        top = .Top + boxHeight + 2
                        left = .Left - 3
                    End With
                    width = System.Math.Max(nrSnapshots * boxWidth * 0.65, 300)

                    height = 16 * boxHeight
                    Dim repObj As Excel.ChartObject = Nothing
                    Call createTrendKPI(repObj, top, left, height, width)

                Else
                    Call MsgBox("es gibt noch keine Projekt-Historie zu " & pName)
                End If

            Else
                Call MsgBox("bitte nur ein Projekt selektieren")
                'For Each singleShp In awinSelection
                '    With singleShp
                '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
                '            nrSelPshp = nrSelPshp + 1
                '            SID = .ID.ToString
                '        End If
                '    End With
                'Next
            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub


    Sub awinShowTimeMachine(control As IRibbonControl)


        Call PBBShowTimeMachine(control)

        ' ''Dim hproj As clsProjekt
        ' ''Dim pName As String, variantName As String
        ' ''Dim vglName As String = " "
        ' ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        ' ''Dim singleShp As Excel.Shape
        ' ''Dim showCharacteristics As New frmShowProjCharacteristics
        '' ''Dim returnValue As DialogResult
        ' ''Dim awinSelection As Excel.ShapeRange
        ' ''Dim grueneAmpel As String = awinPath & "gruen.gif"
        ' ''Dim gelbeAmpel As String = awinPath & "gelb.gif"
        ' ''Dim roteAmpel As String = awinPath & "rot.gif"
        ' ''Dim graueAmpel As String = awinPath & "grau.gif"

        ' ''If timeMachineIsOn Then
        ' ''    Call MsgBox("bitte erst Time Machine beenden ...")
        ' ''    Exit Sub
        ' ''End If

        ' ''Call projektTafelInit()

        ' ''enableOnUpdate = False
        ' ''appInstance.EnableEvents = True


        ' ''Try
        ' ''    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        ' ''Catch ex As Exception
        ' ''    awinSelection = Nothing
        ' ''End Try

        ' ''If Not awinSelection Is Nothing Then


        ' ''    If awinSelection.Count = 1 And isProjectType(kindOfShape(awinSelection.Item(1))) Then
        ' ''        ' jetzt die Aktion durchführen ...
        ' ''        singleShp = awinSelection.Item(1)
        ' ''        hproj = ShowProjekte.getProject(singleShp.Name)
        ' ''        With hproj
        ' ''            pName = .name
        ' ''            variantName = .variantName
        ' ''            'Try
        ' ''            '    variantName = .variantName.Trim
        ' ''            'Catch ex As Exception
        ' ''            '    variantName = ""
        ' ''            'End Try

        ' ''        End With

        ' ''        If Not projekthistorie Is Nothing Then
        ' ''            If projekthistorie.Count > 0 Then
        ' ''                vglName = projekthistorie.First.getShapeText
        ' ''            End If

        ' ''        Else
        ' ''            projekthistorie = New clsProjektHistorie
        ' ''        End If

        ' ''        If vglName <> hproj.getShapeText Then

        ' ''            If request.pingMongoDb() Then
        ' ''                ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
        ' ''                projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
        ' ''                                                                    storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
        ' ''                If projekthistorie.Count <> 0 Then

        ' ''                    projekthistorie.Add(Date.Now, hproj)

        ' ''                End If

        ' ''            Else
        ' ''                Call MsgBox("Datenbank-Verbindung ist unterbrochen")
        ' ''                projekthistorie.clear()
        ' ''            End If

        ' ''        Else
        ' ''            ' der aktuelle Stand hproj muss hinzugefügt werden 
        ' ''            Dim lastElem As Integer = projekthistorie.Count - 1
        ' ''            projekthistorie.RemoveAt(lastElem)
        ' ''            projekthistorie.Add(Date.Now, hproj)
        ' ''        End If


        ' ''        Dim nrSnapshots As Integer = projekthistorie.Count

        ' ''        If nrSnapshots > 0 Then

        ' ''            With showCharacteristics

        ' ''                .Text = "Historie für Projekt " & pName.Trim & vbLf & _
        ' ''                        "( " & projekthistorie.getZeitraum & " )"
        ' ''                .timeSlider.Minimum = 0
        ' ''                .timeSlider.Maximum = nrSnapshots - 1

        ' ''                '.ampelErlaeuterung.Text = kvp.Value.ampelErlaeuterung

        ' ''                'If kvp.Value.ampelStatus = 1 Then
        ' ''                '    .ampelPicture.LoadAsync(grueneAmpel)
        ' ''                'ElseIf kvp.Value.ampelStatus = 2 Then
        ' ''                '    .ampelPicture.LoadAsync(gelbeAmpel)
        ' ''                'ElseIf kvp.Value.ampelStatus = 3 Then
        ' ''                '    .ampelPicture.LoadAsync(roteAmpel)
        ' ''                'Else
        ' ''                '    .ampelPicture.LoadAsync(graueAmpel)
        ' ''                'End If

        ' ''                '.snapshotDate.Text = kvp.Value.timeStamp.ToString
        ' ''                ' das ist ja der aktuelle Stand ..
        ' ''                .snapshotDate.Text = "Aktueller Stand"
        ' ''                ' Designer 
        ' ''                'Dim zE As String = "(" & awinSettings.zeitEinheit & ")"
        ' ''                '.engpass1.Text = "Designer:          " & kvp.Value.getRessourcenBedarf(3).Sum.ToString("###.#") & zE
        ' ''                '.engpass2.Text = "Personalkosten: " & kvp.Value.getAllPersonalKosten.Sum.ToString("###.#") & " (T€)"
        ' ''                '.engpass3.Text = "Sonstige Kosten:   " & kvp.Value.getGesamtAndereKosten.Sum.ToString("###.#") & " (T€)"


        ' ''            End With


        ' ''            ' jetzt wird das Form aufgerufen ... 

        ' ''            'returnValue = showCharacteristics.ShowDialog
        ' ''            showCharacteristics.Show()

        ' ''        Else
        ' ''            Call MsgBox("es gibt noch keine Planungs-Historie")
        ' ''        End If

        ' ''    Else
        ' ''        Call MsgBox("bitte nur ein Projekt selektieren")
        ' ''        'For Each singleShp In awinSelection
        ' ''        '    With singleShp
        ' ''        '        If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Then
        ' ''        '            nrSelPshp = nrSelPshp + 1
        ' ''        '            SID = .ID.ToString
        ' ''        '        End If
        ' ''        '    End With
        ' ''        'Next
        ' ''    End If
        ' ''Else
        ' ''    Call MsgBox("vorher Projekt selektieren ...")
        ' ''End If

        ' ''enableOnUpdate = True


    End Sub



    ' 

    ''' <summary>
    ''' aktuelle Konstellation wird dokumentiert
    ''' Report-Vorlage wird im Formular 'Auswählen der Report-Vorlage' ausgewählt
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub awinAllprojectsReport(control As IRibbonControl)

        Dim getReportVorlage As New frmSelectPPTTempl
        Dim returnValue As DialogResult
        Dim timeZoneWasOff As Boolean = False
        getReportVorlage.calledfrom = "Portfolio1"

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.ScreenUpdating = False
        If showRangeRight - showRangeLeft >= minColumns - 1 Then

            If ShowProjekte.Count > 0 Then

                ' Formular zum Auswählen der Report-Vorlage wird aufgerufen

                returnValue = getReportVorlage.ShowDialog

            Else
                Call MsgBox("Es sind keine Projekte geladen!")
            End If
        Else
            ' automatisch bestimmen 
            timeZoneWasOff = True
            If selectedProjekte.Count > 0 Then
                showRangeLeft = selectedProjekte.getMinMonthColumn
                showRangeRight = selectedProjekte.getMaxMonthColumn
            Else
                showRangeLeft = ShowProjekte.getMinMonthColumn
                showRangeRight = ShowProjekte.getMaxMonthColumn
            End If
            Call awinShowtimezone(showRangeLeft, showRangeRight, True)

        End If

        If timeZoneWasOff Then
            Call awinShowtimezone(showRangeLeft, showRangeRight, False)
            showRangeLeft = 0
            showRangeRight = 0
        End If

        appInstance.ScreenUpdating = True
        enableOnUpdate = True


    End Sub
    Sub PTShowVersions(control As IRibbonControl)

        'Ermittlung der installierten Windows- und der Excelversion
        Call MsgBox("Betriebssystem: " & appInstance.OperatingSystem & Chr(10) & _
        "Excel-Version: " & appInstance.Version, vbInformation, "Info")
        'Call MsgBox("Betriebssystem: " & appInstance.OperatingSystem & Chr(10) & _
        '"Excel-Version: " & My.Settings.ExcelVersion, vbInformation, "Info")
    End Sub
    Sub PTAddMissingInit(control As IRibbonControl, ByRef pressed As Boolean)

        pressed = awinSettings.addMissingPhaseMilestoneDef

    End Sub
    Sub PTAddMissingDefinitions(control As IRibbonControl, ByRef pressed As Boolean)

        If pressed Then
            awinSettings.addMissingPhaseMilestoneDef = True
        Else
            awinSettings.addMissingPhaseMilestoneDef = False
        End If

    End Sub
    Sub PTTestFunktion1(control As IRibbonControl)

        Call MsgBox("Enable Events ist " & appInstance.EnableEvents.ToString)
        Call MsgBox("Screen Updating " & appInstance.ScreenUpdating.ToString)
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


    End Sub

    Sub PTTestFunktion2(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        ''Dim tstCollection As SortedList(Of Date, String)
        Dim anzElements As Integer

        Dim awinSelection As Excel.ShapeRange
        Dim projektHistorien As New clsProjektDBInfos
        Dim todoListe As New clsProjektDBInfos
        Dim i As Integer


        Dim schluessel As String = ""

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True



        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            If awinSelection.Count >= 1 Then
                anzElements = awinSelection.Count

                For i = 1 To anzElements

                    singleShp = awinSelection.Item(i)
                    hproj = ShowProjekte.getProject(singleShp.Name, True)

                    Dim openXMLproj As New clsOpenXML
                    Call openXMLproj.copyFrom(hproj)

                    Dim vglProj As New clsProjekt
                    Call openXMLproj.copyTo(vglProj)

                    vglProj.variantName = "OpenXML"
                    If Not AlleProjekte.Containskey(calcProjektKey(vglProj)) Then
                        AlleProjekte.Add(vglProj)
                    End If


                    Dim unterschiede As New Collection
                    ' jetzt wird festgestellt, ob es Unterschiede gibt 
                    ' 
                    unterschiede = hproj.listOfDifferences(vglProj, True, 0)

                    If unterschiede.Count > 0 Then
                        Dim ergStr As String = ""
                        For ei As Integer = 1 To unterschiede.Count
                            If ei = 1 Then
                                ergStr = CStr(unterschiede.Item(i))
                            Else
                                ergStr = ergStr & "; " & CStr(unterschiede.Item(i))
                            End If

                        Next
                        Call MsgBox("Unterschiede: " & ergStr)

                    Else
                        Call MsgBox(vglProj.name & ": identisch ...")
                    End If


                    'If i = 1 Then
                    '    schluessel = calcProjektKey(hproj)
                    'End If

                    ''If request.pingMongoDb() Then
                    ''    ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                    ''    projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                    ''                                                       storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                    ''Else
                    ''    Call MsgBox("Datenbank-Verbindung ist unterbrochen")
                    ''    projekthistorie.clear()
                    ''End If

                    ''If projekthistorie.Count > 0 Then
                    ''    ' Aufbau der Listen 
                    ''    projektHistorien.Add(projekthistorie)


                    ''End If

                Next
            End If
        End If




        ''tstCollection = projektHistorien.getTimeStamps(schluessel)
        ''anzElements = tstCollection.Count

        ''For i = 1 To anzElements
        ''    ts = tstCollection.ElementAt(0).Key
        ''    projektHistorien.Remove(schluessel, ts)
        ''    todoListe.Add(schluessel, ts)
        ''Next


        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' testet, ob die Hierarchien in den geladenen Projekten alle stimmig sind 
    ''' das heißt, verweisen die Indices tatsächlich auf die richtigen Phasen bzw Meilensteine  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTTestFunktion3(control As IRibbonControl)

        Dim hproj As clsProjekt
        Dim allesInOrdnung As Boolean = True
        Dim anzElements As Integer
        Dim curNode As clsHierarchyNode
        Dim parentNode As clsHierarchyNode
        Dim childNode As clsHierarchyNode
        Dim parentID As String
        Dim curID As String
        Dim childID As String
        Dim elemID As String
        Dim elemName As String
        Dim lfdNr As Integer
        Dim isMilestone As Boolean
        Dim cphase As clsPhase
        Dim cphase2 As clsPhase
        Dim cMilestone As clsMeilenstein
        Dim logMessage As String = ""
        Dim atleastOne As Boolean = False


        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        ' Testreihe 1: ausgehend von der Hierarchie alle Projekte und Varianten 

        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

            hproj = kvp.Value


            ' zuerst wird ausgehend von der Hierarchie gecheckt 
            anzElements = hproj.hierarchy.count

            For ix As Integer = 1 To anzElements

                curID = hproj.hierarchy.getIDAtIndex(ix)
                elemName = elemNameOfElemID(curID)
                lfdNr = lfdNrOfElemID(curID)

                curNode = hproj.hierarchy.nodeItem(ix)

                If curID.StartsWith("1§") Then
                    isMilestone = True
                ElseIf curID.StartsWith("0§") Then
                    isMilestone = False
                Else
                    logMessage = logMessage & vbLf & kvp.Value.getShapeText & ": Node kann nicht identifiziert werden .." & curID
                    atleastOne = True
                End If

                If Not isMilestone Then
                    ' test 1: Zugriff über ID 
                    cphase = hproj.getPhaseByID(curID)

                    If cphase.nameID = curID Then
                        ' ok 
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Node-Zugriff über ID nicht ok " & curID & ", " & cphase.nameID
                        atleastOne = True
                    End If

                    ' Test2: Zugriff über Name und lfd-Nr 
                    elemID = calcHryElemKey(elemName, isMilestone, lfdNr)

                    cphase = hproj.getPhaseByID(elemID)

                    If cphase.nameID = elemID Then
                        ' ok 
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Node-Zugriff über Elem-Name, lfdNr nicht ok " & curID & ", " & cphase.nameID
                        atleastOne = True
                    End If

                Else
                    cMilestone = hproj.getMilestoneByID(curID)

                    If cMilestone.nameID = curID Then
                        ' ok 
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Node-Zugriff über ID nicht ok " & curID & ", " & cMilestone.nameID
                        atleastOne = True
                    End If

                    ' Test2: Zugriff über Name und lfd-Nr 
                    elemID = calcHryElemKey(elemName, isMilestone, lfdNr)

                    cMilestone = hproj.getMilestoneByID(elemID)

                    If cMilestone.nameID = elemID Then
                        ' ok 
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Node-Zugriff über Elem-Name, lfdNr nicht ok " & curID & ", " & cMilestone.nameID
                        atleastOne = True
                    End If


                End If

                ' jetzt wird gecheckt, ob das Element einen parent hat - wenn ja, ob es auch das Kind des Parents ist   
                ' wenn ja, wird gecheckt, ob der Parent-Knoten das aktuelle Element in der Liste der Child-Knoten hat 

                parentID = curNode.parentNodeKey
                If parentID <> "" Then
                    parentNode = hproj.hierarchy.nodeItem(parentID)

                    If Not IsNothing(parentNode) Then
                        Dim found As Boolean
                        For cx As Integer = 1 To parentNode.childCount
                            If curID = parentNode.getChild(cx) Then
                                found = True
                            End If
                        Next
                        If Not found Then
                            logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Eltern-Knoten hat mich nicht als Kind" & parentID & ", Kind:  " & curID
                        End If
                    Else
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "hat keinen Eltern-Knoten: curID: " & curID & ", parentID " & parentID
                    End If
                End If

                ' jetzt wird gecheckt, ob das Element Kinder hat -  
                ' wenn ja, ob jedes Kind das Element als parent hat    

                For cx As Integer = 1 To curNode.childCount

                    childID = curNode.getChild(cx)
                    childNode = hproj.hierarchy.nodeItem(childID)
                    If Not childNode.parentNodeKey = curID Then
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Kind hat mich nicht als Vater:  " & childID & ", CurID " & curID
                    End If

                Next


            Next

            ' jetzt wird ausgehend von den Phasen und den zugehörigen Milestones gecheckt 

            For ix As Integer = 1 To hproj.CountPhases

                cphase = hproj.getPhase(ix)
                curID = cphase.nameID

                ' check in der Hierarchie
                cphase2 = hproj.getPhaseByID(curID)
                If Not IsNothing(cphase2) Then
                    If Not cphase2.nameID = cphase.nameID Then
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Zugriff über ix: " & ix & ": " & cphase.nameID & " <> " & cphase.nameID
                        atleastOne = True
                    End If
                End If


                ' jetzt werden die Meilensteine gecheckt
                For mx As Integer = 1 To cphase.countMilestones
                    cMilestone = cphase.getMilestone(mx)
                    curID = cMilestone.nameID
                    curNode = hproj.hierarchy.nodeItem(curID)
                    If curNode.indexOfElem <> mx Then
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Meilenstein-Zugriff über mx: " & ix & vbLf & _
                                     curNode.indexOfElem & " <> " & mx
                    End If

                    parentID = curNode.parentNodeKey
                    parentNode = hproj.hierarchy.nodeItem(parentID)
                    If Not IsNothing(parentNode) Then
                        If parentNode.indexOfElem <> ix Then
                            logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Phasen-Zugriff über ix: " & ix & vbLf & _
                                         parentNode.indexOfElem & " <> " & mx
                        End If
                    Else
                        If parentID <> "" Then
                            logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Phasen-Zugriff über ix: " & ix & vbLf & _
                                         curID & " hat keinen Parent " & parentID
                        End If
                    End If

                Next


            Next


        Next

        If atleastOne Or logMessage.Length > 1 Then
            atleastOne = False
            Call MsgBox(logMessage)
            logMessage = ""
        End If

        If Not atleastOne Then
            Call MsgBox("alles ok ..")
        End If

        enableOnUpdate = True


    End Sub

    Sub PTTestFunktion7(control As IRibbonControl)

        Dim hproj As clsProjekt

        Dim atleastOne As Boolean = False


        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        ' Testreihe 1: sind die angegebenen Rollen identisch ? 
        Dim duration1 As Integer = 0
        Dim duration2 As Integer = 0


        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

            hproj = kvp.Value

            Dim usedRollen1 As Collection = hproj.getRoleNames
            Dim usedRollen2 As Collection = hproj.rcLists.getRoleNames

            ' Test auf Identität der beiden usedRollen1,2

            If usedRollen1.Count <> usedRollen2.Count Then
                atleastOne = True
            Else
                For ix As Integer = 1 To usedRollen1.Count
                    If Not usedRollen2.Contains(CStr(usedRollen1.Item(ix))) Then
                        Dim name1 As String = CStr(usedRollen1.Item(ix))
                        Dim name2 As String = CStr(usedRollen2.Item(ix))
                        atleastOne = True
                    End If

                Next
            End If


            Dim usedCost1 As Collection = hproj.getCostNames
            Dim usedCost2 As Collection = hproj.rcLists.getCostNames

            If usedCost1.Count <> usedCost2.Count Then
                atleastOne = True
            Else
                For ix As Integer = 1 To usedCost1.Count
                    If Not usedCost2.Contains(CStr(usedCost1.Item(ix))) Then
                        Dim name1 As String = CStr(usedCost1.Item(ix))
                        Dim name2 As String = CStr(usedCost2.Item(ix))
                        atleastOne = True
                    End If

                Next
            End If


        Next

        If atleastOne Then
            Call MsgBox("bei Rollen/Kosten nicht alles ok ...")
        Else
            Call MsgBox("bei Rollen/Kosten alles ok ..")
        End If

        Call MsgBox("jetzt ist es: " & Date.Now.ToLongTimeString)
        atleastOne = False

        ' Test-Zyklus 2 
        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            ' alte Methode .... 

            ' mach es möglichst oft ...

            For iter As Integer = 1 To 1

                For ix As Integer = 1 To RoleDefinitions.Count
                    Dim role As clsRollenDefinition = RoleDefinitions.getRoledef(ix)

                    Dim zeitraumBedarf() As Double = ShowProjekte.getRoleValuesInMonth(role.UID, True)
                    Dim zeitraumBedarf2() As Double = ShowProjekte.getRoleValuesInMonthNew(role.UID, True)

                    If arraysAreDifferent(zeitraumBedarf, zeitraumBedarf2) Then
                        atleastOne = True
                    End If

                Next

                If atleastOne Then
                    Call MsgBox("Rollen-Summen nicht alles ok ...")
                Else
                    Call MsgBox("Rollen-Summen alles ok ..")
                End If
                atleastOne = False

                For ix As Integer = 1 To CostDefinitions.Count
                    Dim cost As clsKostenartDefinition = CostDefinitions.getCostdef(ix)

                    Dim zeitraumBedarf() As Double = ShowProjekte.getCostValuesInMonth(cost.UID)
                    Dim zeitraumBedarf2() As Double = ShowProjekte.getCostValuesInMonthNew(cost.UID)

                    If arraysAreDifferent(zeitraumBedarf, zeitraumBedarf2) Then
                        atleastOne = True
                    End If
                Next

                If atleastOne Then
                    Call MsgBox("Kosten-Summen nicht alles ok ...")
                Else
                    Call MsgBox("Kosten-Summen alles ok ..")
                End If

            Next

            'Call MsgBox("jetzt ist es: " & Date.Now.ToLongTimeString)

            'For iter As Integer = 1 To 400

            '    For ix As Integer = 1 To RoleDefinitions.Count
            '        Dim role As clsRollenDefinition = RoleDefinitions.getRoledef(ix)

            '        Dim zeitraumBedarf() As Double = ShowProjekte.getRoleValuesInMonthNew(role.UID, True)

            '    Next
            'Next

            Call MsgBox("jetzt ist es: " & Date.Now.ToLongTimeString)

        Else
            Call MsgBox("zuerst Zeitraum definieren ...")
        End If



        enableOnUpdate = True


    End Sub

    Public Sub PTTestFunktion4(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True
        Dim yellows As Double = 0.05
        Dim reds As Double = 0.015

        If demoModusHistory And historicDate > StartofCalendar And historicDate < Date.Now Then
            ' es werden nur die Meilensteine verändert, die nach dem historicdate liegen 
            ' oder die, vorher liegen und noch keine Bewertung haben
            Call createInitialRandomBewertungen(yellows, reds, historicDate)
        Else
            ' es werden nur die Meilensteine verändert, die nach dem heutigen Datum  
            ' oder die, die vorher liegen und noch keine Bewertung haben
            Call createInitialRandomBewertungen(yellows, reds, Date.Now)
        End If


        enableOnUpdate = True


    End Sub
    Public Sub PTTestWriteProtect(control As IRibbonControl)

        Call projektTafelInit()
        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        Dim ok2 As Boolean = request.cancelWriteProtections(dbUsername)

        ' ''Dim wpItem As clsWriteProtectionItem

        ' ''For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

        ' ''    wpItem = New clsWriteProtectionItem(kvp.Key, dbUsername, True)
        ' ''    ' ''wpItem = New clsWriteProtectionItem(kvp.Key, dbUsername, False)
        ' ''    ' ''wpItem.isProtected = False
        ' ''    Dim ok As Boolean = request.setWriteProtection(wpItem)
        ' ''    If ok Then
        ' ''        'Call MsgBox("Projekt " & wpItem.pvName & " wurde von User " & wpItem.userName & _
        ' ''        '            "nicht permanent geschützt: Date: " & wpItem.lastDateSet.ToShortDateString)
        ' ''    Else

        ' ''        Dim writeProtections As SortedList(Of String, clsWriteProtectionItem) = request.retrieveWriteProtectionsFromDB(AlleProjekte)
        ' ''        Dim resultstr As String = ""
        ' ''        For Each elem As KeyValuePair(Of String, clsWriteProtectionItem) In writeProtections
        ' ''            resultstr = resultstr & vbLf & elem.Key & elem.Value.userName
        ' ''        Next
        ' ''        Call MsgBox("Projekt " & wpItem.pvName & " konnte nicht für User " & wpItem.userName & _
        ' ''                    " geschützt werden: Date: " & wpItem.lastDateSet.ToShortDateString & vbLf & _
        ' ''                    "writeProtections: " & resultstr)
        ' ''    End If

        ' ''Next

        enableOnUpdate = True

    End Sub

    Public Sub PTCreateLicense(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim frmLizenzen As New frmCreateLicences
        ' ''Dim i As Integer
        ' ''Dim k As Integer
        ' ''Dim VisboLic As New clsLicences
        ' ''Dim clientLic As New clsLicences
        ' ''Dim komponenten() As String = {"Swimlanes2"}
        ' ''Dim users() As String = {"matthias.kaufhold", "thomas.braeutigam", "Ingo.Hanschke", myWindowsName, "BHTC-Domain/thomas.braeutigam", "Ute-Dagmar.Rittinghaus-Koytek"}
        ' ''Dim endDate As Date = DateAdd(DateInterval.Month, 1200, Date.Now)


        Dim returnValue As DialogResult
        returnValue = frmLizenzen.ShowDialog


        ' ''For i = 0 To users.Length - 1

        ' ''    For k = 0 To komponenten.Length - 1

        ' ''        ' Lizenzkey berechnen
        ' ''        Dim licString As String = VisboLic.berechneKey(endDate, users(i), komponenten(k))

        ' ''        ' VsisboListe mit Angabe von username, komponente, endDate
        ' ''        Dim visbokey As String = users(i) & "-" & komponenten(k) & "-" & endDate.ToString
        ' ''        If VisboLic.Liste.ContainsKey(visbokey) Then
        ' ''            Dim ok As Boolean = VisboLic.Liste.Remove(visbokey)
        ' ''        End If
        ' ''        VisboLic.Liste.Add(visbokey, licString)

        ' ''        ' Liste von Lizenzen für den Kunden 
        ' ''        If clientLic.Liste.ContainsKey(licString) Then
        ' ''            Dim ok As Boolean = clientLic.Liste.Remove(licString)
        ' ''        End If
        ' ''        clientLic.Liste.Add(licString, licString)

        ' ''    Next k               'nächste Komponente

        ' ''Next i                   ' nächster User

        '' '' Lizenzen in XML-Dateien speichern
        ' ''Call XMLExportLicences(VisboLic, requirementsOrdner & "visboLicfile.xml")

        ' ''Call XMLExportLicences(clientLic, licFileName)

        enableOnUpdate = True



    End Sub

    Public Sub PTTestLicense(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim frmTestLizenzen As New frmTestLicences

        Dim returnValue As DialogResult
        returnValue = frmTestLizenzen.ShowDialog


        enableOnUpdate = True


    End Sub


    Public Sub PTCreateReportMessages(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim frmCreateReportMsg As New frmCreateReportMeldungen
        Dim returnValue As DialogResult
        returnValue = frmCreateReportMsg.ShowDialog


        enableOnUpdate = True


    End Sub


    Public Sub PTSpracheinstellung(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim frmReportSprache As New frmSelectRepSprache
        Dim returnValue As DialogResult
        returnValue = frmReportSprache.ShowDialog


        enableOnUpdate = True


    End Sub
    ''' <summary>
    ''' Speichern des aktuellen, in CurrentReportProfil gespeicherten ReportProfil auf Platte in Dir awinPath\requirements\ReportProfile
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub PTStoreCurReportProfil(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim profilNameForm As New frmStoreReportProfil
        Dim returnvalue As DialogResult
        returnvalue = profilNameForm.ShowDialog

        enableOnUpdate = True
    End Sub

    Public Sub PTDoReportOfProfil(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        Dim timeZoneWasOff As Boolean = False
        Dim reportAuswahl As New frmReportProfil
        Dim returnvalue As DialogResult

        ' ist ein Timespan selektiert ? 
        ' wenn nein, selektieren ... 
        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            ' alles in Ordnung 
        Else
            ' automatisch bestimmen 
            timeZoneWasOff = True
            If selectedProjekte.Count > 0 Then
                showRangeLeft = selectedProjekte.getMinMonthColumn
                showRangeRight = selectedProjekte.getMaxMonthColumn
            Else
                showRangeLeft = ShowProjekte.getMinMonthColumn
                showRangeRight = ShowProjekte.getMaxMonthColumn
            End If
            Call awinShowtimezone(showRangeLeft, showRangeRight, True)
        End If

        If ShowProjekte.Count > 0 Then

            reportAuswahl.calledFrom = "Multiprojekt-Tafel"
            returnvalue = reportAuswahl.ShowDialog

            Call awinDeSelect()

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please load some projects first ...")
            Else
                Call MsgBox("Aktuell sind keine Projekte geladen. Bitte laden Sie Projekte!")
            End If

        End If

        If timeZoneWasOff Then
            Call awinShowtimezone(showRangeLeft, showRangeRight, False)
            showRangeLeft = 0
            showRangeRight = 0
        End If

        enableOnUpdate = True

    End Sub

    Public Sub PTCreateReportGenTemplate(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False

        If AlleProjekte.Count > 0 Then

            Call createReportGenTemplate()
            Call awinDeSelect()
        Else
            Call MsgBox("Aktuell sind keine Projekte geladen. Bitte laden Sie Projekte!")
        End If


        enableOnUpdate = True

    End Sub
    Public Sub PTExit(control As IRibbonControl)

        enableOnUpdate = False

        appInstance.ActiveWorkbook.Close()

        enableOnUpdate = True

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
