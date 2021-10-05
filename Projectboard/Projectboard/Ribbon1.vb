Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports ProjectboardReports
Imports DBAccLayer
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Security.Principal
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows
Imports System.Net
Imports System
Imports System.IO
Imports System.Globalization
Imports Microsoft.VisualBasic
Imports System.Web



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
    Private tempShowHeaders As Boolean = False

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Microsoft.Office.Core.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("ExcelWorkbook1.Ribbon1.xml")
    End Function


#Region "Menübandrückrufe"
    'Erstellen Sie hier Rückrufmethoden. Weitere Informationen über das Hinzufügen von Rückrufmethoden erhalten Sie, indem Sie das Menüband-XML-Element im Projektmappen-Explorer markieren und dann F1 drücken.
    Public Sub Ribbon_Load(ByVal ribbonUI As Microsoft.Office.Core.IRibbonUI)
        Me.ribbon = ribbonUI
        Me.ribbon.Invalidate()
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
                currentConstellationPvName = constellationName
            End If
        Else
            Call MsgBox("Es sind keine Projekte in der Projekt-Tafel geladen!")
        End If
        ' 
        ' Ende alte Version; vor dem 26.10.14




        enableOnUpdate = True

    End Sub

    Sub PTLoadRemoveConstellationFromSession(control As IRibbonControl)

        Dim ControlID As String = control.Id

        Dim err As New clsErrorCodeMsg

        Dim removeConstFilterFrm As New frmRemoveConstellation
        Dim constFilterName As String
        Dim dbPortfolioNames As New SortedList(Of String, String)
        Dim constellationsToDo As New clsConstellations

        Dim boardWasEmpty As Boolean = ShowProjekte.Count = 0

        Dim returnValue As DialogResult

        Call projektTafelInit()



        Dim deleteFromSession As String = "PT2G3M1B3"
        Dim deleteFilter As String = "Pt6G3B5"
        Dim loadfromSession As String = "PT2G2B2"
        Dim removeFromDB As Boolean

        If ControlID = deleteFromSession Then
            removeConstFilterFrm.frmOption = "ProjConstellation"
            For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste
                dbPortfolioNames.Add(kvp.Key, kvp.Value.vpID)
            Next
            removeConstFilterFrm.dbPortfolioNames = dbPortfolioNames
            removeFromDB = False

        ElseIf ControlID = loadfromSession Then
            removeConstFilterFrm.frmOption = "PortfolioAusSessionLaden"
            For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste
                dbPortfolioNames.Add(kvp.Key, kvp.Value.vpID)
            Next
            removeConstFilterFrm.dbPortfolioNames = dbPortfolioNames
            removeFromDB = False

        ElseIf ControlID = deleteFilter And Not noDB Then
            removeConstFilterFrm.frmOption = "DBFilter"
            removeFromDB = True


            If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
                filterDefinitions.filterListe = CType(databaseAcc, DBAccLayer.Request).retrieveAllFilterFromDB(False)
            Else
                Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
                removeFromDB = False
            End If

        Else
            removeFromDB = False
        End If

        Dim weiterMitFormular As Boolean = True

        If (ControlID = loadfromSession Or ControlID = deleteFromSession) And removeConstFilterFrm.dbPortfolioNames.Count <= 0 Then
            Call MsgBox("es sind keine Portfolios geladen....")
            weiterMitFormular = False
        End If
        If ControlID = deleteFilter And filterDefinitions.filterListe.Count <= 0 Then
            Call MsgBox("es sind keine Filter geladen....")
            weiterMitFormular = False
        End If

        If weiterMitFormular Then

            enableOnUpdate = False

            While returnValue <> DialogResult.OK And returnValue <> DialogResult.Cancel

                ' Formular mit aufgelisteten Portfolios/Filter anzeigen
                returnValue = removeConstFilterFrm.ShowDialog

            End While

            Dim outputCollection As New Collection
            Dim outputLine As String = ""

            If returnValue = DialogResult.OK Then
                If ControlID = deleteFromSession Then

                    appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

                    For ix As Integer = 1 To removeConstFilterFrm.ListBox1.SelectedItems.Count
                        constFilterName = CStr(removeConstFilterFrm.ListBox1.SelectedItems.Item(ix - 1))
                        Dim constFilterName_sav As String = constFilterName
                        ' portfolioName und variantName wieder durch # getrennt
                        Dim hstr() As String = Split(constFilterName, "[")
                        If hstr.Length > 1 Then
                            constFilterName = hstr(0) & "#" & deleteBrackets(hstr(1), "[", "]")
                        End If

                        Dim constvpid As String = dbPortfolioNames(constFilterName)

                        Call awinRemoveConstellation(constFilterName, constvpid, removeFromDB)
                        dbPortfolioNames.Remove(constFilterName)

                        If awinSettings.englishLanguage Then
                            outputLine = constFilterName_sav & " deleted ..."
                        Else
                            outputLine = constFilterName_sav & " wurde gelöscht ..."
                        End If
                        outputCollection.Add(outputLine)
                    Next

                    appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault

                End If

                ' Laden von der Session
                If ControlID = loadfromSession Then

                    appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait

                    For ix As Integer = 1 To removeConstFilterFrm.ListBox1.SelectedItems.Count

                        Try
                            constFilterName = CStr(removeConstFilterFrm.ListBox1.SelectedItems.Item(ix - 1))

                            ' portfolioName und variantName wieder durch # getrennt
                            Dim hstr() As String = Split(constFilterName, "[")
                            If hstr.Length > 1 Then
                                constFilterName = hstr(0) & "#" & deleteBrackets(hstr(1), "[", "]")
                            End If

                            Dim pname As String = getPnameFromKey(constFilterName)
                            Dim vname As String = getVariantnameFromKey(constFilterName)
                            Dim constellation As clsConstellation = projectConstellations.getConstellation(pname, vname)

                            If Not IsNothing(constellation) Then

                                Dim ok As Boolean = False
                                If (Not AlleProjekte.containsAnySummaryProject _
                                    And Not projectConstellations.getConstellation(pname, vname).containsAnySummaryProject) Then
                                    ' alles in Ordnung 
                                    ok = True
                                Else
                                    If Not ShowProjekte.hasAnyConflictsWith(pname, True) Then
                                        ok = True
                                    End If
                                End If

                                If ok Then
                                    ' aufnehmen ...
                                    'Dim constellation As clsConstellation = projectConstellations.getConstellation(pname, vname)

                                    If Not IsNothing(constellation) Then
                                        If Not constellationsToDo.Contains(constellation.constellationName) Then
                                            If Not constellationsToDo.hasAnyConflictsWith(constellation) Then
                                                constellationsToDo.Add(constellation)
                                            Else
                                                Call MsgBox("keine Aufnahme wegen Konflikten (gleiche Projekte enthalten): " & vbLf &
                                                    constellation.constellationName)
                                            End If

                                        End If
                                    End If

                                    ' war vorher ..
                                    If Not IsNothing(constellation) Then
                                        projectConstellations.addToLoadedSessionPortfolios(constellation.constellationName, constellation.variantName)
                                    End If

                                Else
                                    ' Meldung, und dann nicht aufnehmen 
                                    Call MsgBox("Konflikte zwischen Summary Projekten und Projekten ... doppelte Nennungen ..." & vbLf &
                                     "vermeiden Sie es, Platzhalter Summary Projekte und Projekte, die bereits in den Summary Projekten referenziert sind")
                                End If
                            End If

                        Catch ex As Exception
                            Dim tstmsg As String = ex.Message
                        End Try

                        'If awinSettings.englishLanguage Then
                        '    outputLine = constFilterName & " loaded ..."
                        'Else
                        '    outputLine = constFilterName & " wurde geladen ..."
                        'End If
                        'outputCollection.Add(outputLine)
                    Next

                    Dim clearBoard As Boolean = True
                    Dim clearSession As Boolean = False
                    If constellationsToDo.Count > 0 Then
                        Call showConstellations(constellationsToDo, clearBoard, clearSession, Date.Now, showSummaryProject:=False, onlySessionLoad:=(control.Id = loadfromSession))
                    End If

                    ' jetzt muss untersucht werden, ob der Fenster-Ausschnitt einigermaßen passt ... 
                    ' Window so positionieren, dass die Projekte sichtbar sind ...  
                    ' aber nur tun, wenn vorher nisht drin war ...
                    If boardWasEmpty Then
                        If ShowProjekte.Count > 0 Then
                            Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                            If clearBoard Then
                                If leftborder - 12 > 0 Then
                                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                                Else
                                    appInstance.ActiveWindow.ScrollColumn = 1
                                End If
                            End If
                        End If
                    End If


                    appInstance.ScreenUpdating = True

                    Cursor.Current = Cursors.Default
                    appInstance.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault
                End If

                If ControlID = deleteFilter Then

                    Dim removeOK As Boolean = False
                    Dim filter As clsFilter = Nothing

                    constFilterName = removeConstFilterFrm.ListBox1.Text

                    filter = filterDefinitions.retrieveFilter(constFilterName)

                    If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                        ' Filter muss aus der Datenbank gelöscht werden.

                        removeOK = CType(databaseAcc, DBAccLayer.Request).removeFilterFromDB(filter)
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
            ' tk 28.7.19 Beim Löschen von Portfolios ergänzt 
            If ControlID = deleteFromSession Then
                If outputCollection.Count > 0 Then
                    Dim header As String = "Löschen von Portfolios"
                    If awinSettings.englishLanguage Then
                        header = "Delete Portfolios"
                    End If
                    Call showOutPut(outputCollection, header:=header, explanation:="")
                End If
            ElseIf ControlID = loadfromSession Then
                If outputCollection.Count > 0 Then
                    Dim header As String = "Laden von Portfolios"
                    If awinSettings.englishLanguage Then
                        header = "Load Portfolios"
                    End If
                    Call showOutPut(outputCollection, header:=header, explanation:="")
                End If
            End If

        End If




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
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim errmsg As New clsErrorCodeMsg
        Dim DBtimeStamp As Date = Date.Now
        Dim sessionPortfolioNames As New SortedList(Of String, String)
        Dim dbPortfolioNames As New SortedList(Of String, String)
        Dim outPutCollection As New Collection
        Dim outPutLine As String = ""


        If projectConstellations.Liste.Count <= 0 Then
            If awinSettings.englishLanguage Then
                outPutLine = "No Portfolios loaded"
            Else
                outPutLine = "Es ist kein Portfolio geladen"
            End If
            Call MsgBox(outPutLine)
        Else
            ' speichern von Portfolios nur möglich, wenn welche geladen sind
            With storeConstellationFrm
                If awinSettings.englishLanguage Then
                    .Text = "save Portfolio(s) to Datenbase"
                Else
                    .Text = "Portfolio(s) in Datenbank speichern"
                End If
                For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste
                    sessionPortfolioNames.Add(kvp.Key, kvp.Value.variantName)
                Next
                .constellationsToShow = sessionPortfolioNames
                .retrieveFromDB = False
                .lblStandvom.Visible = False
                .requiredDate.Visible = False
                .addToSession.Visible = False
                .loadAsSummary.Visible = False

            End With

            Dim returnValue As DialogResult = storeConstellationFrm.ShowDialog()
            If returnValue = DialogResult.OK Then

                Dim clearBoard As Boolean = Not storeConstellationFrm.addToSession.Checked
                Dim showSummaryProjects As Boolean = storeConstellationFrm.loadAsSummary.Checked


                'If Not IsNothing(storeConstellationFrm.requiredDate.Value) Then
                '    storedAtOrBefore = CDate(storeConstellationFrm.requiredDate.Value).Date.AddHours(23).AddMinutes(59)
                'Else
                '    storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)
                'End If


                Dim constellationsToDo As New clsConstellations

                ' Liste der ausgewählten Portfolio/Variante Paaren (pro Portfolio nur eine Variante)
                Dim constellationsChecked As New SortedList(Of String, String)

                ' WaitCursor einschalten ...
                Cursor.Current = Cursors.WaitCursor

                If clearBoard Then
                    '' nichts zu speichern
                End If
                '' es muss schon unterschieden werden, ob nur von Session geladen werden soll 
                'If loadFromSession Then
                '        currentSessionConstellation.Liste.Clear()
                '    Else
                '        AlleProjekte.Clear(updateCurrentConstellation:=True)
                '    End If

                '    projectConstellations.clearLoadedPortfolios()
                'End If

                ' liste, welche Portfolios und Portfolio-Varianten gespeichert werden soll, wird erstellt
                constellationsChecked = New SortedList(Of String, String)

                For Each tNode As TreeNode In storeConstellationFrm.TreeViewPortfolios.Nodes
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
                            Call MsgBox("Error in Portfolio-Auswahl")
                        End If
                    End If
                Next
                If constellationsChecked.Count = 1 Then
                    ' Liste der Portfolios in der DB
                    dbPortfolioNames = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errmsg)
                    Dim constellationName As String = constellationsChecked.ElementAt(0).Key
                    Dim vname As String = constellationsChecked.ElementAt(0).Value
                    Dim currentConstellation As clsConstellation = projectConstellations.getConstellation(constellationName, vname)
                    Call storeSingleConstellationToDB(outPutCollection, currentConstellation, dbPortfolioNames)
                End If


                '' ur:13.12.2019
                ''Dim dbConstellations As clsConstellations = CType(databaseAcc, DBAccLayer.Request).retrieveConstellationsFromDB(Date.Now, errMsg)
                'dbPortfolioNames = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

                ''For i As Integer = 1 To storeConstellationFrm.ListBox1.SelectedItems.Count

                ''    Dim constellationName As String = CStr(storeConstellationFrm.ListBox1.SelectedItems.Item(i - 1))
                ''    Dim currentConstellation As clsConstellation = projectConstellations.getConstellation(constellationName)

                ''    Call storeSingleConstellationToDB(outPutCollection, currentConstellation, dbPortfolioNames)

                ''Next

                If outPutCollection.Count > 0 Then
                    Dim msgH As String, msgE As String
                    If awinSettings.englishLanguage Then
                        msgH = "Save Portfolios"
                        msgE = "following results:"
                    Else
                        msgH = "Speichern Portfolio(s"
                        msgE = "Rückmeldungen"

                    End If

                    Call showOutPut(outPutCollection, msgH, msgE)

                End If
            End If
        End If

    End Sub

    Sub PTLoadStoreRemoveConstellationFromDB(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

        ' Timer
        Dim sw As clsStopWatch
        sw = New clsStopWatch
        sw.StartTimer()

        Dim load1FromDatenbank As String = "PT5G1B1"
        Dim load2FromDatenbank As String = "PT5G1"
        Dim deleteFromDatenbank As String = "Pt5G3B1"

        Dim loadConstellationFrm As New frmLoadConstellation
        Dim storedAtOrBefore As Date = Date.Now.Date.AddHours(23).AddMinutes(59)
        Dim ControlID As String = control.Id
        Dim timeStampsCollection As New Collection
        'Dim dbConstellations As New clsConstellations
        Dim dbPortfolioNames As New SortedList(Of String, String)
        Dim cTimestamp As Date
        Dim initMessage As String = "Es sind dabei folgende Probleme aufgetreten" & vbLf & vbLf

        Dim deleteFromDB As Boolean = (control.Id = "Pt5G3B1")
        Dim outPutCollection As New Collection
        Dim outputLine As String = ""

        Dim successMessage As String = initMessage
        Dim returnValue As DialogResult

        Dim boardWasEmpty As Boolean = ShowProjekte.Count = 0

        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        Call projektTafelInit()


        ' Wenn das Laden eines Portfolios aus dem Menu Datenbank aufgerufen wird, so werden erneut alle Portfolios aus der Datenbank geholt

        If (ControlID = load1FromDatenbank Or ControlID = load2FromDatenbank Or ControlID = deleteFromDatenbank) _
            And Not noDB Then

            If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                dbPortfolioNames = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, err)
                'dbConstellations = CType(databaseAcc, DBAccLayer.Request).retrieveConstellationsFromDB(Date.Now, err)

                If dbPortfolioNames.Count > 0 Then

                    Try
                        timeStampsCollection = CType(databaseAcc, DBAccLayer.Request).retrieveZeitstempelFromDB()
                        'Dim heute As String = Date.Now.ToString
                        If timeStampsCollection.Count > 0 Then
                            With loadConstellationFrm
                                If deleteFromDB Then
                                    If awinSettings.englishLanguage Then
                                        .Text = "Delete Portfolio"
                                    Else
                                        .Text = "Portfolio Löschen"
                                    End If
                                End If
                                .constellationsToShow = dbPortfolioNames
                                '.constellationsToShow = dbConstellations
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

                    Try
                        enableOnUpdate = False

                        If AlleProjekte.Count > 0 Then
                            loadConstellationFrm.addToSession.Checked = False
                        Else
                            loadConstellationFrm.addToSession.Checked = False
                            loadConstellationFrm.addToSession.Visible = False
                        End If

                        If deleteFromDB Then
                            loadConstellationFrm.addToSession.Checked = False
                            loadConstellationFrm.addToSession.Visible = False
                            loadConstellationFrm.loadAsSummary.Checked = False
                            loadConstellationFrm.loadAsSummary.Visible = False
                        End If

                        'Call MsgBox("PTLadenKonstellation 1st Part took: " & sw.EndTimer & "milliseconds")

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager And AlleProjekte.Count = 0 Then
                            loadConstellationFrm.loadAsSummary.Visible = True
                        Else
                            loadConstellationFrm.loadAsSummary.Visible = False
                        End If

                        returnValue = loadConstellationFrm.ShowDialog

                        sw.StartTimer()

                        If returnValue = DialogResult.OK Then

                            Dim clearBoard As Boolean = Not loadConstellationFrm.addToSession.Checked
                            Dim showSummaryProjects As Boolean = loadConstellationFrm.loadAsSummary.Checked


                            If Not IsNothing(loadConstellationFrm.requiredDate.Value) Then
                                storedAtOrBefore = CDate(loadConstellationFrm.requiredDate.Value).Date.AddHours(23).AddMinutes(59)
                            Else
                                storedAtOrBefore = Date.Now.Date.AddHours(23).AddMinutes(59)
                            End If


                            Dim constellationsToDo As New clsConstellations

                            ' Liste der ausgewählten Portfolio/Variante Paaren (pro Portfolio nur eine Variante)
                            Dim constellationsChecked As New SortedList(Of String, String)

                            ' WaitCursor einschalten ...
                            Cursor.Current = Cursors.WaitCursor

                            If clearBoard Then
                                ' es muss schon unterschieden werden, ob nur von Session geladen werden soll 
                                AlleProjekte.Clear(updateCurrentConstellation:=True)
                                projectConstellations.clearLoadedPortfolios()
                            End If

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
                                        Call MsgBox("Error in Portfolio-Auswahl")
                                    End If
                                End If
                            Next

                            If deleteFromDB Then
                                For Each pvName As KeyValuePair(Of String, String) In constellationsChecked

                                    Dim pName As String = pvName.Key        'portfolio-Name
                                    Dim vName As String = pvName.Value      'variantenName

                                    Try
                                        ' lösche Portfolio (pName,vName) aus der db
                                        Dim result As Boolean = CType(databaseAcc, DBAccLayer.Request).removeConstellationFromDB(pName,
                                                                                                         dbPortfolioNames(pName),
                                                                                                         vName,
                                                                                                        err)
                                        If awinSettings.englishLanguage Then
                                            If result Then
                                                outputLine = pName & "[" & vName & "] deleted"
                                            Else
                                                outputLine = pName & "[" & vName & "] couldn't be deleted"
                                            End If
                                            outPutCollection.Add(outputLine)
                                        Else
                                            If result Then
                                                outputLine = pName & "[" & vName & "] gelöscht"
                                            Else
                                                outputLine = pName & "[" & vName & "] konnte nicht gelöscht werden"
                                            End If
                                            outPutCollection.Add(outputLine)
                                        End If


                                    Catch ex As Exception
                                        outputLine = ex.Message
                                        outPutCollection.Add(outputLine)
                                    End Try
                                Next

                                If outPutCollection.Count > 0 Then
                                    Dim msgH As String, msgE As String
                                    If awinSettings.englishLanguage Then
                                        msgH = "Delete Portfolios"
                                        msgE = "following results:"
                                    Else
                                        msgH = "Löschen Portfolio/s"
                                        msgE = "Rückmeldungen"

                                    End If

                                    Call showOutPut(outPutCollection, msgH, msgE)
                                End If

                            Else

                                For Each pvName As KeyValuePair(Of String, String) In constellationsChecked

                                    Dim pName As String = pvName.Key        'portfolio-Name
                                    Dim vName As String = pvName.Value      'variantenName

                                    ' Plausibilitätsprüfung: darf das geladen werden 
                                    Try
                                        ' Check ...
                                        'Dim checkconst As clsConstellation = projectConstellations.getConstellation(tmpName)
                                        Dim checkconst As clsConstellation = Nothing

                                        ' pName ist nicht mehr in der Session geladen
                                        If IsNothing(checkconst) Then

                                            ' hole Portfolio (pName,vName) aus der db
                                            checkconst = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(pName,
                                                                                                                               dbPortfolioNames(pName),
                                                                                                                               cTimestamp, err,
                                                                                                                               variantName:=vName,
                                                                                                                               storedAtOrBefore:=storedAtOrBefore)

                                            If Not IsNothing(checkconst) Then
                                                ' tmpname in die Session-Liste wieder aufnehmen
                                                projectConstellations.Add(checkconst)
                                            Else
                                                Call MsgBox("Portfolio nicht mehr vorhanden!")
                                            End If

                                        End If

                                        If Not IsNothing(projectConstellations.getConstellation(pName, vName)) Then

                                            Dim ok As Boolean = False
                                            If (Not AlleProjekte.containsAnySummaryProject _
                                                And Not projectConstellations.getConstellation(pName, vName).containsAnySummaryProject _
                                                And Not loadConstellationFrm.loadAsSummary.Checked) Or clearBoard Then
                                                ' alles in Ordnung 
                                                ok = True
                                            Else
                                                If Not ShowProjekte.hasAnyConflictsWith(pName, True) Then
                                                    ok = True
                                                End If
                                            End If

                                            If ok Then
                                                ' aufnehmen ...
                                                Dim constellation As clsConstellation = projectConstellations.getConstellation(pName, vName)

                                                If Not IsNothing(constellation) Then
                                                    If Not constellationsToDo.Contains(constellation.constellationName) Then
                                                        If Not constellationsToDo.hasAnyConflictsWith(constellation) Then
                                                            constellationsToDo.Add(constellation)
                                                        Else
                                                            Call MsgBox("keine Aufnahme wegen Konflikten (gleiche Projekte enthalten): " & vbLf &
                                                                            constellation.constellationName)
                                                        End If

                                                    End If

                                                Else
                                                    ' hole Portfolio (pName,vName) aus den db
                                                    constellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(pName,
                                                                                                                               dbPortfolioNames(pName),
                                                                                                                               cTimestamp, err,
                                                                                                                               variantName:=vName,
                                                                                                                               storedAtOrBefore:=storedAtOrBefore)
                                                    If Not IsNothing(constellation) Then
                                                        If Not constellationsToDo.Contains(constellation.constellationName) Then
                                                            If Not constellationsToDo.hasAnyConflictsWith(constellation) Then
                                                                constellationsToDo.Add(constellation)
                                                            Else
                                                                Call MsgBox("keine Aufnahme wegen Konflikten (gleiche Projekte enthalten): " & vbLf &
                                                                            constellation.constellationName)
                                                            End If

                                                        End If
                                                        projectConstellations.Add(constellation)
                                                    End If


                                                    ' tk jetzt muss für jedes der items, das ein Portfolio ist, dieses in die Liste eintragen 
                                                    'If constellation.containsAnySummaryProject Then
                                                    '    For Each spKvP As KeyValuePair(Of String, clsConstellationItem) In constellation.Liste
                                                    '        Dim tmpProj As clsProjekt = getProjektFromSessionOrDB(spKvP.Value.projectName, spKvP.Value.variantName, AlleProjekte, Date.Now)
                                                    '        If Not IsNothing(tmpProj) Then
                                                    '            If Not AlleProjekte.Containskey(spKvP.Key) Then
                                                    '                AlleProjekte.Add(tmpProj, )
                                                    '            End If
                                                    '        End If
                                                    '        If spKvP.Value.variantName = portfolioVName Then
                                                    '            projectConstellations.addToLoadedSessionPortfolios(spKvP.Key)
                                                    '        End If
                                                    '    Next
                                                    'Else
                                                    '    If Not IsNothing(constellation) Then
                                                    '        projectConstellations.addToLoadedSessionPortfolios(constellation.constellationName)
                                                    '    End If
                                                    'End If

                                                    ' war vorher ..
                                                    If Not IsNothing(constellation) Then
                                                        projectConstellations.addToLoadedSessionPortfolios(constellation.constellationName, constellation.variantName)
                                                    End If
                                                End If

                                            Else
                                                ' Meldung, und dann nicht aufnehmen 
                                                Call MsgBox("Konflikte zwischen Summary Projekten und Projekten ... doppelte Nennungen ..." & vbLf &
                                                             "vermeiden Sie es, Platzhalter Summary Projekte und Projekte, die bereits in den Summary Projekten referenziert sind")
                                            End If
                                        End If

                                    Catch ex As Exception
                                        Dim tstmsg As String = ex.Message
                                    End Try

                                Next

                                sw.StartTimer()

                                'Dim clearSession As Boolean = (((ControlID = load1FromDatenbank) Or (ControlID = load2FromDatenbank)) And clearBoard)
                                Dim clearSession As Boolean = False
                                If constellationsToDo.Count > 0 Then

                                    Call showConstellations(constellationsToDo, clearBoard, clearSession, storedAtOrBefore, showSummaryProject:=showSummaryProjects)

                                    ' Timer
                                    If awinSettings.visboDebug Then
                                        Call MsgBox("PTLadenKonstellation 2nd Part took: " & sw.EndTimer & "milliseconds")
                                    End If


                                    ' jetzt muss die Info zu den Schreibberechtigungen geholt werden 
                                    ' aber nur, wenn es nicht nur von der Session geholt wird  
                                    If Not noDB Then
                                        writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)
                                    End If
                                End If

                                ' jetzt muss untersucht werden, ob der Fenster-Ausschnitt einigermaßen passt ... 
                                ' Window so positionieren, dass die Projekte sichtbar sind ...  
                                If boardWasEmpty Then
                                    If ShowProjekte.Count > 0 Then
                                        Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                                        If clearBoard Then
                                            If leftborder - 12 > 0 Then
                                                appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                                            Else
                                                appInstance.ActiveWindow.ScrollColumn = 1
                                            End If
                                        End If
                                    End If
                                End If


                                appInstance.ScreenUpdating = True

                                Cursor.Current = Cursors.Default
                            End If



                        End If




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
                Call MsgBox("Datenbank-Verbindung ist unterbrochen !")
            End If
        Else
            Call MsgBox("ControlID = " & ControlID)

        End If



    End Sub

    Sub PTAendernKonstellation(control As IRibbonControl)

        Call PBBChangeCurrentPortfolio()

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
            appInstance.Cursor = Excel.XlMousePointer.xlDefault
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

            ' Mouse wieder auf Normalmodus setzen
            appInstance.Cursor = Excel.XlMousePointer.xlDefault

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
        appInstance.ScreenUpdating = False
        Try
            If visboZustaende.projectBoardMode = ptModus.graficboard Then
                Call deleteChartsInSheet(arrWsNames(ptTables.mptPfCharts))
                Call deleteChartsInSheet(arrWsNames(ptTables.mptPrCharts))
                Call deleteChartsInSheet(arrWsNames(ptTables.MPT))
                ' jetzt müssen alle Windows bis auf Window(0) = Multiprojekt-Tafel geschlossen werden 
                ' und mache ProjectboardWindows(mpt) great again ...
                Call closeAllWindowsExceptMPT()

            Else
                Call deleteChartsInSheet(arrWsNames(ptTables.meCharts))
            End If
        Catch ex As Exception
            Dim a As String = ex.Message
        End Try


        appInstance.ScreenUpdating = True

        If Not appInstance.EnableEvents Then
            appInstance.EnableEvents = True
        End If


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

            Dim anzPfDiagrams As Integer = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.mptPfCharts)).ChartObjects, Excel.ChartObjects).Count
            Dim anzPrDiagrams As Integer = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.mptPrCharts)).ChartObjects, Excel.ChartObjects).Count


            If anzPfDiagrams + anzPrDiagrams > 0 Then

                ' hier muss die Auswahl des Names für das Cockpit erfolgen

                returnValue = storeCockpitFrm.ShowDialog  ' Aufruf des Formulars zur Eingabe des Cockpitnamens

                If returnValue = DialogResult.OK Then

                    cockpitName = storeCockpitFrm.ComboBox1.Text

                    'ClearClipboard()

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
                        currentWsName = arrWsNames(ptTables.mptPrCharts)
                        Call deleteChartsInSheet(currentWsName)
                        currentWsName = arrWsNames(ptTables.mptPfCharts)
                        Call deleteChartsInSheet(currentWsName)
                        currentWsName = arrWsNames(ptTables.MPT)
                        Call deleteChartsInSheet(currentWsName)
                    Else
                        currentWsName = arrWsNames(ptTables.meRC)
                        Call deleteChartsInSheet(currentWsName)
                    End If

                    Call closeAllWindowsExceptMPT()

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

                    '' ''Call defineVisboWindowViews()
                    ' ''If thereAreAnyCharts(PTwindows.mpt) Then
                    ' ''    Call showVisboWindow(PTwindows.mpt)
                    ' ''End If

                    '' ''If thereAreAnyCharts(PTwindows.mptpf) Then
                    '' ''    Call showVisboWindow(PTwindows.mptpf)
                    '' ''End If
                    '' ''If thereAreAnyCharts(PTwindows.mptpr) Then
                    '' ''    Call showVisboWindow(PTwindows.mptpr)
                    '' ''End If

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

                ausgabeString = ausgabeString & hproj.name & ": " & hproj.tfZeile.ToString &
                                 " - " & vglWert.ToString & "; " &
                                 calcXCoordToDate(singleShp.Left).ToShortDateString & " vs. " & hproj.startDate.ToShortDateString &
                                 " vs. " & calcXCoordToDate(curCoord(1)).ToShortDateString & singleShp.Left.ToString & vbLf


            Next i


        End If

        Call awinDeSelect()
        Call MsgBox(ausgabeString)

        enableOnUpdate = True



    End Sub

    Sub PT0ShowProjektInfo1(control As IRibbonControl)

        With visboZustaende
            If IsNothing(formProjectInfo1) And (.projectBoardMode = ptModus.massEditRessSkills Or .projectBoardMode = ptModus.massEditCosts) Then

                formProjectInfo1 = New frmProjectInfo1
                Call updateProjectInfo1(visboZustaende.currentProject, visboZustaende.currentProjectinSession)

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

        Dim err As New clsErrorCodeMsg

        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange


        Dim phaseList As New Collection
        Dim milestoneList As New Collection
        Dim neuerVariantenName As String = ""
        Dim ok As Boolean = True
        Dim nameCollection As New Collection
        Dim abbruch As Boolean = False


        ' check ob auch keine Summary Projects selektiert sind ...

        If Not noSummaryProjectsareSelected(nameCollection) Then
            Exit Sub
        End If



        ' hier geht es los , wenn ok ... 
        Try
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

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
                                    If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, "", Date.Now, err) Or
                                        CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, hproj.variantName, Date.Now, err) Then

                                        ok = CType(databaseAcc, DBAccLayer.Request).renameProjectsInDB(pName, newName, dbUsername, err)
                                        If Not ok Then
                                            If awinSettings.englishLanguage Then
                                                Call MsgBox("rename cancelled: there is at least one write-protected variant for Project " & pName)
                                            Else
                                                Call MsgBox("Rename nicht durchgeführt: es gibt mindestens eine schreibgeschützte Variante im Projekt " & pName)
                                            End If
                                        Else
                                            writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)
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

                                        Call clearProjektinPlantafel(pName)

                                        Dim key As String = calcProjektKey(hproj)
                                        ShowProjekte.Remove(pName)

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

    Sub PT2ProjektNeuBasisProjekt(control As IRibbonControl)
        Dim err As New clsErrorCodeMsg
        Dim returnValue As DialogResult
        Dim newProj As clsProjekt = Nothing

        Dim weiterMachen As Boolean = False

        appInstance.EnableEvents = False
        enableOnUpdate = False

        ' tk 14.6 als erstes wird jetzt ein einziges Projekt gewählt ... 
        Dim selectProjectAsVorlage As New frmProjPortfolioAdmin
        Try
            If ShowProjekte.Count > 0 Then

                With selectProjectAsVorlage

                    .aKtionskennung = PTTvActions.loadProjectAsTemplate

                End With

                returnValue = selectProjectAsVorlage.ShowDialog

                If returnValue = DialogResult.OK Then

                    weiterMachen = True
                    Dim hproj As clsProjekt = selectProjectAsVorlage.selProjectAsTemplate
                    Dim idArray() As Integer = myCustomUserRole.getAggregationRoleIDs

                    newProj = hproj.aggregateForPortfolioMgr(idArray)

                    ' jetzt wird das Budget neu gesetzt, und zwar so, das es genau reicht ... 
                    Call newProj.setBudgetAsNeeded()

                    ' jetzt müssen Ampeln, -Bewertungen, %Done, Verantwortlichkeiten 
                    ' wobei - evtl muss das gar nicht gemacht werden, weil das ja im TrageIvProjekte gemacht wird ...  

                    Call newProj.resetTrafficLightsEtc()

                    ' actualDatauntil zurücksetzen
                    newProj.actualDataUntil = Date.MinValue

                Else
                    weiterMachen = False

                End If
            Else
                weiterMachen = False

            End If


        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try

        If weiterMachen Then
            Dim ProjektEingabe As New frmProjektEingabe1
            ProjektEingabe.existingProjAsTemplate = newProj

            Dim zeile As Integer = 0

            Dim pNrDoesNotExistYet As Boolean = True
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            Call projektTafelInit()


            enableOnUpdate = False


            returnValue = ProjektEingabe.ShowDialog

            If returnValue = DialogResult.OK Then
                With ProjektEingabe

                    Dim profitUserAskedFor As String = Nothing
                    If IsNumeric(.profitAskedFor.Text) Or .profitAskedFor.Text = "" Then
                        profitUserAskedFor = .profitAskedFor.Text
                    End If


                    If Not noDB Then

                        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                            If .txtbx_pNr.Text <> "" Then

                                Try
                                    pNrDoesNotExistYet = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(.txtbx_pNr.Text, err).Count = 0
                                Catch ex As Exception

                                End Try

                            End If

                            If pNrDoesNotExistYet Then

                                If Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(projectname:= .projectName.Text, variantname:="", storedAtorBefore:=Date.Now, err:=err) Then

                                    ' Projekt existiert noch nicht in der DB, kann also eingetragen werden



                                    Call TrageivProjektein(newProj, .projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart),
                                                       CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile,
                                                       5.0, 5.0, profitUserAskedFor,
                                                       CStr(.txtbx_description.Text), CStr(.txtbx_pNr.Text))
                                Else
                                    Call MsgBox(" Projekt '" & .projectName.Text & "' existiert bereits in der Datenbank!")
                                End If

                            Else
                                Call MsgBox(" Projekt-Nummer '" & .txtbx_pNr.Text & "' existiert bereits in der Datenbank!")
                            End If


                        Else

                            Call MsgBox("Datenbank- Verbindung ist unterbrochen !")
                            appInstance.ScreenUpdating = True

                            ' Projekt soll trotzdem angezeigt werden
                            Call TrageivProjektein(newProj, .projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart),
                                                   CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile,
                                                   5.0, 5.0, profitUserAskedFor,
                                                   CStr(.txtbx_description.Text), CStr(.txtbx_pNr.Text))

                        End If

                    Else

                        appInstance.ScreenUpdating = True

                        ' Projekt soll trotzdem angezeigt werden
                        Call TrageivProjektein(newProj, .projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart),
                                                   CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile,
                                                   5.0, 5.0, profitUserAskedFor,
                                                   CStr(.txtbx_description.Text), CStr(.txtbx_pNr.Text))

                    End If

                End With
            End If

            ''If Not currentConstellationName.EndsWith("(*)") Then
            ''    currentConstellationName = currentConstellationName & " (*)"
            ''End If

            ' es kam jetzt ein neues Projekt hinzu, also muss das Sort-Kriterium umgesetzt werden auf customtF, massgabe ist jetzt einfach die Zeile, in der die PRojekte stehen 
            currentSessionConstellation.sortCriteria = ptSortCriteria.customTF

            If currentConstellationPvName <> calcLastSessionScenarioName() Then
                currentConstellationPvName = calcLastSessionScenarioName()
            End If
        Else
            If ShowProjekte.Count <= 0 Then
                If awinSettings.englishLanguage Then
                    Call MsgBox("Please, load a project")
                Else
                    Call MsgBox("Bitte laden Sie ein Projekt")
                End If
            End If
        End If


        enableOnUpdate = True
        appInstance.EnableEvents = True


    End Sub

    Sub PT2ProjektNeu(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

        Dim ProjektEingabe As New frmProjektEingabe1
        ' es wird kein existierendes Projekt als Template gewählt .. .
        ProjektEingabe.existingProjAsTemplate = Nothing

        Dim returnValue As DialogResult
        Dim zeile As Integer = 0

        Dim pNrDoesNotExistYet As Boolean = True
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Call projektTafelInit()


        enableOnUpdate = False

        If Projektvorlagen.Count = 0 Then
            Dim msgStr As String = ""
            If awinSettings.englishLanguage Then
                msgStr = "Error: No Template found! - Creating Project from Template is not possible. "
            Else
                msgStr = "Fehler: Keine Vorlage vorhanden! - Neues Projekt auf Basis Vorlage nicht möglich."
            End If
            Call logger(ptErrLevel.logError, "PT2ProjektNeu", msgStr)
            Call MsgBox(msgStr)
        Else

            returnValue = ProjektEingabe.ShowDialog

            If returnValue = DialogResult.OK Then
                With ProjektEingabe

                    Dim profitUserAskedFor As String = Nothing
                    If IsNumeric(.profitAskedFor.Text) Or .profitAskedFor.Text = "" Then
                        profitUserAskedFor = .profitAskedFor.Text
                    End If


                    If Not noDB Then

                        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

                            If .txtbx_pNr.Text <> "" Then

                                Try
                                    pNrDoesNotExistYet = CType(databaseAcc, DBAccLayer.Request).retrieveProjectNamesByPNRFromDB(.txtbx_pNr.Text, err).Count = 0
                                Catch ex As Exception

                                End Try

                            End If

                            If pNrDoesNotExistYet Then

                                If Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(projectname:= .projectName.Text, variantname:="", storedAtorBefore:=Date.Now, err:=err) Then

                                    ' Projekt existiert noch nicht in der DB, kann also eingetragen werden


                                    Call TrageivProjektein(Nothing, .projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart),
                                                   CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile,
                                                   5.0, 5.0, profitUserAskedFor,
                                                   CStr(.txtbx_description.Text), CStr(.txtbx_pNr.Text))
                                Else
                                    Call MsgBox(" Projekt '" & .projectName.Text & "' existiert bereits in der Datenbank!")
                                End If

                            Else
                                Call MsgBox(" Projekt-Nummer '" & .txtbx_pNr.Text & "' existiert bereits in der Datenbank!")
                            End If


                        Else

                            Call MsgBox("Datenbank- Verbindung ist unterbrochen !")
                            appInstance.ScreenUpdating = True

                            ' Projekt soll trotzdem angezeigt werden
                            Call TrageivProjektein(Nothing, .projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart),
                                               CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile,
                                               5.0, 5.0, profitUserAskedFor,
                                               CStr(.txtbx_description.Text), CStr(.txtbx_pNr.Text))

                        End If

                    Else

                        appInstance.ScreenUpdating = True

                        ' Projekt soll trotzdem angezeigt werden
                        Call TrageivProjektein(Nothing, .projectName.Text, .vorlagenDropbox.Text, CDate(.calcProjektStart),
                                               CDate(.calcProjektEnde), CType(.Erloes.Text, Double), zeile,
                                               5.0, 5.0, profitUserAskedFor,
                                               CStr(.txtbx_description.Text), CStr(.txtbx_pNr.Text))

                    End If

                End With
            End If

            ''If Not currentConstellationName.EndsWith("(*)") Then
            ''    currentConstellationName = currentConstellationName & " (*)"
            ''End If

            ' es kam jetzt ein neues Projekt hinzu, also muss das Sort-Kriterium umgesetzt werden auf customtF, massgabe ist jetzt einfach die Zeile, in der die PRojekte stehen 
            currentSessionConstellation.sortCriteria = ptSortCriteria.customTF

            If currentConstellationPvName <> calcLastSessionScenarioName() Then
                currentConstellationPvName = calcLastSessionScenarioName()
            End If

            'Call storeSessionConstellation("Last")
        End If

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

        Dim nameCollection As New Collection

        If Not noSummaryProjectsareSelected(nameCollection) Then
            Exit Sub
        End If

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

        Dim nameCollection As New Collection

        If Not noSummaryProjectsareSelected(nameCollection) Then
            Exit Sub
        End If

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

        Dim nameCollection As New Collection

        If Not noSummaryProjectsareSelected(nameCollection) Then
            Exit Sub
        End If

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

        Dim nameCollection As New Collection

        If Not noSummaryProjectsareSelected(nameCollection) Then
            Exit Sub
        End If


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

                ' jetzt prüfen: macht nur Sinn, wenn es nicht bereits die Base-Variant ist ... 
                If hproj.variantName = "" Then
                    If awinSettings.englishLanguage Then
                        Call MsgBox("The project " & hproj.name & " is already the base-variant")

                    Else
                        Call MsgBox("Projekt " & hproj.name & " ist bereits die Standard-Variante")
                    End If
                Else
                    If tryToprotectProjectforMe(hproj.name, "") Then
                        ' ist erlaubt ...
                        ' das Projekt zur Standard Variante machen 


                        Dim oldvName As String = hproj.variantName
                        Dim newvName As String = ""

                        Try
                            Dim oldStatus As String = getStatusOfBaseVariant(hproj.name, hproj.Status)
                            ' Plausibilitätsprüfung, es dürfen keine abgebrochenen / abgeschlossenen Projekte überschrieben werden   

                            If oldStatus <> ProjektStatus(PTProjektStati.abgebrochen) And
                                oldStatus <> ProjektStatus(PTProjektStati.abgeschlossen) Then

                                ' nur dann darf die Variante übernommen werden ... 
                                If oldStatus = ProjektStatus(PTProjektStati.beauftragt) Then
                                    ' tk 23.4.19 bis auf weiteres soll das ohne ChangeRequest auskommen 
                                    ' muss noch überdacht werden 
                                    'hproj.Status = ProjektStatus(PTProjektStati.ChangeRequest)
                                    hproj.Status = oldStatus
                                Else
                                    hproj.Status = oldStatus
                                End If

                                ' die aktuelle Variante aus der AlleProjekte rausnehmen 
                                key = calcProjektKey(hproj)
                                AlleProjekte.Remove(key)

                                ' das bisherige Standard Projekt aus der AlleProjekte rausnehmen 
                                key = calcProjektKey(hproj.name, "")
                                AlleProjekte.Remove(key)
                                ShowProjekte.Remove(hproj.name)

                                Dim bisherigeBaseVariant As clsProjekt = getProjektFromSessionOrDB(hproj.name, "", AlleProjekte, Date.Now, hproj.kundenNummer)

                                ' wenn es sich um einen Ressourcen-Manager handelt, dann muss das, was er geändert hat in die bisherige Basis Variante gemerged werden 
                                If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then

                                    Dim mergedProj As clsProjekt = Nothing
                                    Dim summaryRoleIDs As New Collection
                                    summaryRoleIDs.Add(myCustomUserRole.specifics)

                                    If Not IsNothing(bisherigeBaseVariant) Then

                                        ' Merge der geänderten Ressourcen => neues Projekt "mergeProj"
                                        mergedProj = bisherigeBaseVariant.deleteAndMerge(summaryRoleIDs, Nothing, hproj)
                                        If Not IsNothing(mergedProj) Then
                                            hproj = mergedProj
                                        End If

                                    End If


                                End If

                                'jetzt die aktuelle Variante zur Standard Variante machen 
                                ' dabei muss sichergestellt sein, dass der Status der bisherigen Basis-Variante übernommen wird 
                                hproj.variantName = ""

                                ' notwendig, um den Speicher-Conflic 409 zu vermeinden 
                                If Not IsNothing(bisherigeBaseVariant) Then
                                    hproj.updatedAt = bisherigeBaseVariant.updatedAt

                                    '' tk 25.7.19 
                                    'If Not hproj.isIdenticalTo(bisherigeBaseVariant) Then
                                    '    hproj.marker = True
                                    'End If
                                End If

                                hproj.timeStamp = Date.Now


                                ' die "neue" Standard Variante in AlleProjekte und ShowProjekte aufnehmen 
                                AlleProjekte.Add(hproj)
                                ShowProjekte.Add(hproj)

                                ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                                ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                                Dim tmpCollection As New Collection
                                Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, hproj.tfZeile, tmpCollection, tmpCollection)

                                ' jetzt müssen noch alle Projekt-Constellationen aktualisiert werden 
                                Call projectConstellations.updateVariantName(hproj.name, oldvName, newvName)


                            Else
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("no changes allowed with finalized / stopped project!")

                                Else
                                    Call MsgBox("an einem abgeschlossenen / abgebrochenen Projekt sind keine Änderungen möglich!")
                                End If
                            End If


                        Catch ex As Exception
                            Call MsgBox(ex.Message)
                        End Try


                    Else
                        ' ist nicht erlaubt ... 
                        If awinSettings.englishLanguage Then
                            Call MsgBox("The base variant of project " & hproj.name & " is protected" & vbLf &
                                        "and cannot be replaced by another variant")

                        Else
                            Call MsgBox("Projekt " & hproj.name & " ist in der Standard-Variante geschützt" & vbLf &
                                        "und kann daher nicht von einer anderen Variante überschrieben werden")
                        End If
                    End If
                End If

                ' jetzt prüfen : die Variante kann nur dann zur Standard-Variante gemacht werden, 
                ' wenn die Standard-Variante nicht geschützt ist ..

            Next i

            If currentConstellationPvName <> calcLastSessionScenarioName() Then
                currentConstellationPvName = calcLastSessionScenarioName()
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
    ''' aktiviert , je nach Modus die entsprechenden Ribbon Controls 
    ''' </summary>
    ''' <param name="modus"></param>
    ''' <remarks></remarks>
    Private Sub enableControls(ByVal modus As ptModus)

        If modus = ptModus.graficboard Then
            visboZustaende.projectBoardMode = modus
            Call visboZustaende.clearAuslastungsArray()


        Else
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
            Case "PT4G1M1-1"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Import Konfiguration"
                Else
                    tmpLabel = "Import Configuration"
                End If
            Case "PT4G1M1-2"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Import Einzelprojekte"
                Else
                    tmpLabel = "Import single Projects"
                End If
            Case "PT4G1M1-3"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Import Portfolios und Projektlisten"
                Else
                    tmpLabel = "Import Portfolios and project mass data"
                End If
            Case "PT4G2M-1"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Export Einzelprojekte"
                Else
                    tmpLabel = "Export single Projects"
                End If
            Case "PT4G2M-2"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Export Portfolios und Projektinformationen"
                Else
                    tmpLabel = "Export Portfolios and projects"
                End If
            Case "PT4G1B12"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Import Ist-Daten"
                Else
                    tmpLabel = "Import Actual Data"
                End If

            Case "PT4G1B13"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Kapazitäten"
                Else
                    tmpLabel = "Capacities"
                End If

            Case "PT4G1B14"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Offline Ressourcen Zuweisung"
                Else
                    tmpLabel = "Offline Resource Assignments"
                End If
            Case "PT4G1B15"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Darstellungsklassen"
                Else
                    tmpLabel = "Appearances"
                End If
            Case "PT4G1B16"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "kundenspez. Einstellungen"
                Else
                    tmpLabel = "Customization"
                End If

            Case "PT4G1B17"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projektvorlagen"
                Else
                    tmpLabel = "Project Templates"
                End If

            Case "PT4G1B8"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Organisation"
                Else
                    tmpLabel = "Organisation"
                End If


            Case "PT4G1B11"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Custom Nutzer Rollen"
                Else
                    tmpLabel = "Custom User Roles"
                End If

            Case "PTproj" ' Project
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt"
                Else
                    tmpLabel = "Project"
                End If
            Case "PTMEC" ' Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Realtime Cockpit"
                Else
                    tmpLabel = "Realtime Cockpit"
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
                    tmpLabel = "Vergleich mit letzter Version"
                Else
                    tmpLabel = "Comparison with last version"
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
                    tmpLabel = "Analyses"
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

                'Case "PT3G1M1" ' Planelemente visualisieren
                '    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '        tmpLabel = "Phasen/Meilensteine..."
                '    Else
                '        tmpLabel = "Phases/Milestones..."
                '    End If

            Case "PTXG1B4" ' Auswahl über Namen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auswahl über Namen..."
                Else
                    tmpLabel = "Select by Names..."
                End If

            Case "PTXG1B5" ' Planelemente visualisieren
                'Case "PT3G1M1" ' Planelemente visualisieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Phasen/Meilensteine..."
                Else
                    tmpLabel = "Phases/Milestones..."
                End If
                'If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '    tmpLabel = "Auswahl über Projekt-Struktur..."
                'Else
                '    tmpLabel = "Select by Structure..."
                'End If
            Case "PTXG1B9" ' Cash-Flow zeigen

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Veränderung Liquidität"
                Else
                    tmpLabel = "Change Liquidity"
                End If

            Case "PTOPTB1" ' Optimieren 
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio Optimieren"
                Else
                    tmpLabel = "Portfolio Optimization"
                End If

            Case "PTPf" ' Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Charts"
                Else
                    tmpLabel = "Charts"
                End If

                'Case "PTXG1M2" ' Engpass Analyse
                '    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '        tmpLabel = "Rollen/Kosten/Meilensteine/Phasen"
                '    Else
                '        tmpLabel = "Ressources/Costs/Milestones/Phases"
                '    End If

                'Case "PTXG1B6" ' Auswahl über Namen
                '    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '        tmpLabel = "Auswahl über Namen..."
                '    Else
                '        tmpLabel = "Select by Names..."
                '    End If

            Case "PTXG1B7" ' Leistbarkeitscharts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Rollen/Kosten/Meilensteine/Phasen"
                Else
                    tmpLabel = "Ressources/Costs/Milestones/Phases"
                End If
                'If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '    tmpLabel = "Auswahl über Projekt-Struktur..."
                'Else
                '    tmpLabel = "Select by Structure..."
                'End If

            Case "PTXG1B10" ' größter Engpass
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "größter Engpass"
                Else
                    tmpLabel = "Worst bottleneck"
                End If

            Case "PTXG1B3" ' Auslastung
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Top 3 Ressourcen Engpässe"
                Else
                    tmpLabel = "Top 3 Resource Bottlenecks"
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
                    tmpLabel = "Analyses"
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
                    tmpLabel = "Milestone Trafficlight"
                End If

            Case "PT0G1M0" ' Planelemente visualisieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Phasen/Meilensteine visualisieren"
                Else
                    tmpLabel = "Visualize Phases/Milestones"
                End If

            Case "PT0G1B8" ' 
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt Filter"
                Else
                    tmpLabel = "Project Filter"
                End If
            Case "PT0G1B9" ' Filter zurücksetzen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Filter zurücksetzen"
                Else
                    tmpLabel = "Delete Filter"
                End If
            Case "PT0G1B10" ' Anzeige der Projekte mit roter ProjektAmpel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekte mit Ampel -rot-"
                Else
                    tmpLabel = "Projects with red flag"
                End If
            Case "PT0G1B11" ' Anzeige der Projektemit ungedeckter Budget-Finanzierung
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekte unterfinanziert"
                Else
                    tmpLabel = "Projects not fully financed"
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
                    tmpLabel = "Resource Needs"
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

            Case "PT2G1M2B3"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Budget ändern"
                Else
                    tmpLabel = "Modify budget"
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
                    tmpLabel = "Portfolio-Charts"
                End If
            Case "PT7G1M1" ' Add Project Charts
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt-Charts"
                Else
                    tmpLabel = "Project-Charts"
                End If
            Case "PT7G1M2" ' Plan/Aktuell
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Plan vs. Aktuell"
                Else
                    tmpLabel = "Plan vs. Actual"
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

            Case "PT0G1B" ' Cockpit
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
                    tmpLabel = "Create Report Profile"
                Else
                    tmpLabel = "Report Profil erstellen"
                End If

            Case "PT1G1M01" ' Einzelprojekt-Berichte
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt Report definieren"
                Else
                    tmpLabel = "Define Project Report"
                End If

            Case "PT1G1M01B0" ' Typ I
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "ohne Element-Auswahl"
                Else
                    tmpLabel = "without element-selection"
                End If

                'Case "PT1G1M1" ' Typ II
                '    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '        tmpLabel = "mit Element-Auswahl"
                '    Else
                '        tmpLabel = "with element-selection"
                '    End If

                'Case "PT1G1M1B1" ' Auswahl über Namen
                '    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '        tmpLabel = "Auswahl über Namen..."
                '    Else
                '        tmpLabel = "Select by Names..."
                '    End If

            Case "PT1G1M1B2" ' Auswahl über Hierarchie
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "mit Element-Auswahl"
                Else
                    tmpLabel = "with element-selection"
                End If
                'If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '    tmpLabel = "Auswahl über Projekt-Struktur..."
                'Else
                '    tmpLabel = "Select by Structure..."
                'End If

            Case "PT1G1M02" ' Multiprojekt-Berichte
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio Report definieren..."
                Else
                    tmpLabel = "Define Portfolio Report..."
                End If

            Case "PT1G1B2" ' Typ I
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "ohne Element-Auswahl"
                Else
                    tmpLabel = "without element-selection"
                End If

                'Case "PT1G1M2" ' Typ II
                'If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '    tmpLabel = "mit Element-Auswahl"
                'Else
                '    tmpLabel = "with element-selection"
                'End If

                'Case "PT1G1M2B1" ' Auswahl über Namen
                '    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '        tmpLabel = "Auswahl über Namen..."
                '    Else
                '        tmpLabel = "Select by Names..."
                '    End If

            Case "PT1G1M2B2" ' Auswahl über Hierarchie
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "mit Element-Auswahl"
                Else
                    tmpLabel = "with element-selection"
                End If

                'If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                '    tmpLabel = "Auswahl über Projekt-Struktur..."
                'Else
                '    tmpLabel = "Select by Structure..."
                'End If

            Case "PT1G1B4" ' letztes Report-Profil speichern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "letzte Report Definition als Profil speichern"
                Else
                    tmpLabel = "Save last Report definition as pre-defined"
                End If

            Case "PT1G1B5" ' Report-Profil ausführen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Report erstellen"
                Else
                    tmpLabel = "Select Report"
                End If

            Case "PT1G1M0B1" ' Report-Profil ausführen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Report erstellen"
                Else
                    tmpLabel = "Select Report"
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
                    tmpLabel = "Single Projects"
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
                    If visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                        tmpLabel = "Kostenart hinzufügen"
                    Else
                        tmpLabel = "Rolle hinzufügen"
                    End If

                Else
                    If visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                        tmpLabel = "Add Cost"
                    Else
                        tmpLabel = "Add Resource/Role"
                    End If

                End If


            Case "PT2G1M2B5" ' Ressource/Kostenart löschen

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    If visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                        tmpLabel = "Kostenart löschen"
                    Else
                        tmpLabel = "Rolle löschen"
                    End If

                Else
                    If visboZustaende.projectBoardMode = ptModus.massEditCosts Then
                        tmpLabel = "Delete Cost"
                    Else
                        tmpLabel = "Delete  Resource/Role"
                    End If

                End If


            Case "PTmassEdit" 'Editieren im MassEdit
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Edit"
                Else
                    tmpLabel = "Edit"
                End If

            Case "PT2G1M2B1" ' Massen-Edit Ressourcen 
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ändern Ressourcen Bedarfe"
                Else
                    tmpLabel = "Modify Resource Needs"
                End If

            Case "PT2G1M2B9" ' Massen-Edit Kosten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ändern Kosten"
                Else
                    tmpLabel = "Modify Cost Needs"
                End If

            Case "PT2G1M2B2" ' Massen-Edit Phasen Meilensteine ändern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Ändern Termine"
                Else
                    tmpLabel = "Modify schedules"
                End If


            Case "PT2G1M2B8" ' Massen Edit Attributes

                If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                        tmpLabel = "Ändern Budget (Baseline)"
                    Else
                        tmpLabel = "Modify Budget (Baseline)"
                    End If
                ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Then
                    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                        tmpLabel = "Ändern Budget (Planungs-Stand)"
                    Else
                        tmpLabel = "Modify Budget (Planning Version)"
                    End If
                    Else
                    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                        tmpLabel = "Ändern Attribute"
                    Else
                        tmpLabel = "Modify Attributes"
                    End If
                End If



            Case "PT4G2M3" ' Export to Excel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Export Projekte in Excel"
                Else
                    tmpLabel = "Export Projects to Excel"
                End If

            Case "PT4G2M3B1" ' Projekte mit einer Übersichtszeile in Excel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Übersicht"
                Else
                    tmpLabel = "Overview"
                End If

            Case "PT4G2M3B2" ' Projekte mit Details in Excel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Details"
                Else
                    tmpLabel = "Details"
                End If

            Case "PT4G2M3B3" ' Projekte mit Details in Excel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Auslastung Rollen"
                Else
                    tmpLabel = "Utilization Roles"
                End If

            Case "PT4G2M3B4" ' Projekte mit Details in Excel
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Offline Planungs Daten"
                Else
                    tmpLabel = "Offline Planning Data"
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
            Case "PT6G2B6" ' Platzhalter Rollen automatisch reduzieren, ohne erneutes Nachfragen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "automatisches Reduzieren ohne Zwischenfrage"
                Else
                    tmpLabel = "Automatically reduce without Asking"
                End If

            Case "PT6G2B5" ' Sortierung ermöglichen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Sortierung ermöglichen"
                Else
                    tmpLabel = "Enable sorting"
                End If

            Case "PT6G2B7" ' Header anzeigen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Header anzeigen"
                Else
                    tmpLabel = "Show Header"
                End If

            Case "PT6G2B8" ' Compare with Last Version
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "mit letzter Version vergleichen"
                Else
                    tmpLabel = "compare with last version"
                End If

            Case "PTfreezeB1" ' Fixieren
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Fixieren"
                Else
                    tmpLabel = "Freeze"
                End If

            Case "PTfreezeB2" ' Fixierung aufheben
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Fixierung aufheben"
                Else
                    tmpLabel = "De-Freeze"
                End If

            Case "PTunmarkBT"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Reset Markierung"
                Else
                    tmpLabel = "Reset Marker"
                End If

            Case "PTMarkBT"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Markieren"
                Else
                    tmpLabel = "Set Mark"
                End If

            Case "PT2G1M1B4" ' Status ändern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt beauftragen"
                Else
                    tmpLabel = "Set Project official"
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
                    tmpLabel = "Beschriftungen..."
                Else
                    tmpLabel = "Annotations..."
                End If

                ''Case "PT2G1B4" ' Beschriften ON
                ''    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                ''        tmpLabel = "Ein"
                ''    Else
                ''        tmpLabel = "ON"
                ''    End If

                ''Case "PT2G1B5" ' Beschriftungen löschen
                ''    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                ''        tmpLabel = "Aus"
                ''    Else
                ''        tmpLabel = "OFF"
                ''    End If

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

            Case "PT2G2M" 'Projekte/Varianten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekte/Varianten"
                Else
                    tmpLabel = "Projects/Variants"
                End If

            Case "PT2G2oa" 'Projekte/Varianten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekte/Varianten"
                Else
                    tmpLabel = "Projects/Variants"
                End If

            Case "PT2G2B2" ' Portfolio/s anzeigen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio/s aus Session anzeigen"
                Else
                    tmpLabel = "Show Session Portfolio/s"
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
                    tmpLabel = "Import"
                End If
            Case "PT4G1M" ' IMPORT
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Import"
                Else
                    tmpLabel = "Import"
                End If

            Case "PT4G2M" ' EXPORT
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Export"
                Else
                    tmpLabel = "Export"
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

            Case "PT4G1B10" ' Import JIRA Projekte
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "JIRA-Projekt"
                Else
                    tmpLabel = "JIRA-project"
                End If

            Case "PT4G1B7" ' Import Projekte (Batch)
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Erzeuge Projekte aus Liste (VISBO)"
                Else
                    tmpLabel = "Create Projects from list (VISBO)"
                End If

            Case "PT4G1B5" ' Import Scenario Definition
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio Definition"
                Else
                    tmpLabel = "Portfolio Definition"
                End If
            Case "PT4G1B9" 'Import Projekte gemäß Konfiguration
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Erzeuge Projekte aus Liste (konfiguriert)"
                Else
                    tmpLabel = "Create Projects from list (customized)"
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
                    tmpLabel = "Auswahl von Meilensteine und Phasen..."
                Else
                    tmpLabel = "Selection of Milestones and Phases..."
                End If

            Case "PT4G2B3" ' Export Priorisierungsliste
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio Definition"
                Else
                    tmpLabel = "Portfolio Definition"
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


            Case "PT5G1" ' Load from Database

                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Laden von VISBO "

                Else
                    tmpLabel = "Load from VISBO"
                End If

            Case "PT5G1B1" ' Portfolio/s
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                        If awinSettings.loadPFV Then
                            tmpLabel = "Portfolio/s (Baselines)"
                        Else
                            tmpLabel = "Portfolio/s (aktuelle Planung)"
                        End If
                    Else
                        tmpLabel = "Portfolio/s"
                    End If
                Else
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                        If awinSettings.loadPFV Then
                            tmpLabel = "Portfolio/s (Baselines)"
                        Else
                            tmpLabel = "Portfolio/s (current planning)"
                        End If
                    Else
                        tmpLabel = "Portfolio/s"
                    End If
                End If

            Case "PT5G1B3" ' Project/s
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    If awinSettings.loadPFV Or awinSettings.filterPFV Then
                        tmpLabel = "Projekt-Baseline laden"
                    Else
                        tmpLabel = "Projekt-Planung laden"
                    End If
                Else
                    If awinSettings.loadPFV Or awinSettings.filterPFV Then
                        tmpLabel = "load Project-Baseline"
                    Else
                        tmpLabel = "load Project-Planning"
                    End If
                End If

            Case "PT5G2" ' Speichern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Speichern in VISBO"
                Else
                    tmpLabel = "Save to VISBO"
                End If

            Case "Pt5G2B1" ' Portfolio/s
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio/s"
                Else
                    tmpLabel = "Portfolio/s"
                End If

            Case "Pt5G2B3" ' Projekt/e
                If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                        tmpLabel = "als Projekt-Baseline speichern"
                    Else
                        tmpLabel = "Store as Project-Baseline"
                    End If
                Else
                    If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                        tmpLabel = "als Projekt-Planung speichern"
                    Else
                        tmpLabel = "store as Project-Planning"
                    End If
                End If


            Case "Pt5G2B4" ' Alles speichern
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Alles als Baseline speichern (Projekte & Portfolios)"
                Else
                    tmpLabel = "Store everything as Baseline (projects & portfolios)"
                End If

            Case "PT5G3" ' Löschen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Löschen in VISBO"
                Else
                    tmpLabel = "Delete in VISBO"
                End If

            Case "Pt5G3B1" ' Multiprojekt-Szenario
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Portfolio/s"
                Else
                    tmpLabel = "Portfolio/s"
                End If

            Case "PT5G3M" ' Löschen aus Datenbank 
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Löschen in VISBO"
                Else
                    tmpLabel = "Delete in VISBO"
                End If

            Case "PT5G3M2" ' Projekte/Varianten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Projekt/e"
                Else
                    tmpLabel = "Project/s"
                End If

            Case "Pt5G3B3" ' Projekte/Varianten/TimeStamps auswählen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    If awinSettings.loadPFV Then
                        tmpLabel = "Baselines/Varianten/TimeStamps"
                    Else
                        tmpLabel = "Planungen/Varianten/TimeStamps"
                    End If

                Else
                    If awinSettings.loadPFV Then
                        tmpLabel = "Baselines/Variants/TimeStamps"
                    Else
                        tmpLabel = "Plans/Variants/TimeStamps"
                    End If

                End If

            Case "Pt5G3B4" ' X Versionen behalten
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "X Versionen behalten..."
                Else
                    tmpLabel = "Keep X Versions..."
                End If


            Case "PT2G2B5" ' Sperre setzen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Schreibschutz setzen/aufheben..."
                Else
                    tmpLabel = "Set/Unset Write-Protection..."
                End If

            Case "PT2G2B5oa" ' Sperre setzen
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Schreibschutz setzen/aufheben..."
                Else
                    tmpLabel = "Set/Unset Write-Protection..."
                End If

            Case "PTedit"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Bearbeiten"
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

            Case "PT6G1B1"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Alphabetisch"
                Else
                    tmpLabel = "Alphabetically"
                End If

            Case "PT6G1B2"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Business Unit-Startdatum-Name"
                Else
                    tmpLabel = "by Business Unit-Startdate-Name"
                End If

            Case "PT6G1B3"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Strategie-Risiko-Profit"
                Else
                    tmpLabel = "by Strategy-Risk-Profit"
                End If

            Case "PT6G1B4"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "nach individuellen Kriterien sortieren"
                Else
                    tmpLabel = "sort by individual criterias"
                End If

            Case "PT6G1B5"
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
            Case "PTTestfunktionen"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "VISBO Testfuntionen"
                Else
                    tmpLabel = "VISBO Testings"
                End If
            Case "PTWebServer"
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "WebServer"
                Else
                    tmpLabel = "WebServer"
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
            Case "PTeinstG1B1" ' Einstellungen für VISBO-Board; MassEdit, Ampel, PropAnpass,Report Sprache
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "VISBO Einstellungen"
                Else
                    tmpLabel = "VISBO Settings"
                End If

            Case "PTeinstG1B4" ' Einstellungen: Custom User Role wechseln 
                If menuCult.Name = ReportLang(PTSprache.deutsch).Name Then
                    tmpLabel = "Custom User Role wechseln"
                Else
                    tmpLabel = "Change Custom User Role"
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

            If myCustomUserRole.isEntitledForMenu(control.Id) Then
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
                    Case "PT6G2B7" ' Header anzeigen
                        chckVisibility = False
                    Case "PThelp" ' Help anzeigen
                        chckVisibility = False
                    Case Else
                        ' alle anderen werden sichtbar gemacht
                        chckVisibility = True
                End Select
            Else
                chckVisibility = False
            End If

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
                Case "PTWebServer"
                    chckVisibility = False
                Case "PTTestfunktionen"
                    chckVisibility = False
                Case "PTlizenz"
                    chckVisibility = False

                Case "PT2G1M2B6" ' Mass-Edit Änderungen verwerfen
                    chckVisibility = False

                Case "PTMEC" ' Charts und Info 
                    If (visboZustaende.projectBoardMode = ptModus.massEditRessSkills Or visboZustaende.projectBoardMode = ptModus.massEditCosts) Then
                        chckVisibility = True
                    Else
                        chckVisibility = False
                    End If


                Case "PTmassEdit" ' Charts und Info 
                    If (visboZustaende.projectBoardMode = ptModus.massEditRessSkills Or visboZustaende.projectBoardMode = ptModus.massEditCosts) Then
                        chckVisibility = Not (myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Or
                        myCustomUserRole.customUserRole = ptCustomUserRoles.InternalViewer Or
                        myCustomUserRole.customUserRole = ptCustomUserRoles.ExternalViewer)
                    Else
                        chckVisibility = False
                    End If

                Case "PTMECsettings" ' Charts und Info 
                    If (visboZustaende.projectBoardMode = ptModus.massEditRessSkills Or visboZustaende.projectBoardMode = ptModus.massEditCosts) Then
                        chckVisibility = True
                    Else
                        chckVisibility = False
                    End If

                Case Else
                    chckVisibility = True
            End Select

        End If

    End Function

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

        Call projektTafelInit()

        enableOnUpdate = False
        ' jetzt auf alle Fälle wieder das MPT Window aktivieren ...
        projectboardWindows(PTwindows.mpt).Activate()

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

                            Call awinShowtimezone(showRangeLeft, showRangeRight, True)
                        Else
                            ' beim alten ShowRangeLeft lassen, wenn es Überlappungen gibt ..
                            Dim newLeft As Integer = ShowProjekte.getMinMonthColumn(todoListe)
                            Dim newRight As Integer = ShowProjekte.getMaxMonthColumn(todoListe)

                            If newLeft >= showRangeRight Or newRight <= showRangeLeft Then
                                ' neu bestimmen 
                                Call awinShowtimezone(showRangeLeft, showRangeRight, False)

                                showRangeLeft = ShowProjekte.getMinMonthColumn(todoListe)
                                showRangeRight = ShowProjekte.getMaxMonthColumn(todoListe)

                                Call awinShowtimezone(showRangeLeft, showRangeRight, True)

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
    Sub Tom2G2MassEdit(control As IRibbonControl)

        ' check ob auch keine Summary Projects selektiert sind ...
        Dim nameCollection As New Collection

        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
            ' es ist dem Portfolio Manager erlaubt, Summary Projekte zu editieren ... 

        ElseIf Not noSummaryProjectsareSelected(nameCollection) Then

            Exit Sub

        End If

        Call massEditRcTeAt(ptModus.massEditRessSkills)


    End Sub

    Sub Tom2G2MassEditC(control As IRibbonControl)
        ' check ob auch keine Summary Projects selektiert sind ...
        Dim nameCollection As New Collection

        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
            ' es ist dem Portfolio Manager erlaubt, Summary Projekte zu editieren ... 

        ElseIf Not noSummaryProjectsareSelected(nameCollection) Then

            Exit Sub

        End If

        Call massEditRcTeAt(ptModus.massEditCosts)
    End Sub

    ''' <summary>
    ''' Online Massen-Edit von Terminen im Visual Board
    ''' </summary>
    ''' <param name="control"></param>
    Sub Tom2G2MassEditTe(control As IRibbonControl)
        Call massEditRcTeAt(ptModus.massEditTermine)
    End Sub

    ''' <summary>
    ''' Online Massen-Edit von Attributen im Visual Board
    ''' </summary>
    ''' <param name="control"></param>
    Sub Tom2G2MassEditAttr(control As IRibbonControl)
        Call massEditRcTeAt(ptModus.massEditAttribute)
    End Sub

    ''' <summary>
    ''' führt in backToProjectBoard die Aktionen durch, die eigentlich in einem deactivate_Event gemacht werden sollten. 
    ''' Da das aber mit den Windows.activate in backtoprojectboard nicht passiert, ist das die Abhilfe  
    ''' </summary>
    ''' <param name="tableTyp">gibt an , ob es sich um Mass-Edit Ressourcen, Termine oder Attribute handelt </param>
    Private Sub performDeactivateActionsFor(ByVal tableTyp As Integer)

        'Dim anzahlMassColSpalten As Integer
        Dim mIX As Integer

        If tableTyp = ptTables.meRC Then
            'anzahlMassColSpalten = 5
            mIX = 0

            If Not IsNothing(formProjectInfo1) Then
                formProjectInfo1.Close()
            End If

        ElseIf tableTyp = ptTables.meTE Then
            mIX = 1
            'anzahlMassColSpalten = 11

        ElseIf tableTyp = ptTables.meAT Then
            mIX = 2
            'anzahlMassColSpalten = 15

        End If


        Dim meWS As Excel.Worksheet =
            CType(CType(appInstance.Workbooks(myProjektTafel), Excel.Workbook) _
            .Worksheets(arrWsNames(tableTyp)), Excel.Worksheet)

        appInstance.EnableEvents = False

        ' jetzt den Schutz aufheben , falls einer definiert ist 
        If meWS.ProtectContents Then
            meWS.Unprotect(Password:="x")
        End If

        Try

            ' jetzt die Spalten Werte merken 
            Try
                massColFontValues(mIX, 0) = CDbl(appInstance.ActiveWindow.Zoom)
                For ik As Integer = 1 To 100
                    massColFontValues(mIX, ik) = CDbl(CType(meWS.Columns(ik), Excel.Range).ColumnWidth)
                Next
            Catch ex As Exception

            End Try


            ' jetzt die Autofilter de-aktivieren ... 
            If CType(meWS, Excel.Worksheet).AutoFilterMode = True Then
                CType(meWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
            End If

            ' jetzt alles löschen 
            Try
                Dim mxZeile As Integer = meWS.UsedRange.Rows.Count
                For i As Integer = 2 To mxZeile
                    CType(meWS.Rows(i), Excel.Range).Delete()
                Next
                ' tk alt ...
                'meWS.UsedRange.Clear()
            Catch ex As Exception

            End Try

        Catch ex As Exception
            Call MsgBox("Fehler beim Filter zurücksetzen " & vbLf & ex.Message)
        End Try

        appInstance.EnableEvents = True

    End Sub
    ''' <summary>
    ''' wird aus Mass-Edit Ressourcen, Termine oder Attibute aufgerufen 
    ''' stellt sicher, dass wieder der Projekt-Tafel Zustand hergestellt wird. 
    ''' der aufruf performDeactivateActions ist notwendig, weil ein table.Deactivate mit Window.Activate  nicht mehr stattfindet 
    ''' </summary>
    ''' <param name="control"></param>
    Sub PTbackToProjectBoard(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg
        Dim reDrawProjects As New Collection

        ' Bildschirm einfrieren ...
        If appInstance.ScreenUpdating = True Then
            appInstance.ScreenUpdating = False
        End If


        Try
            ' hier wieder auf false setzen , in der Multiprojekt-Tafel soll das nicht angezeigt werden ...
            awinSettings.showValuesOfSelected = False

            ' jetzt müssen die Merk- & ggf Rücksetz-Aktionen gemacht werden, die mit dem entsprechenden massEdit Table verbunden sind
            Dim tableTyp As Integer = ptTables.meRC

            If (visboZustaende.projectBoardMode = ptModus.massEditRessSkills Or visboZustaende.projectBoardMode = ptModus.massEditCosts) Then
                tableTyp = ptTables.meRC
                Call deleteChartsInSheet(arrWsNames(ptTables.meCharts))

            ElseIf visboZustaende.projectBoardMode = ptModus.massEditTermine Then
                tableTyp = ptTables.meTE
            ElseIf visboZustaende.projectBoardMode = ptModus.massEditAttribute Then
                tableTyp = ptTables.meAT
            End If

            Call performDeactivateActionsFor(tableTyp)

            ' jetzt muss gecheckt werden, welche dbCache Projekte immer noch identisch zum ShowProjekte Pendant sind
            ' deren temp Schutz muss dann wieder aufgehoben werden ... 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In sessionCacheProjekte.liste

                If ShowProjekte.contains(kvp.Value.name) Then
                    Dim hproj As clsProjekt = ShowProjekte.getProject(kvp.Value.name)
                    Dim pvName As String = calcProjektKey(hproj.name, hproj.variantName)

                    ' wenn PortfolioManager, so muss die PFV Variante Lock aufgehoben werden
                    Dim vNameToUnProtect As String = hproj.variantName
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                        If hproj.variantName <> "" Then
                            vNameToUnProtect = hproj.variantName
                        Else
                            vNameToUnProtect = ptVariantFixNames.pfv.ToString
                        End If
                        pvName = calcProjektKey(hproj.name, vNameToUnProtect)
                    End If

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
                            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName,
                            'dbUsername, dbPasswort)

                            Dim wpItem As New clsWriteProtectionItem(pvName, ptWriteProtectionType.project,
                                                                      dbUsername, False, False)
                            If CType(databaseAcc, DBAccLayer.Request).setWriteProtection(wpItem, err) Then
                                ' erfolgreich
                                writeProtections.upsert(wpItem)
                            Else
                                ' nicht erfolgreich
                                wpItem = CType(databaseAcc, DBAccLayer.Request).getWriteProtection(hproj.name, hproj.variantName, err)
                                writeProtections.upsert(wpItem)
                            End If
                        End If
                    Else
                        If tableTyp = ptTables.meTE Then
                            ' neu Zeichnen des Projektes 
                            If Not reDrawProjects.Contains(hproj.name) Then
                                reDrawProjects.Add(hproj.name)
                            End If
                        End If
                        ' temporär geschützt lassen ...
                    End If
                End If
            Next

            ' zurücksetzen , aber nicht zurücksetzen der currentSessionConstellation
            sessionCacheProjekte.Clear(False)

            ' zurücksetzen der Selektierten Projekte, aber nicht zurücksetzen der currentSessionConstellation
            selectedProjekte.Clear(False)

            'Call projektTafelInit()

            If tempSkipChanges Then
                'Call restoreProjectsFromBackup()
                Call MsgBox("restored ...")
                tempSkipChanges = False
            End If

            Call enableControls(ptModus.graficboard)

            'appInstance.EnableEvents = False
            ' wird ohnehin zu Beginn des MassenEdits ausgeschaltet  
            'enableOnUpdate = False


            ' tk, 16.8.17 Versuch, um das Fenster PRoblem in den Griff zu bekommen 
            appInstance.EnableEvents = True
            Try
                If appInstance.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized Then
                    appInstance.ActiveWindow.WindowState = Excel.XlWindowState.xlNormal
                End If
            Catch ex As Exception

            End Try


            Try

                ' jetzt werden die Windows gelöscht, falls sie überhaupt existieren  ...
                If Not IsNothing(projectboardWindows(PTwindows.massEdit)) Then
                    Try
                        projectboardWindows(PTwindows.massEdit).Close()
                    Catch ex As Exception

                    End Try

                    projectboardWindows(PTwindows.massEdit) = Nothing
                End If

                If Not IsNothing(projectboardWindows(PTwindows.meChart)) Then
                    Try
                        projectboardWindows(PTwindows.meChart).Close()
                    Catch ex As Exception

                    End Try

                    projectboardWindows(PTwindows.meChart) = Nothing
                End If

            Catch ex As Exception

            End Try


            ' jetzt müssen ggf drei Windows wieder angezeigt werden 
            Try
                With projectboardWindows(PTwindows.mpt)
                    .Visible = True
                    If CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Windows.Count = 1 Then
                        .WindowState = Excel.XlWindowState.xlMaximized
                    Else
                        Try
                            If Not IsNothing(projectboardWindows(PTwindows.mptpf)) Then

                                Try
                                    projectboardWindows(PTwindows.mptpf).Visible = True
                                Catch ex As Exception
                                    projectboardWindows(PTwindows.mptpf) = Nothing
                                End Try

                                Try
                                    With CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Worksheets(arrWsNames(ptTables.mptPfCharts)), Excel.Worksheet)
                                        .Activate()
                                    End With
                                Catch ex As Exception

                                End Try


                            End If
                        Catch ex As Exception
                            projectboardWindows(PTwindows.mptpf) = Nothing
                        End Try
                        Try
                            If Not IsNothing(projectboardWindows(PTwindows.mptpr)) Then

                                Try
                                    projectboardWindows(PTwindows.mptpr).Visible = True
                                Catch ex As Exception
                                    projectboardWindows(PTwindows.mptpr) = Nothing
                                End Try

                                Try
                                    With CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Worksheets(arrWsNames(ptTables.mptPrCharts)), Excel.Worksheet)
                                        .Activate()
                                    End With
                                Catch ex As Exception

                                End Try

                            End If
                        Catch ex As Exception
                            projectboardWindows(PTwindows.mptpr) = Nothing
                        End Try
                    End If

                End With
            Catch ex As Exception

            End Try


            enableOnUpdate = True
            appInstance.EnableEvents = True

            Try
                ' mit diesem Befehl wird das dem Window zugeordnete Sheet aktiviert, allerdings ohne die entsprechenden .activate bzw. .deactivate Routinen zu durchlaufen ...
                projectboardWindows(PTwindows.mpt).Activate()
            Catch ex As Exception

            End Try

            Try
                With CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)
                    .Activate()
                End With
            Catch ex As Exception

            End Try



            appInstance.ScreenUpdating = True

            ' jetzt müssen ggf noch die Portfolio Charts neu gezeichnet werden 
            Try
                If Not IsNothing(projectboardWindows(PTwindows.mptpf)) Then
                    If projectboardWindows(PTwindows.mptpf).Visible = True Then
                        Call awinNeuZeichnenDiagramme(2)
                    End If
                End If
            Catch ex As Exception
                projectboardWindows(PTwindows.mptpf) = Nothing
            End Try

            ' jetzt müssen alle ggf in reDrawProjects aufgeführten Projekte neu gezeichnet werden .. 
            If reDrawProjects.Count > 0 Then
                For Each pName As String In reDrawProjects
                    If ShowProjekte.contains(pName) Then
                        Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                        Call replaceProjectVariant(pName, hproj.variantName, False, True, hproj.tfZeile)
                    End If

                Next
            End If
        Catch ex As Exception

            enableOnUpdate = True
            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True
        End Try


    End Sub

    ''' <summary>
    ''' fügt im MassenEdit Sheet eine Zeile ein, macht aber sonst noch nichts, es werden also noch keinerlei Änderungen am 
    ''' betroffenen Projekt vorgenommen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTzeileEinfuegen(control As IRibbonControl)

        Call massEditZeileEinfügen(control.Id)

        If Not appInstance.EnableEvents = True Then
            appInstance.EnableEvents = True
        End If



    End Sub

    ''' <summary>
    ''' löscht im MassenEdit Sheet eine Zeile, das heisst, die Rolle bzw. Kostenart wird rausgenommen 
    ''' es bleibt aber pro Projekt-/Phase eine leere Zeile als Platzhalter stehen  
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTzeileLoeschen(control As IRibbonControl)

        Call massEditZeileLoeschen(control.Id)

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
                            Call MsgBox(hproj.name & ", " & hproj.variantName & " is protected " & vbLf &
                                        "and cannot be modified. You could instead create a variant.")
                        Else
                            Call MsgBox(hproj.name & ", " & hproj.variantName & " ist geschützt " & vbLf &
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
    ' tk 21.8.17 wird nicht mehr aufgerufen 
    '' ''' <summary>
    '' ''' earliest und latest Start eines Projektes ändern 
    '' ''' </summary>
    '' ''' <param name="control"></param>
    '' ''' <remarks></remarks>
    ''Sub Tom2G1EarliestLatestStart(control As IRibbonControl)

    ''    Dim setStartEnd As New frmEarliestLatestStart

    ''    Dim returnValue As DialogResult
    ''    Dim awinSelection As Excel.ShapeRange
    ''    Dim i As Integer
    ''    Dim hproj As clsProjekt
    ''    Dim singleShp As Excel.Shape
    ''    Dim pname As String
    ''    Dim todoListe As New Collection
    ''    Dim errMessage As String = ""
    ''    Dim initMsg As String = "bitte erst eine Variante anlegen"

    ''    Call projektTafelInit()

    ''    ' es wird vbeim Betreten der Tabelle2 nochmal auf False gesetzt ... und insbesondere bei Activate Tabelle1 (!) auf true gesetzt, nicht vorher wieder
    ''    enableOnUpdate = False

    ''    ' Änderung 2.7.14 tk : Vorbedingung sicherstellen: nur Projekte, die noch nicht beauftragt sind, können noch verschoben und 
    ''    ' werden
    ''    '
    ''    Try
    ''        'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
    ''        awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
    ''    Catch ex As Exception
    ''        awinSelection = Nothing
    ''    End Try

    ''    If Not awinSelection Is Nothing Then

    ''        ' Es muss mindestens 1 Projekt selektiert sein
    ''        For i = 1 To awinSelection.Count

    ''            singleShp = awinSelection.Item(i)

    ''            Try
    ''                hproj = ShowProjekte.getProject(singleShp.Name, True)
    ''                pname = hproj.name
    ''            Catch ex As Exception
    ''                Call MsgBox(" Fehler! Projekt " & singleShp.Name & " nicht im Hauptspeicher")
    ''                enableOnUpdate = True
    ''                Exit Sub
    ''            End Try

    ''            If hproj.Status = ProjektStatus(PTProjektStati.geplant) Then
    ''                ' nur dann macht das Setzen von earliest / latest Sinn ...

    ''                todoListe.Add(hproj.name)

    ''                If i = 1 Then

    ''                    ' jetzt die Aktion durchführen ...

    ''                    With setStartEnd

    ''                        .EarliestStart.Value = hproj.earliestStart
    ''                        .LatestStart.Value = hproj.latestStart

    ''                    End With


    ''                Else

    ''                    With setStartEnd

    ''                        If .EarliestStart.Value <> hproj.earliestStart Or .LatestStart.Value <> hproj.latestStart Then

    ''                            .EarliestStart.Value = 0
    ''                            .LatestStart.Value = 0

    ''                        End If

    ''                    End With


    ''                End If
    ''            Else
    ''                errMessage = errMessage & vbLf & hproj.name
    ''            End If

    ''        Next i

    ''        If todoListe.Count > 0 Then

    ''            returnValue = setStartEnd.ShowDialog

    ''            If returnValue = DialogResult.OK Then

    ''                For i = 1 To todoListe.Count

    ''                    pname = CStr(todoListe.Item(i))

    ''                    ' jetzt die Aktion durchführen ...
    ''                    Try
    ''                        hproj = ShowProjekte.getProject(pname)
    ''                        With setStartEnd

    ''                            hproj.earliestStart = .EarliestStart.Value
    ''                            hproj.latestStart = .LatestStart.Value
    ''                            hproj.earliestStartDate = hproj.startDate.AddMonths(.EarliestStart.Value)
    ''                            hproj.latestStartDate = hproj.startDate.AddMonths(.LatestStart.Value)

    ''                        End With
    ''                    Catch ex As Exception
    ''                        Call MsgBox(" Fehler! Projekt " & pname & " earliest/latest kann nicht gesetzt werden")
    ''                        enableOnUpdate = True
    ''                        Exit Sub
    ''                    End Try

    ''                Next i

    ''                Call MsgBox("ok, frühester und spätester Start gesetzt")

    ''            ElseIf returnValue = DialogResult.Cancel Then
    ''                'Call MsgBox("Default soll gelten")

    ''            End If

    ''        End If

    ''        If errMessage.Length > 0 Then
    ''            Call MsgBox(initMsg & vbLf & errMessage)
    ''        End If

    ''    Else

    ''        Call MsgBox("Es muss mindestens ein Projekt selektiert sein")

    ''    End If

    ''    Call awinDeSelect()

    ''    'appInstance.ScreenUpdating = True
    ''    enableOnUpdate = True


    ''End Sub

    Sub sortCurrentConstellation(control As IRibbonControl)

        Dim sortType As Integer = ptSortCriteria.alphabet

        Call projektTafelInit()

        ' Vorbesetzungen 
        appInstance.EnableEvents = False
        enableOnUpdate = False


        Try

            If control.Id = "PT6G1B1" Then
                ' alphabetisch sortieren 
                sortType = ptSortCriteria.alphabet
            ElseIf control.Id = "PT6G1B2" Then
                ' Business Unit StartDate Name
                sortType = ptSortCriteria.buStartName
            ElseIf control.Id = "PT6G1B3" Then
                ' Strategy-Risk-Profit
                sortType = ptSortCriteria.strategyRiskProfitLoss
            ElseIf control.Id = "PT6G1B4" Then
                ' sort by individual criterias
                sortType = ptSortCriteria.customListe
            ElseIf control.Id = "PT6G1B5" Then
                ' custom Formel , zu implementieren 
                sortType = ptSortCriteria.alphabet
            End If

            If Not IsNothing(currentSessionConstellation) Then
                If currentSessionConstellation.Liste.Count <> 0 Then

                    Dim currentSortConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True, cName:="Sort Result")

                    If currentSortConstellation.sortCriteria <> sortType Then
                        appInstance.ScreenUpdating = False
                        Try
                            ' nur dann muss was gemacht werden ...  
                            currentSortConstellation.sortCriteria = sortType

                            Dim tmpConstellation As New clsConstellations
                            tmpConstellation.Add(currentSortConstellation)

                            ' es in der Session Liste verfügbar machen
                            If projectConstellations.Contains(currentSortConstellation.constellationName) Then
                                projectConstellations.Remove(currentSortConstellation.constellationName)
                            End If

                            projectConstellations.Add(currentSortConstellation)

                            Call showConstellations(constellationsToShow:=tmpConstellation,
                                                    clearBoard:=True, clearSession:=False, storedAtOrBefore:=Date.Now)

                            ''If sortType = ptSortCriteria.customListe Then
                            ''    Call awinNeuZeichnenDiagramme(2)
                            ''Else
                            ''    ' in allen anderen Fällen kann sich an der Zahl und Ressourcenbedrag nichts geändert haben 
                            ''End If
                        Catch ex As Exception

                        End Try

                        appInstance.ScreenUpdating = True

                    End If

                Else
                    If awinSettings.englishLanguage Then
                        Call MsgBox("please load projects/portfolios first ...")
                    Else
                        Call MsgBox("bitte zuerst Projekte/Portfolios laden ...")
                    End If
                End If
            Else
                If awinSettings.englishLanguage Then
                    Call MsgBox("please load projects/portfolios first ...")
                Else
                    Call MsgBox("bitte zuerst Projekte/Portfolios laden ...")
                End If
            End If

        Catch ex As Exception
            If appInstance.ScreenUpdating = False Then
                appInstance.ScreenUpdating = True
            End If
        End Try

        appInstance.EnableEvents = True
        enableOnUpdate = True

    End Sub

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



        ' jetzt muss die Behandlung rein, dass ggf das Portfolio oder Project Window angezeigt werden ... 
        If control.Id = "PTXG1B6" Or
            control.Id = "PTXG1B7" Then
            ' Portfolio Charts Ressource/Cost/Phases/Milestones, Auswahl über Namen oder Hierarchie 

            If thereAreAnyCharts(PTwindows.mptpf) Then
                ' jetzt sollte das Window gezeigt werden, wenn es nicht schon sichtbar ist ... 
                Call showVisboWindow(PTwindows.mptpf)
            End If

        End If


    End Sub


    Sub AnalyseLeistbarkeit001(ByVal control As IRibbonControl)


        Call PBBAnalyseLeistbarkeit001(control.Id)



    End Sub

    ''' <summary>
    ''' Projekt fixieren, d.h. vor dem Verschieben schützen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTFreezeProject(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange
        Dim nameCollection As New Collection

        If Not noSummaryProjectsareSelected(nameCollection) Then
            Exit Sub
        End If

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

                            Try
                                Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)

                                If tryToprotectProjectforMe(hproj.name, hproj.variantName) Then

                                    hproj.movable = False
                                    Dim tmpCollection As New Collection
                                    Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, hproj.tfZeile, tmpCollection, tmpCollection)

                                Else
                                    If awinSettings.englishLanguage Then
                                        Call MsgBox(hproj.name & ", " & hproj.variantName & " is protected " & vbLf &
                                                    "and cannot be modified. You could instead create a variant.")
                                    Else
                                        Call MsgBox(hproj.name & ", " & hproj.variantName & " ist geschützt " & vbLf &
                                                    "und kann nicht verändert werden. Sie können jedoch eine Variante anlegen.")
                                    End If
                                End If
                            Catch ex As Exception
                                Call MsgBox(ex.Message)
                            End Try

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
    ''' markiert die selektierten oder alle Projekte, die aktuell angezeigt werden 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTMarkProjects(control As IRibbonControl)
        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange
        Dim atleastOne As Boolean = False

        Call projektTafelInit()

        Dim nameCollection As New Collection

        If Not noSummaryProjectsareSelected(nameCollection) Then
            Exit Sub
        End If

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

                            Try

                                Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)
                                If hproj.marker = False Then
                                    hproj.marker = True
                                    atleastOne = True
                                    Dim tmpCollection As New Collection
                                    Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, hproj.tfZeile, tmpCollection, tmpCollection)
                                End If



                            Catch ex As Exception
                                Call MsgBox(ex.Message)
                            End Try

                        End If

                    End If
                End With
            Next

        Else

            Call markAllProjects(atleastOne)

        End If

        If atleastOne Then
            ' jetzt müssen alle Charts de-selektiert werden ...
            Call unmarkPfDiagrams()

            ' und jetzt muss noch ggf das BubbleDiagramm neu, d.h ohne Markierungen gezeichnet werden 
            Call awinNeuZeichnenDiagramme(99)
        End If


        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
    End Sub

    ''' <summary>
    ''' setzt die Markierungen der Projekte zurück ... 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTUnMarkProject(control As IRibbonControl)
        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange
        Dim atleastOne As Boolean = False

        Call projektTafelInit()

        Dim nameCollection As New Collection

        ' ein Unmark darf auf Summary Projekte gemacht werden 
        'If Not noSummaryProjectsareSelected(nameCollection) Then
        '    Exit Sub
        'End If

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

                            Try

                                Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)
                                hproj.marker = False
                                Dim tmpCollection As New Collection
                                Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, hproj.tfZeile, tmpCollection, tmpCollection)


                            Catch ex As Exception
                                Call MsgBox(ex.Message)
                            End Try

                        End If

                    End If
                End With
            Next

        Else

            Call unMarkAllProjects(atleastOne)
            'If awinSettings.englishLanguage Then
            '    Call MsgBox("select project(s first ...")
            'Else
            '    Call MsgBox("vorher Projekt(e selektieren ...")
            'End If

        End If

        If atleastOne Then
            ' jetzt müssen alle Charts de-selektiert werden ...
            Call unmarkPfDiagrams()

            ' und jetzt muss noch ggf das BubbleDiagramm neu, d.h ohne Markierungen gezeichnet werden 
            Call awinNeuZeichnenDiagramme(99)
        End If


        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

    End Sub

    ''' <summary>
    ''' Projekt-Fixierung aufheben, d.h. es kann verschoben werden 
    ''' darf aber nur für Status = geplant, beauftragt und noch nicht begonnen, alle Stati mit Variante ausser abgebrochen oder abgeschlossen
    ''' gemacht werden 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PTDeFreezeProject(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange

        Call projektTafelInit()

        Dim nameCollection As New Collection

        If Not noSummaryProjectsareSelected(nameCollection) Then
            Exit Sub
        End If

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

                            Try
                                Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)

                                If hproj.Status = ProjektStatus(PTProjektStati.geplant) Or
                                    (hproj.variantName <> "" And Not hproj.Status = ProjektStatus(PTProjektStati.abgebrochen) And
                                     Not hproj.Status = ProjektStatus(PTProjektStati.abgeschlossen)) Then

                                    If tryToprotectProjectforMe(hproj.name, hproj.variantName) Then

                                        hproj.movable = True
                                        Dim tmpCollection As New Collection
                                        Call ZeichneProjektinPlanTafel(tmpCollection, hproj.name, hproj.tfZeile, tmpCollection, tmpCollection)

                                    Else
                                        If awinSettings.englishLanguage Then
                                            Call MsgBox(hproj.name & ", " & hproj.variantName & " is protected " & vbLf &
                                                        "and cannot be modified. You could instead create a variant.")
                                        Else
                                            Call MsgBox(hproj.name & ", " & hproj.variantName & " ist geschützt " & vbLf &
                                                        "und kann nicht verändert werden. Sie können jedoch eine Variante anlegen.")
                                        End If
                                    End If

                                Else
                                    ' nicht erlaubt 
                                    If awinSettings.englishLanguage Then
                                        Call MsgBox(hproj.name & ", " & hproj.variantName & " must not be moved protected.")
                                    Else
                                        Call MsgBox(hproj.name & ", " & hproj.variantName & " darf nicht verschoben / verkürzt / verlängert werden.")
                                    End If
                                End If


                            Catch ex As Exception
                                Call MsgBox(ex.Message)
                            End Try

                        End If

                    End If
                End With
            Next

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("select project(s first ...")
            Else
                Call MsgBox("vorher Projekt(e selektieren ...")
            End If

        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE

    End Sub

    Sub PT2ProjektBeauftragen(control As IRibbonControl)

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
            If awinSelection.Count > 1 Then
                If awinSettings.englishLanguage Then
                    Call MsgBox("select only 1 project, please ...")
                Else
                    Call MsgBox("bitte nur 1 Projekt selektieren, bitte ...")
                End If
            Else


                For Each singleShp In awinSelection

                    Dim shapeArt As Integer
                    shapeArt = kindOfShape(singleShp)

                    With singleShp
                        If isProjectType(shapeArt) Then

                            If ShowProjekte.contains(.Name) Then
                                Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)

                                If tryToprotectProjectforMe(hproj.name, hproj.variantName) Then
                                    Call changeProjectStatus(pname:=hproj.name, type:=PTProjektStati.beauftragt)

                                Else
                                    If awinSettings.englishLanguage Then
                                        Call MsgBox(hproj.name & ", " & hproj.variantName & " is protected " & vbLf &
                                                "and cannot be modified. You could instead create a variant.")
                                    Else
                                        Call MsgBox(hproj.name & ", " & hproj.variantName & " ist geschützt " & vbLf &
                                                "und kann nicht verändert werden. Sie können jedoch eine Variante anlegen.")
                                    End If
                                End If
                            End If

                        End If
                    End With
                Next


            End If

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("first select a project, please ...")
            Else
                Call MsgBox("vorher ein Projekt selektieren, bitte ...")
            End If

        End If

        enableOnUpdate = True
        appInstance.EnableEvents = formerEE
    End Sub

    ''' <summary>
    ''' den Status eines Projekts ändern, aktuell nur auf Projekt-status = 1 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PT2ProjektStatusChange(control As IRibbonControl)

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
                                Call changeProjectStatus(pname:=hproj.name, type:=PTProjektStati.beauftragt)

                            Else
                                If awinSettings.englishLanguage Then
                                    Call MsgBox(hproj.name & ", " & hproj.variantName & " is protected " & vbLf &
                                                "and cannot be modified. You could instead create a variant.")
                                Else
                                    Call MsgBox(hproj.name & ", " & hproj.variantName & " ist geschützt " & vbLf &
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
        Dim listOfArchivFiles As New List(Of String)
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
            dateiName = CStr(getMassenEditImport.selImportFiles.Item(1))

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

                    ' tk importiereMassenEdit wurde auskommentiert - muss ggf gesucht und dann wieder einkommentiert werden ... 
                    'Call importiereMassenEdit()
                    Call MsgBox("aktuell aus dem Funktionsumfang rausgenommen ...")
                    appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                    Dim sessionConstellation As clsConstellation = verarbeiteImportProjekte(scenarioName, noComparison:=True)

                    ' ''If wasNotEmpty Then
                    ' ''    Call awinClearPlanTafel()
                    ' ''End If

                    '' ''Call awinZeichnePlanTafel(True)
                    ' ''Call awinZeichnePlanTafel(False)
                    ' ''Call awinNeuZeichnenDiagramme(2)

                    Dim scenarioPVName As String = calcPortfolioKey(sessionConstellation)
                    If sessionConstellation.count > 0 Then

                        If projectConstellations.Contains(scenarioPVName) Then
                            projectConstellations.Remove(scenarioPVName)
                        End If

                        projectConstellations.Add(sessionConstellation)
                        Call loadSessionConstellation(scenarioPVName, False, True)

                        listOfArchivFiles.Add(dateiName)
                    Else
                        Call MsgBox("keine Projekte importiert ...")
                    End If

                    If ImportProjekte.Count > 0 Then
                        ImportProjekte.Clear(False)
                    End If
                Else

                    Call MsgBox("bitte Datei auswählen ...")
                End If

                ' verschieben der erfolgreich importierten files
                If listOfArchivFiles.Count > 0 Then
                    Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.massenEdit))
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

        Dim noScenarioCreation As Boolean = False
        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult

        Dim listOfArchivFiles As New List(Of String)

        Dim ohneFehler As Boolean = True

        Dim getInventurImport As New frmSelectImportFiles
        Dim wasNotEmpty As Boolean = False

        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' wenn noch etwas in der session ist , warnen ! 
        If AlleProjekte.Count > 0 Then
            If awinSettings.englishLanguage Then
                Call MsgBox("this function is only available with an empty session" & vbLf &
                            "please store and clear your session first")
            Else
                Call MsgBox("diese Funktionalität ist nur möglich mit einer leeren Session" & vbLf &
                            "bitte speichern Sie ggf. ihre Projekte und setzen die Session zurück.")
            End If
        Else
            ' Aktion durchführen ...
            'getInventurImport.menueAswhl = PTImpExp.simpleScen
            getInventurImport.menueAswhl = PTImpExp.batchlists
            returnValue = getInventurImport.ShowDialog

            If returnValue = DialogResult.OK Then
                dateiName = CStr(getInventurImport.selImportFiles.Item(1))

                Try

                    Dim logmsg() As String

                    If My.Computer.FileSystem.FileExists(dateiName) Then

                        If ShowProjekte.Count > 0 Then
                            wasNotEmpty = True
                            'Call storeSessionConstellation("Last")
                            ' hier sollte jetzt auch ein ClearPlan-Tafel gemacht werden ...
                            Call awinClearPlanTafel()
                        End If

                        appInstance.Workbooks.Open(dateiName)

                        Dim scenarioNameP As String = appInstance.ActiveWorkbook.Name
                        Dim scenarioNameS As String = scenarioNameP & " (programs)"
                        Dim positionIX As Integer = scenarioNameP.IndexOf(".xls") - 1
                        Dim tmpName As String = ""
                        For ih As Integer = 0 To positionIX
                            tmpName = tmpName & scenarioNameP.Chars(ih)
                        Next
                        scenarioNameP = tmpName.Trim

                        ' alle Import Projekte erstmal löschen
                        ImportProjekte.Clear(False)
                        Dim isAllianzImport1 As Boolean = False

                        Try
                            'If scenarioNameP.StartsWith("Allianz-Typ 1") Then
                            If scenarioNameP.StartsWith("Rupi-Liste") Then
                                ' das muss noch abgefragt werden ... 

                                Dim startdate As Date = CDate("1.1.2019")
                                Dim enddate As Date = CDate("31.12.2019")

                                ' tk aktuelle Krücke für Allianz um das Einlesen der Batchliste für 2020 zu machen
                                ' dazu muss die Datenbank mit 20 enden ...
                                If awinSettings.databaseName.EndsWith("20") Then
                                    startdate = CDate("1.1.2020")
                                    enddate = CDate("31.12.2020")
                                End If

                                isAllianzImport1 = True
                                Call importAllianzType1(startdate, enddate)

                            ElseIf scenarioNameP.StartsWith("BOBS") Then

                                Dim startdate As Date = CDate("1.1.2020")
                                Dim enddate As Date = CDate("31.12.2020")

                                isAllianzImport1 = True
                                Call importAllianzBOBS(startdate, enddate)

                            ElseIf scenarioNameP.StartsWith("Allianz-Typ 2") Then

                                noScenarioCreation = True
                                Call importAllianzType2()

                            ElseIf scenarioNameP.StartsWith("Istdaten") Then
                                ' immer zwei Monate zurück gehen 
                                ' erst mal immer automatisch auf aktuelles Datum -1  setzen 

                                Dim editActualDataMonth As New frmProvideActualDataMonth

                                If editActualDataMonth.ShowDialog = DialogResult.OK Then

                                    Dim monat As Integer = CInt(editActualDataMonth.valueMonth.Text)

                                    Dim readPastAndFutureData As Boolean = editActualDataMonth.readPastAndFutureData.Checked
                                    Dim createUnknownProjects As Boolean = editActualDataMonth.createUnknownProjects.Checked

                                    Dim outputCollection As New Collection
                                    Call MsgBox("not yet implemented ...")
                                    'Call ImportIstdatenStdFormat(monat, readPastAndFutureData, createUnknownProjects, outputCollection)

                                End If


                            ElseIf scenarioNameP.StartsWith("Allianz-Typ 4") Then
                                Call importAllianzType4()

                            Else
                                Call awinImportProjektInventur()
                            End If

                            listOfArchivFiles.Add(dateiName)

                        Catch ex As Exception

                            Call MsgBox("Fehler bei Import : " & ex.Message)
                            ohneFehler = False

                        End Try

                        appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                        If ohneFehler Then
                            'sessionConstellationP enthält alle Projekte aus dem Import 
                            'Dim sessionConstellationP As clsConstellation = verarbeiteImportProjekte(scenarioNameP, noComparison:=False, considerSummaryProjects:=False)
                            Dim sessionConstellationP As clsConstellation = verarbeiteImportProjekte(scenarioNameP, noComparison:=False, considerSummaryProjects:=False)
                            Dim sessionConstellationS As clsConstellation = Nothing


                            ' tk 8.5.19 das soll jetzt nicht mehr gemacht werden - immer alle Projekte zeigen, die importiert wurden und sich verändert haben 
                            If isAllianzImport1 Then
                                sessionConstellationS = verarbeiteImportProjekte(scenarioNameS, noComparison:=True, considerSummaryProjects:=True)
                            End If

                            'Call sessionConstellationP.calcUnionProject(False)
                            'Call sessionConstellationS.calcUnionProject(False)

                            ' Testen ..
                            ' test
                            If isAllianzImport1 Then
                                Dim everythingOK As Boolean = testUProjandSingleProjs(sessionConstellationP, False)
                                If Not everythingOK Then
                                    ReDim logmsg(1)
                                    logmsg(0) = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben:"
                                    logmsg(1) = sessionConstellationP.constellationName
                                    Call logger(ptErrLevel.logError, "Tom2G4B1InventurImport", logmsg)
                                End If
                                ' ende test


                                ' test
                                everythingOK = testUProjandSingleProjs(sessionConstellationS, False)
                                If Not everythingOK Then
                                    ReDim logmsg(1)
                                    logmsg(0) = "Summary Projekt nicht identisch mit der Liste der Projekt-Vorhaben:"
                                    logmsg(1) = sessionConstellationP.constellationName
                                    Call logger(ptErrLevel.logError, "Tom2G4B1InventurImport", logmsg)
                                End If
                                ' ende test
                            End If


                            ' auch im Fall Allianz sollen die Projekte gezeigt werden - nicht die Summary-Projekte 
                            If projectConstellations.Contains(scenarioNameP) Then
                                projectConstellations.Remove(scenarioNameP)
                            End If

                            If projectConstellations.Contains(scenarioNameS) Then
                                projectConstellations.Remove(scenarioNameS)
                            End If

                            ' tk 22.7.19 es sollen beide Constellations in project-Constellations geschrieben werden ... 
                            ' tk 12.8.19 diese beiden Constellations sollen nicht mehr eingetragen werden , nur noch die Rupi-Liste 

                            If Not IsNothing(sessionConstellationS) Then
                                projectConstellations.Add(sessionConstellationS)
                            End If

                            If Not IsNothing(sessionConstellationP) Then
                                projectConstellations.Add(sessionConstellationP)
                                ' jetzt auf Projekt-Tafel anzeigen 

                                currentConstellationPvName = sessionConstellationP.constellationName
                                ' tk 2.12.19 jetzt wird diese Constellation gezeichnet 
                                ' die andere kann dann über loadConstelaltion gezeichnet werden 
                                Call awinZeichnePlanTafel(sessionConstellationP)
                            End If

                            'Call loadSessionConstellation(scenarioNameP, False, True)

                            '' tk 8.5.19 auskommentiert 
                            'If isAllianzImport1 Then
                            '    If sessionConstellationS.count > 0 Then

                            '        If projectConstellations.Contains(scenarioNameS) Then
                            '            projectConstellations.Remove(scenarioNameS)
                            '        End If

                            '        projectConstellations.Add(sessionConstellationS)
                            '        ' jetzt auf Projekt-Tafel anzeigen 

                            '        Call loadSessionConstellation(scenarioNameS, False, True)

                            '    Else
                            '        Call MsgBox("keine Programmlinien importiert ...")
                            '    End If
                            'Else
                            '    If sessionConstellationP.count > 0 Then

                            '        If projectConstellations.Contains(scenarioNameP) Then
                            '            projectConstellations.Remove(scenarioNameP)
                            '        End If

                            '        projectConstellations.Add(sessionConstellationP)
                            '        ' jetzt auf Projekt-Tafel anzeigen 
                            '        Call loadSessionConstellation(scenarioNameP, False, True)

                            '    Else
                            '        Call MsgBox("keine Projekte importiert ...")
                            '    End If
                            'End If

                            ' ImportDatei ins archive-Directory schieben
                            If listOfArchivFiles.Count > 0 Then
                                Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.batchlists))
                            End If



                            If ImportProjekte.Count > 0 Then
                                ImportProjekte.Clear(False)
                            End If
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

        ' so positionieren, dass die Projekte auch sichtbar sind ...
        If boardWasEmpty Then
            If ShowProjekte.Count > 0 Then
                Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True

        'projectboardWindows(PTwindows.mpt).Activate()

        appInstance.ScreenUpdating = True





    End Sub

    Public Sub Tom2G4B1ScenarioImport(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

        Dim dateiName As String
        Dim myCollection As New Collection
        Dim importDate As Date = Date.Now
        Dim returnValue As DialogResult
        Dim listOfArchivFiles As New List(Of String)

        Dim getScenarioImport As New frmSelectImportFiles
        Dim wasNotEmpty As Boolean = False
        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        ' Aktion durchführen ...
        getScenarioImport.menueAswhl = PTImpExp.scenariodefs
        returnValue = getScenarioImport.ShowDialog

        If returnValue = DialogResult.OK Then
            dateiName = CStr(getScenarioImport.selImportFiles.Item(1))

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

                        projectConstellations.Add(newConstellation)


                        ' Beginn

                        Dim constellationsToDo As New clsConstellations
                        constellationsToDo.Add(newConstellation)

                        Dim clearBoard As Boolean = True
                        Dim clearSession As Boolean = False
                        If constellationsToDo.Count > 0 Then
                            Call showConstellations(constellationsToDo, clearBoard, clearSession, Date.Now)
                        End If

                        ' jetzt muss die Info zu den Schreibberechtigungen geholt werden 
                        If Not noDB Then
                            writeProtections.adjustListe = CType(databaseAcc, DBAccLayer.Request).retrieveWriteProtectionsFromDB(AlleProjekte, err)
                        End If

                    Else
                        Call MsgBox("keine Projekte für Portfolio erkannt ...")
                    End If

                    ' erfolgreich importierte Files aufsammeln
                    listOfArchivFiles.Add(dateiName)

                    If ImportProjekte.Count > 0 Then
                        ImportProjekte.Clear(False)
                    End If
                Else

                    Call MsgBox("bitte Datei auswählen ...")
                End If

                ' verschieben der erfolgreich importierten files
                If listOfArchivFiles.Count > 0 Then
                    Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.scenariodefs))
                End If

            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
            End Try
        Else
            'Call MsgBox(" Import Scenario wurde abgebrochen")
        End If

        ' so positionieren, dass die Projekte auch sichtbar sind ...
        If boardWasEmpty Then
            If ShowProjekte.Count > 0 Then
                Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
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
        Dim listOfArchivFiles As New List(Of String)
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
            dateiName = CStr(getModuleImport.selImportFiles.Item(1))

            Try
                appInstance.Workbooks.Open(dateiName)

                ' alle Import Projekte erstmal löschen
                ImportProjekte.Clear(False)
                Call awinImportModule(myCollection)
                appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
                Call importProjekteEintragen(importDate, True, False, False)

                listOfArchivFiles.Add(dateiName)

            Catch ex As Exception
                appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
            End Try

            ' verschieben der erfolgreich importierten files
            If listOfArchivFiles.Count > 0 Then
                Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.modulScen))
            End If
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
        Dim listOfArchivFiles As New List(Of String)
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
            dateiName = CStr(getModuleImport.selImportFiles.Item(1))
            Dim ruleSet As New clsAddElements
            Dim ok As Boolean = True
            Try
                appInstance.Workbooks.Open(dateiName)

                ' jetzt werden die Regeln ausgelesen ...
                Call awinReadAddOnRules(ruleSet)
                appInstance.ActiveWorkbook.Close(SaveChanges:=True)

                listOfArchivFiles.Add(dateiName)

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

                ' verschieben der erfolgreich importierten files
                If listOfArchivFiles.Count > 0 Then
                    Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.addElements))
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
        Dim listOfArchivFiles As New List(Of String)
        'Dim xlsRplanImport As Excel.Workbook

        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

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

                    ' list of Files, which are imported with success
                    listOfArchivFiles.Add(dateiName)

                    appInstance.ScreenUpdating = True
                    Call importProjekteEintragen(importDate, True, True, True)

                    'Call awinWritePhaseDefinitions()
                    'Call awinWritePhaseMilestoneDefinitions()

                Catch ex As Exception
                    appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                    Call MsgBox("Fehler bei Import " & vbLf & dateiName & vbLf & ex.Message)
                End Try

            Next i

            ' verschieben der erfolgreich importierten files
            If listOfArchivFiles.Count > 0 Then
                Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.rplan))
            End If

        Else
            'Call MsgBox(" Import RPLAN-Projekte wurde abgebrochen")
            'Call logfileSchreiben(" Import RPLAN-Projekte wurde abgebrochen", dateiName, -1)
        End If

        ' so positionieren, dass die Projekte auch sichtbar sind ...
        If boardWasEmpty Then
            If ShowProjekte.Count > 0 Then
                Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
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
        Dim listOfArchivFiles As New List(Of String)
        Dim protokoll As New SortedList(Of Integer, clsProtokoll)

        ' öffnen des LogFiles
        ''Call logfileOpen()

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        'dateiName = awinPath & projektInventurFile

        getRPLANImport.menueAswhl = PTImpExp.rplanrxf
        returnValue = getRPLANImport.ShowDialog

        If returnValue = DialogResult.OK Then
            dateiName = CStr(getRPLANImport.selImportFiles.Item(1))

            Try

                ' alle Import Projekte erstmal löschen
                ImportProjekte.Clear(False)

                Call logger(ptErrLevel.logInfo, "Beginn RXFImport ", dateiName, -1)

                Call RXFImport(myCollection, dateiName, False, protokoll)

                'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))
                Call importProjekteEintragen(importDate, True, True, True)

                ' aufsammeln der zu archivierenden Files
                listOfArchivFiles.Add(dateiName)

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

                ' verschieben der erfolgreich importierten files
                If listOfArchivFiles.Count > 0 Then
                    Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.rplan))
                End If

            Catch ex As Exception

                Call MsgBox(ex.Message & vbLf & dateiName & vbLf & "Fehler bei RXFImport ")
            End Try

        Else
            'Call MsgBox(" RXF-Import RPLAN-Projekte wurde abgebrochen")
            'Call logfileSchreiben(" RXF-Import RPLAN-Projekte wurde abgebrochen", dateiName, -1)
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub


    ''' <summary>
    ''' importiert und speichert die Organisation; wenn mehrere existieren, dann wird ein Formular aufgeschaltet zur Auswahl der Organisation
    ''' </summary>
    ''' <param name="Control"></param>
    Public Sub PTImportOrga(Control As IRibbonControl)

        Dim selectedWB As String = ""
        'Dim dirname As String = My.Computer.FileSystem.CombinePath(awinPath, requirementsOrdner)
        Dim dirname As String = importOrdnerNames(PTImpExp.Orga)
        Dim dateiname As String = ""


        Dim outputCollection As New Collection

        ' ===========================================================
        ' Konfigurationsdatei lesen und Validierung durchführen

        ' wenn es gibt - lesen der ControllingSheet und anderer, die durch configActualDataImport beschrieben sind
        Dim configOrgaImport As String = awinPath & configfilesOrdner & "configOrgaImport.xlsx"
        Dim orgaImportConfig As New SortedList(Of String, clsConfigOrgaImport)
        Dim lastrow As Integer = 0

        ' check Config-File - zum Einlesen der Istdaten gemäß Konfiguration
        ' hier werden Werte für actualDataFile, actualDataConfig gesetzt
        Dim allesOK As Boolean = checkOrgaImportConfig(configOrgaImport, dateiname, orgaImportConfig, lastrow, outputCollection)


        Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname, FileIO.SearchOption.SearchTopLevelOnly, "*rganisation*.xls*")
        Dim anzFiles As Integer = listOfImportfiles.Count

        ' tk by Ute für das Verschieben der Datei in den Archiv-Ordner, wenn erfolgreich 
        Dim listOfArchivFiles As New List(Of String)

        Dim weiterMachen As Boolean = False

        'Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' öffnen des LogFiles
        'Call logfileOpen()

        If anzFiles = 1 Then
            selectedWB = listOfImportfiles.Item(0)
            weiterMachen = True

        ElseIf anzFiles > 1 Then
            Dim getOrgaFile As New frmSelectImportFiles
            getOrgaFile.menueAswhl = PTImpExp.Orga
            Dim returnValue As DialogResult = getOrgaFile.ShowDialog

            If returnValue = DialogResult.OK Then
                selectedWB = CStr(getOrgaFile.selImportFiles.Item(1))
                ' Check if Config or not
                weiterMachen = True
            End If
        Else
            Call MsgBox("keine Organisations-Dateien gefunden ..." & vbLf & "Folder: " & dirname & vbLf & "Dateien müssen folgender Namensgebung genügen *rganisation*.xls*")
        End If


        If weiterMachen Then

            dateiname = My.Computer.FileSystem.CombinePath(dirname, selectedWB)


            Try
                ' hier wird jetzt der Import gemacht 
                Call logger(ptErrLevel.logInfo, "Beginn Import Organisation ", selectedWB, -1)

                ' Öffnen des Organisations-Files
                appInstance.Workbooks.Open(dateiname)

                ' Dim importedOrga As clsOrganisation = ImportOrganisation(outputCollection)
                Dim importedOrga As clsOrganisation = ImportOrganisation(outputCollection, orgaImportConfig)

                Dim wbName As String = My.Computer.FileSystem.GetName(dateiname)

                ' Schliessen des Organisations-Files
                appInstance.Workbooks(wbName).Close(SaveChanges:=True)

                If outputCollection.Count > 0 Then
                    Dim errmsg As String = vbLf & " .. Abbruch .. nicht importiert "
                    outputCollection.Add(errmsg)
                    Call showOutPut(outputCollection, "Organisations-Import", "")

                    Call logger(ptErrLevel.logError, "PTImportOrga: ", outputCollection)

                ElseIf importedOrga.count > 0 Then

                    ' TopNodes und OrgaTeamChilds bauen 
                    Call importedOrga.allRoles.buildTopNodes()

                    ' wird bereits in buildTopNodes gemacht 
                    'Call importedOrga.allRoles.buildOrgaSkillChilds()

                    ' jetzt wird die Orga als Setting weggespeichert ... 
                    Dim err As New clsErrorCodeMsg
                    Dim result As Boolean = False
                    ' ute -> überprüfen bzw. fertigstellen ... 
                    Dim orgaName As String = ptSettingTypes.organisation.ToString

                    ' andere Rollen als Orga-Admin können Orga einlesen, aber eben nicht speichern ! 
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Then
                        result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(importedOrga,
                                                                                    CStr(settingTypes(ptSettingTypes.organisation)),
                                                                                    orgaName,
                                                                                    importedOrga.validFrom,
                                                                                    err)

                    Else
                        result = True
                    End If


                    If result = True Then
                        ' importierte Organisation in die Liste der validOrganisations aufnehmen
                        validOrganisations.addOrga(importedOrga)
                        Dim currentOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)

                        ' Roledefinitions und Costdefinitions neu setzen
                        If Not IsNothing(currentOrga) Then
                            RoleDefinitions = currentOrga.allRoles
                            CostDefinitions = currentOrga.allCosts
                        Else
                            If awinSettings.englishLanguage Then
                                Call MsgBox("You don't have any valid (time now) Organisation in the system!")
                            Else
                                Call MsgBox("Es existiert keine heute gültige Organisation im System!")
                            End If
                        End If

                        listOfArchivFiles.Add(dateiname)

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Then
                            Call MsgBox("ok, Organisation, valid from " & importedOrga.validFrom.ToShortDateString & " stored ...")
                            Call logger(ptErrLevel.logInfo, "Organisation, valid from " & importedOrga.validFrom.ToString & " stored ...", selectedWB, -1)
                        Else
                            Call MsgBox("ok, Organisation, valid from " & importedOrga.validFrom.ToShortDateString & " temporarily loaded ...")
                            Call logger(ptErrLevel.logInfo, "Organisation, valid from " & importedOrga.validFrom.ToShortDateString & " temporarily loaded ...", selectedWB, -1)
                        End If

                    Else
                        Call MsgBox("Error when writing Organisation")
                        Call logger(ptErrLevel.logError, "Error when writing Organisation ...", selectedWB, -1)
                    End If
                End If
            Catch ex As Exception

            End Try
        End If




        ' Schließen des LogFiles
        ''Call logfileSchliessen()

        If listOfArchivFiles.Count > 0 And myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Then
            Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.Orga))
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    ''' <summary>
    ''' ordnet existierenden Projekten pro Phase eine Ressourcen-Summe zu. 
    ''' wenn bereits Istdaten existieren, so werden die angegebenen Summen so verteilt, dass 
    ''' die Formel Istdaten+Prognose = Summe 
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub PTImportOfflineData(control As IRibbonControl)


        Dim dirname As String = My.Computer.FileSystem.CombinePath(awinPath, importOrdnerNames(PTImpExp.offlineData))
        Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname, FileIO.SearchOption.SearchTopLevelOnly, "*ffline*.xls*")
        Dim anzFiles As Integer = listOfImportfiles.Count

        Dim dateiname As String = ""

        Dim weiterMachen As Boolean = False

        'Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False


        Dim getOrgaFile As New frmSelectImportFiles

        If anzFiles > 0 Then

            getOrgaFile.menueAswhl = PTImpExp.offlineData
            Dim returnValue As DialogResult = getOrgaFile.ShowDialog

            If returnValue = DialogResult.OK Then
                weiterMachen = True
            End If
        Else
            Call MsgBox("keine Offline-Daten vorhanden ..." & vbLf & "Folder: " & dirname & vbLf & "Name muss folgender Namensgebung entsprechen: *ffline*.xls*")
            weiterMachen = False
        End If

        If weiterMachen Then

            ' öffnen des LogFiles
            'Call logfileOpen()

            For Each selectedWB As String In getOrgaFile.selImportFiles

                dateiname = My.Computer.FileSystem.CombinePath(dirname, selectedWB)

                Try
                    ' hier wird jetzt der Import gemacht 
                    Call logger(ptErrLevel.logInfo, "Beginn Import Offline-Daten", selectedWB, -1)

                    ' Öffnen des Offline Data -Files
                    appInstance.Workbooks.Open(dateiname)
                    Dim offlineName As String = appInstance.ActiveWorkbook.Name

                    Dim outputCollection As New Collection

                    ' jetzt wird die Aktion durchgeführt ...
                    Call ImportOfflineData(offlineName, outputCollection)

                    ' Schliessen des CustomUser Role-Files
                    appInstance.Workbooks(offlineName).Close(SaveChanges:=True)

                    ' -----------------------------------------------------------------------------
                    ' Start Verarbeitung Import-Projekte verarbeitet 
                    'sessionConstellationP enthält alle Projekte aus dem Import 
                    Dim importScenarioName As String = "offline Data"
                    Dim importConstellation As clsConstellation = verarbeiteImportProjekte(importScenarioName, noComparison:=False, considerSummaryProjects:=False)

                    Dim importScenarioPVName As String = calcPortfolioKey(importScenarioName, "")
                    If importConstellation.count > 0 Then

                        If projectConstellations.Contains(importScenarioPVName) Then
                            projectConstellations.Remove(importScenarioPVName)
                        End If

                        projectConstellations.Add(importConstellation)
                        ' jetzt auf Projekt-Tafel anzeigen 
                        Call loadSessionConstellation(importScenarioPVName, False, True)

                    Else
                        logmessage = "keine Projekte importiert ..."
                        outputCollection.Add(logmessage)
                    End If

                    If ImportProjekte.Count > 0 Then
                        ImportProjekte.Clear(False)
                    End If


                    ' Ende Verarbeitung Import-Projekte
                    ' -----------------------------------------------------------------------------
                    If outputCollection.Count > 0 Then
                        Call showOutPut(outputCollection, "Import Offline Data " & offlineName, "")
                        outputCollection.Clear()
                    End If

                Catch ex As Exception

                End Try
            Next



            ' Schließen des LogFiles
            ''Call logfileSchliessen()

        End If


        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


    End Sub

    ''' <summary>
    ''' importiert und speichert die Istdaten; darf nur Orga-Admin
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub PTImportIstDaten(control As IRibbonControl)

        Dim importDate As Date = Date.Now
        Dim weitermachen As Boolean = False
        Dim selectedWB As String = ""
        Dim actualDataFile As String = ""
        Dim actualDataConfig As New SortedList(Of String, clsConfigActualDataImport)
        Dim outPutCollection As New Collection
        Dim outPutline As String = ""
        Dim lastrow As Integer
        Dim result As Boolean = False
        Dim listOfArchivFiles As New List(Of String)
        Dim listOfArchivFilesAllg As New List(Of String)
        Dim listOfErrorImportFilesAllg As New List(Of String)
        Dim dateiname As String = ""
        Dim dirname As String = My.Computer.FileSystem.CombinePath(awinPath, importOrdnerNames(PTImpExp.actualData))
        Dim anzFiles As Integer = 0

        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

        ' erstmal protokollieren, zu welchen Abteilungen Istdaten gelesen und substituiert werden 
        ' alle Planungen zu den Rollen, die in dieser Referatsliste aufgeführt sind, werden gelöscht 
        Dim istDatenReferatsliste() As Integer

        If awinSettings.ActualdataOrgaUnits = "" Then
            Dim anzTopNodes As Integer = RoleDefinitions.getTopLevelNodeIDs.Count
            ReDim istDatenReferatsliste(anzTopNodes - 1)
            Dim i As Integer = 0
            For i = 0 To anzTopNodes - 1
                istDatenReferatsliste(i) = RoleDefinitions.getTopLevelNodeIDs.Item(i)
            Next
        Else
            istDatenReferatsliste = RoleDefinitions.getIDArray(awinSettings.ActualdataOrgaUnits)
        End If

        ' nimmt auf, zu welcher Orga-Einheit die Ist-Daten erfasst werden ... 
        Dim referatsCollection As New Collection
        Dim msgTxt As String = "Actual Data Departments: "
        Dim first As Boolean = True
        For Each itemID As Integer In istDatenReferatsliste
            Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(itemID)
            If Not IsNothing(tmpRole) Then
                If Not referatsCollection.Contains(tmpRole.name) Then
                    referatsCollection.Add(tmpRole.name, tmpRole.name)
                End If
                If first Then
                    msgTxt = msgTxt & tmpRole.name
                    first = False
                Else
                    msgTxt = msgTxt & ", " & tmpRole.name
                End If
            End If
        Next

        Call logger(ptErrLevel.logInfo, msgTxt, "PTImportIstdaten", anzFehler)

        ' Art und Weise 1: Datei lautet auf "Istdaten*.xlsx

        Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname, FileIO.SearchOption.SearchTopLevelOnly, "Istdaten*.xls*")
        anzFiles = listOfImportfiles.Count

        If anzFiles > 0 Then

            If anzFiles = 1 Then
                selectedWB = listOfImportfiles.Item(0)
                weitermachen = True

            ElseIf anzFiles > 1 Then

                Dim getOrgaFile As New frmSelectImportFiles
                getOrgaFile.menueAswhl = PTImpExp.actualData
                Dim returnValue As DialogResult = getOrgaFile.ShowDialog

                If returnValue = DialogResult.OK Then
                    selectedWB = CStr(getOrgaFile.selImportFiles.Item(1))
                    weitermachen = True
                End If

            End If

            If weitermachen Then

                ' öffnen des LogFiles
                'Call logfileOpen()

                dateiname = My.Computer.FileSystem.CombinePath(dirname, selectedWB)
                result = readActualData(dateiname)
                If result Then
                    listOfArchivFiles.Add(dateiname)
                End If

                ' Schließen des LogFiles
                ''Call logfileSchliessen()

                ' ImportDatei ins archive-Directory schieben

                If listOfArchivFiles.Count > 0 Then
                    Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.actualData))
                End If

                ' es kann nur die eine oder andere Art des Imports geben , falls hier importiert wurde 
            End If

        Else
            ' Konfigurations-Dateien lesen 
            ' ===========================================================
            ' Konfigurationsdatei lesen und Validierung durchführen
            Dim configActualDataImport As String = awinPath & configfilesOrdner & "configActualDataImport.xlsx"

            ' check Config-File - zum Einlesen der Istdaten gemäß Konfiguration
            ' hier werden Werte für actualDataFile, actualDataConfig gesetzt
            Dim allesOK As Boolean = checkActualDataImportConfig(configActualDataImport, actualDataFile, actualDataConfig, lastrow, outPutCollection)

            If allesOK Then

                Call projektTafelInit()

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                enableOnUpdate = False


                Dim listOfImportfilesAllg As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname, FileIO.SearchOption.SearchTopLevelOnly, actualDataFile)
                anzFiles = listOfImportfilesAllg.Count

                If listOfImportfilesAllg.Count >= 1 Then
                    ' Vorbereitungen für die Aufnahme der verschiedenen Excel-File Daten in die unterschiedlichen Projekte
                    Dim editActualDataMonth As New frmInfoActualDataMonth
                    Dim lastValidMonth As Integer = 0  ' angegeben in dem Dialog
                    Dim IstdatenDate As Date
                    Dim curMonth As Integer = 0
                    Dim hrole As New clsRollenDefinition
                    Dim cacheProjekte As New clsProjekteAlle



                    If editActualDataMonth.ShowDialog = DialogResult.OK Then

                        ' Istdaten immer vom Vormonat einlesen
                        IstdatenDate = CDate(editActualDataMonth.MonatJahr.Text).AddMonths(-1)

                        Dim referenzPortfolioName As String = ""
                        If Not IsNothing(editActualDataMonth.comboBxPortfolio.SelectedItem) Then
                            referenzPortfolioName = editActualDataMonth.comboBxPortfolio.SelectedItem.ToString
                        End If

                        Dim curTimeStamp As Date = Date.MinValue
                        Dim err As New clsErrorCodeMsg
                        Dim referenzPortfolio As clsConstellation = Nothing

                        If referenzPortfolioName = "" Then

                            Dim txtMsg As String = "kein Portfolio gewählt - Abbruch!"
                            If awinSettings.englishLanguage Then
                                txtMsg = "no Portfolio selected - Cancelled ..."
                            End If

                            Call MsgBox(txtMsg)
                            ''Call logfileSchliessen()

                            enableOnUpdate = True
                            appInstance.EnableEvents = True
                            appInstance.ScreenUpdating = True

                            Exit Sub

                        End If

                        ' gibt es das Referenz-Portfolio?  
                        referenzPortfolio = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(referenzPortfolioName,
                                                                                                          "",
                                                                                                          curTimeStamp,
                                                                                                          err,
                                                                                                          variantName:="",
                                                                                                          storedAtOrBefore:=Date.Now)

                        If IsNothing(referenzPortfolio) Then
                            Dim txtMsg As String = referenzPortfolioName & ": Portfolio existiert nicht ... "
                            If awinSettings.englishLanguage Then
                                txtMsg = referenzPortfolioName & ": Portfolio does not exist - Cancelled ..."
                            End If

                            Call MsgBox(txtMsg)

                            ''Call logfileSchliessen()

                            enableOnUpdate = True
                            appInstance.EnableEvents = True
                            appInstance.ScreenUpdating = True

                            Exit Sub
                        End If

                        ' jetzt kann weitergemacht werden ... 

                        ' im Key steht der Projekt-Name, im Value steht eine sortierte Liste mit key=Rollen-Name, values die Ist-Werte
                        Dim validProjectNames As New SortedList(Of String, SortedList(Of String, Double()))


                        ' nimmt dann später pro Projekt die vorkommenden Rollen auf - setzt voraus, dass die Datei nach Projekt-Namen, dann nach Jahr, dann nach Monat sortiert ist ...  
                        Dim projectRoleNames(,) As String = Nothing

                        ' nimmt dann die Werte pro Projekt, Rolle und Monat auf  
                        Dim projectRoleValues(,,) As Double = Nothing

                        Dim updatedProjects As Integer = 0

                        Dim logF_Fehler As Integer = 0
                        ' nimmt die Texte für die LogFile Zeile auf
                        ' Array kann beliebig lang werden 
                        Dim logArray() As String
                        Dim logDblArray() As Double



                        For Each tmpDatei As String In listOfImportfilesAllg
                            If awinSettings.englishLanguage Then
                                outPutline = "Reading actual-data " & tmpDatei
                            Else
                                outPutline = "Einlesen der ActualData " & tmpDatei
                            End If

                            ' tk 2.8.2020 das soll nur noch im Logfile erscheinen , aber nicht mehr im Interaktiven Fenster ...
                            'outPutCollection.Add(outPutline)

                            Call logger(ptErrLevel.logInfo, outPutline, "", anzFehler)

                            result = readActualDataWithConfig(actualDataConfig, tmpDatei,
                                                  IstdatenDate,
                                                  cacheProjekte,
                                                  validProjectNames, projectRoleNames,
                                                  projectRoleValues,
                                                  updatedProjects,
                                                  outPutCollection)

                            ' hier weitermachen

                            If result Then
                                ' hier: merken der erfolgreich importierten ActualData Files
                                listOfArchivFilesAllg.Add(tmpDatei)
                                ' Projekt in Importprojekte eintragen
                            Else
                                listOfErrorImportFilesAllg.Add(tmpDatei)
                            End If

                        Next

                        If listOfImportfilesAllg.Count = listOfArchivFilesAllg.Count Then           ' dann sind alle korrekt durchgelaufen

                            ' jetzt kommt die zweite Bearbeitungs-Welle


                            ' jetzt wird hier überprüft 
                            ' gibt es Projekte im Referenz-Portfolio, die keine Ist-Daten erhalten haben - dann sollte jetzt ggf. hier ein Nuller Eintrag im array für diese Projekte erfolgen 
                            ' 
                            ' 

                            ' was hier noch überprüft werden sollte: 
                            ' welche internen Rollen, die im besagten Zeitraum relevant,  haben keine Ist-Daten ? 
                            Dim startFiscalYearTelair As Date
                            Dim endFiscalYearTelair As Date

                            If IstdatenDate.Month - 10 >= 0 Then
                                startFiscalYearTelair = DateSerial(IstdatenDate.Year, 10, 1)
                                endFiscalYearTelair = DateSerial(IstdatenDate.Year + 1, 9, 30)
                            Else
                                startFiscalYearTelair = DateSerial(IstdatenDate.Year - 1, 10, 1)
                                endFiscalYearTelair = DateSerial(IstdatenDate.Year, 9, 30)
                            End If

                            Dim activeinternRoles() As Integer = RoleDefinitions.getActiveInterns(startFiscalYearTelair, endFiscalYearTelair)
                            Dim missingTimeSheets As New List(Of String)


                            For Each tmpUID As Integer In activeinternRoles

                                Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpUID, -1)

                                Dim found As Boolean = False
                                Dim tmprole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(tmpUID)

                                For Each kvp As KeyValuePair(Of String, SortedList(Of String, Double())) In validProjectNames
                                    Try
                                        found = kvp.Value.ContainsKey(roleNameID)
                                        If found Then
                                            Exit For
                                        End If
                                    Catch ex As Exception

                                    End Try
                                Next

                                If Not found Then

                                    If tmprole.entryDate < IstdatenDate And tmprole.exitDate > startFiscalYearTelair Then
                                        missingTimeSheets.Add(tmprole.name)
                                    End If

                                End If

                            Next

                            If missingTimeSheets.Count > 0 Then
                                For Each roleName As String In missingTimeSheets
                                    ReDim logArray(5)
                                    ' ins Protokoll eintragen 
                                    logArray(0) = " Mitarbeiter ohne TimeSheet: "
                                    If awinSettings.englishLanguage Then
                                        logArray(0) = "Employee wothout TimeSheet: "
                                    End If
                                    logArray(1) = ""
                                    logArray(2) = roleName
                                    logArray(4) = ""

                                    Call logger(ptErrLevel.logWarning, "PTImportIstDaten", logArray)

                                    ' 
                                    ' im Output anzeigen ... 
                                    logmessage = logArray(0) & roleName
                                    outPutCollection.Add(logmessage)
                                Next
                            End If

                            ' Ende check : haben alle internen Mitarbeiter ein TimeSheet abgeliefert ... 

                            ' wenn auch externe Rollen Istdaten haben
                            ' welche externen Rollen haben keine Istdaten .. ? 


                            For Each kvp As KeyValuePair(Of String, clsConstellationItem) In referenzPortfolio.Liste
                                Dim tmpPName As String = getPnameFromKey(kvp.Key)
                                If Not validProjectNames.ContainsKey(tmpPName) Then
                                    ' jetzt muss dieses Projekt Null-Istdaten bekommen - wenn es von früher bereits ActualData hat, dann behält es die
                                    ' es werden nur die Monate actualDatuntil+1 .. IstDateDate genullt 
                                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(tmpPName, "", cacheProjekte, Date.Now)
                                    ReDim logArray(5)

                                    If hproj.setNewActualValuesToNull(IstdatenDate, True) Then
                                        Dim jjjj As Integer = Year(IstdatenDate)
                                        Dim mm As Integer = Month(IstdatenDate)
                                        Dim tt As Integer = Day(DateSerial(jjjj, mm + 1, 0)) 'tt ist letzte Tag des Monats mm 

                                        hproj.actualDataUntil = DateSerial(jjjj, mm, tt)

                                        ' jetzt in die Import-Projekte eintragen 
                                        updatedProjects = updatedProjects + 1
                                        ImportProjekte.Add(hproj, updateCurrentConstellation:=False)

                                        ' ins Protokoll eintragen 
                                        logArray(0) = " Projekt ohne Zeiterfassung: Ist-Daten auf Null gesetzt ! "
                                        If awinSettings.englishLanguage Then
                                            logArray(0) = " Project without time sheet records: actual data set to Zero ! "
                                        End If
                                        logArray(1) = ""
                                        logArray(2) = hproj.name
                                        logArray(3) = ""
                                        logArray(4) = ""

                                        Call logger(ptErrLevel.logWarning, "PTImportIstDaten", logArray)

                                    Else
                                        ' Fehler ins Protokoll eintragen 
                                        logArray(0) = " ohne Zeiterfassung: Plan-Werte konnten nicht auf Null gesetzt werden. "
                                        If awinSettings.englishLanguage Then
                                            logArray(0) = " Project without time sheet records: Error : could not set data to Zero ! "
                                        End If
                                        logArray(1) = "Error !"
                                        logArray(2) = hproj.name
                                        logArray(3) = ""
                                        logArray(4) = ""

                                        Call logger(ptErrLevel.logError, "PTImportIstDaten", logArray)
                                    End If

                                    ' im Output anzeigen ... 
                                    logmessage = logArray(0) & hproj.name
                                    outPutCollection.Add(logmessage)

                                End If
                            Next

                            ' jetzt überprüfen, welche Projekte zwar Istdaten bekommen haben, aber nicht im Referenz-Portfolio aufgeführt sind ... 
                            For Each vPKvP As KeyValuePair(Of String, SortedList(Of String, Double())) In validProjectNames

                                ReDim logArray(5)
                                If Not referenzPortfolio.containsProject(vPKvP.Key) Then
                                    ' ins Protokoll eintragen 
                                    logArray(0) = " Projekt hat Ist-Daten, ist aber nicht im Referenz-Portfolio enthalten ... ! "
                                    If awinSettings.englishLanguage Then
                                        logArray(0) = " Project has time sheet records, but is not referenced in active portfolio ... !"
                                    End If
                                    logArray(1) = ""
                                    logArray(2) = vPKvP.Key
                                    logArray(3) = ""
                                    logArray(4) = ""

                                    Call logger(ptErrLevel.logWarning, "PTImportIstDaten", logArray)

                                    ' im Output anzeigen ... 
                                    logmessage = logArray(0) & vPKvP.Key
                                    outPutCollection.Add(logmessage)

                                End If
                            Next

                            ' hier sollte noch ergänzt werdne
                            ' PRotokollieren welche Orga-Units denn ersetzt werden 
                            For Each substituteUnit As String In referatsCollection
                                ReDim logArray(5)
                                ' ins Protokoll eintragen 
                                logArray(0) = " Planwerte für Organisations-Unit werden ersetzt durch Istdaten: "
                                If awinSettings.englishLanguage Then
                                    logArray(0) = " Plan values for organizational unit are being replaced by Actual Data: "
                                End If
                                logArray(1) = ""
                                logArray(2) = substituteUnit
                                logArray(3) = ""
                                logArray(4) = ""

                                Call logger(ptErrLevel.logInfo, "PTImportIstDaten", logArray)

                                ' im Output anzeigen ... 
                                logmessage = logArray(0) & substituteUnit
                                outPutCollection.Add(logmessage)

                            Next


                            'Protokoll schreiben...
                            ' 
                            For Each vPKvP As KeyValuePair(Of String, SortedList(Of String, Double())) In validProjectNames

                                Dim protocolLine As String = ""
                                For Each rVKvP As KeyValuePair(Of String, Double()) In vPKvP.Value

                                    ' jetzt schreiben 
                                    Dim teamID As Integer = -1
                                    Dim hilfsrole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(rVKvP.Key, teamID)
                                    Dim curTagessatz As Double = hrole.tagessatzIntern

                                    ReDim logArray(3)
                                    logArray(0) = "Importiert wurde: "
                                    If awinSettings.englishLanguage Then
                                        logArray(0) = "Imported: "
                                    End If
                                    logArray(1) = ""
                                    logArray(2) = vPKvP.Key
                                    logArray(3) = rVKvP.Key & ": " & hilfsrole.name


                                    ReDim logDblArray(rVKvP.Value.Length - 1)
                                    For j As Integer = 0 To rVKvP.Value.Length - 1
                                        ' umrechnen, damit es mit dem Input File wieder vergleichbar wird 
                                        logDblArray(j) = rVKvP.Value(j) ' * curTagessatz
                                    Next

                                    Call logger(ptErrLevel.logWarning, "PTImportIstDaten", logArray, logDblArray)
                                Next

                            Next
                            ' Protokoll schreiben Ende ... 



                            Dim gesamtIstValue As Double = 0.0

                            For Each vPKvP As KeyValuePair(Of String, SortedList(Of String, Double())) In validProjectNames

                                Dim hproj As clsProjekt = getProjektFromSessionOrDB(vPKvP.Key, "", cacheProjekte, Date.Now)
                                Dim oldPlanValue As Double = 0.0
                                Dim newIstValue As Double = 0.0

                                lastValidMonth = getColumnOfDate(IstdatenDate)

                                If Not IsNothing(hproj) Then
                                    ' es wird pro Projekt eine Variante erzeugt 

                                    ' wenn es noch nicht beauftragt ist ... dann beauftragen 
                                    If hproj.Status = ProjektStatus(PTProjektStati.geplant) Then
                                        Try
                                            hproj.Status = ProjektStatus(PTProjektStati.beauftragt)
                                        Catch ex As Exception

                                        End Try

                                    End If
                                    Dim istDatenVName As String = ptVariantFixNames.acd.ToString
                                    Dim newProj As clsProjekt = hproj.createVariant(istDatenVName, "temporär für Ist-Daten-Aufnahme")

                                    ' es werden in jeder Phase, die einen der actual Monate enthält, die Werte gelöscht ... 
                                    ' gleichzeitig werden die bisherigen Soll-Werte dieser Zeit in T€ gemerkt ...
                                    ' True: die Werte werden auf Null gesetzt 
                                    Dim gesamtvorher As Double = newProj.getGesamtKostenBedarf().Sum * 1000

                                    'oldPlanValue = newProj.getSetRoleCostUntil(referatsCollection, monat, True)
                                    oldPlanValue = newProj.getSetRoleCostUntil(referatsCollection, lastValidMonth - newProj.Start + 1, True)
                                    'Dim checkOldPlanValue As Double = newProj.getSetRoleCostUntil(referatsCollection, monat, False)

                                    newIstValue = calcIstValueOf(vPKvP.Value)

                                    gesamtIstValue = gesamtIstValue + newIstValue

                                    ' die Werte der neuen Rollen in PT werden in der RootPhase eingetragen 
                                    Call newProj.mergeActualValues(rootPhaseName, vPKvP.Value)


                                    Dim gesamtNachher As Double = newProj.getGesamtKostenBedarf().Sum * 1000
                                    Dim checkNachher As Double = gesamtvorher - oldPlanValue + newIstValue
                                    ' Test tk 
                                    'Dim checkIstValue As Double = newProj.getSetRoleCostUntil(referatsCollection, monat, False)
                                    Dim checkIstValue As Double = newProj.getSetRoleCostUntil(referatsCollection, lastValidMonth - newProj.Start + 1, False)

                                    Dim abweichungGesamt As Double = 0.0
                                    If gesamtNachher <> checkNachher Then
                                        abweichungGesamt = System.Math.Abs(gesamtNachher - checkNachher)
                                    End If

                                    Dim abweichungIst As Double = 0.0
                                    If checkIstValue <> newIstValue Then
                                        abweichungIst = System.Math.Abs(checkIstValue - newIstValue)
                                    End If

                                    ' für Test 
                                    'awinSettings.visboDebug = True
                                    If awinSettings.visboDebug Then
                                        If abweichungGesamt > 0.05 Or abweichungIst > 0.05 Then
                                            ReDim logArray(3)
                                            logArray(0) = "Import Istdaten old/new/diff/check1/check2"
                                            If awinSettings.englishLanguage Then
                                                logArray(0) = "Import Actual Data old/new/diff/check1/check2"
                                            End If
                                            logArray(1) = ""
                                            logArray(2) = vPKvP.Key
                                            logArray(3) = ""

                                            ReDim logDblArray(4)
                                            logDblArray(0) = oldPlanValue
                                            logDblArray(1) = newIstValue
                                            logDblArray(2) = oldPlanValue - newIstValue
                                            logDblArray(3) = checkIstValue
                                            logDblArray(4) = gesamtNachher - checkNachher

                                            Call logger(ptErrLevel.logWarning, "PTImportIstDaten", logArray, logDblArray)

                                        End If
                                    End If



                                    Dim jjjj As Integer = Year(IstdatenDate)
                                    Dim mm As Integer = Month(IstdatenDate)
                                    Dim tt As Integer = Day(DateSerial(jjjj, mm + 1, 0)) 'tt ist letzte Tag des Monats mm 

                                    With newProj
                                        newProj.actualDataUntil = DateSerial(jjjj, mm, tt)
                                        .variantName = ""   ' eliminieren von VariantenName acd
                                        .variantDescription = ""
                                    End With

                                    ' jetzt in die Import-Projekte eintragen 
                                    updatedProjects = updatedProjects + 1
                                    ImportProjekte.Add(newProj, updateCurrentConstellation:=False)

                                Else
                                    ReDim logArray(5)
                                    logArray(0) = "Projekt existiert nicht !!?? ... darf nicht sein ..."
                                    logArray(1) = ""
                                    logArray(2) = vPKvP.Key
                                    logArray(3) = ""
                                    logArray(4) = ""

                                    Call logger(ptErrLevel.logError, "PTImportIstDaten", logArray)
                                End If

                            Next

                            ' tk Test 
                            If awinSettings.visboDebug Then
                                ReDim logArray(3)
                                logArray(0) = "Import von insgesamt " & updatedProjects & " Projekten (Gesamt-Euro): "
                                If awinSettings.englishLanguage Then
                                    logArray(0) = "Import of total " & updatedProjects & " Projects (Sum in Euro): "
                                End If
                                logArray(1) = ""
                                logArray(2) = ""
                                logArray(3) = ""

                                ReDim logDblArray(0)
                                logDblArray(0) = gesamtIstValue
                                Call logger(ptErrLevel.logWarning, "PTImportIstDaten", logArray, logDblArray)
                            End If


                            logmessage = vbLf & "Projekte aktualisiert: " & updatedProjects
                            If awinSettings.englishLanguage Then
                                logmessage = vbLf & "Projects updated: " & updatedProjects
                            End If
                            outPutCollection.Add(logmessage)

                            logmessage = vbLf & "detailllierte Protokollierung LogFile ./logfiles/logfile*.txt"
                            If awinSettings.englishLanguage Then
                                logmessage = vbLf & "Details see LogFile ./logfiles/logfile*.txt"
                            End If
                            outPutCollection.Add(logmessage)

                            If outPutCollection.Count > 0 Then
                                If awinSettings.englishLanguage Then
                                    Call showOutPut(outPutCollection, "Import Actual Data", "please check the notifications ...")
                                Else
                                    Call showOutPut(outPutCollection, "Import Istdaten", "folgende Probleme sind aufgetaucht")
                                End If

                            End If



                            '' Cursor auf Default setzen
                            Cursor.Current = Cursors.Default


                            ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                            Try
                                Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False,
                                                 getSomeValuesFromOldProj:=False, calledFromActualDataImport:=True)


                                ' ImportDatei ins archive-Directory schieben

                                If listOfArchivFilesAllg.Count > 0 Then
                                    Call moveFilesInArchiv(listOfArchivFilesAllg, importOrdnerNames(PTImpExp.actualData))
                                End If

                            Catch ex As Exception
                                If awinSettings.englishLanguage Then
                                    Call MsgBox("Error at Import: " & vbLf & ex.Message)
                                Else
                                    Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                                End If

                            End Try

                        Else

                            logmessage = vbLf & "detailllierte Protokollierung LogFile ./logfiles/logfile*.txt"
                            outPutCollection.Add(logmessage)

                            If outPutCollection.Count > 0 Then
                                If awinSettings.englishLanguage Then
                                    Call showOutPut(outPutCollection, "no Import because of errors", "please check the notifications ...")
                                Else
                                    Call showOutPut(outPutCollection, "Kein Import wegen Fehler", "folgende Probleme sind aufgetaucht")
                                End If

                            End If
                        End If

                    Else
                        ' nichts weiter tun ... auch kein Logfile schreiben  Logfile schreiben 

                    End If



                Else
                    If awinSettings.englishLanguage Then
                        outPutline = "No file to import actual data"
                    Else
                        outPutline = "Es gibt keine Datei zum Importieren von Istdaten"
                    End If

                    Call MsgBox(outPutline)

                    Call logger(ptErrLevel.logWarning, outPutline, "PTImportIstdaten", anzFehler)
                End If



            Else
                ' Fehlermeldung für Konfigurationsfile nicht vorhanden
                If awinSettings.englishLanguage Then
                    outPutline = "Error: either no configuration file found or worng definitions !"
                Else
                    outPutline = "Fehler: entweder fehlt die Konfigurations-Datei oder sie enthält fehlerhafte Definitionen!"
                End If
                Call logger(ptErrLevel.logError, outPutline, "PTImportIstdaten", anzFehler)

                Call MsgBox(outPutline)

            End If    ' allesOK

        End If


        ' so positionieren, dass die Projekte auch sichtbar sind ...
        If boardWasEmpty Then
            If ShowProjekte.Count > 0 Then
                Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
        End If


        ' Schließen des LogFiles
        ''Call logfileSchliessen()

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    ''' <summary>
    ''' importiert und speichert die CustomUserRoles
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub PTImportCustomUserRoles(control As IRibbonControl)

        Dim selectedWB As String = ""
        'Dim dirname As String = My.Computer.FileSystem.CombinePath(awinPath, requirementsOrdner)

        Dim dirname As String = importOrdnerNames(PTImpExp.customUserRoles)

        Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname, FileIO.SearchOption.SearchTopLevelOnly, "*roles*.xls*")
        Dim anzFiles As Integer = listOfImportfiles.Count

        Dim dateiname As String = ""

        ' tk by Ute für das Verschieben de rDatei nin den Archiv-Ordner wenn erfolgreich 
        Dim listOfArchivFiles As New List(Of String)

        Dim weiterMachen As Boolean = False

        'Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' öffnen des LogFiles
        'Call logfileOpen()


        If anzFiles = 1 Then
            selectedWB = listOfImportfiles.Item(0)
            weiterMachen = True

        ElseIf anzFiles > 1 Then
            Dim getUserRoleFile As New frmSelectImportFiles
            getUserRoleFile.menueAswhl = PTImpExp.customUserRoles
            Dim returnValue As DialogResult = getUserRoleFile.ShowDialog

            If returnValue = DialogResult.OK Then
                selectedWB = CStr(getUserRoleFile.selImportFiles.Item(1))
                weiterMachen = True
            End If
        Else
            Call MsgBox("keine Dateien vorhanden ..." & vbLf & "Folder: " & dirname & vbLf & "Datei muss diesen Teilstring enthalten: '" & "roles'")
        End If

        If weiterMachen Then

            dateiname = My.Computer.FileSystem.CombinePath(dirname, selectedWB)

            Try
                ' hier wird jetzt der Import gemacht 
                Call logger(ptErrLevel.logInfo, "Beginn Import Custom User Roles", selectedWB, -1)

                ' Öffnen des Organisations-Files
                appInstance.Workbooks.Open(dateiname)

                Dim outputCollection As New Collection
                Dim importedRoles As clsCustomUserRoles = ImportCustomUserRoles(outputCollection)

                Dim wbName As String = My.Computer.FileSystem.GetName(dateiname)

                ' Schliessen des CustomUser Role-Files
                appInstance.Workbooks(wbName).Close(SaveChanges:=True)

                If outputCollection.Count > 0 Then
                    Dim errmsg As String = vbLf & " .. Abbruch .. nicht importiert "
                    outputCollection.Add(errmsg)
                    Call showOutPut(outputCollection, "User Role Import", "")

                    Call logger(ptErrLevel.logError, "PTImportCustomUserRoles: ", outputCollection)

                ElseIf importedRoles.count > 0 Then
                    ' jetzt wird die Orga als Setting weggespeichert ... 
                    ' alles ok 
                    Dim err As New clsErrorCodeMsg
                    Dim result As Boolean = False
                    result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(importedRoles,
                                                                                    CStr(settingTypes(ptSettingTypes.customroles)),
                                                                                    CStr(settingTypes(ptSettingTypes.customroles)),
                                                                                    Nothing,
                                                                                    err)

                    If result = True Then
                        Call MsgBox("ok, Custom User Roles stored ...")
                        Call logger(ptErrLevel.logInfo, "Custom User Roles stored ...", selectedWB, -1)
                    Else
                        Call MsgBox("Error when writing Custom User Roles")
                        Call logger(ptErrLevel.logError, "Error when writing Custom User Roles ...", selectedWB, -1)
                    End If

                    listOfArchivFiles.Add(dateiname)

                Else
                    Call MsgBox("no roles found ...")
                End If
            Catch ex As Exception

            End Try
        End If




        ' Schließen des LogFiles
        ''Call logfileSchliessen()

        If listOfArchivFiles.Count > 0 Then
            Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.customUserRoles))
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    ''' <summary>
    ''' importiert und speichert die Kapazitäten 
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub PTImportKapas(control As IRibbonControl)

        Dim actualDataFile As String = ""
        Dim actualDataConfig As New SortedList(Of String, clsConfigActualDataImport)
        Dim outPutline As String = ""
        Dim lastrow As Integer = 0
        Dim listofArchivUrlaub As New List(Of String)
        Dim listofArchivConfig As New List(Of String)

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' öffnen des LogFiles
        'Call logfileOpen()

        Dim outputCollection As New Collection

        Dim changedOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)

        If Not IsNothing(changedOrga) Then

            If changedOrga.allRoles.Count > 0 Then

                RoleDefinitions = changedOrga.allRoles
                CostDefinitions = changedOrga.allCosts

                ' Liste enthält die Datei-Namen der erfolgreich eingelesenen externen Kapazitäts-Files 
                Dim listOfArchivExtern As New List(Of String)
                ' wenn es gibt - lesen der Modifier Kapas, wo interne wie externe angegeben sein können ..
                Call readMonthlyModifierKapas(outputCollection, listOfArchivExtern)

                ' wenn es gibt - lesen der Externen Verträge 
                Call readMonthlyExternKapasEV(outputCollection, listOfArchivExtern)

                '' wenn es gibt - lesen der Urlaubslisten DateiName "Urlaubsplaner*.xlsx
                listofArchivUrlaub = readInterneAnwesenheitslisten(outputCollection)

                ''  check Config-File - zum Einlesen der Istdaten gemäß Konfiguration -
                ''  - hier benötigt um den Kalender von IstDaten und Urlaubsdaten aufeinander abzustimmen
                Dim configActualDataImport As String = awinPath & configfilesOrdner & "configActualDataImport.xlsx"
                Dim allesOK As Boolean = checkActualDataImportConfig(configActualDataImport, actualDataFile, actualDataConfig, lastrow, outputCollection)

                ' wenn es gibt - lesen der Zeuss- listen und anderer, die durch configCapaImport beschrieben sind
                Dim configCapaImport As String = awinPath & configfilesOrdner & "configCapaImport.xlsx"
                If My.Computer.FileSystem.FileExists(configCapaImport) Then

                    listofArchivConfig = readInterneAnwesenheitslistenAllg(configCapaImport, actualDataConfig, outputCollection)
                Else
                    outPutline = "There is no Config-File for the capacities!"
                    Call logger(ptErrLevel.logWarning, outPutline, "PTImportKapas", anzFehler)
                End If

                If listofArchivUrlaub.Count > 0 Or listofArchivConfig.Count > 0 Or listOfArchivExtern.Count > 0 Then

                    changedOrga.allRoles = RoleDefinitions

                    If outputCollection.Count = 0 Then
                        ' keine Fehler aufgetreten ... 
                        ' jetzt wird die Orga als Setting weggespeichert ... 
                        Dim err As New clsErrorCodeMsg
                        Dim result As Boolean = False
                        ' ute -> überprüfen bzw. fertigstellen ... 
                        Dim orgaName As String = ptSettingTypes.organisation.ToString

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Then

                            result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(changedOrga,
                                                                                CStr(settingTypes(ptSettingTypes.organisation)),
                                                                                orgaName,
                                                                                changedOrga.validFrom,
                                                                                err)

                            If result = True Then
                                Call MsgBox("ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " updated ...")
                                Call logger(ptErrLevel.logInfo, "ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " updated ...", "", -1)

                                ' verschieben der Kapa-Dateien Kapazität* Modifier  in den ArchivOrdner
                                Call moveFilesInArchiv(listOfArchivExtern, importOrdnerNames(PTImpExp.Kapas))
                                ' verschieben der Kapa-Dateien Urlaubsplaner*.xlsx in den ArchivOrdner
                                Call moveFilesInArchiv(listofArchivUrlaub, importOrdnerNames(PTImpExp.Kapas))
                                ' verschieben der Kapa-Dateien,die durch configCapaImport.xlsx beschrieben sind, in den ArchivOrdner
                                Call moveFilesInArchiv(listofArchivConfig, importOrdnerNames(PTImpExp.Kapas))

                            Else
                                Call MsgBox("Error when writing Organisation to Database")
                                Call logger(ptErrLevel.logError, "Error when writing Organisation to Database...", "", -1)
                            End If

                        Else
                            Call MsgBox("ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " temporarily updated ...")
                            Call logger(ptErrLevel.logInfo, "ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " temporarily updated ...", "", -1)
                            ' verschieben der Kapa-Dateien Urlaubsplaner*.xlsx in den ArchivOrdner
                            'Call moveFilesInArchiv(listofArchivUrlaub, importOrdnerNames(PTImpExp.Kapas))
                            '' verschieben der Kapa-Dateien,die durch configCapaImport.xlsx beschrieben sind, in den ArchivOrdner
                            'Call moveFilesInArchiv(listofArchivAllg, importOrdnerNames(PTImpExp.Kapas))
                        End If

                    Else

                        Call showOutPut(outputCollection, "Importing Capacities", "... mit Fehlern abgebrochen ...")
                        Call logger(ptErrLevel.logError, "PTImportKapas: ", outputCollection)

                    End If
                Else
                    If outputCollection.Count > 0 Then

                        Call showOutPut(outputCollection, "Importing Capacities", "... mit Fehlern abgebrochen ...")
                        Call logger(ptErrLevel.logError, "PTImportKapas: ", outputCollection)
                    Else

                        If awinSettings.englishLanguage Then
                            Call MsgBox("no Files to import ...")
                        Else
                            Call MsgBox("es gab keine Dateien zum Einlesen ... ")

                        End If
                    End If

                End If

            Else
                If awinSettings.englishLanguage Then
                    Call MsgBox("No valid roles! Please import one first!")
                Else
                    Call MsgBox("Die gültige Organisation beinhaltet keine Rollen! ")

                End If
            End If

        Else

            If awinSettings.englishLanguage Then
                Call MsgBox("No valid organization! Please import one first!")
            Else
                Call MsgBox("Es existiert keine gültige Organisation! Bitte zuerst Organisation importieren")
            End If


            Dim errMsg As String = "Kapazitäten wurden nicht aktualisiert - bitte erst die Import-Dateien korrigieren ... "
            outputCollection.Add(errMsg)
            Call showOutPut(outputCollection, "Importing Capacities", "")
            Call logger(ptErrLevel.logError, "PTImportKapas: ", outputCollection)

        End If

        ' Schließen des LogFiles
        ''Call logfileSchliessen()

        enableOnUpdate = True
        appInstance.EnableEvents = True

        With CType(CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook).Worksheets(arrWsNames(ptTables.MPT)), Excel.Worksheet)
            .Activate()
        End With
        appInstance.ScreenUpdating = True

    End Sub
    ''' <summary>
    ''' importiert und speichert die Custom-Einstellungen
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub PTImportCustomization(control As IRibbonControl)

        Dim selectedWB As String = ""
        Dim dirname As String = My.Computer.FileSystem.CombinePath(awinPath, requirementsOrdner)


        Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname, FileIO.SearchOption.SearchTopLevelOnly, "Project Board Customization*.xls*")
        Dim anzFiles As Integer = listOfImportfiles.Count

        Dim dateiname As String = ""

        Dim weiterMachen As Boolean = False


        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' öffnen des LogFiles
        'Call logfileOpen()


        If anzFiles = 1 Then
            selectedWB = listOfImportfiles.Item(0)
            weiterMachen = True

        ElseIf anzFiles > 1 Then
            Dim getCustomizationFile As New frmSelectImportFiles
            getCustomizationFile.menueAswhl = PTImpExp.customization
            Dim returnValue As DialogResult = getCustomizationFile.ShowDialog

            If returnValue = DialogResult.OK Then
                selectedWB = CStr(getCustomizationFile.selImportFiles.Item(1))
                weiterMachen = True
            End If
        Else
            Call MsgBox("keine Dateien vorhanden ..." & vbLf & "Folder: " & dirname & vbLf & "Datei muss diesen Teilstring enthalten: '" & "Customization'")
        End If

        If weiterMachen Then

            dateiname = My.Computer.FileSystem.CombinePath(dirname, selectedWB)

            Try
                ' hier wird jetzt der Import gemacht 
                Call logger(ptErrLevel.logInfo, "Beginn Import kundenspezifischer Einstellungen", selectedWB, -1)

                ' Öffnen des Customization-Files
                appInstance.Workbooks.Open(dateiname)

                Dim outputCollection As New Collection
                Dim importedCustomization As clsCustomization = ImportCustomization(outputCollection)

                ' vorher zurücksetzen ...
                If customFieldDefinitions.count > 0 Then
                    customFieldDefinitions = New clsCustomFieldDefinitions
                End If
                Dim customFieldDefs As clsCustomFieldDefinitions = ImportCustomFieldDefinitions(outputCollection)

                Dim wbName As String = My.Computer.FileSystem.GetName(dateiname)

                ' Schliessen des Customizations-Files
                appInstance.Workbooks(wbName).Close(SaveChanges:=True)

                If outputCollection.Count > 0 Then
                    Dim errmsg As String = vbLf & " .. Abbruch .. nicht importiert "
                    outputCollection.Add(errmsg)
                    Call showOutPut(outputCollection, "Einstellungen Import", "")

                    Call logger(ptErrLevel.logError, "PTImportCustomization: ", outputCollection)

                ElseIf Not IsNothing(importedCustomization) Then
                    ' jetzt werden die Einstellungen als Setting weggespeichert ... 
                    ' alles ok 
                    Dim err As New clsErrorCodeMsg
                    Dim ts As Date = CDate("1.1.1900")
                    Dim result As Boolean = False
                    Dim result1 As Boolean = False

                    result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(importedCustomization,
                                                                                    CStr(settingTypes(ptSettingTypes.customization)),
                                                                                    CStr(settingTypes(ptSettingTypes.customization)),
                                                                                    ts,
                                                                                    err)


                    If Not IsNothing(customFieldDefs) Then
                        ' jetzt werden die Einstellungen als Setting weggespeichert ... 
                        ' alles ok 
                        Dim err1 As New clsErrorCodeMsg


                        result1 = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(customFieldDefs,
                                                                                        CStr(settingTypes(ptSettingTypes.customfields)),
                                                                                        CStr(settingTypes(ptSettingTypes.customfields)),
                                                                                        ts,
                                                                                        err1)
                    End If

                    If result And result1 Then
                        Call MsgBox("ok, Customizations and CustomFieldDefinitions stored ...")
                        Call logger(ptErrLevel.logInfo, "Customizations and CustomFieldDefinitions stored ...", selectedWB, -1)
                    Else
                        Call MsgBox("Error when writing Customizations or CustomfieldDefinitions")
                        Call logger(ptErrLevel.logError, "Error when writing Customizations or Customfielddefinitions ...", selectedWB, -1)
                    End If


                Else
                    Call MsgBox("no customizations found ...")
                End If




            Catch ex As Exception
                Dim resultMessage As String = ex.Message
                Call MsgBox(resultMessage)
                Call logger(ptErrLevel.logError, "Error when writing Customizations ...", resultMessage, -1)
            End Try
        End If


        ' Schließen des LogFiles
        ''Call logfileSchliessen()

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub
    ''' <summary>
    ''' importiert und speichert die Darstellungsklassen
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub PTImportAppearances(control As IRibbonControl)

        Dim selectedWB As String = ""
        Dim dirname As String = My.Computer.FileSystem.CombinePath(awinPath, requirementsOrdner)

        'Dim dirname As String = importOrdnerNames(PTImpExp.customization)

        Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname, FileIO.SearchOption.SearchTopLevelOnly, "Project Board Customization*.xls*")
        Dim anzFiles As Integer = listOfImportfiles.Count

        Dim dateiname As String = ""

        Dim weiterMachen As Boolean = False

        'Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' öffnen des LogFiles
        'Call logfileOpen()


        If anzFiles = 1 Then
            selectedWB = listOfImportfiles.Item(0)
            weiterMachen = True

        ElseIf anzFiles > 1 Then
            Dim getAppearanceFile As New frmSelectImportFiles
            getAppearanceFile.menueAswhl = PTImpExp.customization
            Dim returnValue As DialogResult = getAppearanceFile.ShowDialog

            If returnValue = DialogResult.OK Then
                selectedWB = CStr(getAppearanceFile.selImportFiles.Item(1))
                weiterMachen = True
            End If
        Else
            Call MsgBox("keine Dateien vorhanden ..." & vbLf & "Folder: " & dirname & vbLf & "Datei muss diesen Teilstring enthalten: '" & "Customization'")
        End If

        If weiterMachen Then

            dateiname = My.Computer.FileSystem.CombinePath(dirname, selectedWB)

            Try
                ' hier wird jetzt der Import gemacht 
                Call logger(ptErrLevel.logInfo, "Beginn Import Appearances", selectedWB, -1)

                ' Öffnen des Customization-Files
                appInstance.Workbooks.Open(dateiname)

                Dim outputCollection As New Collection
                '' ??? sollen die appearances von grund auf aufgebaut werden, oder nur verändert?
                Dim importedAppearances As SortedList(Of String, clsAppearance) = ImportAppearances(outputCollection)

                Dim wbName As String = My.Computer.FileSystem.GetName(dateiname)

                ' Schliessen des CustomUser Role-Files
                appInstance.Workbooks(wbName).Close(SaveChanges:=True)

                If outputCollection.Count > 0 Then
                    Dim errmsg As String = vbLf & " .. Abbruch .. nicht importiert "
                    outputCollection.Add(errmsg)
                    Call showOutPut(outputCollection, "Appearances Import", "")

                    Call logger(ptErrLevel.logError, "PTImportAppearances: ", outputCollection)

                ElseIf Not IsNothing(importedAppearances) And importedAppearances.Count > 0 Then
                    ' jetzt wird die Appearances als Setting weggespeichert ... 
                    ' alles ok 
                    Dim err As New clsErrorCodeMsg
                    Dim result As Boolean = False
                    result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(importedAppearances,
                                                                                    CStr(settingTypes(ptSettingTypes.appearance)),
                                                                                    CStr(settingTypes(ptSettingTypes.appearance)),
                                                                                    Nothing,
                                                                                    err)

                    If result = True Then
                        Call MsgBox("ok, appearances stored ...")
                        Call logger(ptErrLevel.logInfo, "appearances stored ...", selectedWB, -1)
                    Else
                        Call MsgBox("Error when writing appearances")
                        Call logger(ptErrLevel.logError, "Error when writing appearances ...", selectedWB, -1)
                    End If
                Else
                    Call MsgBox("no appearances found ...")
                End If
            Catch ex As Exception

            End Try
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub
    ''' <summary>
    ''' Importiert werden die auf Platte verfügbaren Projekt-Templates
    ''' diese werden in DB als VP mit Projekt-Typ= 2 (Vorlage) gespeichert
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub PTImportProjectTemplates(control As IRibbonControl)

        Dim i As Integer = 0
        Dim outputCollection As New Collection
        Dim isIdentical As Boolean = False
        Dim msgStr As String = ""

        ' hier wird jetzt der Import gemacht 
        Call logger(ptErrLevel.logInfo, "Beginn Import Templates", "PTImportProjectTemplates", -1)

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        ' jetzt werden die Projekt-Vorlagen ausgelesen 
        Call readVorlagen(False)
        Dim anzahlTemplates As Integer = Projektvorlagen.Count

        Call logger(ptErrLevel.logInfo, anzahlTemplates & " Templates erfolgreich eingelesen", "PTImportProjectTemplates", -1)

        For Each kvp As KeyValuePair(Of String, clsProjektvorlage) In Projektvorlagen.Liste
            Dim vproj As clsProjektvorlage = kvp.Value
            Dim template As New clsProjekt
            ' mache aus clsprojektVorlage ein 'clsProjekt'
            Dim startDate As Date = StartofCalendar
            Dim endDate As Date = startDate.AddDays(vproj.dauerInDays - 1)
            Dim myProject As clsProjekt = Nothing
            template = erstelleProjektAusVorlage(myProject, vproj.VorlagenName, vproj.VorlagenName, startDate, endDate, vproj.Erloes, 0, 5.0, 5.0, "0", vproj.VorlagenName, "", "", True)

            ' ur: 28.2.2021: nicht mehr benötigt, da eine ganzes Projekt angelegt wird und im ReSt-Server als vorlage dient.
            ' vproj.copyTo(template)


            If Not IsNothing(template) Then
                template.name = vproj.VorlagenName
                template.projectType = ptPRPFType.projectTemplate
                Dim erfolgreich As Boolean = storeSingleProjectToDB(template, outputCollection, isIdentical)
                If Not erfolgreich Then
                    If awinSettings.englishLanguage Then
                        msgStr = "Error when writing Template: " & template.name
                    Else
                        msgStr = "Fehler beim Speichern der Vorlage: " & template.name
                    End If
                    outputCollection.Add(msgStr)
                    Call logger(ptErrLevel.logError, msgStr, "PTImportProjectTemplates", -1)
                Else
                    If awinSettings.englishLanguage Then
                        msgStr = "Template: " & template.name & " stored"
                    Else
                        msgStr = "Vorlage: " & template.name & " gespeichert"
                    End If
                    Call logger(ptErrLevel.logInfo, msgStr, "PTImportProjectTemplates", -1)
                End If

            Else
                If awinSettings.englishLanguage Then
                    msgStr = "Error when reading/writing Template: (conflict) " & vproj.VorlagenName
                Else
                    msgStr = "Fehler beim Lesen/Schreiben der Vorlage: (Konflikt) " & vproj.VorlagenName
                End If
                outputCollection.Add(msgStr)
            End If
        Next

        If outputCollection.Count > 0 Then
            Dim errmsg As String = vbLf & " .. Information ... to Import "
            outputCollection.Add(errmsg)
            Call showOutPut(outputCollection, "Projekt-Templates Import", "")

            Call logger(ptErrLevel.logError, "PTImportProjectTemplates: ", outputCollection)

        Else
            If awinSettings.englishLanguage Then
                msgStr = "All Templates stored successfully"
            Else
                msgStr = "Alle Vorlagen erfolgreich gespeichert"
            End If
            Call MsgBox(msgStr)
            Call logger(ptErrLevel.logInfo, msgStr, "PTImportProjectTemplates", -1)
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Public Sub Tom2G4M1Import(control As IRibbonControl)

        If Not noDB Then
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
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

        Dim listOfArchivFiles As New List(Of String)

        Dim getVisboImport As New frmSelectImportFiles
        Dim returnValue As DialogResult

        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

        ''Call logfileOpen()

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
                        Call logger(ptErrLevel.logInfo, "Beginn Import: " & dateiName, "Tom2G4M1Import", -1)

                    Catch ex1 As Exception
                        Call logger(ptErrLevel.logError, "Fehler bei Öffnen der Datei: " & dateiName, "Tom2G4M1Import", -1)
                        skip = True
                    End Try

                    If Not skip Then
                        pname = ""
                        hproj = New clsProjekt
                        Try
                            Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)

                            Try
                                Dim keyStr As String = calcProjektKey(hproj)
                                ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                                myCollection.Add(calcProjektKey(hproj))
                                listOfArchivFiles.Add(dateiName)

                            Catch ex2 As Exception
                                Call MsgBox("Projekt kann nicht zweimal importiert werden ...")
                            End Try

                            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                            ' liste der Dateien, die nach archive verschoben werden sollen


                        Catch ex1 As Exception
                            appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                            Call logger(ptErrLevel.logError, ex1.Message, "Tom2G4M1Import", anzFehler)
                            Call MsgBox(ex1.Message)
                            'Call MsgBox("Fehler bei Import von Projekt " & hproj.name & vbCrLf & "Siehe Logfile")
                        End Try



                    End If



                End If


            Next i


            Try
                Call importProjekteEintragen(importDate, True, False, True)
                'Call importProjekteEintragen(myCollection, importDate, ProjektStatus(1))

                ' ImportDatei ins archive-Directory schieben
                If listOfArchivFiles.Count > 0 Then
                    Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.visbo))
                End If

            Catch ex As Exception
                Call MsgBox("Fehler bei Import : " & vbLf & ex.Message)
            End Try

        Else

            'Call logfileSchreiben("Import wurde abgebrochen", "", -1)

        End If

        ' so positionieren, dass die Projekte auch sichtbar sind ...
        If boardWasEmpty Then
            If ShowProjekte.Count > 0 Then
                Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
        End If



        '' Call logfileSchliessen()

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


    End Sub

    Public Sub Tom2G4M2ImportMSProject(control As IRibbonControl)

        If Not noDB Then
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        End If
        Dim hproj As New clsProjekt
        Dim cproj As New clsProjekt
        Dim vglName As String = " "
        Dim outputString As String = ""
        Dim dateiName As String
        Dim getMSImport As New frmSelectImportFiles
        Dim returnValue As DialogResult

        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False



        getMSImport.menueAswhl = PTImpExp.msproject
        returnValue = getMSImport.ShowDialog

        If returnValue = DialogResult.OK Then


            Dim importDate As Date = Date.Now
            Dim listOfArchivFiles As New List(Of String)

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

            '' Cursor auf HourGlass setzen
            Cursor.Current = Cursors.WaitCursor

            ' jetzt müssen die Projekte ausgelesen werden, die in dateiListe stehen 
            Dim i As Integer

            Dim outPutCollection As New Collection
            Dim outputLine As String = ""

            For i = 1 To listofVorlagen.Count
                dateiName = listofVorlagen.Item(i).ToString

                hproj = New clsProjekt

                ' Definition für ein eventuelles Mapping
                Dim mapProj As clsProjekt = Nothing

                Try
                    Call awinImportMSProject("", dateiName, hproj, mapProj, importDate)

                    Try
                        Dim keyStr As String = calcProjektKey(hproj)
                        ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                        myCollection.Add(calcProjektKey(hproj))

                        If Not IsNothing(mapProj) Then
                            keyStr = calcProjektKey(mapProj)
                            ImportProjekte.Add(mapProj, updateCurrentConstellation:=False)
                            myCollection.Add(calcProjektKey(mapProj))

                        End If
                    Catch ex2 As Exception
                        Call MsgBox("Projekt kann nicht zweimal importiert werden ...")
                    End Try

                    ' ''appInstance.ActiveWorkbook.Close(SaveChanges:=False)

                Catch ex1 As Exception
                    ''appInstance.ActiveWorkbook.Close(SaveChanges:=False)
                    Call MsgBox(ex1.Message)
                    Call MsgBox("Fehler bei Import von Projekt " & hproj.name)
                End Try

                ' erfolgreich importiertes msproject-File in Liste zum Archivieren speichern
                listOfArchivFiles.Add(dateiName)

            Next i

            If missingRoleDefinitions.Count > 0 Or missingCostDefinitions.Count > 0 Then

                For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In missingRoleDefinitions.liste
                    If awinSettings.englishLanguage Then
                        outputLine = "unknown Role: " & kvp.Value.name
                    Else
                        outputLine = "unbekannte Rolle: " & kvp.Value.name
                    End If

                    outPutCollection.Add(outputLine)
                Next

                For Each kvp As KeyValuePair(Of Integer, clsKostenartDefinition) In missingCostDefinitions.liste
                    If awinSettings.englishLanguage Then
                        outputLine = "unknown Cost: " & kvp.Value.name
                    Else
                        outputLine = "unbekannte Kostenart: " & kvp.Value.name
                    End If

                    outPutCollection.Add(outputLine)
                Next

                'Call logfileOpen()
                Call logger(ptErrLevel.logError, "Tom2G42ImportMSProject: ", outPutCollection)
                ''Call logfileSchliessen()

                If awinSettings.englishLanguage Then
                    Call showOutPut(outPutCollection, "unknown Elements:", "please modify organisation-file or input ...")
                Else
                    Call showOutPut(outPutCollection, "Unbekannte Elemente:", "bitte in Organisations-Datei korrigieren")
                End If


            Else

            End If

            '' Cursor auf Default setzen
            Cursor.Current = Cursors.Default

            ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
            Try
                Call importProjekteEintragen(importDate, True, True, True)

                ' verschieben der erfolgreich importierten files
                If listOfArchivFiles.Count > 0 Then
                    Call moveFilesInArchiv(listOfArchivFiles, importOrdnerNames(PTImpExp.msproject))
                End If

            Catch ex As Exception
                If awinSettings.englishLanguage Then
                    Call MsgBox("Error at Import: " & vbLf & ex.Message)
                Else
                    Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                End If

            End Try



        End If

        ' so positionieren, dass die Projekte auch sichtbar sind ...
        If boardWasEmpty Then
            If ShowProjekte.Count > 0 Then
                Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
        End If



        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub
    Public Sub PTImportProjectsWithConfig(control As IRibbonControl)

        Dim projectConfig As New SortedList(Of String, clsConfigProjectsImport)
        Dim projectsFile As String = ""
        Dim lastrow As Integer = 0
        Dim outputString As String = ""
        Dim dateiName As String = ""
        Dim listofArchivAllg As New List(Of String)
        Dim outPutCollection As New Collection

        Dim outputLine As String = ""

        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

        ' Konfigurationsdatei lesen und Validierung durchführen

        ' wenn es gibt - lesen der Zeuss- listen und anderer, die durch configCapaImport beschrieben sind
        Dim configProjectsImport As String = awinPath & configfilesOrdner & "configProjectImport.xlsx"

        ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
        Dim allesOK As Boolean = checkProjectImportConfig(configProjectsImport, projectsFile, projectConfig, lastrow, outPutCollection)

        If allesOK Then

            Dim getProjConfigImport As New frmSelectImportFiles

            Dim returnValue As DialogResult


            Call projektTafelInit()

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False
            enableOnUpdate = False

            getProjConfigImport.menueAswhl = PTImpExp.projectWithConfig
            getProjConfigImport.importFileNames = projectsFile
            returnValue = getProjConfigImport.ShowDialog

            If returnValue = DialogResult.OK Then

                Dim ok As Boolean = False
                Dim importDate As Date = Date.Now
                'Dim importDate As Date = "31.10.2013"
                ''Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String)
                Dim listofVorlagen As Collection
                listofVorlagen = getProjConfigImport.selImportFiles


                '' ''dirName = awinPath & msprojectFilesOrdner
                ' ''dirName = importOrdnerNames(PTImpExp.msproject)
                ' ''listOfVorlagen = My.Computer.FileSystem.GetFiles(dirName, FileIO.SearchOption.SearchTopLevelOnly, "*.mpp")

                ' alle Import Projekte erstmal löschen
                ImportProjekte.Clear(False)

                '' Cursor auf HourGlass setzen
                Cursor.Current = Cursors.WaitCursor

                'Call logfileOpen()

                ' jetzt müssen die Projekte ausgelesen werden, die in dateiListe stehen 
                Dim i As Integer

                For i = 1 To listofVorlagen.Count
                    dateiName = listofVorlagen.Item(i).ToString


                    listofArchivAllg = readProjectsAllg(listofVorlagen, projectConfig, outPutCollection)

                    If listofArchivAllg.Count > 0 Then
                        Call moveFilesInArchiv(listofArchivAllg, importOrdnerNames(PTImpExp.projectWithConfig))
                    End If

                Next

                'Call logfileSchreiben(outPutCollection)
                ''Call logfileSchliessen()


                ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                Try
                    ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                    ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                    Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False, getSomeValuesFromOldProj:=False, calledFromActualDataImport:=False)

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Error at Import: " & vbLf & ex.Message)
                    Else
                        Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                    End If

                End Try

                '' Cursor auf Default setzen
                Cursor.Current = Cursors.Default

            End If

        End If


        outputString = vbLf & "detailllierte Protokollierung LogFile ./logfiles/logfile*.txt"
        outPutCollection.Add(outputString)

        If outPutCollection.Count > 0 Then
            If awinSettings.englishLanguage Then
                Call showOutPut(outPutCollection, "Import Projects", "please check the notifications ...")
            Else
                Call showOutPut(outPutCollection, "Einlesen Projekte", "folgende Probleme sind aufgetaucht")
            End If
        End If

        ' so positionieren, dass die Projekte auch sichtbar sind ...
        If boardWasEmpty Then
            If ShowProjekte.Count > 0 Then
                Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
        End If


        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

    End Sub

    Public Sub PTImportJIRAProjects(control As IRibbonControl)

        Dim JIRAProjectsConfig As New SortedList(Of String, clsConfigProjectsImport)
        Dim projectsFile As String = ""
        Dim lastrow As Integer = 0
        Dim outputString As String = ""
        Dim dateiName As String = ""
        Dim listofArchivAllg As New List(Of String)
        Dim outPutCollection As New Collection

        Dim outputLine As String = ""

        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

        ' Konfigurationsdatei lesen und Validierung durchführen

        ' wenn es gibt - lesen der Zeuss- listen und anderer, die durch configCapaImport beschrieben sind
        Dim configJIRAProjects As String = awinPath & configfilesOrdner & "configJIRAProjectImport.xlsx"

        ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
        Dim allesOK As Boolean = checkProjectImportConfig(configJIRAProjects, projectsFile, JIRAProjectsConfig, lastrow, outPutCollection)

        If allesOK Then

            Dim getProjConfigImport As New frmSelectImportFiles

            Dim returnValue As DialogResult


            Call projektTafelInit()

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False
            enableOnUpdate = False

            getProjConfigImport.menueAswhl = PTImpExp.JiraProjects
            getProjConfigImport.importFileNames = projectsFile
            returnValue = getProjConfigImport.ShowDialog

            If returnValue = DialogResult.OK Then

                Dim ok As Boolean = False
                Dim importDate As Date = Date.Now
                'Dim importDate As Date = "31.10.2013"
                ''Dim listOfVorlagen As Collections.ObjectModel.ReadOnlyCollection(Of String)
                Dim listofVorlagen As Collection
                listofVorlagen = getProjConfigImport.selImportFiles


                '' ''dirName = awinPath & msprojectFilesOrdner
                ' ''dirName = importOrdnerNames(PTImpExp.msproject)
                ' ''listOfVorlagen = My.Computer.FileSystem.GetFiles(dirName, FileIO.SearchOption.SearchTopLevelOnly, "*.mpp")

                ' alle Import Projekte erstmal löschen
                ImportProjekte.Clear(False)

                '' Cursor auf HourGlass setzen
                Cursor.Current = Cursors.WaitCursor

                'Call logfileOpen()

                ' jetzt müssen die Projekte ausgelesen werden, die in dateiListe stehen 
                Dim i As Integer

                For i = 1 To listofVorlagen.Count
                    dateiName = listofVorlagen.Item(i).ToString


                    listofArchivAllg = readProjectsJIRA(listofVorlagen, JIRAProjectsConfig, outPutCollection)

                Next
                Try
                    ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                    ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                    Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=True, getSomeValuesFromOldProj:=True, calledFromActualDataImport:=True)

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Error at Import: " & vbLf & ex.Message)
                    Else
                        Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                    End If

                End Try

                If listofArchivAllg.Count > 0 Then
                    Call moveFilesInArchiv(listofArchivAllg, importOrdnerNames(PTImpExp.JiraProjects))
                End If


                '' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                'Try
                '    ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                '    ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                '    Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False, getSomeValuesFromOldProj:=False, calledFromActualDataImport:=False)

                'Catch ex As Exception
                '    If awinSettings.englishLanguage Then
                '        Call MsgBox("Error at Import: " & vbLf & ex.Message)
                '    Else
                '        Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                '    End If

                'End Try

                '' Cursor auf Default setzen
                Cursor.Current = Cursors.Default

            End If

        End If


        If outPutCollection.Count > 0 Then

            outputString = vbLf & "detailllierte Protokollierung LogFile ./logfiles/logfile*.txt"
            outPutCollection.Add(outputString)

            If awinSettings.englishLanguage Then
                Call showOutPut(outPutCollection, "Import Projects", "please check the notifications ...")
            Else
                Call showOutPut(outPutCollection, "Einlesen Projekte", "folgende Probleme sind aufgetaucht")
            End If
        End If

        ' so positionieren, dass die Projekte auch sichtbar sind ...
        If boardWasEmpty Then
            If ShowProjekte.Count > 0 Then
                Dim leftborder As Integer = ShowProjekte.getMinMonthColumn
                If leftborder - 12 > 0 Then
                    appInstance.ActiveWindow.ScrollColumn = leftborder - 12
                Else
                    appInstance.ActiveWindow.ScrollColumn = 1
                End If
            End If
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
                        outputString = outputString & hproj.name & " nicht erfolgreich .." & vbLf &
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
                appInstance.ActiveWorkbook.Close(SaveChanges:=True, Filename:=exportOrdnerNames(PTImpExp.rplan) & "\" &
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

    ''' <summary>
    ''' schreibt die Prio Liste 
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub awinWritePrioList(control As IRibbonControl)

        Call projektTafelInit()

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Dim roleCostCollection As New Collection
        Dim costCollection As New Collection
        Try
            Call writeProjektsForSequencing(roleCostCollection, costCollection)
        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try


        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
    End Sub

    Public Sub ExportExcelPlanning(control As IRibbonControl)
        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()

        If ok Then

        End If
        Call projektTafelInit()

        Dim frmMERoleCost As New frmMEhryRoleCost
        With frmMERoleCost
            .hproj = Nothing
            .phaseName = ""
            .phaseNameID = rootPhaseName
            .pName = ""
            .vName = ""
            .rcName = ""
        End With

        Dim returnValue As DialogResult = frmMERoleCost.ShowDialog()

        If returnValue = DialogResult.OK Then

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False
            enableOnUpdate = False

            ' jetzt muss die myCollection aus den rolesToAdd und costsToAdd aufgebaut werden ; 
            ' erstmal werden nur die rolesToAdd berücksichtigt

            Dim roleCollection As New Collection

            For Each element As String In frmMERoleCost.rolesToAdd

                Dim teamID As Integer = -1
                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(element, teamID)
                If Not IsNothing(tmpRole) Then
                    roleCollection.Add(tmpRole.name)
                End If

            Next

            Dim costCollection As New Collection
            For Each element As String In frmMERoleCost.costsToAdd
                Dim tmpCost As clsKostenartDefinition = CostDefinitions.getCostdef(element)
                If Not IsNothing(tmpCost) Then
                    costCollection.Add(tmpCost.name)
                End If
            Next


            Try
                Call writeYearInitialPlanningSupportToExcel(showRangeLeft, showRangeRight, roleCollection, costCollection, PTEinheiten.hrs)
            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

            enableOnUpdate = True
            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True

        End If

    End Sub


    ''' <summary>
    ''' schreibt pro Projekt eine Zeile ...
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub exportExcelSumme(control As IRibbonControl)

        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()

        Call projektTafelInit()

        Dim frmMERoleCost As New frmMEhryRoleCost
        With frmMERoleCost
            .hproj = Nothing
            .phaseName = ""
            .phaseNameID = rootPhaseName
            .pName = ""
            .vName = ""
            .rcName = ""
        End With

        Dim returnValue As DialogResult = frmMERoleCost.ShowDialog()

        If returnValue = DialogResult.OK Then

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False
            enableOnUpdate = False

            ' jetzt muss die myCollection aus den rolesToAdd und costsToAdd aufgebaut werden ; 
            ' erstmal werden nur die rolesToAdd berücksichtigt

            Dim roleCollection As New Collection

            For Each element As String In frmMERoleCost.rolesToAdd

                Dim teamID As Integer = -1
                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(element, teamID)
                If Not IsNothing(tmpRole) Then
                    roleCollection.Add(tmpRole.name)
                End If

            Next

            Dim costCollection As New Collection
            For Each element As String In frmMERoleCost.costsToAdd
                Dim tmpCost As clsKostenartDefinition = CostDefinitions.getCostdef(element)
                If Not IsNothing(tmpCost) Then
                    costCollection.Add(tmpCost.name)
                End If
            Next


            Try
                Call writeProjektsForSequencing(roleCollection, costCollection)
            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

            enableOnUpdate = True
            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True

        End If

    End Sub
    Public Sub exportExcelKuGDetails(control As IRibbonControl)
        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()

        Call projektTafelInit()

        Try
            Call writeProjektKuGDetailsToExcel(showRangeLeft, showRangeRight)
        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
    End Sub
    ''' <summary>
    ''' schreibt pro Projekt alle ausgewählten Rollen / Kosten weg 
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub exportExcelDetails(control As IRibbonControl)

        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()

        Call projektTafelInit()

        Dim frmMERoleCost As New frmMEhryRoleCost
        With frmMERoleCost
            .hproj = Nothing
            .phaseName = ""
            .phaseNameID = rootPhaseName
            .pName = ""
            .vName = ""
            .rcName = ""
        End With

        Dim returnValue As DialogResult = frmMERoleCost.ShowDialog()

        If returnValue = DialogResult.OK Then

            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False
            enableOnUpdate = False


            Dim myCollectionR As New Collection
            Dim myCollectionC As New Collection

            ' erstmal werden hier nur die 
            For Each element As String In frmMERoleCost.rolesToAdd
                Dim teamID As Integer
                Dim roleUID As Integer = RoleDefinitions.parseRoleNameID(element, teamID)
                If roleUID > 0 Then
                    ' es ist eine Rolle 
                    If Not myCollectionR.Contains(element) Then
                        myCollectionR.Add(element, element)
                    End If
                ElseIf CostDefinitions.containsName(element) Then
                    If Not myCollectionC.Contains(element) Then
                        myCollectionC.Add(element, element)
                    End If
                End If

                'End If

            Next


            Try
                Call writeProjektDetailsToExcel(showRangeLeft, showRangeRight, myCollectionR, myCollectionC)
            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

            enableOnUpdate = True
            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True

        End If

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
                Call showOutPut(outPutCollection,
                                 "Exportieren Steckbriefe",
                                 "erfolgreich exportierte Dateien liegen in " & vbLf &
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

    ' tk 21.8.2017 wird nicht mehr verwendet 
    '''' <summary>
    '''' erstellt die Summary Zuordnungs-Datei 
    '''' </summary>
    '''' <param name="control"></param>
    '''' <remarks></remarks>
    'Sub Tom2G4M2B1ZuordnungRP(control As IRibbonControl)


    '    Dim fileName As String
    '    Dim zeile As Integer = 2
    '    Dim ok As Boolean

    '    Call projektTafelInit()

    '    appInstance.EnableEvents = False
    '    appInstance.ScreenUpdating = False
    '    enableOnUpdate = False


    '    fileName = "Vorlage Zuordnung.xlsx"

    '    ' öffnen der Excel Datei 
    '    Try
    '        appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & fileName)
    '    Catch ex As Exception
    '        Call MsgBox("File " & fileName & " nicht gefunden ... Abbruch")
    '        appInstance.EnableEvents = True
    '        appInstance.ScreenUpdating = True
    '        enableOnUpdate = True
    '        Exit Sub
    '    End Try




    '    Call awinExportRessZuordnung(0, " ")


    '    Try

    '        appInstance.ActiveWorkbook.SaveAs(awinPath & projektRessOrdner & "\Summary.xlsx", _
    '                                  ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)
    '        ok = True
    '        appInstance.ActiveWorkbook.Close()

    '    Catch ex As Exception
    '        ok = False
    '        appInstance.ActiveWorkbook.Close()
    '    End Try


    '    If ok Then
    '        Call MsgBox("ok, Datei erstellt ...")
    '    Else
    '        Call MsgBox("Fehler bei Save as ..\summary.xlsx")
    '    End If


    '    appInstance.EnableEvents = True
    '    appInstance.ScreenUpdating = True
    '    enableOnUpdate = True


    'End Sub

    ' tk 21.8.17 wird nicht mehr verwendet 
    '''' <summary>
    '''' erstellt die Zuordnungs-Datei Ressourcen -> Projekt
    '''' </summary>
    '''' <param name="control"></param>
    '''' <remarks></remarks>
    'Sub Tom2G4M2B2ZuordnungRP(control As IRibbonControl)

    '    Dim initialeVorlageName As String, kapaFileName As String
    '    Dim zeile As Integer = 2
    '    Dim anzRollen As Integer
    '    Dim i As Integer
    '    Dim initMessage As String = "bitte die Kapazitäten eintragen zu folgenden Rollen" & vbLf
    '    Dim infoMessage As String = initMessage
    '    Dim zuordnungsOrdner As String = projektRessOrdner & "\" & "Projekt Zuordnungen"

    '    Call projektTafelInit()

    '    appInstance.EnableEvents = False
    '    appInstance.ScreenUpdating = False
    '    enableOnUpdate = False



    '    ' für jede Ressource eine eigene Datei machen
    '    anzRollen = RoleDefinitions.Count

    '    Dim ok As Boolean = True
    '    Dim roleName As String

    '    For i = 1 To anzRollen

    '        roleName = RoleDefinitions.getRoledef(i).name.Trim
    '        kapaFileName = roleName & " Kapazität.xlsx"

    '        ' öffnen der Excel Datei 
    '        Try

    '            appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & kapaFileName)
    '            ok = True

    '        Catch ex As Exception

    '            initialeVorlageName = "template Kapazität.xlsx"
    '            ok = False

    '            Try
    '                appInstance.Workbooks.Open(awinPath & projektRessOrdner & "\" & initialeVorlageName)
    '                Try
    '                    appInstance.ActiveWorkbook.SaveAs(awinPath & projektRessOrdner & "\" & kapaFileName, _
    '                                  ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)

    '                    infoMessage = infoMessage & kapaFileName & vbLf
    '                Catch ex2 As Exception

    '                End Try



    '            Catch ex1 As Exception
    '                Call MsgBox("File " & initialeVorlageName & " nicht gefunden ... Abbruch" & vbLf & vbLf & _
    '                            "dieses File muss im Ordner " & awinPath & projektRessOrdner & "abgelegt werden")
    '                appInstance.EnableEvents = True
    '                appInstance.ScreenUpdating = True
    '                enableOnUpdate = True
    '                Exit Sub
    '            End Try

    '        End Try


    '        If ok Then

    '            Dim curFilename As String = roleName & " Projekt-Zuordnung" & " " & Date.Now.ToString("MMM yy") & ".xlsx"


    '            Try
    '                Call awinExportRessZuordnung(1, roleName)
    '                'appInstance.ActiveWorkbook.Save()

    '                appInstance.ActiveWorkbook.SaveAs(Filename:=awinPath & zuordnungsOrdner & "\" & curFilename, _
    '                                                  ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges)


    '            Catch ex As Exception

    '                Call MsgBox("Fehler bei Zuordnung " & roleName)
    '            End Try

    '        End If


    '        appInstance.ActiveWorkbook.Close(SaveChanges:=False)



    '    Next

    '    If infoMessage.Length > initMessage.Length Then
    '        ' in diesem Fall wurden  nur die Kapazität-Zuordnungs-Files erstellt 
    '        infoMessage = infoMessage & vbLf & vbLf & "es wurden noch keine Zuordnungs-Dateien erstellt!"
    '        Call MsgBox(infoMessage)
    '    Else
    '        Call MsgBox("ok, Dateien erstellt ...")
    '    End If



    '    appInstance.EnableEvents = True
    '    appInstance.ScreenUpdating = True
    '    enableOnUpdate = True


    'End Sub

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

    Sub PTTestAPI_Client(control As IRibbonControl)

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt

        Dim clientValues As Double()
        Dim APIvalues As List(Of Double)

        Dim outputString As String = ""
        Dim outPutCollection As New Collection

        Dim err As New clsErrorCodeMsg

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

                    Dim shapeArt As Integer
                    shapeArt = kindOfShape(singleShp)

                    With singleShp
                        If isProjectType(shapeArt) Then

                            Try
                                hproj = ShowProjekte.getProject(singleShp.Name, True)

                                ' jetzt wird dieses Projekt exportiert ... 
                                Try
                                    ' hier muss nun die Berechnung der Personalkosten im Client aufgerufen werden
                                    clientValues = hproj.getAllPersonalKosten

                                    ' hier muss nun die Berechnung der Personaltkosten im Server aufgerufen werden
                                    APIvalues = CType(databaseAcc, DBAccLayer.Request).evaluateCostsOfProject(hproj.name, hproj.variantName, Date.Now, dbUsername, err)

                                    ' die beiden werden nun verglichen

                                    outputString = hproj.getShapeText & " erfolgreich .."
                                    outPutCollection.Add(outputString)

                                    outputString = "Vergleich API - Client"
                                    outPutCollection.Add(outputString)

                                    ' Ausgabe des Ergebnisses
                                    Dim i As Integer = 0
                                    For Each apival As Double In APIvalues
                                        outputString = apival.ToString & "   -   " & clientValues(i)
                                        outPutCollection.Add(outputString)
                                        i = i + 1
                                    Next

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

                Catch ex As Exception

                    outputString = "Fehler in TestAPI_Client"
                    outPutCollection.Add(outputString)

                End Try


            Next

            If outPutCollection.Count > 0 Then
                Call showOutPut(outPutCollection,
                                 "Berechnung Kosten API - Client",
                                 "berechnete Werte im Vergleich im Folgenden")
            End If

        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If


        Call awinDeSelect()
        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True




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

    Public Function PTCompareLast(control As IRibbonControl) As Boolean
        PTCompareLast = awinSettings.meCompareWithLastVersion
    End Function

    Sub awinCompareLast(control As IRibbonControl, ByRef pressed As Boolean)

        awinSettings.meCompareWithLastVersion = pressed

        ' jetzt muss der Auslastungs-Array neu aufgebaut werden 
        visboZustaende.clearAuslastungsArray()

        ' jetzt müssen die Charts aktualisiert werden 
        Call aktualisiereCharts(visboZustaende.currentProject, False)

    End Sub

    Public Function PTProzAuslastung(control As IRibbonControl) As Boolean
        PTProzAuslastung = awinSettings.mePrzAuslastung
    End Function

    Sub awinPTProzAuslastung(control As IRibbonControl, ByRef pressed As Boolean)

        awinSettings.mePrzAuslastung = pressed

        ' jetzt muss der Auslastungs-Array neu aufgebaut werden 
        visboZustaende.clearAuslastungsArray()

        If awinSettings.meExtendedColumnsView Then
            'Call deleteColorFormatMassEdit()
            'Call updateMassEditAuslastungsValues(showRangeLeft, showRangeRight, Nothing)
            'Call colorFormatMassEditRC()
        End If


    End Sub

    Public Function PTSkipChanges(control As IRibbonControl) As Boolean
        PTSkipChanges = tempSkipChanges
    End Function

    Sub awinPTSkipChanges(control As IRibbonControl, ByRef pressed As Boolean)
        tempSkipChanges = pressed
    End Sub

    Public Function PTshowHeader(control As IRibbonControl) As Boolean
        PTshowHeader = tempShowHeaders
    End Function

    Public Function PTmeLastPlanCompare(control As IRibbonControl) As Boolean
        PTmeLastPlanCompare = awinSettings.meCompareVsLastPlan
    End Function

    ''' <summary>
    ''' wenn Header gezeigt werden , können Spaltenbreiten verändert werden ..
    ''' </summary>
    ''' <param name="control"></param>
    ''' <param name="pressed"></param>
    Public Sub awinPTshowHeader(control As IRibbonControl, ByRef pressed As Boolean)
        tempShowHeaders = pressed

        If tempShowHeaders Then
            appInstance.ActiveWindow.DisplayHeadings = True
        Else
            appInstance.ActiveWindow.DisplayHeadings = False
        End If
    End Sub

    Public Sub awinMELastOrBasline(control As IRibbonControl, ByRef pressed As Boolean)
        awinSettings.meCompareVsLastPlan = pressed

        If awinSettings.meCompareVsLastPlan = True Then
            Dim provideDate As New frmProvideDate
            If provideDate.ShowDialog() = DialogResult.OK Then
                ' do nothing , already all set
            Else
                'do nothing, already all set
            End If
        End If
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
                .Protect(Password:="x", UserInterfaceOnly:=True,
                         AllowFormattingCells:=True,
                         AllowFormattingColumns:=True,
                         AllowInsertingColumns:=False,
                         AllowInsertingRows:=True,
                         AllowDeletingColumns:=False,
                         AllowDeletingRows:=True,
                         AllowSorting:=True,
                         AllowFiltering:=True)
                .EnableSelection = Excel.XlEnableSelection.xlNoRestrictions
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
    Public Function PTdontAskWhenAutoReduce(control As IRibbonControl) As Boolean
        PTdontAskWhenAutoReduce = awinSettings.meDontAskWhenAutoReduce
    End Function

    Sub awinPTdontAskWhenAutoReduce(control As IRibbonControl, ByRef pressed As Boolean)
        awinSettings.meDontAskWhenAutoReduce = pressed
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
    ''' zeigt eine Zusammenstellung der wichtigsten Projekt-Charts
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub TomMostImportantProjectCharts(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

        Dim singleShp As Excel.Shape
        Dim ok As Boolean = True
        'Dim SID As String
        Dim hproj As clsProjekt = Nothing
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 1
        Dim top As Double, left As Double, width As Double, height As Double
        Dim myCollection As New Collection
        Call projektTafelInit()

        ' wird für das LEsen des Vergleichs-Projekts gebraucht ... 
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim vglProjekt As clsProjekt = Nothing
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            singleShp = awinSelection.Item(1)

            Try
                hproj = ShowProjekte.getProject(singleShp.Name, True)
                If IsNothing(hproj) Then
                    ok = False
                Else
                    myCollection.Add(hproj.name)
                End If

            Catch ex As Exception
                ok = False
                hproj = Nothing
            End Try


        Else
            If ShowProjekte.Count > 0 Then
                hproj = ShowProjekte.getProject(1)
                If IsNothing(hproj) Then
                    ok = False
                Else
                    ok = True
                    myCollection.Add(hproj.name)
                End If
            Else
                If awinSettings.englishLanguage Then
                    Call MsgBox("no projects loaded ...")
                Else
                    Call MsgBox("es sind keine Projekte geladen ...")
                End If

                ok = False
            End If

        End If

        If ok And Not IsNothing(hproj) Then

            ' bei Projekten, egal ob standard Projekt oder Portfolio Projekt wird immer mit der Vorgaben-Variante verglichen
            Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString
            ' tk 28.12.18 deprecated
            'If hproj.projectType = ptPRPFType.portfolio Then
            '    tmpVariantName = portfolioVName
            'End If

            Dim repObj As Excel.ChartObject
            appInstance.EnableEvents = False
            appInstance.ScreenUpdating = False

            repObj = Nothing

            Dim tmpAnzRollen As Integer = 0
            Dim tmpAnzCosts As Integer = 0

            Try
                tmpAnzRollen = hproj.getRoleNames.Count
                tmpAnzCosts = hproj.getCostNames.Count
            Catch ex As Exception

            End Try



            Try
                ' Projekt-Ergebnis
                Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 0, top, left, width, height)

                Call createProjektErgebnisCharakteristik2(hproj, repObj, PThis.current,
                                                         top, left, width, height, False)


                Dim scInfo As New clsSmartPPTChartInfo
                With scInfo
                    .hproj = hproj
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                        .q2 = myCustomUserRole.specifics
                    Else
                        .q2 = ""
                    End If

                    .vergleichsArt = PTVergleichsArt.beauftragung
                    .prPF = ptPRPFType.project
                    .chartTyp = PTChartTypen.Balken
                    .einheit = PTEinheiten.euro
                    .elementTyp = ptElementTypen.roles
                End With

                Dim compareWithLast As Boolean = awinSettings.meCompareWithLastVersion
                Dim compareTyp As Integer
                Try

                    If compareWithLast Then
                        scInfo.vergleichsTyp = PTVergleichsTyp.letzter
                        scInfo.detailID = PTprdk.PersonalBalken2
                        vglProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, Date.Now, err)
                        compareTyp = PTprdk.PersonalBalken2
                    Else
                        scInfo.vergleichsTyp = PTVergleichsTyp.erster
                        scInfo.detailID = PTprdk.PersonalBalken

                        vglProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)
                        compareTyp = PTprdk.PersonalBalken
                    End If

                Catch ex As Exception
                    vglProjekt = Nothing
                End Try


                scInfo.vglProj = vglProjekt

                ' Rollen-Balken
                Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)
                'Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, tmpAnzRollen, top, left, width, height)

                auswahl = 2 ' zeige Personalbedarfe, aber in T€
                Call createRessBalkenOfProject(scInfo, auswahl, repObj, top, left, height, width, False, calledFromMassEdit:=False)

                ' Kosten-Balken
                Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)
                'Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, tmpAnzCosts, top, left, width, height)

                auswahl = 1 ' zeige Sonstige Kosten
                scInfo.elementTyp = ptElementTypen.costs
                scInfo.q2 = ""
                If compareWithLast Then
                    scInfo.vergleichsTyp = PTVergleichsTyp.letzter
                    scInfo.detailID = PTprdk.KostenBalken2

                Else
                    scInfo.vergleichsTyp = PTVergleichsTyp.erster
                    scInfo.detailID = PTprdk.KostenBalken


                End If

                'If compareWithLast Then
                '    compareTyp = PTprdk.KostenBalken2
                'Else
                '    compareTyp = PTprdk.KostenBalken
                'End If
                'Call createCostBalkenOfProject(hproj, vglProjekt, repObj, auswahl, top, left, height, width, False, compareTyp)
                Call createRessBalkenOfProject(scInfo, auswahl, repObj, top, left, height, width, False, calledFromMassEdit:=False)

                ' Strategie / Risiko / Marge
                Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 0, top, left, width, height)

                Call awinCreatePortfolioDiagrams(myCollection, repObj, True, PTpfdk.FitRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)

                If thereAreAnyCharts(PTwindows.mptpr) Then
                    Dim tmpmsg As String = hproj.getShapeText & " (" & hproj.timeStamp.ToString & ")"
                    Call showVisboWindow(PTwindows.mptpr, tmpmsg)
                End If

            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

            appInstance.EnableEvents = True
            appInstance.ScreenUpdating = True

        End If

        enableOnUpdate = True


    End Sub

    ''' <summary>
    ''' Charakteristik Personal Bedarfe
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B2Resources(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

        Dim singleShp As Excel.Shape
        'Dim SID As String
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 2
        Dim top As Double, left As Double, width As Double, height As Double

        ' wird für das LEsen des Vergleichs-Projekts gebraucht ... 
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim vglProjekt As clsProjekt = Nothing

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

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                Catch ex As Exception
                    ok = False
                    hproj = Nothing
                End Try

                If ok Then


                    ' tk 28.12.18 deprecated
                    'If hproj.projectType = ptPRPFType.portfolio Then
                    '    tmpVariantName = portfolioVName
                    'End If

                    Dim repObj As Excel.ChartObject
                    appInstance.EnableEvents = False
                    appInstance.ScreenUpdating = False

                    repObj = Nothing

                    Dim tmpAnzRollen As Integer = 0
                    Dim tmpAnzCosts As Integer = 0

                    Try
                        tmpAnzRollen = hproj.getRoleNames.Count
                        tmpAnzCosts = hproj.getCostNames.Count
                    Catch ex As Exception

                    End Try

                    Try

                        Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)
                        'Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, tmpAnzRollen, top, left, width, height)

                        Dim scInfo As New clsSmartPPTChartInfo
                        With scInfo
                            .hproj = hproj
                            If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                                .q2 = myCustomUserRole.specifics
                            Else
                                .q2 = ""
                            End If

                            .vergleichsArt = PTVergleichsArt.beauftragung
                            .prPF = ptPRPFType.project
                            .chartTyp = PTChartTypen.Balken
                            .einheit = PTEinheiten.euro
                            .elementTyp = ptElementTypen.roles
                        End With

                        Dim compareWithLast As Boolean = awinSettings.meCompareWithLastVersion
                        Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString
                        Try

                            If compareWithLast Then
                                scInfo.vergleichsTyp = PTVergleichsTyp.letzter
                                scInfo.detailID = PTprdk.PersonalBalken2
                                vglProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, Date.Now, err)

                            Else
                                scInfo.vergleichsTyp = PTVergleichsTyp.erster
                                scInfo.detailID = PTprdk.PersonalBalken

                                vglProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)

                            End If

                        Catch ex As Exception
                            vglProjekt = Nothing
                        End Try


                        scInfo.vglProj = vglProjekt

                        'Call createRessBalkenOfProject(hproj, vglProjekt, repObj, auswahl, top, left, height, width, False, vglTyp:=compareTyp)
                        Call createRessBalkenOfProject(scInfo, auswahl, repObj, top, left, height, width, False)
                        ' jetzt wird das Pie-Diagramm gezeichnet 
                        'Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, tmpAnzRollen, top, left, width, height)
                        Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)

                        repObj = Nothing
                        Call createRessPieOfProject(hproj, repObj, auswahl, top, left, height, width, False)


                        If thereAreAnyCharts(PTwindows.mptpr) Then
                            Dim tmpmsg As String = hproj.getShapeText & " (" & hproj.timeStamp.ToString & ")"
                            Call showVisboWindow(PTwindows.mptpr, tmpmsg)
                        End If

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
    ''' Charakteristik Gesamtkosten
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub Tom2G2M1B5GKosten(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt, vglProj As clsProjekt = Nothing
        Dim awinSelection As Excel.ShapeRange
        Dim auswahl As Integer = 1 ' das steuert , dass die sonstigen Kosten angezeigt werden  
        Dim top As Double, left As Double, width As Double, height As Double

        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

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
                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try


                ' tk 28.12.18 deprecated
                'If hproj.projectType = ptPRPFType.portfolio Then
                '    tmpVariantName = portfolioVName
                'End If

                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                Dim repObj As Excel.ChartObject = Nothing

                Dim tmpAnzCosts As Integer = 0

                Try
                    tmpAnzCosts = hproj.getCostNames.Count
                Catch ex As Exception

                End Try


                Try


                    Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)
                    'Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, tmpAnzCosts, top, left, width, height)

                    Dim compareWithLast As Boolean = awinSettings.meCompareWithLastVersion
                    Dim compareTyp As Integer

                    Try
                        ' bei Projekten, egal ob standard Projekt oder Portfolio Projekt wird immer mit der Vorlagen-Variante verglichen
                        Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString

                        If compareWithLast Then
                            vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, Date.Now, err)
                            compareTyp = PTprdk.KostenBalken2
                        Else
                            vglProj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)
                            compareTyp = PTprdk.KostenBalken
                        End If

                    Catch ex As Exception
                        vglProj = Nothing
                    End Try


                    Call createCostBalkenOfProject(hproj, vglProj, repObj, auswahl, top, left, height, width, False, compareTyp)

                    ' jetzt wird das Pie-Diagramm gezeichnet 
                    'Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, tmpAnzCosts, top, left, width, height)
                    Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)

                    repObj = Nothing
                    Call createCostPieOfProject(hproj, repObj, auswahl, top, left, height, width, False)

                    If thereAreAnyCharts(PTwindows.mptpr) Then
                        Dim tmpmsg As String = hproj.getShapeText & " (" & hproj.timeStamp.ToString & ")"
                        Call showVisboWindow(PTwindows.mptpr, tmpmsg)
                    End If

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
        Dim hproj As clsProjekt

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

            If awinSelection.Count = 1 Then
                ' jetzt die Aktion durchführen ...
                Dim ok As Boolean = True
                singleShp = awinSelection.Item(1)

                Try
                    hproj = ShowProjekte.getProject(singleShp.Name, True)
                    myCollection.Add(hproj.name)
                Catch ex As Exception
                    ok = False
                    hproj = Nothing
                End Try

                If ok Then

                    Dim repObj As Excel.ChartObject

                    repObj = Nothing

                    Try
                        Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)

                        Call awinCreatePortfolioDiagrams(myCollection, repObj, True, PTpfdk.FitRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)

                        If thereAreAnyCharts(PTwindows.mptpr) Then
                            Dim tmpmsg As String = hproj.getShapeText & " (" & hproj.timeStamp.ToString & ")"
                            Call showVisboWindow(PTwindows.mptpr, tmpmsg)
                        End If

                    Catch ex As Exception
                        Call MsgBox(ex.Message)
                    End Try

                End If

            Else
                Call MsgBox("vorher Projekt selektieren ...")
            End If
        End If


        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True


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
            Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.FitRisikoDependency, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
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

            Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.FitRisikoVol, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
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
                Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.Dependencies, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
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
                Call awinCreatePortfolioDiagrams(myCollection, obj, True, PTpfdk.ComplexRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
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


        ' alt ....
        Call projektTafelInit()
        ' auswahl steuert , dass die Personal-Kosten angezeigt werden 
        Dim auswahl As Integer = 1
        Dim vglBaseline As Boolean = True
        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "
        Try
            Call awinSollIstVergleich(auswahl, typ, vglBaseline)
            Call showVisboWindow(PTwindows.mptpr)
        Catch ex As Exception
            Call MsgBox(ex.Message)
        End Try



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
        Call showVisboWindow(PTwindows.mptpr)

    End Sub


    Sub Tom2G2M2M3B2SollIstGKosten(control As IRibbonControl)

        ' auswahl steuert , welche Kosten angezeigt werden
        Dim auswahl As Integer = 3
        Dim vglBaseline As Boolean = True

        ' typ steuert, ob Summenbetrachtung oder Curve angezeigt wird
        Dim typ As String = " "

        Call projektTafelInit()

        Call awinSollIstVergleich(auswahl, typ, vglBaseline)
        Call showVisboWindow(PTwindows.mptpr)

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

        Dim err As New clsErrorCodeMsg

        Dim singleShp As Excel.Shape
        Dim hproj As clsProjekt
        Dim awinSelection As Excel.ShapeRange
        Dim top As Double, left As Double, width As Double, height As Double
        Dim reportobj As Excel.ChartObject
        Dim heute As Date = Date.Now
        Dim vglName As String = " "
        Dim pName As String = ";"
        Dim variantName As String = ""
        Dim bproj As clsProjekt = Nothing

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

                Catch ex As Exception
                    Call MsgBox("Projekt " & singleShp.Name & " nicht gefunden ...")
                    enableOnUpdate = True
                    Exit Sub
                End Try

                ' tk, 7.8.18 wird nicht mehr gebraucht .... wurde ersetzt durch retrieveFirstContracted ...
                ''If Not projekthistorie Is Nothing Then
                ''    If projekthistorie.Count > 0 Then
                ''        vglName = projekthistorie.First.getShapeText
                ''    End If
                ''Else
                ''    projekthistorie = New clsProjektHistorie
                ''End If

                ''With hproj
                ''    pName = .name
                ''    variantName = .variantName
                ''End With


                ''If vglName <> hproj.getShapeText Then
                ''    If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
                ''        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                ''        projekthistorie.liste = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:="",
                ''                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                ''        projekthistorie.Add(Date.Now, hproj)
                ''    Else
                ''        Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
                ''        projekthistorie.clear()
                ''    End If


                ''Else
                ''    ' der aktuelle Stand hproj muss hinzugefügt werden 
                ''    Dim lastElem As Integer = projekthistorie.Count - 1
                ''    projekthistorie.RemoveAt(lastElem)
                ''    projekthistorie.Add(Date.Now, hproj)
                ''End If
                ''Dim nrSnapshots As Integer = projekthistorie.Count

                ' bei normalen Projekten wird immer mit der Basis-Variante verglichen, bei Portfolio Projekten mit dem Portfolio Name
                Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString
                ' tk 28.12.18 deprecated
                'If hproj.projectType = ptPRPFType.portfolio Then
                '    tmpVariantName = portfolioVName
                'End If

                ' das bproj bestimmen 
                bproj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)



                appInstance.EnableEvents = False
                appInstance.ScreenUpdating = False
                reportobj = Nothing

                Dim qualifier As String = " "

                Try

                    Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)

                    If typ = "Curve" Then
                        Call createSollIstCurveOfProject(hproj, bproj, reportobj, heute, auswahl, qualifier, vglBaseline, top, left, height, width)
                    Else
                        Call createSollIstOfProject(hproj, reportobj, heute, auswahl, qualifier, vglBaseline, top, left, height, width, False)
                    End If
                Catch ex As Exception

                End Try

                appInstance.EnableEvents = True
                appInstance.ScreenUpdating = True

            Else
                enableOnUpdate = True
                'Call MsgBox("bitte nur ein Projekt selektieren")
                Throw New ArgumentException("bitte nur ein Projekt selektieren")
            End If
        Else
            enableOnUpdate = True
            'Call MsgBox("vorher Projekt selektieren ...")
            Throw New ArgumentException("vorher Projekt selektieren ...")
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
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
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
                Call awinCreateStatusDiagram1(projektliste, tmpObj, compareTyp, auswahl, qualifier, True, True,
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

    ''' <summary>
    ''' zeigt die wichtigsten Portfolio Charts ...
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub TomMostImportantPortfolioCharts(control As IRibbonControl)

        Dim selectionType As Integer = PTpsel.alle ' Keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim obj As Excel.ChartObject = Nothing
        Dim myCollection As New Collection

        Call projektTafelInit()

        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()

        If ok Then
            appInstance.ScreenUpdating = False
            appInstance.EnableEvents = False
            enableOnUpdate = False


            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

            If myCollection.Count > 0 Then

                ' Portfolio Ergebnis
                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 2, top, left, width, height)

                Call awinCreateBudgetErgebnisDiagramm(obj, top, left, width, height, False, False)

                ' Top 3 Bottlenecks
                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 2, top, left, width, height)

                Try

                    'top = top + height + 10
                    Call createAuslastungsDetailPie(obj, 1, top, left, height, width, False)

                    ' jetzt sollen hier die bis zu drei Top Bottlenecks gezeigt werden 
                    If Not IsNothing(obj) Then
                        Dim top3Collection As New Collection
                        If DiagramList.contains(obj.Name) Then
                            ' in der gsCollection ist die Info drin, um welche Rollen es sich handelt ...
                            top3Collection = DiagramList.getDiagramm(obj.Name).gsCollection
                            For i As Integer = 1 To top3Collection.Count
                                Dim roleName As String = CStr(top3Collection.Item(i))
                                Dim tmpCollection As New Collection
                                tmpCollection.Add(roleName)
                                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 2, top, left, width, height)
                                Call awinCreateprcCollectionDiagram(tmpCollection, obj, top, left, width, height, False, DiagrammTypen(1), False)
                            Next
                        End If

                    End If


                    If thereAreAnyCharts(PTwindows.mptpf) Then
                        Call showVisboWindow(PTwindows.mptpf)
                    End If

                    ' jetzt Unterauslastung
                    ''top = top + height + 10
                    ''Call createAuslastungsDetailPie(obj, 2, top, left, height, width, False)

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
                        Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
                                    "gibt es keine Projekte ")
                    End If
                End If

            End If



            appInstance.ScreenUpdating = True
            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please load projects/portfolios first ...")
            Else
                Call MsgBox("bitte zuerst Projekte/Portfolios laden ...")
            End If
        End If



    End Sub

    Sub PT0ShowCashFlow(control As IRibbonControl)

        Dim selectionType As Integer = PTpsel.alle ' Keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim obj As Excel.ChartObject = Nothing
        Dim myCollection As New Collection

        Call projektTafelInit()



        If ShowProjekte.Count > 0 Then

            If Not (showRangeRight > showRangeLeft) Then
                showRangeLeft = getColumnOfDate(Date.Now) + 1
                showRangeRight = showRangeLeft + 5
                Call awinShowtimezone(showRangeLeft, showRangeRight, True)
            ElseIf showRangeLeft <> getColumnOfDate(Date.Now) + 1 Or showRangeRight <> showRangeLeft + 5 Then
                Call awinShowtimezone(showRangeLeft, showRangeRight, False)
                showRangeLeft = getColumnOfDate(Date.Now) + 1
                showRangeRight = showRangeLeft + 5
                Call awinShowtimezone(showRangeLeft, showRangeRight, True)
            End If

            appInstance.ScreenUpdating = False
            appInstance.EnableEvents = False
            enableOnUpdate = False


            myCollection.Add("Cashflow")

            Try
                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 2, top, left, width, height)
                Call awinCreateprcCollectionDiagram(myCollection, obj, top, left, width, height, False, DiagrammTypen(9), False)

                If thereAreAnyCharts(PTwindows.mptpf) Then
                    Call showVisboWindow(PTwindows.mptpf)
                End If

            Catch ex As Exception
                Call MsgBox("keine Information vorhanden")
            End Try

            appInstance.ScreenUpdating = True
            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please load projects/portfolios first ...")
            Else
                Call MsgBox("bitte zuerst Projekte/Portfolios laden ...")
            End If
        End If

    End Sub
    Sub PT0ShowAuslastung(control As IRibbonControl)

        Dim selectionType As Integer = PTpsel.alle ' Keine Einschränkung
        Dim top As Double, left As Double, width As Double, height As Double
        Dim obj As Excel.ChartObject = Nothing
        Dim myCollection As New Collection

        Call projektTafelInit()

        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()

        If ok Then
            appInstance.ScreenUpdating = False
            appInstance.EnableEvents = False
            enableOnUpdate = False


            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

            If myCollection.Count > 0 Then

                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 2, top, left, width, height)

                Try

                    'top = top + height + 10
                    Call createAuslastungsDetailPie(obj, 1, top, left, height, width, False)

                    ' jetzt sollen hier die bis zu drei Top Bottlenecks gezeigt werden 
                    If Not IsNothing(obj) Then
                        Dim top3Collection As New Collection
                        If DiagramList.contains(obj.Name) Then
                            ' in der gsCollection ist die Info drin, um welche Rollen es sich handelt ...
                            top3Collection = DiagramList.getDiagramm(obj.Name).gsCollection
                            For i As Integer = 1 To top3Collection.Count
                                Dim roleName As String = CStr(top3Collection.Item(i))
                                Dim tmpCollection As New Collection
                                tmpCollection.Add(roleName)
                                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 2, top, left, width, height)
                                Call awinCreateprcCollectionDiagram(tmpCollection, obj, top, left, width, height, False, DiagrammTypen(1), False)
                            Next
                        End If

                    End If


                    If thereAreAnyCharts(PTwindows.mptpf) Then
                        Call showVisboWindow(PTwindows.mptpf)
                    End If

                    ' jetzt Unterauslastung
                    ''top = top + height + 10
                    ''Call createAuslastungsDetailPie(obj, 2, top, left, height, width, False)

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
                        Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
                                    "gibt es keine Projekte ")
                    End If
                End If

            End If



            appInstance.ScreenUpdating = True
            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please load projects/portfolios first ...")
            Else
                Call MsgBox("bitte zuerst Projekte/Portfolios laden ...")
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

                    'height = awinSettings.ChartHoehe1
                    height = maxScreenHeight / 4 - 3
                    'top = 180
                    top = 3
                    left = 3
                    width = maxScreenWidth / 5 - 3
                    'If von > 1 Then
                    '    left = showRangeRight * boxWidth + 4
                    'Else
                    '    left = 0
                    'End If

                    'Dim breite As Integer = System.Math.Max(bis - von, 6)

                    'width = 265 + (bis - von - 12 + 1) * boxWidth + (bis - von) * screen_correct

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


    Sub PT0ShowStrategieRisiko(control As IRibbonControl)

        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim myCollection As New Collection
        Dim top As Double, left As Double, width As Double, height As Double
        'Dim sichtbarerBereich As Excel.Range

        Call projektTafelInit()

        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()
        If ok Then

            appInstance.EnableEvents = False
            enableOnUpdate = False


            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

            If myCollection.Count > 0 Then

                Dim obj As Excel.ChartObject = Nothing


                ' bestimme Position ... 
                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 2, top, left, width, height)

                Try
                    Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisiko, PTpfdk.AmpelFarbe, False, False, True, top, left, width, height, False)
                Catch ex As Exception

                End Try

                If thereAreAnyCharts(PTwindows.mptpf) Then
                    Call showVisboWindow(PTwindows.mptpf)
                End If

            Else

                If ShowProjekte.Count = 0 Then
                    Call MsgBox("es sind keine Projekte angezeigt")

                Else
                    Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
                                "gibt es keine Projekte")
                End If


            End If



            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please load project/portfolios first ...")
            Else
                Call MsgBox("bitte zuerst Projekte/Portfolios laden ...")
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
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoDependency, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoVol, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
                'Call awinCreateStratRiskVolumeDiagramm(myCollection, obj, False, False, True, True, top, left, width, height)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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

        Dim err As New clsErrorCodeMsg

        ' whether or there need to be three or four charts
        Dim withSkills As Boolean = RoleDefinitions.getAllSkillIDs.Count > 0

        ' das Ganze nur machen, wenn das Chart nicht ohnehin schon gezeigt wird ... 
        Try
            If Not IsNothing(projectboardWindows(PTwindows.meChart)) Then
                Exit Sub
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

        With projectboardWindows(PTwindows.massEdit)
            .WindowState = Excel.XlWindowState.xlNormal
            .EnableResize = True
            ' wird ja bereits in der Vorbereitung zu MassEdit gemacht 
            '.SplitRow = 1
            '.FreezePanes = True
            '.DisplayFormulas = False
            '.DisplayHeadings = False
            '.DisplayGridlines = True
            '.GridlineColor = RGB(220, 220, 220)
            '.DisplayWorkbookTabs = False
            '.Caption = bestimmeWindowCaption(PTwindows.massEdit)
            ''.Caption = windowNames(PTwindows.massEdit)
        End With


        projectboardWindows(PTwindows.meChart) = appInstance.ActiveWindow.NewWindow

        visboWorkbook.Worksheets.Item(arrWsNames(ptTables.meCharts)).activate()
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
            '.Caption = windowNames(PTwindows.meChart)
        End With


        ''jetzt das Ursprungs-Window ausblenden ...
        'For Each tmpWindow As Excel.Window In visboWorkbook.Windows
        '    If (CStr(tmpWindow.Caption) <> windowNames(4)) And (CStr(tmpWindow.Caption) <> windowNames(1)) Then
        '        tmpWindow.Visible = False
        '    End If
        'Next


        visboWorkbook.Windows.Arrange(Excel.XlArrangeStyle.xlArrangeStyleHorizontal)

        ' in Abhängigkeit von der Resolution soll jetzt mehr oder weniger prozentualer Platz spendiert werden 
        Dim teilungsfaktor As Double = 0.7
        If maxScreenHeight < 520 Then
            teilungsfaktor = 0.6
        End If

        ' jetzt die Größen anpassen 
        With projectboardWindows(PTwindows.massEdit)
            .Top = 0
            .Left = 1.0
            '.Height = 3 / 4 * maxScreenHeight
            .Height = teilungsfaktor * maxScreenHeight
            .Width = maxScreenWidth - 7.0        ' -7.0, damit der Scrollbar angeklickt werden kann
        End With

        ' jetzt die Größen anpassen 
        With projectboardWindows(PTwindows.meChart)
            .Top = teilungsfaktor * maxScreenHeight + 1
            .Left = 1.0
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


            Dim stdBreite As Double = (projectboardWindows(PTwindows.meChart).UsableWidth - 12) / 3
            Dim showFourDiagrams As Boolean = (withSkills And visboZustaende.projectBoardMode = ptModus.massEditRessSkills)
            If showFourDiagrams Then
                stdBreite = (projectboardWindows(PTwindows.meChart).UsableWidth - 12) / 4
            End If

            Dim chWidth As Double = stdBreite
            Dim chHeight As Double = projectboardWindows(PTwindows.meChart).UsableHeight - 2

            If chHeight < 0.8 * projectboardWindows(PTwindows.meChart).Height Then
                chHeight = 0.8 * projectboardWindows(PTwindows.meChart).Height
            End If

            Dim chTop As Double = 5

            ' show the project Profit/Lost Diagram
            If ShowProjekte.contains(pName) Then
                hproj = ShowProjekte.getProject(pName)
                Call createProjektErgebnisCharakteristik2(hproj, dummyObj, PThis.current,
                                                                     chTop, chLeft, chWidth, chHeight, False, True)

                selectedProjekte.Clear(False)
                selectedProjekte.Add(hproj, False)
            End If


            ' now show Utilization Chart
            ' das Auslastungs-Chart Orga-Einheit
            Dim repObj As Excel.ChartObject = Nothing
            chLeft = chLeft + chWidth + 2
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
            Dim obj As Excel.ChartObject = Nothing
            chLeft = chLeft + chWidth + 2
            chWidth = stdBreite

            ' hier muss jetzt das lproj bestimmt werden 
            Dim lproj As clsProjekt = Nothing


            Dim comparisonTyp As Integer
            Dim qualifier2 As String = ""
            Dim teamID As Integer = -1

            Dim scInfo As New clsSmartPPTChartInfo
            With scInfo
                .hproj = hproj
                .vergleichsArt = PTVergleichsArt.beauftragung
                .einheit = PTEinheiten.euro
                .prPF = ptPRPFType.project
                .elementTyp = ptElementTypen.roles
                .chartTyp = PTChartTypen.Balken
                .detailID = PTprdk.KostenBalken2
            End With

            Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString

            If awinSettings.meCompareVsLastPlan Then
                Dim vpID As String = ""
                lproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(hproj.name, hproj.variantName, vpID, awinSettings.meDateForLastPlan, err)
                comparisonTyp = PTprdk.KostenBalken2

                scInfo.vergleichsArt = PTVergleichsArt.planungsstand
                scInfo.einheit = PTEinheiten.personentage
                scInfo.vergleichsDatum = awinSettings.meDateForLastPlan
                scInfo.vglProj = lproj
                scInfo.vergleichsTyp = PTVergleichsTyp.standVom
                scInfo.q2 = rcName
                scInfo.detailID = PTprdk.KostenBalken2
            Else

                If awinSettings.meCompareWithLastVersion Then

                    lproj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, Date.Now, err)
                    comparisonTyp = PTprdk.KostenBalken2

                    scInfo.vglProj = lproj
                    scInfo.vergleichsTyp = PTVergleichsTyp.letzter
                    scInfo.q2 = ""
                    scInfo.detailID = PTprdk.KostenBalken2

                    If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                        If myCustomUserRole.specifics.Length > 0 Then
                            If RoleDefinitions.containsNameOrID(myCustomUserRole.specifics) Then

                                comparisonTyp = PTprdk.PersonalBalken2
                                scInfo.q2 = RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, teamID).name

                            End If
                        End If

                    ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Then

                        ' wenn es ein Team-Member ist , soll nachgesehen werden, ob es für das Team Vorgaben gibt 
                        ' wenn nein, dann soll die Kostenstelle der Person genommen werden, sofern sie 
                        If rcName <> "" Then
                            Dim potentialParents() As Integer = RoleDefinitions.getIDArray(myCustomUserRole.specifics)

                            If Not IsNothing(potentialParents) And Not IsNothing(lproj) Then

                                Dim tmpParentName As String = ""

                                If rcNameTeamID = -1 Then
                                    tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
                                Else
                                    Dim tmpTeamName As String = RoleDefinitions.getRoleDefByID(rcNameTeamID).name
                                    tmpParentName = RoleDefinitions.chooseParentFromList(tmpTeamName, potentialParents)
                                    If tmpParentName = "" Then
                                        tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
                                    Else
                                        Dim tmpParentNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpParentName, "")
                                        If lproj.containsRoleNameID(tmpParentNameID) Then
                                            ' passt bereits 
                                        Else
                                            tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
                                        End If

                                    End If
                                End If

                                If tmpParentName <> "" Then
                                    scInfo.q2 = tmpParentName
                                End If
                            End If


                        End If

                    End If



                Else
                    lproj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)
                    comparisonTyp = PTprdk.KostenBalken

                    scInfo.vglProj = lproj
                    scInfo.vergleichsTyp = PTVergleichsTyp.erster
                    scInfo.q2 = ""
                    scInfo.detailID = PTprdk.KostenBalken

                    If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
                        If myCustomUserRole.specifics.Length > 0 Then
                            If RoleDefinitions.containsNameOrID(myCustomUserRole.specifics) Then

                                comparisonTyp = PTprdk.PersonalBalken
                                scInfo.q2 = RoleDefinitions.getRoleDefByIDKennung(myCustomUserRole.specifics, teamID).name

                            End If
                        End If

                    ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung Then

                        If rcName <> "" Then
                            Dim potentialParents() As Integer = RoleDefinitions.getIDArray(myCustomUserRole.specifics)

                            If Not IsNothing(potentialParents) And Not IsNothing(lproj) Then

                                Dim tmpParentName As String = ""

                                If rcNameTeamID = -1 Then
                                    tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
                                Else
                                    Dim tmpTeamName As String = RoleDefinitions.getRoleDefByID(rcNameTeamID).name
                                    tmpParentName = RoleDefinitions.chooseParentFromList(tmpTeamName, potentialParents)
                                    If tmpParentName = "" Then
                                        tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
                                    Else
                                        Dim tmpParentNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpParentName, "")
                                        If lproj.containsRoleNameID(tmpParentNameID) Then
                                            ' passt bereits 
                                        Else
                                            tmpParentName = RoleDefinitions.chooseParentFromList(rcName, potentialParents)
                                        End If

                                    End If
                                End If

                                If tmpParentName <> "" Then
                                    scInfo.q2 = tmpParentName
                                End If

                            End If


                        End If
                    End If

                End If

            End If

            'Dim vglBaseline As Boolean = Not IsNothing(lproj)
            Dim reportObj As Excel.ChartObject = Nothing


            Try

                Call createRessBalkenOfProject(scInfo, 2, reportObj, chTop, chLeft, chHeight, chWidth, True,
                                                   calledFromMassEdit:=True)

                ' alt, am 20.2. durch obiges ersetzt 
                'If scInfo.q2 = "" Then
                '    Call createCostBalkenOfProject(hproj, lproj, reportObj, 2, chTop, chLeft, chHeight, chWidth, False, comparisonTyp)
                'Else
                '    Call createRessBalkenOfProject(scInfo, 2, reportObj, chTop, chLeft, chHeight, chWidth, True,
                '                                   calledFromMassEdit:=True)
                'End If


            Catch ex As Exception

            End Try


        Else
            ' sie sind schon da 

        End If

        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True
        enableOnUpdate = True

        ' jetzt das Chart-Window aktivieren (sonst bleibt Ribbon stehen)
        projectboardWindows(PTwindows.meChart).Activate()


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
                    Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.Dependencies, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
                Else
                    Call MsgBox(" es gibt in diesem Zeitraum keine Projekte mit Abhängigkeiten")
                End If


            Catch ex As Exception

            End Try

        Else
            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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
                        Call awinCreateBetterWorsePortfolio(ProjektListe:=myCollection, repChart:=obj, showAbsoluteDiff:=True, isTimeTimeVgl:=False, vglTyp:=1,
                                                        charttype:=PTpfdk.betterWorseL, bubbleColor:=0, bubbleValueTyp:=PTbubble.strategicFit, showLabels:=True, chartBorderVisible:=True,
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
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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
                        Call awinCreateBetterWorsePortfolio(ProjektListe:=myCollection, repChart:=obj, showAbsoluteDiff:=True, isTimeTimeVgl:=False, vglTyp:=1,
                                                        charttype:=PTpfdk.betterWorseB, bubbleColor:=0, bubbleValueTyp:=PTbubble.strategicFit, showLabels:=True, chartBorderVisible:=True,
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
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.ComplexRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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
                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.ZeitRisiko, PTpfdk.ProjektFarbe, False, True, True, top, left, width, height, False)
            Catch ex As Exception

            End Try

        Else

            If ShowProjekte.Count = 0 Then
                Call MsgBox("es sind keine Projekte angezeigt")

            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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

        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()

        If ok Then

            appInstance.EnableEvents = False
            enableOnUpdate = False

            Dim formerES As Boolean = awinSettings.meEnableSorting

            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

            If myCollection.Count > 0 Then

                Call bestimmeChartPositionAndSize(ptTables.mptPfCharts, 2, top, left, width, height)

                Dim obj As Excel.ChartObject = Nothing
                Call awinCreateBudgetErgebnisDiagramm(obj, top, left, width, height, False, False)

                If thereAreAnyCharts(PTwindows.mptpf) Then
                    ' jetzt sollte das Window gezeigt werden, wenn es nicht schon sichtbar ist ... 
                    Call showVisboWindow(PTwindows.mptpf)
                End If

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
                            Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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
                Call MsgBox("please load projects/portfolios first ...")
            Else
                Call MsgBox("bitte erst Projekte/Portfolios laden ...")
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
                    Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
                            "gibt es keine Projekte ")
                End If



            End If

            appInstance.EnableEvents = True
            enableOnUpdate = True

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("there are no projects in Timeframe " & textZeitraum(showRangeLeft, showRangeRight))
            Else
                Call MsgBox("im angezeigten Zeitraum " & textZeitraum(showRangeLeft, showRangeRight) & vbLf &
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
                        Dim top As Double = 0
                        Dim left As Double = 0
                        Dim height As Double = 0
                        Dim width As Double = 0

                        Call bestimmeChartPositionAndSize(ptTables.mptPrCharts, 2, top, left, width, height)

                        Call createProjektErgebnisCharakteristik2(hproj, dummyObj, PThis.current,
                                                                 top, left, width, height, False)

                        If thereAreAnyCharts(PTwindows.mptpr) Then
                            Dim tmpmsg As String = hproj.getShapeText & " (" & hproj.timeStamp.ToString & ")"
                            Call showVisboWindow(PTwindows.mptpr, tmpmsg)
                        End If

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
            Call MsgBox("ein Projekt selektieren, um mit Vorlage zu vergleichen" & vbLf &
                        " oder zwei Projekte für den Vergleich untereinander")
        End If

        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub

    Sub PT3G1B2PhasenVgl(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

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


        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

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
                            projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName,
                                                                                storedEarliest:=StartofCalendar, storedLatest:=Date.Now, err:=err)
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

        Dim err As New clsErrorCodeMsg

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

        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

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
                        projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName,
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now, err:=err)
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

        Dim err As New clsErrorCodeMsg

        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
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

        If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then

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
                        projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName,
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now, err:=err)
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

        Dim err As New clsErrorCodeMsg

        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
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
                    If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName,
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now, err:=err)
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

            End If
        Else
            Call MsgBox("vorher Projekt selektieren ...")
        End If

        enableOnUpdate = True
        appInstance.ScreenUpdating = True


    End Sub


    Sub awinShowTimeMachine(control As IRibbonControl)

        Call PBBShowTimeMachine(control)

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
        Dim timeZoneWasOff As Boolean = (showRangeLeft = 0 Or showRangeRight = 0)
        getReportVorlage.calledfrom = "Portfolio1"

        Call projektTafelInit()


        Dim ok As Boolean = setTimeZoneIfTimeZonewasOff()


        enableOnUpdate = False
        appInstance.ScreenUpdating = False
        If ok Then

            If ShowProjekte.Count > 0 Then

                ' Formular zum Auswählen der Report-Vorlage wird aufgerufen

                returnValue = getReportVorlage.ShowDialog

            Else
                Call MsgBox("Es sind keine Projekte geladen!")
            End If
        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please load projects/portfoliso first ...")
            Else
                Call MsgBox("bitte erst Projekt/Portfolios laden ...")
            End If

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
        Call MsgBox("Betriebssystem: " & appInstance.OperatingSystem & Chr(10) &
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
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
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

                Next
            End If
        End If



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
                        logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Meilenstein-Zugriff über mx: " & ix & vbLf &
                                     curNode.indexOfElem & " <> " & mx
                    End If

                    parentID = curNode.parentNodeKey
                    parentNode = hproj.hierarchy.nodeItem(parentID)
                    If Not IsNothing(parentNode) Then
                        If parentNode.indexOfElem <> ix Then
                            logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Phasen-Zugriff über ix: " & ix & vbLf &
                                         parentNode.indexOfElem & " <> " & mx
                        End If
                    Else
                        If parentID <> "" Then
                            logMessage = logMessage & vbLf & kvp.Value.getShapeText & "Phasen-Zugriff über ix: " & ix & vbLf &
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


        Dim outputcollection As Collection = TestAggregateMethod()

        If outputcollection.Count > 0 Then
            Call showOutPut(outputcollection, "Test Aggregate-Funktion", "")
        End If


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

        Dim err As New clsErrorCodeMsg

        Call projektTafelInit()
        enableOnUpdate = False
        appInstance.EnableEvents = True
        If CType(databaseAcc, DBAccLayer.Request).cancelWriteProtections(dbUsername, err) Then
            If awinSettings.visboDebug Then
                Call MsgBox("Ihre vorübergehenden Schreibsperren wurden aufgehoben")
            End If
        End If
        'Dim ok2 As Boolean = CType(databaseAcc, DBAccLayer.Request).cancelWriteProtections(dbUsername, err)

        enableOnUpdate = True

    End Sub

    Public Sub PTCreateLicense(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim frmLizenzen As New frmCreateLicences


        Dim returnValue As DialogResult
        returnValue = frmLizenzen.ShowDialog


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


    ''' <summary>
    ''' erzeugt die Report Messages ...
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub PTCreateReportMessages(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim frmCreateReportMsg As New frmCreateReportMeldungen
        Dim returnValue As DialogResult
        returnValue = frmCreateReportMsg.ShowDialog


        enableOnUpdate = True


    End Sub


    ''' <summary>
    ''' Spracheinstellungen für die Reports
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
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
    ''' Einstellungen zum Visual Board von VISBO
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Public Sub PTVisboSettings(control As IRibbonControl)

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Dim frmVisboEinst As New frmEinstellungen
        Dim returnValue As DialogResult
        returnValue = frmVisboEinst.ShowDialog

        enableOnUpdate = True

        Me.ribbon.Invalidate()


    End Sub

    Public Sub PTChgUserRole(control As IRibbonControl)

        Dim meldungen As Collection = New Collection

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True

        Try
            ' Lesen der CustomUserRoles aus VCSetting in DB
            Call setUserRoles(meldungen)
        Catch ex As Exception
            Call MsgBox("Error when changing User Role: " & ex.Message)
        End Try


        enableOnUpdate = True

        Me.ribbon.Invalidate()
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

        If Not IsNothing(currentReportProfil) Then

            If currentReportProfil.PPTTemplate <> "" Then
                Dim returnvalue As DialogResult
                returnvalue = profilNameForm.ShowDialog
            End If

        End If
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


    ' ----------------------------------------------------------------
    ' Ab hier sind die MenuPunkte für WebServer
    ' ----------------------------------------------------------------
    Public Sub PTWebRequestLogin(control As IRibbonControl)
        Try

            Dim loginerfolgreich As Boolean = logInToMongoDB(True)

        Catch ex As Exception

        End Try
    End Sub
    Public Sub PTWebVCSettingTest(control As IRibbonControl)
        Try

            Dim listofCURs As New clsCustomUserRoles
            Dim type As String = CStr(settingTypes(ptSettingTypes.customroles))
            Dim name As String = type
            Dim err As New clsErrorCodeMsg

            Dim result As Boolean = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(listofCURs, type, name, Date.Now, err)
        Catch ex As Exception

        End Try
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
