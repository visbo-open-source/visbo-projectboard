
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports DBAccLayer
Imports ProjectboardReports
'Imports WPFPieChart ' wird nicht verwendet 
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Security.Principal
Imports System.Diagnostics
Imports System.Drawing
Imports System.Windows.Forms


Public Module PBBModules


    ''' <summary>
    ''' wird aus der Multiprojekt-Tafel zum Testen der Report Erstellungs-Routinen 
    ''' und aus dem MS Project AddIn aufgerufen 
    ''' </summary>
    ''' <param name="controlID"></param>
    ''' <remarks></remarks>

    Sub PBBBHTCHierarchySelAction(controlID As String, ByVal reportprofil As clsReportAll)

        Dim hryFormular As New frmHierarchySelection
        Dim returnValue As DialogResult
        Dim formerSettings(3) As Boolean

        If controlID <> "BHTC" Then
            hryFormular.calledFrom = "Multiprojekt-Tafel"

            If Not IsNothing(reportprofil) Then
                hryFormular.repProfil = New clsReportAll
                reportprofil.CopyTo(hryFormular.repProfil)
            End If

            With awinSettings
                formerSettings(0) = .mppExtendedMode
                formerSettings(1) = .mppShowAllIfOne
                formerSettings(2) = .mppShowAmpel
                formerSettings(3) = .mppFullyContained
            End With

            With awinSettings
                .mppExtendedMode = True
                .mppShowAllIfOne = False
                .mppShowAmpel = False
                .mppFullyContained = False
            End With

        Else
            hryFormular.calledFrom = "MS-Project"

            hryFormular.repProfil = New clsReportAll
            reportprofil.CopyTo(hryFormular.repProfil)
        End If

        ' Dim formerSettings(3) As Boolean
        With awinSettings
            formerSettings(0) = .mppExtendedMode
            formerSettings(1) = .mppShowAllIfOne
            formerSettings(2) = .mppShowAmpel
            formerSettings(3) = .mppFullyContained
        End With

        With awinSettings
            .mppExtendedMode = True
            .mppShowAllIfOne = False
            .mppShowAmpel = False
            .mppFullyContained = False
        End With

        awinSettings.useHierarchy = True
        With hryFormular


            Try
                If .calledFrom = "MS-Project" Then


                    .menuOption = PTmenue.reportBHTC

                    ' hier müssen die für BHTC nicht wählbaren Optionen gesetzt werden 
                    With awinSettings
                        .mppShowProjectLine = False
                        .mppShowAmpel = False
                        .mppShowAllIfOne = False
                        .mppSortiertDauer = False
                        .mppExtendedMode = True
                        '.eppExtendedMode = True
                    End With

                    If Not IsNothing(reportprofil) Then
                        .filterDropbox.Text = reportprofil.name
                    Else
                        .filterDropbox.Text = ""
                    End If


                    Dim lic As New clsLicences
                    Try
                        lic = XMLImportLicences(licFileName)
                    Catch ex As Exception

                    End Try

                    ' nur mit dem Recht für ProjectAdmin können ReportProfile gespeichert werden
                    If lic.validLicence(myWindowsName, LizenzKomponenten(PTSWKomp.ProjectAdmin)) Then

                        .auswSpeichern.Visible = True
                        .filterDropbox.Enabled = True
                    Else
                        .auswSpeichern.Visible = False
                        .filterDropbox.Enabled = False
                    End If

                ElseIf .calledFrom = "Multiprojekt-Tafel" Then

                    .menuOption = PTmenue.reportMultiprojektTafel

                    If Not IsNothing(reportprofil) Then
                        .filterDropbox.Text = reportprofil.name
                    Else
                        .filterDropbox.Text = ""
                    End If

                    .auswSpeichern.Visible = True
                    .filterDropbox.Enabled = True

                Else
                    .auswSpeichern.Visible = False
                    .filterDropbox.Enabled = False
                End If

            Catch ex As Exception
                .auswSpeichern.Visible = False
                .filterDropbox.Enabled = False
            End Try


            ' bei Verwendung Background Worker muss Aufruf so erfolgen: 
            returnValue = .ShowDialog
        End With


        With awinSettings
            .mppExtendedMode = formerSettings(0)
            .mppShowAllIfOne = formerSettings(1)
            .mppShowAmpel = formerSettings(2)
            .mppFullyContained = formerSettings(3)
        End With


    End Sub
   

    ''' <summary>
    ''' wird aus der Multiprojekt-Tafel aufgerufen 
    ''' </summary>
    ''' <param name="controlID"></param>
    ''' <remarks></remarks>
    Sub PBBNameHierarchySelAction(controlID As String)


        Dim nameFormular As New frmNameSelection
        Dim hryFormular As New frmHierarchySelection
        Dim awinSelection As Excel.ShapeRange
        Dim returnValue As DialogResult
        Dim timeZoneWasOff As Boolean = True

        Call projektTafelInit()

        hryFormular.calledFrom = "Multiprojekt-Tafel"


        ' gibt es überhaupt Objekte, zu denen man was anzeigen kann ? 
        'If ShowProjekte.Count > 0 And showRangeRight - showRangeLeft > 5 Then

        If controlID = "Pt6G3M1B1" Then
            ' normale, volle Auswahl des filters ; Namens-Definition
            ' With nameFormular
            awinSettings.useHierarchy = True
            With hryFormular

                .menuOption = PTmenue.filterdefinieren
                returnValue = .ShowDialog

            End With

        ElseIf controlID = "Pt6G3M1B2" Then

            awinSettings.useHierarchy = True

            With hryFormular

                .menuOption = PTmenue.filterdefinieren
                returnValue = .ShowDialog

            End With

        ElseIf controlID = "PT0G1B8" Then
            ' Menupunkt Filter  

            Dim currentFilterConstellation As clsConstellation = currentSessionConstellation.copy(False, "Filter Result")
            beforeFilterConstellation = currentSessionConstellation.copy(False, "beforeFilter")

            Dim formerEoU As Boolean = enableOnUpdate
            enableOnUpdate = False
            Dim filter As clsFilter = Nothing

            Try
                '  With nameFormular

                awinSettings.useHierarchy = True

                With hryFormular

                    Dim anzP As Integer = ShowProjekte.Count
                    .menuOption = PTmenue.sessionFilterDefinieren
                    '.actionCode = PTTvActions.chgInSession


                    returnValue = .ShowDialog
                    filter = filterDefinitions.retrieveFilter("Last")

                    ' Anzeigen ...
                    Dim removeList As New Collection


                    For Each kvp As KeyValuePair(Of String, clsConstellationItem) In currentFilterConstellation.Liste

                        If ShowProjekte.contains(kvp.Value.projectName) Then
                            Dim hproj As clsProjekt = ShowProjekte.getProject(kvp.Value.projectName)

                            If filter.doesNotBlock(hproj) Then
                                ' nichts tun 
                            Else
                                If Not removeList.Contains(kvp.Key) Then
                                    removeList.Add(kvp.Key, kvp.Key)
                                End If
                            End If

                        Else
                            If Not removeList.Contains(kvp.Key) Then
                                removeList.Add(kvp.Key, kvp.Key)
                            End If
                        End If

                    Next

                    ' jetzt die Liste bereinigen ...
                    For Each tmpPvName As String In removeList
                        currentFilterConstellation.remove(tmpPvName)
                    Next

                    If currentFilterConstellation.sortCriteria = ptSortCriteria.customTF Then
                        ' jetzt wird das SortCriteria umgesetzt, weil andernfalls, bei customTF, die 
                        ' Zeilen unverändert bleiben ... 
                        currentFilterConstellation.sortCriteria = ptSortCriteria.customListe
                    End If

                    ' jetzt müssen die tfZeile neu besetzt werden;
                    '  nach standard, d.h 0 bedeutet einfach sortiert nach Name 
                    ' tk 21.3.17: ab jetzt nicht mehr .... jetzt wird ja in der _sortlist alles mitgeführt 
                    ''currentBrowserConstellation.setTfZeilen(0)

                    If removeList.Count > 0 Then


                        ' erst am Ende alle Diagramme neu machen ...

                        Dim tmpConstellation As New clsConstellations
                        tmpConstellation.Add(currentFilterConstellation)

                        ' es in der Session Liste verfügbar machen 
                        ' es in der Session Liste verfügbar machen
                        If projectConstellations.Contains(currentFilterConstellation.constellationName) Then
                            projectConstellations.Remove(currentFilterConstellation.constellationName)
                        End If

                        projectConstellations.Add(currentFilterConstellation)

                        Call showConstellations(constellationsToShow:=tmpConstellation,
                                                clearBoard:=True, clearSession:=False, storedAtOrBefore:=Date.Now)

                        ''Call awinNeuZeichnenDiagramme(2)

                    End If


                End With
            Catch ex As Exception

            End Try

            enableOnUpdate = formerEoU

        ElseIf controlID = "PT0G1B9" Then

            Dim formerEoU As Boolean = enableOnUpdate
            enableOnUpdate = False
            Dim filter As clsFilter = Nothing

            Try
                If IsNothing(beforeFilterConstellation) Then

                    If awinSettings.visboDebug Then

                        If awinSettings.englishLanguage Then
                            Call MsgBox("There is no active filter!")
                        Else
                            Call MsgBox("Es ist kein Filter gesetzt!")
                        End If

                    End If

                Else

                    If beforeFilterConstellation.Liste.Count = 0 Then

                        If awinSettings.visboDebug Then

                            If awinSettings.englishLanguage Then
                                Call MsgBox("There is no active filter!")
                            Else
                                Call MsgBox("Es ist kein Filter gesetzt!")
                            End If

                        End If

                    Else

                        ' erst am Ende alle Diagramme neu machen ...
                        Dim tmpConstellations As New clsConstellations
                        tmpConstellations.Add(beforeFilterConstellation)

                        '' es in der Session Liste verfügbar machen
                        'If projectConstellations.Contains(beforeFilterConstellation.constellationName) Then
                        '    projectConstellations.Remove(beforeFilterConstellation.constellationName)
                        'End If

                        'projectConstellations.Add(beforeFilterConstellation)

                        Call showConstellations(constellationsToShow:=tmpConstellations,
                                            clearBoard:=True, clearSession:=False, storedAtOrBefore:=Date.Now)

                    End If

                End If

            Catch ex As Exception

                If awinSettings.visboDebug Then
                    Call MsgBox("Fehler beim Zurücksetzen des Filters")
                End If

            End Try

            enableOnUpdate = formerEoU


        ElseIf ShowProjekte.Count > 0 Then

            If awinSettings.isHryNameFrmActive Then
                Call MsgBox("es kann nur ein Fenster zur Hierarchie- bzw. Namenauswahl geöffnet sein ...")

            ElseIf controlID = "PTXG1B4" Or controlID = "PT0G1B8" Then
                ' Namen auswählen, Visualisieren
                Dim ok As Boolean = True
                If controlID = "PTXG1B4" Then
                    ' Multiprojekt Sicht erfordert Zeitraum 

                    timeZoneWasOff = setTimeZoneIfTimeZonewasOff()

                    'If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                    '    ' alles ok 
                    'Else
                    '    timeZoneWasOff = True
                    '    If selectedProjekte.Count > 0 Then
                    '        showRangeLeft = selectedProjekte.getMinMonthColumn
                    '        showRangeRight = selectedProjekte.getMaxMonthColumn
                    '    Else
                    '        showRangeLeft = ShowProjekte.getMinMonthColumn
                    '        showRangeRight = ShowProjekte.getMaxMonthColumn
                    '    End If
                    '    Call awinShowtimezone(showRangeLeft, showRangeRight, True)
                    '    ' wurde jetzt ersetzt durch automatische Selektion
                    '    'ok = False
                    '    'If awinSettings.englishLanguage Then
                    '    '    Call MsgBox("please define timeframe first ...")
                    '    'Else
                    '    '    Call MsgBox("bitte zuerst den Zeitraum definieren ...")
                    '    'End If
                    'End If
                End If

                If ok Then
                    awinSettings.useHierarchy = False
                    With nameFormular

                        .menuOption = PTmenue.visualisieren
                        ' Nicht Modal anzeigen
                        .Show()
                        'returnValue = .ShowDialog

                    End With
                End If


            ElseIf controlID = "PTXG1B5" Or controlID = "PT0G1B9" Then
                ' Hierarchie auswählen, visualisieren
                Dim ok As Boolean = True
                If controlID = "PTXG1B5" Then
                    ' Multiprojekt Sicht erfordert Zeitraum 
                    If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                        ' alles ok 
                    Else
                        timeZoneWasOff = setTimeZoneIfTimeZonewasOff()

                        'ok = False
                        'If awinSettings.englishLanguage Then
                        '    Call MsgBox("please define timeframe first ...")
                        'Else
                        '    Call MsgBox("bitte zuerst den Zeitraum definieren ...")
                        'End If
                    End If
                End If

                If ok Then
                    awinSettings.useHierarchy = True

                    With hryFormular

                        .menuOption = PTmenue.visualisieren
                        ' Nicht Modal anzeigen
                        .Show()
                        'returnValue = .ShowDialog

                    End With
                End If

            ElseIf controlID = "PTXG1B6" Or controlID = "PTMEC1" Then
                ' Namen auswählen, Leistbarkeit
                Dim ok As Boolean = True
                If controlID = "PTXG1B6" Then
                    ' Multiprojekt Sicht erfordert Zeitraum 
                    If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                        ' alles ok 
                    Else
                        timeZoneWasOff = setTimeZoneIfTimeZonewasOff()
                        '
                        'ok = False
                        'If awinSettings.englishLanguage Then
                        '    Call MsgBox("please define timeframe first ...")
                        'Else
                        '    Call MsgBox("bitte zuerst den Zeitraum definieren ...")
                        'End If
                    End If
                End If

                If ok Then
                    awinSettings.useHierarchy = False
                    With nameFormular

                        .ribbonButtonID = controlID
                        .menuOption = PTmenue.leistbarkeitsAnalyse
                        ' Nicht Modal anzeigen
                        '.Show()
                        returnValue = .ShowDialog

                    End With
                End If


            ElseIf controlID = "PTXG1B7" Then



                ' Multiprojekt Sicht erfordert Zeitraum 
                If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                    ' alles ok 
                    ' Hierarchie auswählen, Leistbarkeit
                Else

                    timeZoneWasOff = setTimeZoneIfTimeZonewasOff()


                End If

                awinSettings.useHierarchy = True
                With hryFormular

                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    ' Nicht Modal anzeigen
                    '.Show()
                    returnValue = .ShowDialog

                End With

            ElseIf controlID = "PT1G1M1B1" Then
                ' Namen auswählen, Einzelprojekt Berichte 

                Try
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                If awinSelection Is Nothing Then
                    Call MsgBox("vorher Projekt/e selektieren ...")
                Else

                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' false, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 
                    appInstance.ScreenUpdating = False
                    appInstance.EnableEvents = False

                    With nameFormular


                        .menuOption = PTmenue.einzelprojektReport
                        '.Show()
                        ' bei Reports mit der Background Worker Behandlung 
                        returnValue = .ShowDialog()

                    End With

                    appInstance.ScreenUpdating = True
                    appInstance.EnableEvents = True

                End If

            ElseIf controlID = "PT1G1M1B2" Then

                Try
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                If awinSelection Is Nothing Then
                    Call MsgBox("vorher Projekt/e selektieren ...")
                Else


                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' dalse, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 
                    appInstance.ScreenUpdating = False
                    appInstance.EnableEvents = False

                    awinSettings.useHierarchy = True
                    With hryFormular

                        .menuOption = PTmenue.einzelprojektReport
                        ' bei Verwendung Background Worker muss Modal erfolgen 
                        '.Show()
                        returnValue = .ShowDialog

                    End With

                    appInstance.ScreenUpdating = True
                    appInstance.EnableEvents = True
                End If

            ElseIf controlID = "PT1G1M2B1" Then


                If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                    ' Namen Auswahl, Multiprojekt Report
                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' dalse, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 


                Else

                    timeZoneWasOff = setTimeZoneIfTimeZonewasOff()

                End If

                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                With nameFormular

                    .menuOption = PTmenue.multiprojektReport
                    ' .show; bei Verwendung mit Background Worker Funktion muss das modal erfolgen
                    returnValue = .ShowDialog

                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True

            ElseIf controlID = "PT1G1M2B2" Then

                If showRangeLeft > 0 And showRangeRight > showRangeLeft Then

                    ' Hierarchie Auswahl, Multiprojekt Report
                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' dalse, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 

                Else

                    timeZoneWasOff = setTimeZoneIfTimeZonewasOff()

                End If

                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                awinSettings.useHierarchy = True
                With hryFormular

                    .menuOption = PTmenue.multiprojektReport
                    ' .show; bei Verwendung mit Background Worker Funktion muss das modal erfolgen
                    returnValue = .ShowDialog

                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True

            ElseIf controlID = "PT4G1M0B1" Then
                ' Auswahl über Namen, Typ II Export
                ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                ' dalse, dann auf True gesetzt werden
                ' bei .show darf das nicht gemacht werden ! 
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                With nameFormular


                    .menuOption = PTmenue.excelExport
                    returnValue = .ShowDialog

                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True

            ElseIf controlID = "PT4G1M0B2" Then

                ' Auswahl über Hierarchie, Typ II Export
                ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                ' dalse, dann auf True gesetzt werden
                ' bei .show darf das nicht gemacht werden ! 
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                awinSettings.useHierarchy = True

                With hryFormular

                    .menuOption = PTmenue.excelExport
                    ' Nicht Modal anzeigen
                    '.Show()
                    returnValue = .ShowDialog

                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True

            ElseIf controlID = "PT4G1M2B1" Then
                ' Auswahl über Namen, Vorlagen erzeugen
                ' Auswahl über Hierarchie, Typ II Export
                ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                ' dalse, dann auf True gesetzt werden
                ' bei .show darf das nicht gemacht werden ! 
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                With nameFormular


                    .menuOption = PTmenue.vorlageErstellen
                    returnValue = .ShowDialog

                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True


            ElseIf controlID = "PT4G1M2B2" Then
                ' Auswahl über Hierarchie, Vorlagen Export

                ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                ' dalse, dann auf True gesetzt werden
                ' bei .show darf das nicht gemacht werden ! 
                appInstance.ScreenUpdating = False
                appInstance.EnableEvents = False

                awinSettings.useHierarchy = True
                With hryFormular

                    .menuOption = PTmenue.vorlageErstellen
                    ' Nicht Modal anzeigen
                    '.Show()
                    returnValue = .ShowDialog

                End With

                appInstance.ScreenUpdating = True
                appInstance.EnableEvents = True

            ElseIf controlID = "PT0G1M2B7" Then
                ' Auswahl über Namen, Meilensteine für Meilenstein Trendanalyse
                Try
                    awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
                Catch ex As Exception
                    awinSelection = Nothing
                End Try

                If awinSelection Is Nothing Then
                    Call MsgBox("vorher Projekt/e selektieren ...")
                Else

                    ' wenn nachher .showdialog aufgerufen wird, müssen die beiden Settings erst auf 
                    ' dalse, dann auf True gesetzt werden
                    ' bei .show darf das nicht gemacht werden ! 
                    appInstance.ScreenUpdating = False
                    appInstance.EnableEvents = False

                    With nameFormular

                        .menuOption = PTmenue.meilensteinTrendanalyse
                        returnValue = .ShowDialog()

                    End With

                    appInstance.ScreenUpdating = True
                    appInstance.EnableEvents = True

                End If


            End If
        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please load projects/portfolios first ...")
            Else
                Call MsgBox("bitte zuerst Projekte/Portfolios laden ...")
            End If
        End If

        ' darf nicht zurückgenommen werden, weil manche Fenster nicht modal angezeigt werden, d.h bevor irgendeine Aktion passiert 
        ' wird der TimeFrame wieder zurückgesetzt ...
        'If timeZoneWasOff Then
        '    Call awinShowtimezone(showRangeLeft, showRangeRight, False)
        '    showRangeLeft = 0
        '    showRangeRight = 0
        'End If

        ' oben ist es de-aktiviert 
        'appInstance.EnableEvents = True
        'enableOnUpdate = True

    End Sub

    Sub PBBAnalyseLeistbarkeit001(ByVal ControlID As String)

        Dim namensFormular As New frmNameSelection
        Dim hierarchieFormular As New frmHierarchySelection
        Dim returnValue As DialogResult


        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = False

        ' gibt es überhaupt Objekte, zu denen man was anzeigen kann ? 
        If ShowProjekte.Count > 0 And showRangeRight - showRangeLeft >= minColumns - 1 Then

            If ControlID = "PTXG1B6" Then
                ' Auswahl über Namen

                With namensFormular

                    .ribbonButtonID = ControlID
                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    '.Show()
                    returnValue = .ShowDialog

                End With


            Else
                ' Auswahl über Hierarchie
                ' Hierarchie
                awinSettings.useHierarchy = True
                With hierarchieFormular

                    .menuOption = PTmenue.leistbarkeitsAnalyse
                    '.Show()
                    returnValue = .ShowDialog

                End With

            End If

        ElseIf ShowProjekte.Count = 0 Then

            Call MsgBox("Es sind keine Projekte geladen!  ")

        ElseIf showRangeRight - showRangeLeft < minColumns - 1 Then

            Call MsgBox("bitte zuerst einen Zeitraum markieren! ")

        End If



        appInstance.EnableEvents = True
        enableOnUpdate = True



    End Sub
    ''' <summary>
    ''' prüft ob auch kein Summary Projekt selektiert wurde ..
    ''' viele aktionen sind darauf nicht definiert 
    ''' </summary>
    ''' <param name="nameCollection"></param>
    ''' <returns></returns>
    Public Function noSummaryProjectsareSelected(ByRef nameCollection As Collection) As Boolean

        Dim tmpResult As Boolean = True
        Dim awinSelection As Excel.ShapeRange
        Dim okCollection As New Collection
        Dim tmpCollection As New Collection
        Dim pName As String

        Dim errMsg As String = ""
        If awinSettings.englishLanguage Then
            errMsg = "no summary projects allowed (Exception: Portfolio Manager)..."
        Else
            errMsg = "keine Summary Projekte zugelassen (ausser für Portfolio Manager)..."
        End If

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not awinSelection Is Nothing Then

            For i As Integer = 1 To awinSelection.Count
                pName = awinSelection.Item(i).Name

                Try
                    Dim hproj As clsProjekt = ShowProjekte.getProject(pName)
                    If hproj.projectType = ptPRPFType.portfolio And Not myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
                        tmpResult = False
                        Call MsgBox(errMsg)
                        Exit For
                    Else
                        okCollection.Add(pName)
                    End If
                Catch ex As Exception

                End Try
            Next

            If Not tmpResult Then
                okCollection.Clear()
            End If

        End If

        nameCollection = okCollection
        noSummaryProjectsareSelected = tmpResult

    End Function
    ''' <summary>
    ''' eine neue Variante anlegen 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBVarianteNeu(control As IRibbonControl)

        Dim hproj As clsProjekt

        Dim neueVariante As New frmCreateNewVariant
        Dim resultat As DialogResult
        ''Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim newproj As clsProjekt
        'Dim key As String
        Dim phaseList As New Collection
        Dim milestoneList As New Collection
        Dim neuerVariantenName As String = ""
        Dim ok As Boolean = True
        Dim zaehler As Integer = 1
        Dim nameCollection As New Collection
        Dim abbruch As Boolean = False

        Dim variantDescription As String = ""


        Call projektTafelInit()

        enableOnUpdate = False



        If control.Id = "PT2G1M1B0" Then
            ' neue Variante anlegen 
            Dim errMsg As String = ""

            If noSummaryProjectsareSelected(nameCollection) Then
                If nameCollection.Count = 0 Then
                    If awinSettings.englishLanguage Then
                        errMsg = "please select project(s ..."
                    Else
                        errMsg = "bitte mind. ein Projekt selektieren ..."
                    End If
                    Call MsgBox(errMsg)
                Else
                    While zaehler <= nameCollection.Count And Not abbruch

                        ' jetzt die Aktion durchführen ...
                        Dim pName As String = CStr(nameCollection.Item(zaehler))

                        Try
                            hproj = ShowProjekte.getProject(pName)
                            pName = hproj.name
                            phaseList = projectboardShapes.getPhaseList(hproj.name)
                            milestoneList = projectboardShapes.getMilestoneList(hproj.name)
                        Catch ex As Exception
                            Call MsgBox("Projekt " & pName & " nicht gefunden ...")
                            enableOnUpdate = True
                            Exit Sub
                        End Try

                        ' enableevents wird hier nicht false gesetzt; wenn dann wird das im Formular gemacht 
                        ' screenupdating wird hier ebenso nicht auf false gesetzt 

                        ' jetzt wird hier das Formular aufgerufen, wo eine neue Variante eingegeben werden kann 
                        With neueVariante
                            .txtDescription.Text = variantDescription
                            .projektName.Text = hproj.name
                            .variantenName.Text = hproj.variantName
                            .newVariant.Text = neuerVariantenName
                        End With

                        resultat = neueVariante.ShowDialog
                        If resultat = DialogResult.OK Then

                            With neueVariante
                                neuerVariantenName = .newVariant.Text
                                variantDescription = .txtDescription.Text
                            End With

                            newproj = hproj.createVariant(neuerVariantenName, variantDescription)
                            ' alt - wurde ersetzt durch obigen Aufruf ..
                            ''newproj = New clsProjekt
                            ''hproj.copyTo(newproj)

                            ''If newproj.dauerInDays <> hproj.dauerInDays Then
                            ''    'Call MsgBox("ungleich: " & newproj.dauerInDays & " versus " & hproj.dauerInDays)
                            ''End If

                            ''With neueVariante
                            ''    neuerVariantenName = .newVariant.Text
                            ''    variantDescription = .txtDescription.Text
                            ''End With


                            ''With newproj
                            ''    .name = hproj.name
                            ''    .variantName = neuerVariantenName
                            ''    .variantDescription = variantDescription
                            ''    .ampelErlaeuterung = hproj.ampelErlaeuterung
                            ''    .ampelStatus = hproj.ampelStatus
                            ''    .timeStamp = Date.Now
                            ''    .shpUID = hproj.shpUID
                            ''    .tfZeile = hproj.tfZeile

                            ''End With

                            If currentConstellationPvName <> calcLastSessionScenarioName() Then
                                currentConstellationPvName = calcLastSessionScenarioName()
                            End If

                            If Not AlleProjekte.hasAnyConflictsWith(calcProjektKey(newproj), False) Then
                                ' jetzt muss die bisherige Variante aus Showprojekte rausgenommen werden ..
                                ShowProjekte.Remove(hproj.name)

                                ' die neue Variante wird aufgenommen
                                'key = calcProjektKey(newproj)
                                AlleProjekte.Add(newproj, checkOnConflicts:=True)
                                ShowProjekte.Add(newproj)

                                ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                                ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                                Try

                                    Dim tmpCollection As New Collection
                                    Call ZeichneProjektinPlanTafel(tmpCollection, newproj.name, newproj.tfZeile, phaseList, milestoneList)

                                Catch ex As Exception

                                    Call MsgBox("Konflikte zw. Summary Projekt und Variante " & ex.Message)

                                End Try
                            End If


                            zaehler = zaehler + 1
                        Else
                            abbruch = True
                        End If

                    End While
                End If

            End If
        End If


        If currentConstellationPvName <> calcLastSessionScenarioName() Then
            currentConstellationPvName = calcLastSessionScenarioName()
        End If

        enableOnUpdate = True

    End Sub
    ''' <summary>
    ''' Es werden Projekte, die Varianten haben angezeigt in einem TreeView
    ''' Hier können Varianten ausgewählt werden, die gelöscht werden sollen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBVarianteLoeschen(control As IRibbonControl)

        Dim deletedProj As Integer = 0
        'Dim returnValue As DialogResult

        'Dim activateVariant As New frmDeleteProjects
        Dim deleteVariant As New frmProjPortfolioAdmin

        Try

            With deleteVariant

                .aKtionskennung = PTTvActions.deleteV

            End With

            'returnValue = activateVariant.ShowDialog
            deleteVariant.Show()

            'If returnValue = DialogResult.OK Then
            '    'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

            'Else
            '    ' returnValue = DialogResult.Cancel

            'End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try

        If currentConstellationPvName <> calcLastSessionScenarioName() Then
            currentConstellationPvName = calcLastSessionScenarioName()
        End If

    End Sub
    ''' <summary>
    ''' Projekt löschen
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBLoeschen(control As IRibbonControl)

        Dim bestaetigeLoeschen As New frmconfirmDeletePrj
        Dim singleShp As Excel.Shape
        Dim awinSelection As Excel.ShapeRange
        Dim returnValue As DialogResult
        Dim outputCollection As New Collection
        Call projektTafelInit()

        appInstance.EnableEvents = False
        enableOnUpdate = False

        Try
            'awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not IsNothing(awinSelection) Then

            If awinSettings.englishLanguage Then
                bestaetigeLoeschen.botschaft = "please confirm deleting in session ..." & vbLf &
                                            "Attention: all variants will get deleted as well ..."
            Else
                bestaetigeLoeschen.botschaft = "Bitte bestätigen Sie das Löschen" & vbLf &
                                            "Vorsicht: alle Varianten werden mitgelöscht ..."
            End If

            returnValue = bestaetigeLoeschen.ShowDialog

            If returnValue = DialogResult.Cancel Then

                appInstance.EnableEvents = True
                enableOnUpdate = True
                Exit Sub

            End If



            ' jetzt die Aktion durchführen ...


            For Each singleShp In awinSelection


                Dim shapeArt As Integer
                shapeArt = kindOfShape(singleShp)

                With singleShp
                    If isProjectType(shapeArt) Then

                        Try
                            Dim hproj As clsProjekt = ShowProjekte.getProject(.Name)
                            If Not IsNothing(hproj) Then

                                Call awinDeleteProjectInSession(pName:= .Name)
                                ' Änderung tk: bei dem Löschen in der Session soll keine Restriktion gelten;
                                ' ausserdem soll es konsistent zu Löschen aus Session über Portfolio Browser sein 
                                'If notReferencedByAnyPortfolio(hproj.name, hproj.variantName) Then
                                '    Call awinDeleteProjectInSession(pName:=.Name)
                                'Else
                                '    Dim outputline As String = "Löschen verweigert " & hproj.name & " wird in Szenarien referenziert: "
                                '    outputline = outputline & projectConstellations.getSzenarioNamesWith(hproj.name, hproj.variantName)
                                '    outputCollection.Add(outputline)
                                'End If
                            End If


                        Catch ex As Exception
                            Exit For
                        End Try

                    End If
                End With


            Next

            If currentConstellationPvName <> calcLastSessionScenarioName() Then
                currentConstellationPvName = calcLastSessionScenarioName()
            End If

            ' ein oder mehrere Projekte wurden gelöscht  - typus = 3
            Call awinNeuZeichnenDiagramme(3)

            If outputCollection.Count > 0 Then
                Call showOutPut(outputCollection, "Löschen von Projekten", "folgende Fehler sind aufgetreten:")
            End If

        Else

            Dim deletedProj As Integer = 0

            If AlleProjekte.Count = 0 Then
                If awinSettings.englishLanguage Then
                    Call MsgBox("no projects in session ...")
                Else
                    Call MsgBox("es sind keine Projekte in der Session geladen ...")
                End If

            Else

                'Dim deleteProjects As New frmDeleteProjects
                Dim deleteProjects As New frmProjPortfolioAdmin
                Try

                    With deleteProjects

                        .aKtionskennung = PTTvActions.delFromSession

                    End With

                    returnValue = deleteProjects.ShowDialog

                    If returnValue = DialogResult.OK Then

                        ' das war vorherin frmProjPortfolioAdmin, im Click 
                        Call awinNeuZeichnenDiagramme(2)

                    Else
                        ' returnValue = DialogResult.Cancel

                    End If

                Catch ex As Exception

                    Call MsgBox(ex.Message)
                End Try

            End If



        End If



        Call awinDeSelect()

        If currentConstellationPvName <> calcLastSessionScenarioName() Then
            currentConstellationPvName = calcLastSessionScenarioName()
        End If


        enableOnUpdate = True
        appInstance.EnableEvents = True

    End Sub



    ''' <summary>
    ''' lädt die gewählten Projekte und gewählten Varianten in die Session
    ''' </summary>
    ''' <param name="Control"></param>
    ''' <remarks></remarks>
    Public Sub PBBDatenbankLoadProjekte(Control As IRibbonControl)

        Dim deletedProj As Integer = 0
        Dim returnValue As DialogResult

        'Dim deleteProjects As New frmDeleteProjects
        Dim loadProjectsForm As New frmProjPortfolioAdmin

        Try

            With loadProjectsForm

                .aKtionskennung = PTTvActions.loadPV

                '' '' ''.portfolioName.Visible = False
                '' '' ''.Label1.Visible = False
            End With

            returnValue = loadProjectsForm.ShowDialog

            If returnValue = DialogResult.OK Then
                'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

                ' tk 7.10.19 das war vorher in Click-Aktion von frmProjPortfolioAdmin
                Call awinNeuZeichnenDiagramme(2)

            Else
                ' returnValue = DialogResult.Cancel

            End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try

        If currentConstellationPvName <> calcLastSessionScenarioName() Then
            currentConstellationPvName = calcLastSessionScenarioName()
        End If


    End Sub

    ''' <summary>
    ''' alle aktuell in AlleProjekte geladenen PRojekte und Varianten werden angezeigt und können 
    ''' aktiv / de-aktiv gesetzt werden 
    ''' on-the-fly werden evtl gezeigte Portfolio Charts aktualisiert
    ''' erst mit OK werden die Projekte gezeichnet und als lastConstellation gespeichert , oder unter dem angegebenen Namen
    ''' </summary>
    ''' <remarks></remarks>
    Sub PBBChangeCurrentPortfolio()

        ' verhindert, dass das mehrmals aufgerufen wird, da das Formular im nicht-modalen Modus aufgerufen wird 
        If Not awinSettings.isChangePortfolioFrmActive Then

            Call activateProjectBoard()

            Dim changePortfolio As New frmProjPortfolioAdmin

            Call awinDeSelect(True)

            If AlleProjekte.Count > 0 Then
                ' das letzte Portfolio speichern 
                'Call storeSessionConstellation("Last")

                Try

                    With changePortfolio

                        .aKtionskennung = PTTvActions.chgInSession

                    End With

                    'Call awinClearPlanTafel()

                    changePortfolio.Show()

                    ' diese Variable zeigt an, dass das Formular zu Bearbeiten des Portfolios bereits aktiv ist
                    awinSettings.isChangePortfolioFrmActive = True

                Catch ex As Exception

                    Call MsgBox(ex.Message)
                End Try
            Else
                If awinSettings.englishLanguage Then
                    Call MsgBox("no projects loaded ...")
                Else
                    Call MsgBox("keine Projekte geladen ...")
                End If

            End If
        Else

            ' das Formular zum Bearbeiten des Portfolios nicht erneut anzeigen

        End If



    End Sub
    ''' <summary>
    ''' löscht die ausgewählten Projekte aus der Datenbank 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBDeleteProjectsInDB(control As IRibbonControl)


        Dim deletedProj As Integer = 0
        Dim returnValue As DialogResult

        'Dim deleteProjects As New frmDeleteProjects
        Dim deleteProjects As New frmProjPortfolioAdmin

        Try

            With deleteProjects

                If control.Id = "Pt5G3B4" Then
                    .aKtionskennung = PTTvActions.delAllExceptFromDB
                ElseIf control.Id = "Pt5G3B3" Then
                    .aKtionskennung = PTTvActions.delFromDB
                End If

            End With

            returnValue = deleteProjects.ShowDialog

            ' die Operation ist bereits ausgeführt - deswegen muss hier nichts mehr unterschieden werden 

            If returnValue = DialogResult.OK Then
                ' everything is done ... 

            Else
                ' everything is done ... 

            End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try



    End Sub

    Sub PBBWriteProtections(ByVal control As IRibbonControl, ByVal setFlag As Boolean)



        Dim writeProtectProjects As New frmProjPortfolioAdmin
        If AlleProjekte.Count > 0 Then

            Try

                With writeProtectProjects
                    .aKtionskennung = PTTvActions.setWriteProtection
                End With

                writeProtectProjects.Show()

            Catch ex As Exception

                Call MsgBox(ex.Message)
            End Try

        Else
            If awinSettings.englishLanguage Then
                Call MsgBox("please load projects/variants first ...")
            Else
                Call MsgBox("bitte erst Projekte / Varianten laden ...")
            End If
        End If


    End Sub
    ''' <summary>
    ''' aktiviert die selektierte Variante 
    ''' </summary>
    ''' <param name="control"></param>
    ''' <remarks></remarks>
    Sub PBBVarianteAktiv(ByVal control As IRibbonControl)

        Dim deletedProj As Integer = 0
        'Dim returnValue As DialogResult

        'Dim activateVariant As New frmDeleteProjects
        Dim activateVariant As New frmProjPortfolioAdmin

        Try

            With activateVariant

                .aKtionskennung = PTTvActions.activateV

            End With

            'returnValue = activateVariant.ShowDialog
            activateVariant.Show()

            'If returnValue = DialogResult.OK Then
            '    'deletedProj = RemoveSelectedProjectsfromDB(deleteProjects.selectedItems)    ' es werden die selektierten Projekte in der DB gespeichert, die Anzahl gespeicherter Projekte sind das Ergebnis

            'Else
            '    ' returnValue = DialogResult.Cancel

            'End If

        Catch ex As Exception

            Call MsgBox(ex.Message)
        End Try

        If currentConstellationPvName <> calcLastSessionScenarioName() Then
            currentConstellationPvName = calcLastSessionScenarioName()
        End If

    End Sub

    Sub PBBShowTimeMachine(control As IRibbonControl)

        Dim err As New clsErrorCodeMsg

        Dim hproj As clsProjekt
        Dim pName As String, variantName As String
        Dim vglName As String = " "
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As Excel.Shape
        Dim showCharacteristics As New frmShowProjCharacteristics
        'Dim returnValue As DialogResult
        Dim awinSelection As Excel.ShapeRange
        Dim grueneAmpel As String = awinPath & "gruen.gif"
        Dim gelbeAmpel As String = awinPath & "gelb.gif"
        Dim roteAmpel As String = awinPath & "rot.gif"
        Dim graueAmpel As String = awinPath & "grau.gif"

        If timeMachineIsOn Then
            Call MsgBox("bitte erst Time Machine beenden ...")
            Exit Sub
        End If

        Call projektTafelInit()

        enableOnUpdate = False
        appInstance.EnableEvents = True


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
                    'Try
                    '    variantName = .variantName.Trim
                    'Catch ex As Exception
                    '    variantName = ""
                    'End Try

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
                        ' projekthistorie muss nur dann neu geladen werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName,
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now, err:=err)
                        If projekthistorie.Count <> 0 Then

                            projekthistorie.Add(Date.Now, hproj)

                        End If

                    Else
                        Call MsgBox("Datenbank-Verbindung ist unterbrochen")
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

                    With showCharacteristics

                        .Text = "Historie für Projekt " & pName.Trim & vbLf &
                                "( " & projekthistorie.getZeitraum & " )"
                        .timeSlider.Minimum = 0
                        .timeSlider.Maximum = nrSnapshots - 1

                        '.ampelErlaeuterung.Text = kvp.Value.ampelErlaeuterung

                        'If kvp.Value.ampelStatus = 1 Then
                        '    .ampelPicture.LoadAsync(grueneAmpel)
                        'ElseIf kvp.Value.ampelStatus = 2 Then
                        '    .ampelPicture.LoadAsync(gelbeAmpel)
                        'ElseIf kvp.Value.ampelStatus = 3 Then
                        '    .ampelPicture.LoadAsync(roteAmpel)
                        'Else
                        '    .ampelPicture.LoadAsync(graueAmpel)
                        'End If

                        '.snapshotDate.Text = kvp.Value.timeStamp.ToString
                        ' das ist ja der aktuelle Stand ..
                        .snapshotDate.Text = "Aktueller Stand"
                        ' Designer 
                        'Dim zE As String = "(" & awinSettings.zeitEinheit & ")"
                        '.engpass1.Text = "Designer:          " & kvp.Value.getRessourcenBedarf(3).Sum.ToString("###.#") & zE
                        '.engpass2.Text = "Personalkosten: " & kvp.Value.getAllPersonalKosten.Sum.ToString("###.#") & " (T€)"
                        '.engpass3.Text = "Sonstige Kosten:   " & kvp.Value.getGesamtAndereKosten.Sum.ToString("###.#") & " (T€)"


                    End With


                    ' jetzt wird das Form aufgerufen ... 

                    'returnValue = showCharacteristics.ShowDialog
                    showCharacteristics.Show()

                Else
                    Call MsgBox("es gibt noch keine Planungs-Historie")
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


    End Sub
End Module
