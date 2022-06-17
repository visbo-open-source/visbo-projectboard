Imports ProjectBoardDefinitions
Imports System.Text.RegularExpressions
Imports DBAccLayer
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports System.Runtime.Serialization
Imports System.Xml
Imports System.Xml.Serialization

Public Module agm4

    Private Enum ptTL
        Pname = 0
        PhaseName = 1
        Who = 2
        What = 3
        Prio = 4
        StartDate = 5
        EndDate = 6
        TotalEffort = 7
        ForecastEffort = 8
        CommentTxt = 9
    End Enum

    Private Enum ptMMX
        pstart = 0
        pEnd = 1
        phStart = 2
        phEnd = 3
        tpstart = 4
        tpend = 5
        tphstart = 6
        tphend = 7
    End Enum

    Public Sub ImportTaskLists(ByVal offlineName As String, ByRef outputCollection As Collection, ByVal modifyDates As Boolean)
        '
        ' check validity of filestructure

        '
        ' first iteration: get all projects , create a tempVariant from it ; define min/max dates of projects / phases; write all projects first need to be created


        '
        ' then: check, if dates are ok, if not then correct dates of projects, phases 

        '
        ' second iteration: for each line in Excel add the according effort value 


        ' now start 


        ' check validity of filestructure
        Dim lastRow As Integer = -1
        Dim updatedProjects As Integer = 0

        Dim logF_Fehler As Integer = 0
        ' nimmt die Texte für die LogFile Zeile auf
        ' Array kann beliebig lang werden 
        Dim logArray() As String
        'Dim logDblArray() As Double




        ' im Key steht der Projekt-Name, im Value steht eine sortierte Liste mit key=Rollen-Name, values die Ist-Werte
        Dim identifiedProjects As New clsProjekte
        Dim pName As String = ""
        Dim phaseName As String = ""
        Dim who As String = ""
        Dim what As String = ""
        Dim prio As String = ""
        Dim totalEffort As Double = 0.0
        Dim forecastEffort As Double = 0.0
        Dim startDate As Date = Date.MinValue
        Dim endDate As Date = Date.MinValue
        Dim commentTxt As String = ""

        Dim heute As Date = Date.Now




        Try

            ' jetzt kommt die eigentliche Import Behandlung 
            Dim currentWS As Excel.Worksheet = Nothing
            Try
                ' just consider first worksheet as the relevant worksheet
                currentWS = CType(appInstance.ActiveWorkbook.Worksheets(1),
                                                           Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                logmessage = "Keine Tabelle mit Namen 'Istdaten gefunden' ... Abbruch"
                outputCollection.Add(logmessage)
                Exit Sub
            End Try


            ' tk, 2.8.2018 Behandlung LookupTable 
            Dim lookUpTableWS As Excel.Worksheet = Nothing

            ' die lookupTable nimmt die Projekt-Nummer als KEy auf und den korrespondierenden NAmen aus der Rupi-Liste
            ' bei Aufbau der lookupTable werden die Rupi-Liste NAmen bereits in valide Namen gewandelt ... 
            Dim lookupTable As SortedList(Of String, String) = Nothing

            Try
                lookUpTableWS = CType(appInstance.ActiveWorkbook.Worksheets("lookupTable"),
                                                           Global.Microsoft.Office.Interop.Excel.Worksheet)
            Catch ex As Exception
                lookUpTableWS = Nothing
            End Try

            ' wenn jetzt eine Tabelle vorhanden ist, dann muss die LookupTable aufgebaut werden 
            If Not IsNothing(lookUpTableWS) Then

                With lookUpTableWS

                    Dim lupTLastZeile As Integer = CType(.Cells(60000, "B"), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
                    Dim lupTZeile As Integer = 2
                    If lupTLastZeile >= lupTZeile Then

                        lookupTable = New SortedList(Of String, String)

                        While lupTZeile <= lupTLastZeile
                            Try
                                Dim pNameInThis As String = CStr(CType(.Cells(lupTZeile, 1), Excel.Range).Value).Trim
                                Dim betterName As String = CStr(CType(.Cells(lupTZeile, 2), Excel.Range).Value).Trim

                                If Not isValidPVName(betterName) Then
                                    betterName = makeValidProjectName(betterName)
                                End If

                                If pNameInThis <> "" Then
                                    If Not lookupTable.ContainsKey(pNameInThis) Then
                                        lookupTable.Add(pNameInThis, betterName)
                                    End If
                                Else
                                    If Not lookupTable.ContainsKey("0") Then
                                        lookupTable.Add("0", betterName)
                                    End If
                                End If

                            Catch ex As Exception

                            End Try

                            lupTZeile = lupTZeile + 1

                        End While

                    End If
                End With


            End If

            Dim lookupsExist As Boolean = False
            If Not IsNothing(lookupTable) Then
                lookupsExist = (lookupTable.Count > 0)
            End If


            ' hat die Datei die richtige Header-Struktur ? 
            Dim firstZeile As Excel.Range = currentWS.Rows(1)

            If Not isCorrecttaskListStructure(currentWS, lastRow) Then
                logmessage = "files can not be recognized as task lists!"
                outputCollection.Add(logmessage)
                Exit Sub
            End If


            With currentWS


                ' welche Werte sollen ausgelesen werden, wo stehen die 
                Dim colParams(9) As Integer

                colParams(ptTL.Pname) = CType(.Range("A1"), Excel.Range).Column
                colParams(ptTL.PhaseName) = CType(.Range("B1"), Excel.Range).Column
                colParams(ptTL.Who) = CType(.Range("C1"), Excel.Range).Column
                colParams(ptTL.What) = CType(.Range("D1"), Excel.Range).Column
                colParams(ptTL.Prio) = CType(.Range("E1"), Excel.Range).Column
                colParams(ptTL.StartDate) = CType(.Range("F1"), Excel.Range).Column
                colParams(ptTL.EndDate) = CType(.Range("G1"), Excel.Range).Column
                colParams(ptTL.TotalEffort) = CType(.Range("H1"), Excel.Range).Column
                colParams(ptTL.ForecastEffort) = CType(.Range("I1"), Excel.Range).Column
                colParams(ptTL.CommentTxt) = CType(.Range("J1"), Excel.Range).Column


                ' im key steht der Name aus der Datei , im Value steht der Name aus AlleProjekte 
                Dim handledNames As New SortedList(Of String, String)
                ' nimmt die unbekannten / nicht erkannten Role-Names auf 
                Dim unKnownRoleNames As New SortedList(Of String, Boolean)

                ' tk 2.1.2021 Vorab Schleife 
                ' 1. Schleife find out which project has actualdata from ... to ... 
                ' und die dazugehörigen Min - Max Dates Integer = columnOfDates
                Dim MinMaxInformations As New SortedList(Of String, Integer())


                Dim minValue As Integer = 1000000000
                Dim maxValue As Integer = 0

                Dim zeile As Integer = 2

                While zeile <= lastRow
                    Try

                        Call readParametersFromRow(zeile, lookupTable, colParams,
                                                   pName, phaseName, who, what, prio,
                                                   startDate, endDate,
                                                   totalEffort, forecastEffort,
                                                   commentTxt)

                        If pName.Length > 0 Then
                            ' does the project exist ? 

                            If Not identifiedProjects.contains(pName) Then

                                Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, "", AlleProjekte, heute)
                                If Not IsNothing(hproj) Then
                                    identifiedProjects.Add(hproj, updateCurrentConstellation:=False)

                                    ' now set MinMaxInformations
                                    Dim skey As String = hproj.name & "[" & phaseName
                                    Dim startCol As Integer = getColumnOfDate(hproj.startDate)
                                    Dim endCol As Integer = getColumnOfDate(hproj.endeDate)
                                    Dim startPhCol As Integer = startCol
                                    Dim endPhCol As Integer = endCol

                                    If phaseName <> "." Then
                                        Dim cPhase As clsPhase = hproj.getPhase(phaseName)
                                        If Not IsNothing(cPhase) Then
                                            startPhCol = startCol + cPhase.relStart - 1
                                            endPhCol = startCol + cPhase.relEnde - 1
                                        End If
                                    End If

                                    Dim myCols As Integer()
                                    ReDim myCols(7)


                                End If
                            Else
                                ' set the 
                            End If

                        End If

                    Catch ex As Exception

                    End Try

                    zeile = zeile + 1

                End While



            End With


        Catch ex As Exception
            ReDim logArray(1)
            logArray(0) = "Exception aufgetreten 100457: "
            logArray(1) = ex.Message
            Call logger(ptErrLevel.logError, "ImportAllianzIstdaten", logArray)
            Throw New Exception("Fehler in Import-Datei Typ 3" & ex.Message)
        End Try


    End Sub

    Private Function isCorrecttaskListStructure(ByVal ws As Excel.Worksheet, ByRef lastRow As Integer) As Boolean

        Dim tmpResult As Boolean = False
        lastRow = 1

        Dim headerCheck() As String = {"Projekt", "Phase", "Wer", "Was", "Prio", "Start", "Ende", "total", "still-to-do", "Comment"}
        Dim colCheck() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10}

        For i As Integer = 1 To headerCheck.Length
            Dim tmpLastRow As Integer = CType(ws.Cells(60000, i), Global.Microsoft.Office.Interop.Excel.Range).End(XlDirection.xlUp).Row
            lastRow = System.Math.Max(lastRow, tmpLastRow)
        Next

        Try
            tmpResult = True ' initiale Vorbesetzung 
            Dim ix As Integer = 0
            Do While tmpResult = True And ix <= headerCheck.Length - 1
                tmpResult = tmpResult And CStr(CType(ws.Cells(1, colCheck(ix)), Excel.Range).Value).Contains(headerCheck(ix))
                ix = ix + 1
            Loop

        Catch ex As Exception
            tmpResult = False
        End Try

        isCorrecttaskListStructure = tmpResult

    End Function

    ''' <summary>
    ''' reads the according parameters from one row of the tasklist File
    ''' </summary>
    ''' <param name="zeile"></param>
    ''' <param name="colValue"></param>
    ''' <param name="pName"></param>
    ''' <param name="phaseName"></param>
    ''' <param name="who"></param>
    ''' <param name="what"></param>
    ''' <param name="prio"></param>
    ''' <param name="startDate"></param>
    ''' <param name="endDate"></param>
    ''' <param name="totalEffort"></param>
    ''' <param name="forecastEffort"></param>
    ''' <param name="commentTxt"></param>
    Private Sub readParametersFromRow(ByVal zeile As Integer,
                                      ByVal lookUptable As SortedList(Of String, String),
                                      ByVal colValue As Integer(),
                                      ByRef pName As String,
                                      ByRef phaseName As String,
                                      ByRef who As String,
                                      ByRef what As String,
                                      ByRef prio As String,
                                      ByRef startDate As Date,
                                      ByRef endDate As Date,
                                      ByRef totalEffort As Double,
                                      ByRef forecastEffort As Double,
                                      ByRef commentTxt As String)

        Try
            ' just consider first worksheet as the relevant worksheet
            Dim ws As Excel.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets(1),
                                                           Global.Microsoft.Office.Interop.Excel.Worksheet)

            If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.Pname)), Excel.Range).Value) Then
                pName = CStr(CType(ws.Cells(zeile, colValue(ptTL.Pname)), Excel.Range).Value).Trim
            Else
                pName = ""
            End If

            ' is there a lookUptable defined ? 
            If Not IsNothing(lookUptable) Then
                If lookUptable.Count > 0 Then
                    Dim searchName As String = pName
                    If pName = "" Then
                        searchName = "-"
                    End If
                    If lookUptable.ContainsKey(searchName) Then
                        pName = lookUptable.Item(searchName)
                    End If
                End If
            End If

            If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.PhaseName)), Excel.Range).Value) Then
                phaseName = CStr(CType(ws.Cells(zeile, colValue(ptTL.PhaseName)), Excel.Range).Value).Trim
            Else
                phaseName = ""
            End If

            If phaseName = "" Then
                phaseName = "."
            End If

            If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.Who)), Excel.Range).Value) Then
                who = CStr(CType(ws.Cells(zeile, colValue(ptTL.Who)), Excel.Range).Value).Trim
            Else
                who = ""
            End If

            If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.What)), Excel.Range).Value) Then
                what = CStr(CType(ws.Cells(zeile, colValue(ptTL.What)), Excel.Range).Value).Trim
            Else
                what = ""
            End If

            If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.Prio)), Excel.Range).Value) Then
                prio = CStr(CType(ws.Cells(zeile, colValue(ptTL.Prio)), Excel.Range).Value).Trim
            Else
                prio = ""
            End If

            Try
                If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.StartDate)), Excel.Range).Value) Then
                    startDate = CDate(CType(ws.Cells(zeile, colValue(ptTL.StartDate)), Excel.Range).Value)
                Else
                    startDate = Date.MinValue
                End If
            Catch ex As Exception
                startDate = Date.MinValue
            End Try

            Try
                If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.EndDate)), Excel.Range).Value) Then
                    endDate = CDate(CType(ws.Cells(zeile, colValue(ptTL.EndDate)), Excel.Range).Value)
                Else
                    endDate = Date.MinValue
                End If
            Catch ex As Exception
                endDate = Date.MinValue
            End Try

            Try
                If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.TotalEffort)), Excel.Range).Value) Then
                    totalEffort = System.Math.Max(0, CDbl(CType(ws.Cells(zeile, colValue(ptTL.TotalEffort)), Excel.Range).Value))
                Else
                    totalEffort = 0.0
                End If
            Catch ex As Exception
                totalEffort = 0.0
            End Try

            Try
                If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.ForecastEffort)), Excel.Range).Value) Then
                    forecastEffort = System.Math.Max(0, CDbl(CType(ws.Cells(zeile, colValue(ptTL.ForecastEffort)), Excel.Range).Value))
                    totalEffort = System.Math.Max(totalEffort, forecastEffort)
                Else
                    forecastEffort = 0.0
                End If
            Catch ex As Exception
                forecastEffort = 0.0
            End Try

            If Not IsNothing(CType(ws.Cells(zeile, colValue(ptTL.CommentTxt)), Excel.Range).Value) Then
                commentTxt = CStr(CType(ws.Cells(zeile, colValue(ptTL.CommentTxt)), Excel.Range).Value).Trim
            Else
                commentTxt = ""
            End If

        Catch ex As Exception
            Call MsgBox("in readParametersFromRow - Kein Worksheet(1) gefunden ...")

            Exit Sub
        End Try


    End Sub



    ''' <summary>
    ''' sorts the currentSessionConstellation according BU, Start-Date, End-Date
    ''' if name ist provided then Constellation gets that name
    ''' </summary>
    ''' <param name="name"></param>
    Public Sub sortCurrentConst(ByVal name As String, ByVal vName As String)

        Dim sortType As Integer = ptSortCriteria.buStartName

        ' Vorbesetzungen 
        Dim formerEE As Boolean = appInstance.EnableEvents
        Dim formerSU As Boolean = appInstance.ScreenUpdating

        appInstance.EnableEvents = False
        enableOnUpdate = False


        Try

            If Not IsNothing(currentSessionConstellation) Then
                If currentSessionConstellation.Liste.Count <> 0 Then

                    Dim currentSortConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True, cName:="Sort Result")

                    If currentSortConstellation.sortCriteria <> sortType Then
                        appInstance.ScreenUpdating = False
                        Try
                            ' nur dann muss was gemacht werden ...  
                            currentSortConstellation.sortCriteria = sortType

                            If name <> "" Then
                                currentSortConstellation.constellationName = name
                                currentSortConstellation.variantName = vName
                            End If

                            Dim tmpConstellation As New clsConstellations
                            tmpConstellation.Add(currentSortConstellation)

                            ' es in der Session Liste verfügbar machen
                            If projectConstellations.Contains(currentSortConstellation.constellationName) Then
                                projectConstellations.Remove(currentSortConstellation.constellationName)
                            End If

                            projectConstellations.Add(currentSortConstellation)

                            Call showConstellations(constellationsToShow:=tmpConstellation,
                                                    clearBoard:=True, clearSession:=False, storedAtOrBefore:=Date.Now)


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

        appInstance.EnableEvents = formerEE
        enableOnUpdate = True
    End Sub

    ''' <summary>
    ''' reading the VCSetting "customization" if stored in the actual VC
    ''' </summary>
    ''' <returns></returns>
    Public Function readCustomizations(Optional ByVal customizations As clsCustomization = Nothing) As Date

        Dim result As Date = Date.MinValue
        Dim err As New clsErrorCodeMsg
        Dim msgTxt As String

        If IsNothing(customizations) Then
            '
            ' Read Customizations 
            customizations = CType(databaseAcc, DBAccLayer.Request).retrieveCustomizationFromDB("", Date.Now, False, err)
        End If


        If Not IsNothing(customizations) Then

            StartofCalendar = customizations.kalenderStart
            Call logger(ptErrLevel.logInfo, "readCustomizations", " StartOfCalendar: " & StartofCalendar.ToString)

            businessUnitDefinitions = customizations.businessUnitDefinitions

            PhaseDefinitions = customizations.phaseDefinitions

            MilestoneDefinitions = customizations.milestoneDefinitions

            showtimezone_color = customizations.showtimezone_color
            noshowtimezone_color = customizations.noshowtimezone_color
            calendarFontColor = customizations.calendarFontColor
            nrOfDaysMonth = customizations.nrOfDaysMonth
            farbeInternOP = customizations.farbeInternOP
            farbeExterne = customizations.farbeExterne
            iProjektFarbe = customizations.iProjektFarbe
            iWertFarbe = customizations.iWertFarbe
            vergleichsfarbe0 = customizations.vergleichsfarbe0
            vergleichsfarbe1 = customizations.vergleichsfarbe1

            awinSettings.SollIstFarbeB = customizations.SollIstFarbeB
            awinSettings.SollIstFarbeL = customizations.SollIstFarbeL
            awinSettings.SollIstFarbeC = customizations.SollIstFarbeC
            awinSettings.AmpelGruen = customizations.AmpelGruen

            awinSettings.AmpelGelb = customizations.AmpelGelb
            awinSettings.AmpelRot = customizations.AmpelRot
            awinSettings.AmpelNichtBewertet = customizations.AmpelNichtBewertet
            awinSettings.glowColor = customizations.glowColor

            awinSettings.timeSpanColor = customizations.timeSpanColor
            awinSettings.showTimeSpanInPT = customizations.showTimeSpanInPT

            awinSettings.gridLineColor = customizations.gridLineColor

            awinSettings.missingDefinitionColor = customizations.missingDefinitionColor

            awinSettings.ActualdataOrgaUnits = customizations.allianzIstDatenReferate
            awinSettings.ActualdataOrgaUnits = customizations.isActualDataRelevant

            awinSettings.onePersonOneRole = customizations.onePersonOneRole
            awinSettings.autoSetActualDataDate = customizations.autoSetActualDataDate

            awinSettings.actualDataMonth = customizations.actualDataMonth
            ergebnisfarbe1 = customizations.ergebnisfarbe1
            ergebnisfarbe2 = customizations.ergebnisfarbe2
            weightStrategicFit = customizations.weightStrategicFit
            awinSettings.kalenderStart = customizations.kalenderStart
            awinSettings.zeitEinheit = customizations.zeitEinheit
            awinSettings.kapaEinheit = customizations.kapaEinheit
            awinSettings.offsetEinheit = customizations.offsetEinheit
            awinSettings.EinzelRessExport = customizations.EinzelRessExport
            awinSettings.zeilenhoehe1 = customizations.zeilenhoehe1
            awinSettings.zeilenhoehe2 = customizations.zeilenhoehe2
            awinSettings.spaltenbreite = customizations.spaltenbreite
            awinSettings.autoCorrectBedarfe = customizations.autoCorrectBedarfe
            awinSettings.propAnpassRess = customizations.propAnpassRess
            awinSettings.showValuesOfSelected = customizations.showValuesOfSelected

            awinSettings.enableInvoices = customizations.enableInvoices
            awinSettings.noNewCalculation = customizations.noNewCalculation

            awinSettings.mppProjectsWithNoMPmayPass = customizations.mppProjectsWithNoMPmayPass
            awinSettings.fullProtocol = customizations.fullProtocol
            awinSettings.addMissingPhaseMilestoneDef = customizations.addMissingPhaseMilestoneDef
            awinSettings.alwaysAcceptTemplateNames = customizations.alwaysAcceptTemplateNames
            awinSettings.eliminateDuplicates = customizations.eliminateDuplicates
            awinSettings.importUnknownNames = customizations.importUnknownNames
            awinSettings.createUniqueSiblingNames = customizations.createUniqueSiblingNames

            awinSettings.readWriteMissingDefinitions = customizations.readWriteMissingDefinitions
            awinSettings.meExtendedColumnsView = customizations.meExtendedColumnsView
            awinSettings.meDontAskWhenAutoReduce = customizations.meDontAskWhenAutoReduce
            awinSettings.readCostRolesFromDB = customizations.readCostRolesFromDB

            awinSettings.importTyp = customizations.importTyp

            awinSettings.meAuslastungIsInclExt = customizations.meAuslastungIsInclExt

            awinSettings.englishLanguage = customizations.englishLanguage

            awinSettings.showPlaceholderAndAssigned = customizations.showPlaceholderAndAssigned
            awinSettings.considerRiskFee = customizations.considerRiskFee

            StartofCalendar = awinSettings.kalenderStart

            historicDate = StartofCalendar
            Try
                If awinSettings.englishLanguage Then
                    menuCult = ReportLang(PTSprache.englisch)
                    repCult = menuCult
                    awinSettings.kapaEinheit = "PD"
                Else
                    awinSettings.kapaEinheit = "PT"
                    menuCult = ReportLang(PTSprache.deutsch)
                    repCult = menuCult
                End If
            Catch ex As Exception
                awinSettings.englishLanguage = False
                awinSettings.kapaEinheit = "PT"
                menuCult = ReportLang(PTSprache.deutsch)
                repCult = menuCult
            End Try
            result = Date.Now
        Else
            msgTxt = "No customization in VISBO"
            Call logger(ptErrLevel.logWarning, "readCustomizations", msgTxt)
            result = Date.MinValue
        End If
        readCustomizations = result
    End Function


    ''' <summary>
    ''' gets the newest Organisation from now
    ''' </summary>
    ''' <returns>date of last reading</returns>
    Public Function readOrganisations() As Date

        Dim result As Date = Date.MinValue
        Dim err As New clsErrorCodeMsg
        Dim msgTxt As String
        'Read Organisation

        Dim currentOrga As clsOrganisation = CType(databaseAcc, DBAccLayer.Request).retrieveTSOrgaFromDB("organisation", Date.Now, err, False, True, True)

        ' ur: old ReSt-Call
        'Dim currentOrga As clsOrganisation = CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", Date.Now, False, Err)

        If Not IsNothing(currentOrga) Then
            If currentOrga.count > 0 Then

                If currentOrga.count > 0 Then
                    validOrganisations.addOrga(currentOrga)
                End If

                CostDefinitions = currentOrga.allCosts
                RoleDefinitions = currentOrga.allRoles

                Dim tmpActDataString As String = currentOrga.allRoles.getActualdataOrgaUnits
                If tmpActDataString = "" And awinSettings.ActualdataOrgaUnits <> "" Then
                    ' do nothing, leave it as is 
                Else
                    awinSettings.ActualdataOrgaUnits = tmpActDataString
                End If
                result = Date.Now

            Else
                msgTxt = "No organisation in VISBO"
                Call logger(ptErrLevel.logError, "readOrganisations", msgTxt)
                result = Date.MinValue

            End If
        Else
            msgTxt = "No organisation in VISBO"
            Call logger(ptErrLevel.logError, "readOrganisations", msgTxt)
            result = Date.MinValue

        End If

        readOrganisations = result

    End Function


    ''' <summary>
    ''' Read the projectTemplates from the actual VisboCenter 
    ''' </summary>
    ''' <returns></returns>
    Public Function readProjectTemplates() As Date

        Dim result As Date = Date.MinValue
        Dim err As New clsErrorCodeMsg


        ' lesen der templates des akt. VC
        Dim projectTemplates As clsProjekteAlle = CType(databaseAcc, DBAccLayer.Request).retrieveProjectTemplatesFromDB(err)

        If err.errorCode = 200 Then

            Dim projVorlage As clsProjektvorlage
            For Each kvp As KeyValuePair(Of String, clsProjekt) In projectTemplates.liste

                projVorlage = createTemplateOfProject(kvp.Value)
                If Not IsNothing(projVorlage) Then
                    ' hiermit wird die _Dauer gesetzt
                    Dim vorlagenDauer = projVorlage.dauerInDays

                    Projektvorlagen.Add(projVorlage)

                Else
                    Call logger(ptErrLevel.logError, "readProjectTemplates", "Creating a project template fromm project " & kvp.Value.name & " crashed")
                    result = Date.MinValue
                End If
            Next
            If projectTemplates.liste.Count > 0 Then
                If projectTemplates.liste.Count = Projektvorlagen.Count Then
                    result = Date.Now
                End If
            Else
                Call logger(ptErrLevel.logWarning, "readProjectTemplates", "No project templates in this VC: " & awinSettings.databaseName)
                result = Date.MinValue
            End If

        Else
            Call logger(ptErrLevel.logWarning, "readProjectTemplates", "Getting project templates from Server finished with warning: " & err.errorMsg)
            result = Date.MinValue
        End If

        readProjectTemplates = result

    End Function

End Module
