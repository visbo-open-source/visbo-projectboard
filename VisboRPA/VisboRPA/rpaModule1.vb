Imports xlns = Microsoft.Office.Interop.Excel

Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Newtonsoft.Json
Imports System.IO
Imports DBAccLayer
Imports WebServerAcc
Imports System.Security.Principal

Imports System.Diagnostics
Module rpaModule1

    ' Name des TempFiles
    Public tempFile As String = My.Computer.FileSystem.GetTempFileName()

    Public myActivePortfolio As String
    Public myVC As String
    Public inputvalues As clsRPASetting

    Public rpaPath As String
    Public swPath As String

    Public errMsgCode As clsErrorCodeMsg
    Public msgTxt As String
    Public errMessages As Collection

    Public completedOK As Boolean = False
    Public result As Boolean = False

    Public rpaFolder As String
    Public successFolder As String
    Public failureFolder As String
    Public collectFolder As String
    Public logfileFolder As String
    Public unknownFolder As String
    Public settingsFolder As String
    Public settingJsonFile As String

    Public lastReadingCustomization As Date = Date.MinValue
    Public lastReadingOrganisation As Date = Date.MinValue
    Public lastReadingCustomFields As Date = Date.MinValue
    Public lastReadingProjectTemplates As Date = Date.MinValue

    Public watchDialog As VisboRPAStart

    Public Sub Main()
        ' reads the VISBO RPA folder und treats each file it finds there appropriately
        ' in most cases new project and portfolio versions will be written 
        ' suggestions for Team Members will follow 
        ' automation in resource And team allocation will follow

        Dim actDir = My.Computer.FileSystem.CurrentDirectory

        'Call MsgBox("TempFile:" & tempFile)

        ' name des aktuell laufenden Clients
        visboClient = "VISBO RPA /"

        logfileNamePath = tempFile

        Call logger(ptErrLevel.logInfo, "VisboRPA_Main", "Before Test if VisboRPA.exe is already running")

        ' check if the VisboRPA is already running
        If IsProcessRunning("VisboRPA") Then
            Call MsgBox("VisboRPA is already running")
            Exit Sub
        End If

        ' to reset all settings to the beginning
        'My.Settings.Reset()
        Call logger(ptErrLevel.logInfo, "VisboRPA_Main", "Before reading my.settings...")

        ' Zugriff zu Daten über den VisboServer
        awinSettings.visboServer = True
        ' default Plattform
        awinSettings.databaseURL = My.Settings.VisboURL
        ' default VisboCenter
        awinSettings.databaseName = My.Settings.VisboCenter
        ' user password merken
        ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
        awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        If My.Settings.rememberUserPWD Then
            awinSettings.userNamePWD = My.Settings.userNamePWD
        Else
            awinSettings.userNamePWD = ""
        End If
        ' proxy Server URL
        awinSettings.proxyURL = My.Settings.proxyURL
        ' Default Path für RPA
        rpaPath = My.Settings.rpaPath
        ' Default Portfolio zu verwenden
        myActivePortfolio = My.Settings.activePortfolio

        myVC = My.Settings.VisboCenter

        Call logger(ptErrLevel.logInfo, "VisboRPA_Main", "default RPA-path will be set to '" & rpaPath & "'")

        Dim defaultPath = rpaPath
        'Dim defaultPath = "C:\VISBO\VISBO Config Data\RPA"

        Call logger(ptErrLevel.logInfo, "VisboRPA_Main", "Init the ReSt-Server/database Access - Schnittstelle")
        'Init the ReSt-Server/database Access - Schnittstelle
        If IsNothing(databaseAcc) Then
            databaseAcc = New DBAccLayer.Request
        End If

        Call logger(ptErrLevel.logInfo, "VisboRPA_Main", "Excel-application mit Parameter festlegen ")
        ' Parameter für die Excel Instance festlegen
        appInstance = New xlns.Application
        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        appInstance.Visible = False
        appInstance.DisplayAlerts = False

        'rpaPath not yet defined, therefore the defaultPath is used
        If rpaPath = "" Then
            rpaPath = defaultPath
        End If

        ' create DefaultDirectories if they are not exist 
        If Not My.Computer.FileSystem.DirectoryExists(rpaPath) Then
            My.Computer.FileSystem.CreateDirectory(rpaPath)
        End If

        rpaFolder = rpaPath
        If Not My.Computer.FileSystem.DirectoryExists(rpaFolder) Then
            My.Computer.FileSystem.CreateDirectory(rpaFolder)
        End If

        ' create Formula for Input of other RPA-Folder
        watchDialog = New VisboRPAStart




        Dim err As New clsErrorCodeMsg
        noDB = False

        Call logger(ptErrLevel.logInfo, "VisboRPA_Main", "check if the user/pwd is remembered")

        If awinSettings.userNamePWD <> "" Then

            Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)

            dbUsername = visboCrypto.getUserNameFromCipher(awinSettings.userNamePWD)
            dbPasswort = visboCrypto.getPwdFromCipher(awinSettings.userNamePWD)
            

            If IsNothing(awinSettings.VCid) Then
                awinSettings.VCid = ""
            End If

            If IsNothing(databaseAcc) Then
                databaseAcc = New DBAccLayer.Request
            End If

            'ur:08022022: soll mit Default erfragt werden
            'Try
            '    loginErfolgreich = CType(databaseAcc, DBAccLayer.Request).login(awinSettings.databaseURL, awinSettings.databaseName, awinSettings.VCid, dbUsername, dbPasswort, err)
            'Catch ex As Exception
            '    loginErfolgreich = False
            'End Try

            If Not loginErfolgreich Then

                Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "Login for starting")
                loginErfolgreich = logInToMongoDB(True)

            End If

        Else
            loginErfolgreich = logInToMongoDB(True)
        End If


        If loginErfolgreich Then

            Call logger(ptErrLevel.logInfo, "VisboRPA_Main", "login successfull")

            ' FileNamen für logging zusammenbauen
            logfileNamePath = createLogfileName(rpaFolder, "")
            'Call MsgBox("logfile Name and path changed: " & logfileNamePath)


            ' hierin wird der eigentliche Import erledigt
            watchDialog.ShowDialog()
        Else
            ' FileNamen für logging zusammenbauen
            logfileNamePath = createLogfileName(rpaFolder, "")

            Call logger(ptErrLevel.logInfo, "VisboRPA: proxyURL", awinSettings.proxyURL)
            Call logger(ptErrLevel.logInfo, "VisboRPA: Visbo Plattform", awinSettings.databaseURL)
            Call logger(ptErrLevel.logInfo, "VisboRPA: User", dbUsername)
            Call logger(ptErrLevel.logInfo, "VisboRPA: Visbo Center", awinSettings.databaseName)
            Call logger(ptErrLevel.logInfo, "VisboRPA: active Portfolio", myActivePortfolio)
            Call logger(ptErrLevel.logInfo, "VisboRPA: RPA Folder", rpaFolder)

            msgTxt = "VISBO Robotic Process automation cancelled: For more details have a look at the logfiles ....  " & rpaFolder & "\logfiles"
            Call MsgBox(msgTxt)
            ' Console.WriteLine(msgTxt)
            Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)
            ' Fehlermeldung für login-Error
        End If


    End Sub
    ''' <summary>
    ''' Check if the process is already running
    ''' </summary>
    ''' <param name="process"></param>
    ''' <returns></returns>
    Public Function IsProcessRunning(process As String) As Boolean

        Dim p() As Process
        p = System.Diagnostics.Process.GetProcessesByName(process)
        If p.Count > 1 Then
            ' Process is running
            IsProcessRunning = True
        Else
            ' Process is not running
            IsProcessRunning = False
        End If

    End Function


    ''' <summary>
    ''' Import of file fname, category rpaCat time importDate to the VisboCenter awinSettings.databaseName
    ''' </summary>
    ''' <param name="fname"></param>
    ''' <param name="rpaCat"></param>
    ''' <param name="importDate"></param>
    ''' <returns></returns>
    Public Function importOneProject(ByVal fname As String, ByVal rpaCat As PTRpa, ByVal importDate As Date) As Boolean


        Dim myName As String = My.Computer.FileSystem.GetName(fname)
        Dim currentWB As xlns.Workbook = Nothing
        Dim allOk As Boolean = False

        errMessages = New Collection

        Try

            If Not rpaCat = PTRpa.visboMPP _
                                And Not rpaCat = PTRpa.visboActualData1 _
                                And Not rpaCat = PTRpa.visboActualData2 Then

                appInstance.DisplayAlerts = False
                currentWB = appInstance.Workbooks.Open(fname)
            End If

            ' now clear Session 
            Call emptyAllVISBOStructures()

            logfileNamePath = createLogfileName(rpaFolder, myName)
            Select Case rpaCat
                Case CInt(PTRpa.visboProjectList)

                    allOk = processProjectList(myName, myActivePortfolio)

                Case CInt(PTRpa.visboFindProjectStart)

                    allOk = processFindProjectStart(myName)

                Case CInt(PTRpa.visboFindProjectStartPM)

                    allOk = processFindProjectStart(myName, PTRpa.visboFindProjectStartPM)

                Case CInt(PTRpa.visboMPP)

                    allOk = processMppFile(fname, importDate)

                Case CInt(PTRpa.visboProject)

                    allOk = processVisboBrief(myName, importDate, errMessages)

                Case CInt(PTRpa.visboJira)

                    allOk = processVisboJira(fname, importDate)

                Case CInt(PTRpa.visboDefaultCapacity)

                    allOk = processVisboUrlaubsplaner(fname, importDate, errMessages)

                Case CInt(PTRpa.visboInitialOrga)

                    allOk = processInitialOrga(myName)

                Case CInt(PTRpa.visboRoundtripOrga)
                    ' this will no longer be support -> error message
                    allOk = processRoundTripOrga(myName)

                Case CInt(PTRpa.visboModifierCapacities)

                    allOk = True
                    Call logger(ptErrLevel.logError, "import Modifier Capacities", " not yet implemented !")

                Case CInt(PTRpa.visboExternalContracts)

                    allOk = True
                    Call logger(ptErrLevel.logError, "import external Contracts", " not yet implemented !")


                Case CInt(PTRpa.visboActualData1)

                    allOk = processVisboActualData1(fname, importDate)

                Case CInt(PTRpa.visboActualData2)

                    logfileNamePath = createLogfileName(rpaFolder)

                    'Dim completionFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(rpaFolder, FileIO.SearchOption.SearchTopLevelOnly, "Timesheet_completed*.*")
                    ' in collectFolder verschieben
                    Dim newDestination As String = My.Computer.FileSystem.CombinePath(collectFolder, myName)
                    My.Computer.FileSystem.MoveFile(fname, newDestination, True)
                    Call logger(ptErrLevel.logInfo, "collect: ", myName)
                    ' nachsehen ob collect vollständig
                    If completedOK Then
                        allOk = processVisboActualData2(fname, myActivePortfolio, collectFolder, importDate)
                    End If

                Case CInt(PTRpa.visboActualData3)
                    allOk = True
                    Call logger(ptErrLevel.logError, "Import of actual data ???", " do not exist so far !")

                Case CInt(PTRpa.visboNewTagetik)
                    allOk = True
                    Call logger(ptErrLevel.logError, "Import new Projects of Tagetik", " not yet integrated !")

                Case CInt(PTRpa.visboUpdateTagetik)
                    allOk = True
                    Call logger(ptErrLevel.logError, "Import Project-update of Tagetik", " not yet integrated !")

                Case CInt(PTRpa.visboEGeckoCapacity)
                    allOk = True
                    Call logger(ptErrLevel.logError, "Import Capacities coming from eGecko", " not yet integrated !")

                Case CInt(PTRpa.visboInstartProposal)
                    allOk = processInstartProposal(fname, myActivePortfolio, collectFolder, importDate)
                    'Call logger(ptErrLevel.logError, "Import Calc-Sheet of Instart", " not yet integrated !")

                Case CInt(PTRpa.visboProposal)
                    allOk = True
                    Call logger(ptErrLevel.logError, "Import Cost-Assertion Sheet Telair", " not yet integrated !")

                Case CInt(PTRpa.visboZeussCapacity)
                    allOk = True
                    Call logger(ptErrLevel.logError, "Import Zeuss-Capacities Telair", " not yet integrated !")

                Case CInt(PTRpa.visboFindfeasiblePortfolio)

                    allOk = readListIntoStorage(PTRpa.visboFindfeasiblePortfolio)

                    If allOk Then
                        allOk = defineFeasiblePortfolio()
                    End If

                Case CInt(PTRpa.visboAutoAdjust)

                    allOk = processAutoAdjustPortfolio()

                Case CInt(PTRpa.visboSuggestResourceAllocation)

                    allOk = processAutoAllocatePortfolio()

                Case CInt(PTRpa.visboCreateHedgedVariant)

                    allOk = readListIntoStorage(PTRpa.visboCreateHedgedVariant)

                    If allOk Then
                        allOk = processCreateHedgedVariants()
                    Else

                    End If


                Case Else
                    Call logger(ptErrLevel.logError, "ImportType is not known so far !", " unknown !")

            End Select

            ' Sendet eine Email an den User
            'Dim result_sendEmail As Boolean = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("files abgearbeitet", errMsgCode)

            Try
                If Not (rpaCat = PTRpa.visboMPP Or
                                        rpaCat = PTRpa.visboActualData1 Or
                                        rpaCat = PTRpa.visboActualData2) Then

                    'If allOk Then
                    '    If IsNothing(currentWB) Then
                    '        ' workbook bereits wieder geschlossen
                    '        appInstance.DisplayAlerts = False
                    '        currentWB = appInstance.Workbooks.Open(fname)
                    '    End If
                    '    CType(currentWB.Worksheets(1), xlns.Worksheet).Cells(1, 1).interior.color = visboFarbeGreen
                    'Else
                    '    CType(currentWB.Worksheets(1), xlns.Worksheet).Cells(1, 1).interior.color = visboFarbeRed
                    'End If
                    currentWB.Close(SaveChanges:=False)
                End If
            Catch ex As Exception

            End Try

            ' here the logfiles and the importfiles will be moved to a folder depending on the result of the import
            If Not rpaCat = PTRpa.visboActualData2 Then
                Call processResult(fname, allOk, errMessages)
            Else
                Call processResult(fname, allOk, errMessages)
            End If

        Catch ex As Exception

            If awinSettings.englishLanguage Then
                msgTxt = "Error importing: " & ex.Message
            Else
                msgTxt = "Fehler beim Import von: " & ex.Message
            End If
            Call logger(ptErrLevel.logsevereError, msgTxt, myName & "/" & rpaCat.ToString)
        End Try

        importOneProject = allOk


    End Function


    ''' <summary>
    ''' clear the Session
    ''' </summary>
    Private Sub emptyRPASession()
        ImportProjekte.Clear()
        ShowProjekte.Clear(False)
        AlleProjekte.Clear(False)
        AlleProjektSummaries.Clear(False)
    End Sub

    Private Function setWriteProtection(ByVal hproj As clsProjekt, ByVal writeProtect As Boolean) As Boolean

        Dim err As New clsErrorCodeMsg

        Dim pvName As String = calcProjektKey(hproj.name, hproj.variantName)

        Dim wpItem As New clsWriteProtectionItem(pvName, ptWriteProtectionType.project,
                                                dbUsername, False, writeProtect)

        setWriteProtection = CType(databaseAcc, DBAccLayer.Request).setWriteProtection(wpItem, err)

    End Function


    ''' <summary>
    ''' Read the projectTemplates from the actual VisboCenter 
    ''' </summary>
    ''' <returns></returns>
    Private Function readProjectTemplates() As Date

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
                Call logger(ptErrLevel.logWarning, "readProjectTemplates", "No project templates in this VC: " & myVC)
                result = Date.MinValue
            End If

        Else
            Call logger(ptErrLevel.logWarning, "readProjectTemplates", "Getting project templates from Server finished with warning: " & err.errorMsg)
            result = Date.MinValue
        End If

        readProjectTemplates = result

    End Function
    ''' <summary>
    ''' reading the VCSetting "customization" if stored in the actual VC
    ''' </summary>
    ''' <returns></returns>
    Private Function readCustomizations() As Date

        Dim result As Date = Date.MinValue
        Dim err As New clsErrorCodeMsg
        '
        ' Read Customizations 
        Dim customizations As clsCustomization = CType(databaseAcc, DBAccLayer.Request).retrieveCustomizationFromDB("", Date.Now, False, err)

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

    Private Function storeConstellationFromProjectList(ByVal projectList As clsProjekteAlle,
                                                    ByVal portfolioName As String, ByVal variantName As String) As Boolean

        Dim result = True

        Try

            AlleProjekte.Clear()
            ' now make sure all projects are in AlleProjekte
            For Each ppair As KeyValuePair(Of String, clsProjekt) In projectList.liste
                If Not AlleProjekte.Containskey(ppair.Key) Then
                    AlleProjekte.Add(ppair.Value)
                End If
            Next

            ShowProjekte.Clear()
            For Each ppair As KeyValuePair(Of String, clsProjekt) In projectList.liste

                If Not ShowProjekte.contains(ppair.Value.name) Then
                    ShowProjekte.Add(ppair.Value)
                End If

            Next


            ' currentSessionConstellation is build by alle the Showprojekte.add and AlleProjekte.add Commands ...
            ' create form that a portfolio, only containing the show-Elements 
            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:=variantName)

            ' now store the Portfolio , with name portfolioName
            Dim errMsg As New clsErrorCodeMsg
            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

            Dim outputCollection As New Collection
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)

            ' then empty ShowProjekte again 
            ShowProjekte.Clear()

        Catch ex As Exception
            result = False
            Call logger(ptErrLevel.logError, "failure in store Portfolio: " & portfolioName & vbLf & ex.Message, PTRpa.visboProjectList.ToString)
        End Try

        storeConstellationFromProjectList = result

    End Function

    ''' <summary>
    ''' stores all projects in ImportProjekte, then clears ImportProjekte
    ''' returns true, if all went ok
    ''' </summary>
    ''' <returns></returns>
    Private Function storeImportProjekte() As Boolean

        Dim saveUserRole As ptCustomUserRoles = myCustomUserRole.customUserRole
        Dim jetzt As Date = Date.Now
        Dim ok As Boolean = True

        Try
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste
                Dim outputCollection As New Collection
                Dim hproj As clsProjekt = Nothing
                Dim Err As New clsErrorCodeMsg
                If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(kvp.Value.name, kvp.Value.variantName, jetzt, Err) Then
                    hproj = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(kvp.Value.name, kvp.Value.variantName, "", jetzt, Err)
                End If

                If IsNothing(hproj) Then
                    ' does not yet exist .. 
                    If Not AlleProjekte.Containskey(calcProjektKey(kvp.Value)) Then
                        ' necessary because store ruft writeProtections für AllePRojekte Projekte auf 
                        AlleProjekte.Add(kvp.Value, False)
                    End If

                    myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager

                    If storeSingleProjectToDB(kvp.Value, outputCollection) Then
                        ok = ok And True
                        Call logger(ptErrLevel.logInfo, "project stored: ", kvp.Value.getShapeText)
                        'Console.WriteLine("project stored: " & kvp.Value.getShapeText)
                    Else
                        ok = ok And False
                        Call logger(ptErrLevel.logError, "project store failed: ", outputCollection)
                        'Console.WriteLine("!! ... project store failed: " & kvp.Value.getShapeText)
                    End If

                Else
                    ' hproj in alleProjekte schieben, damit writeProtection gecheckt werden kann.
                    If Not AlleProjekte.Containskey(calcProjektKey(kvp.Value)) Then
                        ' necessary because store ruft writeProtections für AllePRojekte Projekte auf 
                        AlleProjekte.Add(kvp.Value, False)
                    End If

                    myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung

                    If storeSingleProjectToDB(kvp.Value, outputCollection) Then
                        ok = ok And True
                        Call logger(ptErrLevel.logInfo, "project updated: ", kvp.Value.getShapeText)
                        'Console.WriteLine("project updated: " & kvp.Value.getShapeText)
                    Else
                        ok = ok And False
                        Call logger(ptErrLevel.logError, "project update failed: ", outputCollection)
                        'Console.WriteLine("!! ... project update failed: " & kvp.Value.getShapeText)
                    End If

                End If

            Next

        Catch ex As Exception
            ok = False
            Call logger(ptErrLevel.logError, "Store Projects from List failed", ex.Message)
            'Console.WriteLine("!!!! Store Projects from List failed" & ex.Message)
        End Try

        storeImportProjekte = ok
    End Function
    ''' <summary>
    ''' start of the RPA with VisboCenter , at url, rpaPath and perhaps a proxy-Server
    ''' </summary>
    ''' <param name="mongoName"></param>
    ''' <param name="url"></param>
    ''' <param name="path"></param>
    ''' <param name="proxy"></param>
    ''' <returns></returns>
    Public Function startUpRPA(ByVal mongoName As String, ByVal url As String, ByVal path As String, ByVal proxy As String) As Boolean

        Dim result As Boolean = False

        ' ggf hier noch die appInstance setzen ... 
        appInstance = New xlns.Application

        Try

            'If readawinSettings(path) Then

            result = True
            ' independent of what is given in projectboardConfig.xml
            awinSettings.databaseName = mongoName
            awinSettings.databaseURL = url
            awinSettings.proxyURL = proxy
            ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
            'awinSettings.rememberUserPwd = True
            'awinSettings.userNamePWD = My.Settings.userNamePWD

            awinSettings.visboServer = True
            ' returns false if anything goes wrong .. 
            result = rpaSetTypen()

            'ElseIf readawinSettings(swPath) Then
            '    result = True
            '    ' independent of what is given in projectboardConfig.xml
            '    awinSettings.databaseName = mongoName
            '    awinSettings.databaseURL = url
            '    awinSettings.proxyURL = proxy
            '    ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
            '    awinSettings.rememberUserPwd = True
            '    awinSettings.userNamePWD = My.Settings.userNamePWD

            '    awinSettings.visboServer = True

            '    ' returns false if anything goes wrong .. 
            '    result = rpaSetTypen()
            'End If


        Catch ex As Exception
            Call logger(ptErrLevel.logError, "startUpRPA", ex.Message)
            result = False
        End Try

        startUpRPA = result
    End Function

    ''' <summary>
    ''' when called, all awinSetting Variables are set .. 
    ''' </summary>
    ''' <returns></returns>
    Private Function rpaSetTypen() As Boolean

        Dim result As Boolean = False

        Try
            Dim err As New clsErrorCodeMsg


            Dim anzIEOrdner As Integer = [Enum].GetNames(GetType(PTImpExp)).Length
            ReDim importOrdnerNames(anzIEOrdner - 1)
            ReDim exportOrdnerNames(anzIEOrdner - 1)

            ' Auslesen des Window Namens 
            Dim accountToken As IntPtr = WindowsIdentity.GetCurrent().Token
            Dim myUser As New WindowsIdentity(accountToken)
            myWindowsName = myUser.Name

            ' tk: in RPA there is no necessity to have a globalPath
            globalPath = awinSettings.globalPath
            globalPath = ""
            awinPath = "C:\VISBO\VISBO Config Data"

            Dim curUserDir As String = "C:\VISBO"


            If My.Settings.rpaPath = "" Then
                ' tk 12.12.18 damit wird sichergestellt, dass bei einer Installation die Demo Daten einfach im selben Directory liegen können
                ' im ProjectBoardConfig kann demnach entweder der leere String stehen oder aber ein relativer Pfad, der vom User/Home Directory ausgeht ... 
                'Dim locationOfProjectBoard = My.Computer.FileSystem.GetParentPath(appInstance.ActiveWorkbook.FullName)
                Dim locationOfRPAExe As String = My.Computer.FileSystem.CurrentDirectory
                locationOfRPAExe = "C:\VISBO"
                Dim stdConfigDataName As String = "VISBO Config Data"

                awinPath = My.Computer.FileSystem.CombinePath(locationOfRPAExe, stdConfigDataName)

                If My.Computer.FileSystem.DirectoryExists(awinPath) Then
                    ' alles ok
                Else
                    awinPath = My.Computer.FileSystem.CombinePath(curUserDir, stdConfigDataName)
                    If My.Computer.FileSystem.DirectoryExists(awinPath) Then
                        ' alles ok 
                    End If
                End If
            ElseIf My.Computer.FileSystem.DirectoryExists(My.Settings.rpaPath) Then
                awinPath = My.Settings.rpaPath
            Else
                awinPath = My.Computer.FileSystem.CombinePath(curUserDir, awinSettings.awinPath)
            End If


            If Not awinPath.EndsWith("\") Then
                awinPath = awinPath & "\"
            End If


            ' Debug-Mode?
            ' Logfile schreiben: 
            'Call logger(ptErrLevel.logInfo, "startUpRPA", "localPath:" & awinPath)
            'Call logger(ptErrLevel.logInfo, "startUpRPA", "GlobalPath:" & globalPath)


            If globalPath <> "" Then

                If Not globalPath.EndsWith("\") Then
                    globalPath = globalPath & "\"
                End If

                ' Synchronization von Globalen und Lokalen Pfad

                If awinPath <> globalPath And My.Computer.FileSystem.DirectoryExists(globalPath) Then

                    Call synchronizeGlobalToLocalFolder()
                    Call logger(ptErrLevel.logInfo, "rpaSetTypen", "Synchronized localPath with globalPath")

                Else

                    Call logger(ptErrLevel.logInfo, "rpaSetTypen", "no Synchronization between localPath and globalPath")

                End If

            End If

            StartofCalendar = StartofCalendar.Date

            'ur:07.02.2022 auskommentiert
            'DiagrammTypen(0) = "Phase"
            'DiagrammTypen(1) = "Rolle"
            'DiagrammTypen(2) = "Kostenart"
            'DiagrammTypen(3) = "Portfolio"
            'DiagrammTypen(4) = "Ergebnis"
            'DiagrammTypen(5) = "Meilenstein"
            'DiagrammTypen(6) = "Meilenstein Trendanalyse"
            'DiagrammTypen(7) = "Phasen-Kategorie"
            'DiagrammTypen(8) = "Meilenstein-Kategorie"
            'DiagrammTypen(9) = "Cash-Flow"


            'Try
            '    repMessages = XMLImportReportMsg(repMsgFileName, awinSettings.ReportLanguage)
            '    Call setLanguageMessages()
            'Catch ex As Exception

            'End Try


            'ur:07.02.2022 auskommentiert ---
            'autoSzenarioNamen(0) = "before Optimization"
            'autoSzenarioNamen(1) = "1. Optimum"
            'autoSzenarioNamen(2) = "2. Optimum"
            'autoSzenarioNamen(3) = "3. Optimum"

            ''
            '' die Namen der Worksheets Ressourcen und Portfolio verfügbar machen
            '' die Zahlen müssen korrespondieren mit der globalen Enumeration ptTables 
            'arrWsNames(1) = "repCharts" ' Tabellenblatt zur Aufnahme der Charts für Reports 
            'arrWsNames(2) = "Vorlage" ' depr
            '' arrWsNames(3) = 
            'arrWsNames(ptTables.MPT) = "MPT"                          ' Multiprojekt-Tafel 
            'arrWsNames(4) = "Einstellungen"                ' in Customization File 
            '' arrWsNames(5) = 
            'arrWsNames(ptTables.meRC) = "meRC"                          ' Edit Ressourcen
            'arrWsNames(6) = "meTE"                          ' Edit Termine
            'arrWsNames(7) = "Darstellungsklassen"           ' wird in awinsettypen hinter MPT kopiert; nimmt für die Laufzeit die Darstellungsklassen auf 
            'arrWsNames(8) = "Phasen-Mappings"               ' in Customization
            'arrWsNames(9) = "meAT"                          ' Edit Attribute 
            'arrWsNames(10) = "Meilenstein-Mappings"         ' in Customization
            '' arrWsNames(11) = 
            'arrWsNames(ptTables.meCharts) = "meCharts"                     ' Massen-Edit Charts 
            'arrWsNames(ptTables.mptPfCharts) = "mptPfCharts"                     ' vorbereitet: Portfolio Charts 
            'arrWsNames(ptTables.mptPrCharts) = "mptPrCharts"                     ' vorbereitet: Projekt Charts 
            'arrWsNames(14) = "Objekte" ' depr
            'arrWsNames(15) = "missing Definitions"          ' in Customization File 


            'awinSettings.applyFilter = False

            'showRangeLeft = 0
            'showRangeRight = 0
            'ur:07.02.2022 auskommentiert ---


            ' always needs to be database / VISBO Server access 
            noDB = False

            If awinSettings.userNamePWD <> "" Then

                Dim visboCrypto As New clsVisboCryptography(visboCryptoKey)

                dbUsername = visboCrypto.getUserNameFromCipher(awinSettings.userNamePWD)
                dbPasswort = visboCrypto.getPwdFromCipher(awinSettings.userNamePWD)


                If IsNothing(awinSettings.VCid) Then
                    awinSettings.VCid = ""
                End If

                If IsNothing(databaseAcc) Then
                    databaseAcc = New DBAccLayer.Request
                End If

                If Not loginErfolgreich Then
                    loginErfolgreich = logInToMongoDB(True)
                End If

            Else
                If Not loginErfolgreich Then
                    loginErfolgreich = logInToMongoDB(True)
                End If
            End If

            ''ûr: 10032022: not needed for RPA
            '' Read appearance Definitions
            'appearanceDefinitions.liste = CType(databaseAcc, DBAccLayer.Request).retrieveAppearancesFromDB("", Date.Now, False, err)
            'If IsNothing(appearanceDefinitions.liste) Or appearanceDefinitions.liste.Count > 0 Then
            '    ' user has no access to any VISBO Center 
            '    msgTxt = "No appearance Definitions in VISBO"
            '    Call logger(ptErrLevel.logInfo, "rpaSetTypen", "")
            '    'Throw New ArgumentException(msgTxt)
            'End If

            ''
            '' Read Customizations 
            lastReadingCustomization = readCustomizations()

            'Dim customizations As clsCustomization = CType(databaseAcc, DBAccLayer.Request).retrieveCustomizationFromDB("", Date.Now, False, err)

            'If Not IsNothing(customizations) Then
            '    StartofCalendar = customizations.kalenderStart
            '    Call logger(ptErrLevel.logInfo, "rpaSetTypen", " StartOfCalendar: " & StartofCalendar.ToString)

            '    businessUnitDefinitions = customizations.businessUnitDefinitions

            '    PhaseDefinitions = customizations.phaseDefinitions

            '    MilestoneDefinitions = customizations.milestoneDefinitions

            '    showtimezone_color = customizations.showtimezone_color
            '    noshowtimezone_color = customizations.noshowtimezone_color
            '    calendarFontColor = customizations.calendarFontColor
            '    nrOfDaysMonth = customizations.nrOfDaysMonth
            '    farbeInternOP = customizations.farbeInternOP
            '    farbeExterne = customizations.farbeExterne
            '    iProjektFarbe = customizations.iProjektFarbe
            '    iWertFarbe = customizations.iWertFarbe
            '    vergleichsfarbe0 = customizations.vergleichsfarbe0
            '    vergleichsfarbe1 = customizations.vergleichsfarbe1

            '    awinSettings.SollIstFarbeB = customizations.SollIstFarbeB
            '    awinSettings.SollIstFarbeL = customizations.SollIstFarbeL
            '    awinSettings.SollIstFarbeC = customizations.SollIstFarbeC
            '    awinSettings.AmpelGruen = customizations.AmpelGruen

            '    awinSettings.AmpelGelb = customizations.AmpelGelb
            '    awinSettings.AmpelRot = customizations.AmpelRot
            '    awinSettings.AmpelNichtBewertet = customizations.AmpelNichtBewertet
            '    awinSettings.glowColor = customizations.glowColor

            '    awinSettings.timeSpanColor = customizations.timeSpanColor
            '    awinSettings.showTimeSpanInPT = customizations.showTimeSpanInPT

            '    awinSettings.gridLineColor = customizations.gridLineColor

            '    awinSettings.missingDefinitionColor = customizations.missingDefinitionColor

            '    awinSettings.ActualdataOrgaUnits = customizations.allianzIstDatenReferate
            '    awinSettings.ActualdataOrgaUnits = customizations.isActualDataRelevant

            '    awinSettings.onePersonOneRole = customizations.onePersonOneRole
            '    awinSettings.autoSetActualDataDate = customizations.autoSetActualDataDate

            '    awinSettings.actualDataMonth = customizations.actualDataMonth
            '    ergebnisfarbe1 = customizations.ergebnisfarbe1
            '    ergebnisfarbe2 = customizations.ergebnisfarbe2
            '    weightStrategicFit = customizations.weightStrategicFit
            '    awinSettings.kalenderStart = customizations.kalenderStart
            '    awinSettings.zeitEinheit = customizations.zeitEinheit
            '    awinSettings.kapaEinheit = customizations.kapaEinheit
            '    awinSettings.offsetEinheit = customizations.offsetEinheit
            '    awinSettings.EinzelRessExport = customizations.EinzelRessExport
            '    awinSettings.zeilenhoehe1 = customizations.zeilenhoehe1
            '    awinSettings.zeilenhoehe2 = customizations.zeilenhoehe2
            '    awinSettings.spaltenbreite = customizations.spaltenbreite
            '    awinSettings.autoCorrectBedarfe = customizations.autoCorrectBedarfe
            '    awinSettings.propAnpassRess = customizations.propAnpassRess
            '    awinSettings.showValuesOfSelected = customizations.showValuesOfSelected

            '    awinSettings.mppProjectsWithNoMPmayPass = customizations.mppProjectsWithNoMPmayPass
            '    awinSettings.fullProtocol = customizations.fullProtocol
            '    awinSettings.addMissingPhaseMilestoneDef = customizations.addMissingPhaseMilestoneDef
            '    awinSettings.alwaysAcceptTemplateNames = customizations.alwaysAcceptTemplateNames
            '    awinSettings.eliminateDuplicates = customizations.eliminateDuplicates
            '    awinSettings.importUnknownNames = customizations.importUnknownNames
            '    awinSettings.createUniqueSiblingNames = customizations.createUniqueSiblingNames

            '    awinSettings.readWriteMissingDefinitions = customizations.readWriteMissingDefinitions
            '    awinSettings.meExtendedColumnsView = customizations.meExtendedColumnsView
            '    awinSettings.meDontAskWhenAutoReduce = customizations.meDontAskWhenAutoReduce
            '    awinSettings.readCostRolesFromDB = customizations.readCostRolesFromDB

            '    awinSettings.importTyp = customizations.importTyp

            '    awinSettings.meAuslastungIsInclExt = customizations.meAuslastungIsInclExt

            '    awinSettings.englishLanguage = customizations.englishLanguage

            '    awinSettings.showPlaceholderAndAssigned = customizations.showPlaceholderAndAssigned
            '    awinSettings.considerRiskFee = customizations.considerRiskFee

            '    StartofCalendar = awinSettings.kalenderStart

            '    historicDate = StartofCalendar
            '    Try
            '        If awinSettings.englishLanguage Then
            '            menuCult = ReportLang(PTSprache.englisch)
            '            repCult = menuCult
            '            awinSettings.kapaEinheit = "PD"
            '        Else
            '            awinSettings.kapaEinheit = "PT"
            '            menuCult = ReportLang(PTSprache.deutsch)
            '            repCult = menuCult
            '        End If
            '    Catch ex As Exception
            '        awinSettings.englishLanguage = False
            '        awinSettings.kapaEinheit = "PT"
            '        menuCult = ReportLang(PTSprache.deutsch)
            '        repCult = menuCult
            '    End Try
            'Else
            '    msgTxt = "No customization in VISBO"
            '    Call logger(ptErrLevel.logInfo, "rpaSetTypen", msgTxt)
            '    'Throw New ArgumentException(msgTxt)
            'End If

            '
            ' now read Organisation 
            ''
            '' Read Customizations 
            lastReadingOrganisation = readOrganisations()

            'Dim currentOrga As clsOrganisation = CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", Date.Now, False, err)

            'If Not IsNothing(currentOrga) Then
            '    If currentOrga.count > 0 Then

            '        If currentOrga.count > 0 Then
            '            validOrganisations.addOrga(currentOrga)
            '        End If

            '        CostDefinitions = currentOrga.allCosts
            '        RoleDefinitions = currentOrga.allRoles

            '        Dim tmpActDataString As String = currentOrga.allRoles.getActualdataOrgaUnits
            '        If tmpActDataString = "" And awinSettings.ActualdataOrgaUnits <> "" Then
            '            ' do nothing, leave it as is 
            '        Else
            '            awinSettings.ActualdataOrgaUnits = tmpActDataString
            '        End If

            '    Else
            '        msgTxt = "No organisation in VISBO"
            '        Call logger(ptErrLevel.logInfo, "rpaSetTypen", msgTxt)
            '        'Throw New ArgumentException("msgTxt")
            '    End If
            'Else
            '    msgTxt = "No organisation in VISBO"
            '    Call logger(ptErrLevel.logInfo, "rpaSetTypen", msgTxt)
            '    'Throw New ArgumentException("msgTxt")
            'End If

            '
            ' now read customFieldDefinitions; is allowed to be empty
            customFieldDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCustomFieldsFromDB(err)

            If IsNothing(customFieldDefinitions) Then
                customFieldDefinitions = New clsCustomFieldDefinitions
                Call logger(ptErrLevel.logInfo, "rpaSetTypen", "no CustomFieldDefinitions found")
            End If

            '
            ' myCustomUserRole wird by Default auf <Alles> gesetzt 
            myCustomUserRole = New clsCustomUserRole

            With myCustomUserRole
                .customUserRole = ptCustomUserRoles.Alles
                .specifics = ""
                .userName = dbUsername
            End With

            '
            ' now read Vorlagen - maybe Empty
            lastReadingProjectTemplates = readProjectTemplates()

            result = True

        Catch ex As Exception

            result = False
            Call logger(ptErrLevel.logError, "rpaSetTypen", ex.Message)

        End Try

        rpaSetTypen = result

    End Function



    Public Function bestimmeRPACategory(ByVal fileName As String) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown

        ' Open fileName 
        If Not IsNothing(fileName) Then

            If My.Computer.FileSystem.FileExists(fileName) Then

                Try
                    appInstance.DisplayAlerts = False
                    Dim currentWB As xlns.Workbook = Module1.appInstance.Workbooks.Open(fileName, UpdateLinks:=0)
                    currentWB.Final = False
                    appInstance.DisplayAlerts = True

                    Try
                        ' Check auf Project Batch-List
                        If result = PTRpa.visboUnknown Then
                            result = checkProjectBatchList(currentWB)
                        End If

                        ' Check auf Project Find Best Start
                        If result = PTRpa.visboUnknown Then
                            result = checkFindBestStarts(currentWB)
                        End If

                        ' Check auf define Feasible Portfolio 
                        If result = PTRpa.visboUnknown Then
                            result = checkfeasiblePortfolio(currentWB)
                        End If

                        ' Check auf Auto-Allocate 
                        If result = PTRpa.visboUnknown Then
                            result = checkAutoAllocate(currentWB)
                        End If

                        ' Check auf Sales Pipeline 
                        If result = PTRpa.visboUnknown Then
                            result = checkCreateHedgedVariants(currentWB)
                        End If

                        ' Check auf VISBO Project Brief
                        If result = PTRpa.visboUnknown Then
                            result = checkProjectBrief(currentWB)
                        End If

                        ' Check auf Auto Adjust Resource Bottlenecks
                        If result = PTRpa.visboUnknown Then
                            result = checkAutoAdjustPortfolio(currentWB)
                        End If

                        ' Check auf Organisation 
                        If result = PTRpa.visboUnknown Then
                            result = checkOrganisation(currentWB)
                        End If

                        ' Check auf Visbo Center Organisation 
                        If result = PTRpa.visboUnknown Then
                            result = checkVCOrganisation(currentWB)
                        End If

                        ' Check auf Jira Ausleitung
                        If result = PTRpa.visboUnknown Then
                            result = checkJiraProjects(currentWB)
                        End If

                        ' Check auf VISBO Project Template  

                        ' Check auf Urlaubskalender 
                        If result = PTRpa.visboUnknown Then
                            result = checkUrlaubsplaner(currentWB)
                        End If

                        ' Check auf Modifier Kapazitäten

                        ' Check auf externe Rahmenverträge 
                        If result = PTRpa.visboUnknown Then
                            result = checkExtRahmenvertr(currentWB)
                        End If

                        ' Check auf Instart eGecko Urlaube ...(Instart) 

                        ' Check auf Zeuss Kapazitäten... (Telair)

                        ' Check auf Ist-Daten 
                        If result = PTRpa.visboUnknown Then
                            result = checkActualData1(currentWB)
                        End If

                        ' Check auf Telair TimeSheets
                        If result = PTRpa.visboUnknown Then
                            'result = checkActualData2(currentWB)
                        End If

                        ' Check auf Tagetik new Project List 
                        If result = PTRpa.visboUnknown Then
                            result = checkTagetikProjectList(currentWB)
                        End If
                        ' Check auf Tagetik update projects 

                        ' Check auf Instart Calculation Template 
                        If result = PTRpa.visboUnknown Then
                            result = checkInstartProposal(currentWB)
                        End If
                        ' Check auf VISBO Calculation Template 

                        currentWB.Close(SaveChanges:=False)

                    Catch ex As Exception

                        currentWB.Close(SaveChanges:=False)

                    End Try


                Catch ex As Exception
                    Dim fileN As String = My.Computer.FileSystem.GetName(fileName)
                    Call logger(ptErrLevel.logWarning, "Error when determining category of RPA", fileN)
                End Try

            End If

        End If


        bestimmeRPACategory = result
    End Function
    Private Function checkFindBestStarts(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName1 As String = "Find Best Start"
        Dim blattName1A As String = "Find Best Start PM"
        Dim blattName2 As String = "Parameters"


        Try

            Dim currentWS As xlns.Worksheet = Nothing

            Try
                currentWS = CType(currentWB.Worksheets.Item(blattName1), xlns.Worksheet)
            Catch ex As Exception
                currentWS = Nothing
            End Try

            If IsNothing(currentWS) Then
                Try
                    currentWS = CType(currentWB.Worksheets.Item(blattName1A), xlns.Worksheet)
                Catch ex As Exception
                    currentWS = Nothing
                End Try
            End If

            If Not IsNothing(currentWS) Then

                Dim paramWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName2), xlns.Worksheet)

                If IsNothing(currentWS) Or IsNothing(paramWS) Then
                    result = PTRpa.visboUnknown
                Else
                    Dim ersteZeile As xlns.Range = CType(currentWS.Rows.Item(1), xlns.Range)
                    Try

                        verifiedStructure = ersteZeile.Cells(1, 1).value.trim = "Name" And
                            CStr(ersteZeile.Cells(1, 2).value).Trim = "Variant"

                    Catch ex As Exception
                        verifiedStructure = False
                    End Try


                    If verifiedStructure Then

                        If currentWS.Name = blattName1 Then
                            result = PTRpa.visboFindProjectStart
                        Else
                            result = PTRpa.visboFindProjectStartPM
                        End If


                        ' Aktiviere das Worksheet 
                        If CType(currentWB.ActiveSheet, xlns.Worksheet).Name <> currentWS.Name Then
                            currentWS.Activate()
                        End If

                        Dim mymessages As New Collection
                        Dim infomsg As String = "File to find best start dates Phases/Milestones detected: " & currentWB.Name
                        If currentWS.Name = blattName1 Then
                            infomsg = "File to find best start dates Roles/Skills detected: " & currentWB.Name
                        End If
                        Call logger(ptErrLevel.logInfo, infomsg, mymessages)
                        'Console.WriteLine(infomsg)
                    Else
                        result = PTRpa.visboUnknown
                    End If


                End If
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkFindBestStarts = result
    End Function

    Private Function checkAutoAllocate(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim blattName0 As String = "VISBO Auto-Allocate"
        Dim blattName1 As String = "Parameters"

        Dim blattExist(1) As Boolean
        blattExist(0) = False
        blattExist(1) = False


        Try

            For Each ws As xlns.Worksheet In currentWB.Worksheets
                If (ws.Name = blattName0) Then
                    blattExist(0) = True
                End If

                If (ws.Name = blattName1) Then
                    blattExist(1) = True
                End If

            Next

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If blattExist(0) And blattExist(1) Then
            result = PTRpa.visboSuggestResourceAllocation
        End If

        checkAutoAllocate = result
    End Function

    Private Function checkAutoAdjustPortfolio(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim blattName0 As String = "Exception List"
        Dim blattName1 As String = "Parameters"
        Dim blattName2 As String = "Ranking List"

        Dim blattExist(1) As Boolean
        blattExist(0) = False
        blattExist(1) = False

        Try

            For Each ws As xlns.Worksheet In currentWB.Worksheets
                If (ws.Name = blattName0) Then
                    blattExist(0) = True
                End If

                If (ws.Name = blattName1) Then
                    blattExist(1) = True
                End If

            Next

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If blattExist(0) And blattExist(1) Then
            result = PTRpa.visboAutoAdjust
        End If

        checkAutoAdjustPortfolio = result
    End Function


    ''' <summary>
    ''' checks whether or not the file is a findFeasiblePortfolio file
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkfeasiblePortfolio(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim blattName1 As String = "VISBO"
        Dim blattName2 As String = "Parameters"

        Dim hasVISBO As Boolean = False
        Dim hasParameters As Boolean = False

        Try

            For Each ws As xlns.Worksheet In currentWB.Worksheets
                If (ws.Name = blattName1) Then

                    Dim ersteZeile As xlns.Range = CType(ws.Rows.Item(1), xlns.Range)
                    Try
                        hasVISBO = ersteZeile.Cells(1, 1).value.trim = "Name" And
                        CStr(ersteZeile.Cells(1, 2).value).Trim = "Variant"
                    Catch ex As Exception

                    End Try
                End If

                If (ws.Name = blattName2) Then
                    hasParameters = True
                End If

            Next

            If (hasVISBO And hasParameters) Then
                result = PTRpa.visboFindfeasiblePortfolio

                Dim mymessages As New Collection
                Dim infomsg As String = "File to define feasible portfolio detected: " & currentWB.Name
                Call logger(ptErrLevel.logInfo, infomsg, mymessages)
                'Console.WriteLine(infomsg)
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkfeasiblePortfolio = result

    End Function


    ''' <summary>
    ''' creates hedged variants for all the sales Pipeline projects
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkCreateHedgedVariants(ByVal currentWB As xlns.Workbook) As PTRpa

        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattPartName As String = "Pipeline"

        Try

            For Each ws As xlns.Worksheet In currentWB.Worksheets
                If (ws.Name.Contains(blattPartName)) Then

                    Dim ersteZeile As xlns.Range = CType(ws.Rows.Item(2), xlns.Range)
                    Try
                        If Not IsNothing(ersteZeile.Cells(1, 3).value) Then
                            If IsNumeric(ersteZeile.Cells(1, 3).value) Then
                                Dim tstValue As Double = CDbl(ersteZeile.Cells(1, 3).value)
                                Dim tstDate As Date = CDate(ersteZeile.Cells(1, 4).value)
                                If tstValue > 0 And tstValue <= 1.0 And tstDate <> Date.MinValue Then
                                    verifiedStructure = True
                                    Exit For
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        verifiedStructure = False
                    End Try
                End If


            Next

            If verifiedStructure Then

                result = PTRpa.visboCreateHedgedVariant
                Dim mymessages As New Collection
                Dim infomsg As String = "File to create hedged variants detected: " & currentWB.Name
                Call logger(ptErrLevel.logInfo, infomsg, mymessages)
                'Console.WriteLine(infomsg)
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkCreateHedgedVariants = result

    End Function




    ''' <summary>
    ''' checks whether or not a file is a visbo project list 
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkProjectBatchList(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "Batch List"

        Try

            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName), xlns.Worksheet)

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                Dim ersteZeile As xlns.Range = CType(currentWS.Rows.Item(1), xlns.Range)
                Try

                    verifiedStructure = ersteZeile.Cells(1, 1).value.trim = "Name" And
                        CStr(ersteZeile.Cells(1, 2).value).Trim = "Variant" And
                        CStr(ersteZeile.Cells(1, 3).value).Trim = "Template" And
                        CStr(ersteZeile.Cells(1, 4).value).Trim = "Responsible" And
                        CStr(ersteZeile.Cells(1, 5).value).Trim = "Start" And
                        CStr(ersteZeile.Cells(1, 6).value).Trim = "End" And
                        CStr(ersteZeile.Cells(1, 7).value).Trim.StartsWith("Duration") And
                        CStr(ersteZeile.Cells(1, 8).value).Trim.StartsWith("Budget") And
                        CStr(ersteZeile.Cells(1, 9).value).Trim.Contains("Resources") And
                        CStr(ersteZeile.Cells(1, 10).value).Trim.Contains("Other Cost") And
                        CStr(ersteZeile.Cells(1, 11).value).Trim = "Risk" And
                        CStr(ersteZeile.Cells(1, 12).value).Trim = "Strategy" And
                        CStr(ersteZeile.Cells(1, 13).value).Trim = "Business Unit" And
                        CStr(ersteZeile.Cells(1, 14).value).Trim = "Description"

                Catch ex As Exception
                    verifiedStructure = False
                End Try

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboProjectList
        Else
            result = PTRpa.visboUnknown
        End If

        checkProjectBatchList = result
    End Function

    ''' <summary>
    ''' checks whether or not it is a default VISBO project brief with Stammdaten, Termine and the like 
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkProjectBrief(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure() As Boolean
        Dim possibleTableNames() As String

        ReDim verifiedStructure(3)
        verifiedStructure(0) = False
        verifiedStructure(1) = False
        verifiedStructure(2) = False
        verifiedStructure(3) = False

        ReDim possibleTableNames(4)
        possibleTableNames(0) = "Stammdaten"
        possibleTableNames(1) = "Ressourcen"
        possibleTableNames(2) = "Ressourcenbedarfe"
        possibleTableNames(3) = "Termine"
        possibleTableNames(4) = "Attribute"


        Try

            If IsNothing(currentWB) Then
                result = PTRpa.visboUnknown
            Else

                For Each tmpSheet As xlns.Worksheet In CType(currentWB.Worksheets, xlns.Sheets)

                    If tmpSheet.Name = possibleTableNames(0) Then
                        verifiedStructure(0) = True
                    End If

                    If tmpSheet.Name = possibleTableNames(1) Or tmpSheet.Name = possibleTableNames(2) Then
                        verifiedStructure(1) = True
                    End If

                    If tmpSheet.Name = possibleTableNames(3) Then
                        verifiedStructure(2) = True
                    End If

                    If tmpSheet.Name = possibleTableNames(4) Then
                        verifiedStructure(3) = True
                    End If

                Next

                ' that will be done in the Import Routine itself - here it is about recognizing which kind of Import Method should be choosen ... 
                'If verifiedStructure(0) Then

                '    ' valid StammDaten ? 

                '    Try
                '        Dim pName As String = CStr(currentWB.range("Projekt_Name").value)
                '        Dim sd As String = CDate(currentWB.range("StartDatum").value)
                '        Dim ed As String = CDate(currentWB.range("EndeDatum").value)

                '    Catch ex As Exception
                '        verifiedStructure(0) = False
                '    End Try

                'End If

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure(0) And verifiedStructure(2) Then
            result = PTRpa.visboProject
        Else
            result = PTRpa.visboUnknown
        End If

        checkProjectBrief = result

    End Function

    ''' <summary>
    ''' checks the original / initial VISBO Excel Organisation Structure
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkOrganisation(ByVal currentWB As xlns.Workbook) As PTRpa

        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure() As Boolean
        Dim possibleRangeNames = {"awin_Rollen_Definition", "awin_Gruppen_Definition", "awin_Kosten_Definition"}
        Dim blattName As String = "Organisation"

        Try
            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

            If Not IsNothing(currentWS) Then
                ReDim verifiedStructure(2)
                verifiedStructure(0) = False
                verifiedStructure(1) = False
                verifiedStructure(2) = False

                Dim i As Integer = 0
                For Each rngName As String In possibleRangeNames

                    Dim myRange As xlns.Range = currentWS.Range(rngName)

                    If Not IsNothing(myRange) Then
                        verifiedStructure(i) = True
                    Else
                        Call logger(ptErrLevel.logError, "CheckOrganisation - missing range:", rngName)
                    End If

                    i = i + 1

                Next

                If verifiedStructure(0) And verifiedStructure(1) And verifiedStructure(2) Then
                    result = PTRpa.visboInitialOrga
                End If

            End If


        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkOrganisation = result

    End Function

    ''' <summary>
    ''' checks whether or not it is a downloaded and edited VisboCenterOrganisation 
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkExtRahmenvertr(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim possibleTableNames() As String = {"externe Vertraege", "externe Rahmenvertraege"}
        Dim verifiedStructure As Boolean = False
        Try

            Dim currentWS As xlns.Worksheet = Nothing
            Dim found As Boolean = False

            For Each tmpSheet As xlns.Worksheet In CType(currentWB.Worksheets, xlns.Worksheets)

                For Each tblname As String In possibleTableNames
                    If tmpSheet.Name.StartsWith(tblname) Then
                        found = True
                        currentWS = tmpSheet
                        Exit For
                    End If
                Next

            Next

            If found Then
                result = PTRpa.visboExternalContracts
            End If


        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try


        checkExtRahmenvertr = result
    End Function


    ''' <summary>
    ''' checks whether or not it is a VISBO Urlaubskalender 
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkUrlaubsplaner(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim possibleTableNames() As String = {"1.Halbjahr", "2.Halbjahr"}
        Dim verifiedStructure As Boolean = False
        Try

            Dim currentWS As xlns.Worksheet = Nothing
            Dim found As Boolean = False


            If IsNothing(currentWB) Then
                result = PTRpa.visboUnknown
            Else
                For Each tmpSheet As xlns.Worksheet In CType(currentWB.Worksheets, xlns.Sheets)
                    For Each tblname As String In possibleTableNames
                        If tmpSheet.Name.StartsWith(tblname) Then
                            found = True
                            currentWS = tmpSheet
                            Exit For
                        End If
                    Next
                Next
            End If

            If found Then
                result = PTRpa.visboDefaultCapacity
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkUrlaubsplaner = result
    End Function


    Private Function checkVCOrganisation(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim possibleTableNames() As String = {"VisboCenterOrganisation"}
        Dim verifiedStructure As Boolean = False
        Try

            Dim currentWS As xlns.Worksheet = Nothing
            Dim found As Boolean = False
            Dim wb As xlns.Workbook = currentWB


            For Each tmpSheet As xlns.Worksheet In currentWB.Worksheets

                For Each tblname As String In possibleTableNames
                    If tmpSheet.Name.StartsWith(tblname) Then
                        found = True
                        currentWS = tmpSheet
                        Exit For
                    End If
                Next

                If found Then
                    verifiedStructure = CStr(currentWS.Cells(1, 1).value).Trim = "name" And
                                        CStr(currentWS.Cells(1, 2).value).Trim = "uid"
                    If verifiedStructure Then
                        result = PTRpa.visboRoundtripOrga
                    End If
                    Exit For
                End If

            Next


        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try


        checkVCOrganisation = result
    End Function

    ''' <summary>
    ''' returns form Parameters the Portfolio-Name and Vname 
    ''' </summary>
    ''' <returns></returns>
    ''' 
    Public Function getNameList(ByVal blattName As String) As Collection
        Dim result As New Collection


        Try

            Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

            Dim zeile As Integer = 2
            Dim spalte As Integer = 1



            If Not IsNothing(currentWS) Then
                With currentWS
                    Dim lastRow As Integer = CType(.Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    While zeile <= lastRow
                        Dim pName As String = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        If Not IsNothing(pName) Then
                            pName = pName.Trim

                            If pName <> "" Then
                                If Not result.Contains(pName) Then
                                    result.Add(pName, pName)
                                End If
                            End If

                            zeile = zeile + 1
                        End If
                    End While

                End With
            End If
        Catch ex As Exception

        End Try

        getNameList = result
    End Function

    ''' <summary>
    ''' returns empty string for all roles / skills 
    ''' noConsideration = true: all roles / skills which need to be excluded from calculation ; empty: nothing is going to excluded 
    ''' noConsideration = false: all roles / skills which should be taken into cosideration, independent what roles / skills are ecisting; empty: all roles/skill will be considered, except those in exclusion 
    '''  
    ''' </summary>
    ''' <param name="excludedNames"></param>
    ''' <returns></returns>
    Public Function getConsiderationList(ByVal excludedNames As Boolean, ByVal Optional isRoleSkills As Boolean = True) As Collection

        Dim result As New Collection
        Dim blattName As String = "Parameters"

        Dim zeile As Integer = 3
        Dim spalte As Integer = 2
        If excludedNames Then
            zeile = 3
        Else
            zeile = 4
        End If

        Try
            Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)
            Dim lastColumn As Integer = CType(currentWS.Cells(zeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlToLeft).Column


            For columnIndex As Integer = 2 To lastColumn
                If Not IsNothing(CType(currentWS.Cells(zeile, columnIndex), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                    Dim myName As String = CStr(CType(currentWS.Cells(zeile, columnIndex), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                    If isRoleSkills Then

                        Dim myNameID As String = ""
                        If RoleDefinitions.containsName(myName) Then
                            Dim myRoleSkill As clsRollenDefinition = RoleDefinitions.getRoledef(myName)
                            If Not IsNothing(myRoleSkill) And Not result.Contains(myName) Then

                                If myRoleSkill.isSkill Then
                                    Dim containingRoleID As Integer = RoleDefinitions.getContainingRoleOfSkillMembers(myRoleSkill.UID).UID
                                    myNameID = RoleDefinitions.bestimmeRoleNameID(containingRoleID, myRoleSkill.UID)
                                Else
                                    Dim skillID As Integer = -1
                                    myNameID = RoleDefinitions.bestimmeRoleNameID(myRoleSkill.UID, skillID)
                                End If

                                result.Add(myNameID, myNameID)

                            End If
                        End If
                    Else
                        If myName <> "" And Not result.Contains(myName) Then
                            result.Add(myName, myName)
                        End If
                    End If

                End If
            Next


        Catch ex As Exception

        End Try

        getConsiderationList = result
    End Function

    Public Function getJobParameters(ByVal blattName As String, ByVal myKennung As PTRpa) As clsJobParameters
        Dim result As New clsJobParameters

        Try

            Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

            If Not IsNothing(currentWS) Then
                With currentWS

                    Select Case myKennung
                        Case PTRpa.visboFindfeasiblePortfolio

                            result.allowedOverloadMonth = CDbl(.Cells(1, 2).value)
                            result.allowedOverloadTotal = CDbl(.Cells(2, 2).value)

                            ' zeile 3
                            result.donotConsiderRoleSkills = getConsiderationList(True, True)
                            ' zeile 4
                            result.considerRoleSkills = getConsiderationList(False, True)


                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.defaultLatestEnd = CDate(.Cells(7, 2).value)
                            End If


                        Case PTRpa.visboFindProjectStart

                            result.allowedOverloadMonth = CDbl(.Cells(1, 2).value)
                            result.allowedOverloadTotal = CDbl(.Cells(2, 2).value)

                            ' zeile 3
                            result.donotConsiderRoleSkills = getConsiderationList(True, True)
                            ' zeile 4
                            result.considerRoleSkills = getConsiderationList(False, True)

                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            Else
                                result.portfolioName = ""
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            Else
                                result.portfolioVariantName = "new projects"
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            Else
                                result.projectVariantName = "fbs"
                            End If

                            If Not IsNothing(.Cells(8, 2).value) Then
                                result.defaultLatestEnd = CDate(.Cells(8, 2).value)
                            Else
                                result.defaultLatestEnd = DateSerial(Date.Now.Year + 1, 12, 31)
                            End If

                        Case PTRpa.visboFindProjectStartPM

                            result.limitPhases = CDbl(.Cells(1, 2).value)
                            result.limitMilestones = CDbl(.Cells(2, 2).value)

                            ' zeile 3
                            result.phases = getConsiderationList(True, False)
                            ' zeile 4
                            result.milestones = getConsiderationList(False, False)

                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            Else
                                result.portfolioName = ""
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            Else
                                result.portfolioVariantName = "new projects"
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            Else
                                result.projectVariantName = "fbs"
                            End If

                            If Not IsNothing(.Cells(8, 2).value) Then
                                result.defaultLatestEnd = CDate(.Cells(8, 2).value)
                            Else
                                result.defaultLatestEnd = DateSerial(Date.Now.Year + 1, 12, 31)
                            End If

                            If Not IsNothing(.Cells(9, 2).value) Then
                                result.defaultDeltaInDays = CInt(.Cells(9, 2).value)
                            End If

                        Case PTRpa.visboSuggestResourceAllocation
                            ' same as feasible Portfolio , except Line 7: project variantName
                            result.allowedOverloadMonth = CDbl(.Cells(1, 2).value)
                            result.allowedOverloadTotal = CDbl(.Cells(2, 2).value)

                            ' zeile 3
                            result.donotConsiderRoleSkills = getConsiderationList(True, True)
                            ' zeile 4
                            result.considerRoleSkills = getConsiderationList(False, True)


                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            End If


                        Case PTRpa.visboCreateHedgedVariant

                            ' only Zeile = 5 Portfolio Name is of relevance
                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            End If

                        Case PTRpa.visboAutoAdjust

                            ' only Zeile = 5 Portfolio Name is of relevance
                            If Not IsNothing(.Cells(5, 2).value) Then
                                result.portfolioName = CStr(.Cells(5, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(6, 2).value) Then
                                result.portfolioVariantName = CStr(.Cells(6, 2).value).Trim
                            End If

                            If Not IsNothing(.Cells(7, 2).value) Then
                                result.projectVariantName = CStr(.Cells(7, 2).value).Trim
                            End If


                    End Select


                End With
            Else
                Call logger(ptErrLevel.logError, "GetJobParameters: missing table Parameters in File ", currentWB.Name)
            End If
        Catch ex As Exception
            Call logger(ptErrLevel.logsevereError, "GetJobParameters: Severe Error ", ex.Message)
        End Try

        getJobParameters = result

    End Function

    Public Function getPortfolioNames() As String()

        Dim result As String()
        ReDim result(2)
        result(0) = "NoName"
        result(1) = ""
        result(2) = ""

        Dim blattName As String = "Parameters"

        Try

            Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

            If Not IsNothing(currentWS) Then
                With currentWS
                    If Not IsNothing(.Cells(5, 2).value) Then
                        result(0) = CStr(.Cells(5, 2).value).Trim
                    End If
                    If Not IsNothing(.Cells(6, 2).value) Then
                        result(1) = CStr(.Cells(6, 2).value).Trim
                    End If
                    If Not IsNothing(.Cells(7, 2).value) Then
                        result(2) = CStr(.Cells(7, 2).value).Trim
                    End If

                End With
            End If
        Catch ex As Exception

        End Try

        getPortfolioNames = result

    End Function


    Public Function getOverloadParams() As Double()

        Dim blattName As String = "Parameters"
        Dim result As Double()
        ReDim result(1)
        result(0) = 1.05
        result(1) = 1.0

        Try
            Dim currentWB As xlns.Workbook = CType(Module1.appInstance.ActiveWorkbook,
                                                            Global.Microsoft.Office.Interop.Excel.Workbook)

            Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

            If Not IsNothing(currentWS) Then
                With currentWS
                    result(0) = CDbl(.Cells(1, 2).value)
                    result(1) = CDbl(.Cells(2, 2).value)
                End With
            End If
        Catch ex As Exception
            result(0) = 1.05
            result(1) = 1.0
        End Try

        getOverloadParams = result

    End Function


    ''' <summary>
    ''' checks whether or not a file is a visbo project list 
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkJiraProjects(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "Tabelle1"

        Try

            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName), xlns.Worksheet)

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                Dim ersteZeile As xlns.Range = CType(currentWS.Rows.Item(1), xlns.Range)
                Try
                    verifiedStructure = ersteZeile.Cells(1, 1).value.trim = "Vorgangstyp" And
                        CStr(ersteZeile.Cells(1, 2).value).Trim = "Schlüssel" And
                        CStr(ersteZeile.Cells(1, 3).value).Trim = "Zusammenfassung" And
                        CStr(ersteZeile.Cells(1, 4).value).Trim = "Zugewiesene Person" And
                        CStr(ersteZeile.Cells(1, 5).value).Trim = "Autor" And
                        CStr(ersteZeile.Cells(1, 6).value).Trim = "Priorität" And
                        CStr(ersteZeile.Cells(1, 7).value).Trim = "Status" And
                        CStr(ersteZeile.Cells(1, 8).value).Trim.StartsWith("Lösung") And
                        CStr(ersteZeile.Cells(1, 9).value).Trim.StartsWith("Erstellt") And
                        CStr(ersteZeile.Cells(1, 10).value).Trim.Contains("Story Point-Schätzung") And
                        CStr(ersteZeile.Cells(1, 11).value).Trim.Contains("Aktualisiert") And
                        CStr(ersteZeile.Cells(1, 12).value).Trim = "Fälligkeitsdatum" And
                        CStr(ersteZeile.Cells(1, 13).value).Trim = "Fortschritt" And
                        CStr(ersteZeile.Cells(1, 14).value).Trim = "Erledigt" And
                        CStr(ersteZeile.Cells(1, 15).value).Trim.Contains("Übergeordnet") And
                        ersteZeile.Cells(1, 16).value.trim = "Verknüpfte Vorgänge" And
                        ersteZeile.Cells(1, 17).value.trim = "Area" And
                        ersteZeile.Cells(1, 18).value.trim = "Sprint.name" And
                        ersteZeile.Cells(1, 19).value.trim = "Sprint.startDate" And
                        ersteZeile.Cells(1, 20).value.trim = "Sprint.endDate" And
                        ersteZeile.Cells(1, 21).value.trim = "Sprint.completeDate" And
                        ersteZeile.Cells(1, 22).value.trim = "Sprint.goal" And
                        ersteZeile.Cells(1, 23).value.trim = "Start date" And
                        ersteZeile.Cells(1, 24).value.trim = "Rank" And
                        ersteZeile.Cells(1, 25).value.trim = "Projekt"

                Catch ex As Exception
                    verifiedStructure = False
                End Try

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboJira
        Else
            result = PTRpa.visboUnknown
        End If

        checkJiraProjects = result
    End Function


    ''' <summary>
    ''' checks whether or not a file is a visbo project list 
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkActualData1(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "Istdaten"

        Try

            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName), xlns.Worksheet)

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                Dim ersteZeile As xlns.Range = CType(currentWS.Rows.Item(1), xlns.Range)
                Try

                    verifiedStructure = ersteZeile.Cells(1, 1).value.trim = "Projektnummer" And
                        CStr(ersteZeile.Cells(1, 2).value).Trim = "Projekt" And
                        CStr(ersteZeile.Cells(1, 3).value).Trim = "Vorgang/Aktivität" And
                        CStr(ersteZeile.Cells(1, 4).value).Trim = "Intern/Extern" And
                        CStr(ersteZeile.Cells(1, 5).value).Trim = "Ressource/Personal-Nummer" And
                        CStr(ersteZeile.Cells(1, 6).value).Trim = "Jahr" And
                        CStr(ersteZeile.Cells(1, 7).value).Trim = "Monat" And
                        CStr(ersteZeile.Cells(1, 8).value).Trim.StartsWith("IST (PT)") And
                        CStr(ersteZeile.Cells(1, 9).value).Trim.StartsWith("IST (Euro)")

                Catch ex As Exception
                    verifiedStructure = False
                End Try

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboActualData1

        Else
            result = PTRpa.visboUnknown
        End If

        checkActualData1 = result
    End Function

    ''' <summary>
    ''' checks whether or not a file is a Timesheet of Telair
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkTagetikProjectList(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "Instructions"

        Try

            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName), xlns.Worksheet)

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                Dim zweiteZeile As xlns.Range = CType(currentWS.Rows.Item(2), xlns.Range)
                Try

                    verifiedStructure = CStr(zweiteZeile.Cells(1, 2).value).Trim.Contains("TIMESHEET")

                    ' hier muss noch geprüft werden, ob alle timesheets vorhanden, sonst in separates Dir schieben und erst wenn Timesheet-completed - file vorhanden, dann alle einlesen


                Catch ex As Exception
                    verifiedStructure = False
                End Try

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboActualData2

        Else
            result = PTRpa.visboUnknown
        End If

        checkTagetikProjectList = result
    End Function

    ''' <summary>
    ''' checks whether or not a file is a Instart Proposal - CalcSheet
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkInstartProposal(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "VISBO Summary"

        Try
            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName), xlns.Worksheet)

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                Dim firstUsefullLine As Integer = CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlDown).Row
                Dim zweiteZeile As xlns.Range = CType(currentWS.Rows.Item(firstUsefullLine), xlns.Range)
                Try

                    verifiedStructure = CStr(zweiteZeile.Cells(1, 2).value).Trim.Contains("Phase/Arbeitspaket")

                    ' hier muss noch geprüft werden, ob alle timesheets vorhanden, sonst in separates Dir schieben und erst wenn Timesheet-completed - file vorhanden, dann alle einlesen


                Catch ex As Exception
                    verifiedStructure = False
                End Try

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboProposal
        Else
            result = PTRpa.visboUnknown
        End If

        checkInstartProposal = result
    End Function

    ''' <summary>
    ''' retrieves the Portfolio Variant Definitions 
    ''' </summary>
    ''' <param name="kennung"></param>
    ''' <param name="blattname"></param>
    ''' <returns></returns>
    Public Function getPortfolioDefinitions(ByVal kennung As PTRpa,
                               Optional ByVal blattname As String = "") As clsPortfolioDefinitions

        Dim result As New clsPortfolioDefinitions

        Dim zeile As Integer = 2
        Dim spalte As Integer = 1

        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattname = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattname),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            Dim lastRow As Integer
            Dim firstZeile As xlns.Range


            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    firstZeile = CType(.Rows(1), xlns.Range)
                    lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    Dim portfolioZeile As Integer = 2

                    ' now read the portfolio definitions, if there are any 

                    If kennung = PTRpa.visboCreateHedgedVariant Then

                        Dim firstPFColumn As Integer = spalte + 4
                        Dim currentColumn As Integer = firstPFColumn

                        If Not IsNothing(CType(.Cells(1, firstPFColumn), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                            Dim myPortfolioVariantName As String = CStr(CType(.Cells(1, currentColumn), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                            If myPortfolioVariantName <> "" Then

                                Do While myPortfolioVariantName <> ""

                                    Dim abbruch As Boolean = False
                                    Dim myPName As String = ""
                                    Dim myVName As String = ""
                                    Dim myUniqueList As New Collection
                                    Dim myPortfolioList As New List(Of String)

                                    zeile = 2

                                    Do While Not abbruch

                                        Try
                                            myPName = CStr(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                            myVName = CStr(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                        Catch ex As Exception
                                            myVName = ""
                                        End Try



                                        ' now check whether or not there is a 'x' in myCurrentColumn
                                        If Not IsNothing(CType(.Cells(zeile, currentColumn), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                            Dim signal As String = CStr(CType(.Cells(zeile, currentColumn), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                                            If signal.ToLower = "x" Then

                                                If Not myUniqueList.Contains(myPName) Then
                                                    myUniqueList.Add(myPName, myPName)
                                                    Dim Key As String = calcProjektKey(myPName, myVName)
                                                    myPortfolioList.Add(Key)
                                                End If

                                            End If
                                        End If

                                        zeile = zeile + 1
                                        abbruch = (zeile > lastRow)
                                    Loop

                                    ' now add Portfolio definition 

                                    If Not result.contains(myPortfolioVariantName) Then
                                        result.addPortfolio(myPortfolioVariantName, myPortfolioList)
                                    End If


                                    ' now consider next column ...
                                    currentColumn = currentColumn + 1
                                    myPortfolioVariantName = CStr(CType(.Cells(1, currentColumn), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                Loop

                            End If


                        End If


                    End If

                End With

            End If

        Catch ex As Exception

        End Try

        getPortfolioDefinitions = result
    End Function


    Public Function getRanking(ByVal kennung As PTRpa,
                               Optional ByVal blattname As String = "") As SortedList(Of Integer, clsRankingParameters)

        Dim result As New SortedList(Of Integer, clsRankingParameters)

        Dim zeile As Integer = 2
        Dim spalte As Integer = 1

        Dim projectName As String = ""
        Dim projectVariantName As String = ""

        Dim earliestStart As Date = Date.Now.AddDays(-1 * Date.Now.Day + 1).AddMonths(1)
        Dim latestEnd As Date = Date.Now.AddDays(-1 * Date.Now.Day + 1).AddMonths(13)


        ' given in Percentage
        Dim shortestDuration As Double = 1.0
        Dim longestDuration As Double = 1.0

        Dim aktDateTime As Date = Date.Now
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If


        ' the hedge Factorm 
        Dim hedgeFactor As Double = 1.0

        Dim lastRow As Integer
        Dim firstZeile As xlns.Range



        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattname = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattname),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    firstZeile = CType(.Rows(1), xlns.Range)
                    lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    Dim portfolioZeile As Integer = 2

                    While zeile <= lastRow

                        Dim myCurrentParams As New clsRankingParameters

                        myCurrentParams.projectName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        myCurrentParams.projectVariantName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)

                        Select Case kennung

                            Case PTRpa.visboCreateHedgedVariant

                                Try
                                    Dim tmpHedgeFactor As Double = CDbl(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If tmpHedgeFactor >= 0 And tmpHedgeFactor <= 1.0 Then
                                        myCurrentParams.hedgeFactor = tmpHedgeFactor
                                    Else
                                        myCurrentParams.hedgeFactor = 0.0
                                    End If
                                Catch ex As Exception
                                    myCurrentParams.hedgeFactor = 0.0
                                End Try

                                Try
                                    Dim tmpStartDate As Date = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                    If DateDiff(DateInterval.Day, tmpStartDate, Date.Now) > 0 Then
                                        myCurrentParams.newStartDate = Date.MinValue
                                    Else
                                        myCurrentParams.newStartDate = tmpStartDate
                                    End If
                                Catch ex As Exception
                                    myCurrentParams.newStartDate = Date.MinValue
                                End Try


                            Case PTRpa.visboFindProjectStart

                                myCurrentParams.earliestStart = CDate(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.latestEnd = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.shortestDuration = CDbl(CType(.Cells(zeile, spalte + 4), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.longestDuration = CDbl(CType(.Cells(zeile, spalte + 5), Global.Microsoft.Office.Interop.Excel.Range).Value)


                            Case PTRpa.visboFindProjectStartPM

                                myCurrentParams.earliestStart = CDate(CType(.Cells(zeile, spalte + 2), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.latestEnd = CDate(CType(.Cells(zeile, spalte + 3), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.shortestDuration = CDbl(CType(.Cells(zeile, spalte + 4), Global.Microsoft.Office.Interop.Excel.Range).Value)
                                myCurrentParams.longestDuration = CDbl(CType(.Cells(zeile, spalte + 5), Global.Microsoft.Office.Interop.Excel.Range).Value)

                            Case PTRpa.visboSuggestResourceAllocation

                                Dim peopleIDs As New SortedList(Of String, Double)
                                Dim lastColumn As Integer = CType(.Cells(zeile, 2000), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlToLeft).Column

                                For i As Integer = 3 To lastColumn
                                    If Not IsNothing(CType(.Cells(zeile, i), Global.Microsoft.Office.Interop.Excel.Range).Value) Then

                                        Dim myName As String = CStr(CType(.Cells(zeile, i), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                        Dim teamID As Integer = -1

                                        If RoleDefinitions.containsNameOrID(myName) Then
                                            Dim myRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(myName, teamID)
                                            Dim myNameID As String = RoleDefinitions.bestimmeRoleNameID(myRole.UID, teamID)

                                            If Not peopleIDs.ContainsKey(myNameID) Then
                                                peopleIDs.Add(myNameID, 0)
                                            End If

                                        End If
                                    End If
                                Next

                                myCurrentParams.peopleSuggestions = peopleIDs

                            Case Else

                        End Select

                        result.Add(zeile, myCurrentParams)

                        zeile = zeile + 1


                    End While


                End With
            End If

        Catch ex As Exception

            Throw New Exception("Fehler In Portfolio-Datei" & ex.Message)
        End Try

        getRanking = result

    End Function

    ''' <summary>
    ''' returns the sequence of the project-names 
    ''' there is only one project-variant per ranking allowed
    ''' </summary>
    ''' <returns></returns>
    Public Function getRanking(Optional ByVal blattname As String = "") As SortedList(Of Integer, String)

        Dim rankingList As New SortedList(Of Integer, String)
        Dim nameList As New SortedList(Of String, String)
        Dim key As String

        Dim zeile As Integer, spalte As Integer


        Dim tfZeile As Integer = 2

        Dim pName As String = ""
        Dim variantName As String = ""

        Dim lastRow As Integer


        Dim geleseneProjekte As Integer


        Dim firstZeile As xlns.Range


        zeile = 2
        spalte = 1
        geleseneProjekte = 0




        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattname = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattname),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    firstZeile = CType(.Rows(1), xlns.Range)
                    lastRow = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    Dim portfolioZeile As Integer = 2

                    While zeile <= lastRow


                        pName = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
                        variantName = CStr(CType(.Cells(zeile, spalte + 1), Global.Microsoft.Office.Interop.Excel.Range).Value)

                        key = calcProjektKey(pName, variantName)

                        If Not nameList.ContainsKey(pName) Then
                            nameList.Add(pName, key)
                            If Not rankingList.ContainsKey(zeile) Then
                                rankingList.Add(zeile, key)
                            End If
                        End If

                        zeile = zeile + 1


                    End While


                End With
            End If

        Catch ex As Exception

            Throw New Exception("Fehler In Portfolio-Datei" & ex.Message)
        End Try

        getRanking = rankingList
    End Function


    Private Function processVisboBrief(ByVal myName As String, ByVal importDate As Date, ByRef errMessages As Collection) As Boolean

        Dim allOK As Boolean = False
        Dim aktDateTime As Date = Date.Now

        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboProject.ToString, myName)

        ' ist hier eine Projektvorlage zu importieren?
        Dim isTemplate As Boolean = LCase(myName).Contains("template")

        ' cache löschen
        Dim result As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()

        ' project brief do not need any template

        'If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 24 Then
        lastReadingOrganisation = readOrganisations()
        'End If


        'read Project Brief and put it into ImportProjekte
        Try
            Dim hproj As clsProjekt = Nothing
            Dim vproj As clsProjektvorlage = Nothing

            Dim wsGeneralInformation As xlns.Worksheet = CType(appInstance.ActiveWorkbook.Worksheets("Stammdaten"),
                    Global.Microsoft.Office.Interop.Excel.Worksheet)

            ' read the file and import into hproj


            Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)

            If isTemplate And Not IsNothing(hproj) Then
                hproj.projectType = ptPRPFType.projectTemplate
            End If

            allOK = Not IsNothing(hproj)

            If allOK Then
                Try
                    Dim keyStr As String = calcProjektKey(hproj)
                    ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                    'AlleProjekte.Add(hproj, updateCurrentConstellation:=False)

                    Call importProjekteEintragen(importDate, drawPlanTafel:=False, fileFrom3rdParty:=True, getSomeValuesFromOldProj:=True, calledFromActualDataImport:=False, calledFromRPA:=True)
                Catch ex2 As Exception
                    allOK = False
                    Call logger(ptErrLevel.logError, "RPA Error importing project brief file " & PTRpa.visboProject.ToString, ex2.Message)
                End Try
            Else
                Call logger(ptErrLevel.logError, "RPA Error importing project brief file " & PTRpa.visboProject.ToString, myName)
            End If

            ' store Project 
            If allOK Then
                allOK = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboProject.ToString, myName)

        Catch ex1 As Exception
            allOK = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Excel Brief ", ex1.Message)
        End Try

        processVisboBrief = allOK

    End Function

    Private Function readListIntoStorage(ByVal kennung As PTRpa) As Boolean
        Dim allOk As Boolean = False
        ' jetzt alle Projekte aus der Liste holen und die OverloadParams holen 

        Try
            Call logger(ptErrLevel.logInfo, "start Processing: " & kennung.ToString, "Read Projects")

            Dim listOfProjs As SortedList(Of Integer, String) = getRanking()

            For Each kvp As KeyValuePair(Of Integer, String) In listOfProjs

                Dim pname As String = getPnameFromKey(kvp.Value)
                Dim vname As String = getVariantnameFromKey(kvp.Value)
                Dim today As Date = Date.Now
                Dim hproj As clsProjekt = getProjektFromSessionOrDB(pname, vname, AlleProjekte, today)

                If Not IsNothing(hproj) Then
                    ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                End If

            Next
            allOk = True

        Catch ex As Exception
            allOk = False
        End Try
        readListIntoStorage = allOk

    End Function


    Private Function processFindProjectStart(ByVal myName As String, Optional ByVal myKennung As PTRpa = PTRpa.visboFindProjectStart) As Boolean

        Dim allOk As Boolean = False

        Try
            Dim jobParameters As clsJobParameters = getJobParameters("Parameters", myKennung)

            If jobParameters.portfolioName = "" Then
                jobParameters.portfolioName = myActivePortfolio
            End If

            Dim portfolioName As String = jobParameters.portfolioName

            Dim aggregationList As New List(Of String)
            Dim skillList As New List(Of String)


            If myKennung = PTRpa.visboFindProjectStart Then
                Call logger(ptErrLevel.logInfo, "start Processing find best Start with regard to roles & skills: ", myName)
            Else
                Call logger(ptErrLevel.logInfo, "start Processing Find best Start with regard to Phases , Milestones: ", myName)
            End If


            Dim readProjects As Integer = 0
            Dim createdProjects As Integer = 0
            'Dim importedProjects As Integer = ImportProjekte.Count

            Dim outPutCollection As New Collection



            ' jetzt alle Projekte aus der Liste holen und die OverloadParams holen 
            Try
                'Dim listOfProjs As SortedList(Of Integer, String) = getRanking()
                Dim listOfProjs As SortedList(Of Integer, clsRankingParameters) = getRanking(myKennung)

                If listOfProjs.Count > 0 Then
                    allOk = True
                Else
                    Dim msgTxt As String = "no new project names were given "
                    outPutCollection.Add(msgTxt)
                    allOk = False
                End If
                For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In listOfProjs

                    Dim pname As String = getPnameFromKey(kvp.Value.projectName)
                    Dim vname As String = getVariantnameFromKey(kvp.Value.projectVariantName)
                    Dim today As Date = Date.Now
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pname, vname, AlleProjekte, today)

                    If Not IsNothing(hproj) Then
                        ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                        allOk = allOk And True
                    Else
                        Dim msgTxt As String = "could not find " & hproj.getShapeText
                        outPutCollection.Add(msgTxt)
                        allOk = False
                    End If

                Next

            Catch ex As Exception
                allOk = False
            End Try


            If allOk Then
                Call logger(ptErrLevel.logInfo, "Project List imported: " & myName, ImportProjekte.Count & " read; ")
            Else
                Call logger(ptErrLevel.logError, "failure in reading new projects: " & myName, outPutCollection)
            End If

            If allOk Then

                Dim noActivePortfolio As Boolean = True
                Dim dbPortfolioNames As New SortedList(Of String, String)

                ' if Portfolio with active Projects is given and exists:  
                ' then we probably do have a brownfield
                If portfolioName <> "" Then

                    Dim errMsg As New clsErrorCodeMsg
                    dbPortfolioNames = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)
                    noActivePortfolio = Not dbPortfolioNames.ContainsKey(portfolioName)
                End If

                If noActivePortfolio Then
                    Call logger(ptErrLevel.logError, "no active Portfolio: " & portfolioName, myKennung.ToString)
                Else
                    ' check whether and how projects are fitting to the already existing Portfolio 
                    allOk = processProjectListWithActivePortfolio(jobParameters, myKennung)

                End If

            Else
                ' no additional logger necessary - is done in storeImportProjekte
            End If


            ' now empty the complete session  
            Call emptyRPASession()
            Call logger(ptErrLevel.logInfo, "end Processing: " & myKennung.ToString, myName)

        Catch ex As Exception
            Call logger(ptErrLevel.logError, "errors occurred when processing: " & myKennung.ToString, myName & ": " & ex.Message)
        End Try

        processFindProjectStart = allOk

    End Function

    Private Function processProjectList(ByVal myName As String, ByVal myActivePortfolio As String) As Boolean

        Dim allOk As Boolean = False
        Dim aktDateTime As Date = Date.Now

        'check the pre-conditions
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If

        'If DateDiff(DateInterval.Hour, lastReadingProjectTemplates, aktDateTime) > 24 Then
        lastReadingProjectTemplates = readProjectTemplates()
        'End If

        lastReadingCustomization = readCustomizations()

        ' cache löschen
        Dim result As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()

        Try

            Dim portfolioName As String = myName.Substring(0, myName.IndexOf(".xls"))

            Dim overloadAllowedinMonths As Double = 1.05
            Dim overloadAllowedTotal As Double = 1.0

            Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboProjectList.ToString, myName)
            Dim readProjects As Integer = 0
            Dim createdProjects As Integer = 0
            Dim importedProjects As Integer = ImportProjekte.Count

            ' now get the aggregation Roles
            Dim aggregationRoles As SortedList(Of Integer, String) = RoleDefinitions.getAggregationRoles()
            Dim aggregationList As New List(Of String)
            Dim skillList As New List(Of String)
            Dim teamID As Integer = -1

            ' checkout aggregation Roles
            For Each ar As KeyValuePair(Of Integer, String) In aggregationRoles
                Dim tmpStrID As String = RoleDefinitions.bestimmeRoleNameID(ar.Key, teamID)
                If Not aggregationList.Contains(tmpStrID) Then
                    aggregationList.Add(tmpStrID)
                End If
            Next


            Dim anzTemplates As Integer = Projektvorlagen.Count

            allOk = awinImportProjektInventur(readProjects, createdProjects)
            If allOk Then
                Call logger(ptErrLevel.logInfo, "Project List imported: " & myName, readProjects & " read; " & createdProjects & " created")
                allOk = storeImportProjekte()
            Else
                Call logger(ptErrLevel.logError, "failure in Processing: " & myName, PTRpa.visboProjectList.ToString)
            End If

        Catch ex As Exception
            Call logger(ptErrLevel.logError, "errors occurred when processing: " & PTRpa.visboProjectList.ToString, myName & ": " & ex.Message)
        End Try

        processProjectList = allOk

    End Function


    ''' <summary>
    ''' creates hedged variants for existing projects
    ''' projects need to be imported already with readListIntoStorage
    ''' </summary>
    ''' <returns></returns>
    Private Function processCreateHedgedVariants() As Boolean

        Dim result As Boolean = False
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""


        Try

            ' now get Portfolio Name and Variante-NAme 
            Dim params() As String = getPortfolioNames()
            Dim portfolioName As String = params(0)
            Dim portfolioVariantName As String = params(1)

            If portfolioName = "" Then
                portfolioName = myActivePortfolio
            End If

            Dim rankingList As SortedList(Of Integer, clsRankingParameters) = getRanking(PTRpa.visboCreateHedgedVariant)

            Dim projectVariantName As String = "hedged"
            If params(2) <> "" Then
                projectVariantName = params(2).Trim
            End If

            Dim outPutCollection As New Collection

            For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In rankingList

                Dim projectstoreRequired As Boolean = False
                Dim variantStoreRequired As Boolean = False

                Dim key As String = calcProjektKey(kvp.Value.projectName, kvp.Value.projectVariantName)
                Dim hproj As clsProjekt = ImportProjekte.getProject(key)
                If Not IsNothing(hproj) Then
                    ' now move the project 
                    If hproj.vpStatus = "initial" Or hproj.vpStatus = "initialized" Or hproj.vpStatus = "vorgeschlagen" Or hproj.vpStatus = "proposed" Then

                        Dim newStartDate As Date = kvp.Value.newStartDate
                        If DateDiff(DateInterval.Day, Date.Now, newStartDate) >= 0 Then

                            Dim deltaInDays As Integer = DateDiff(DateInterval.Day, hproj.startDate, newStartDate)

                            If deltaInDays <> 0 Then
                                Dim newEndDate As Date = hproj.endeDate.AddDays(deltaInDays)
                                If hproj.movable And Not hproj.hasActualValues Then
                                    Dim tmpProj As clsProjekt = moveProject(hproj, newStartDate, newEndDate)

                                    If Not IsNothing(tmpProj) Then
                                        If (DateDiff(DateInterval.Day, tmpProj.startDate, newStartDate) + DateDiff(DateInterval.Day, tmpProj.endeDate, newEndDate)) = 0 Then
                                            hproj = tmpProj
                                            projectstoreRequired = True
                                            msgTxt = "project moved to: " & newStartDate.ToShortDateString & " - " & newEndDate.ToShortDateString
                                            Call logger(ptErrLevel.logInfo, "process Create Hedged Variants, ", msgTxt)
                                        Else
                                            msgTxt = "project could not be moved correctly: " & hproj.getShapeText
                                            Call logger(ptErrLevel.logWarning, "process Create Hedged Variants, ", msgTxt)
                                        End If
                                    Else
                                        msgTxt = "project could not be moved at all"
                                    End If
                                Else
                                    If hproj.hasActualValues Then
                                        msgTxt = "Project with actual data can not be moved " & hproj.getShapeText
                                    Else
                                        msgTxt = "Project with Status other than initial or proposed can not be moved " & hproj.getShapeText
                                    End If

                                    Call logger(ptErrLevel.logWarning, "process Create Hedged Variants, ", msgTxt)
                                End If

                            End If

                            ' now create the variant with appropriate hedgeFactor 
                            Dim variantProj As clsProjekt = Nothing
                            If kvp.Value.hedgeFactor > 0 And kvp.Value.hedgeFactor < 1 Then
                                variantProj = hproj.createHedgedVariant(kvp.Value.hedgeFactor)
                                If Not IsNothing(variantProj) Then
                                    variantStoreRequired = True
                                End If
                            End If


                            If projectstoreRequired Then
                                outPutCollection.Clear()

                                msgTxt = hproj.getShapeText & " : " & hproj.startDate.ToShortDateString

                                ' make sure it is in AlleProjekte, becaue store Method requires it being in AlleProjekte 
                                Dim didExist As Boolean = AlleProjekte.Containskey(calcProjektKey(hproj))
                                If Not didExist Then
                                    AlleProjekte.Add(hproj, False)
                                End If

                                If storeSingleProjectToDB(hproj, outPutCollection) Then
                                    Call logger(ptErrLevel.logInfo, "project with new startDate stored: ", msgTxt)
                                    'Console.WriteLine("project with new startDate stored: " & msgTxt)
                                    'If Not setWriteProtection(hproj, False) Then
                                    '    Call logger(ptErrLevel.logWarning, "Aufheben Write PRotection did not work ...  ", hproj.getShapeText)
                                    'End If
                                Else
                                    Call logger(ptErrLevel.logError, "project store with new startDate failed: " & msgTxt, outPutCollection)
                                    'Console.WriteLine("!! ... project store with new startDate failed: " & msgTxt)
                                End If

                                If Not didExist Then
                                    AlleProjekte.Remove(calcProjektKey(hproj), False)
                                End If

                            End If

                            If variantStoreRequired Then
                                outPutCollection.Clear()

                                msgTxt = variantProj.getShapeText & " : " & variantProj.variantDescription & vbLf & " hedgeFactor: " & kvp.Value.hedgeFactor * 100 & "%"

                                Dim didExist As Boolean = AlleProjekte.Containskey(calcProjektKey(variantProj))
                                If Not didExist Then
                                    AlleProjekte.Add(variantProj, False)
                                End If
                                If storeSingleProjectToDB(variantProj, outPutCollection) Then
                                    Call logger(ptErrLevel.logInfo, "hedged variant stored: ", msgTxt)
                                Else
                                    Call logger(ptErrLevel.logError, "hedged variant store failed: " & msgTxt, outPutCollection)
                                    'Console.WriteLine("!! ... hedged variant store failed: " & msgTxt)
                                End If

                                If Not didExist Then
                                    AlleProjekte.Remove(calcProjektKey(variantProj), False)
                                End If

                            End If
                        Else
                            msgTxt = "not a appropriate new Startdate : " & hproj.getShapeText & newStartDate.ToShortDateString
                            Call logger(ptErrLevel.logInfo, "process Create Hedged Variants, ", msgTxt)
                        End If

                    Else
                        msgTxt = "Project with Status other than initial or proposed can not be moved " & hproj.getShapeText
                        Call logger(ptErrLevel.logWarning, "process Create Hedged Variants, ", msgTxt)
                    End If
                Else
                    msgTxt = "Project does not exist " & kvp.Value.projectName
                    Call logger(ptErrLevel.logWarning, "process Create Hedged Variants, ", msgTxt)
                End If

            Next

            ' now read the new portfolios 
            Dim myPortfolioVariants As clsPortfolioDefinitions = getPortfolioDefinitions(PTRpa.visboCreateHedgedVariant)

            If myPortfolioVariants.portfolioListe.Count > 0 Then

                ' get all projects of Active Portfolio , put them into AlleActiveProjects
                Dim activePortfolioProjects As New clsProjekteAlle
                If putPortfolioIntoSession(myActivePortfolio, "", activePortfolioProjects) Then

                    ' 2: now for each portfolio Variant : 
                    ' get all projects of AlleActiveProjects , put them into AlleProjekte and in ShowProjekte 



                    For Each kvp As KeyValuePair(Of String, List(Of String)) In myPortfolioVariants.portfolioListe

                        Try
                            AlleProjekte.Clear()
                            ShowProjekte.Clear()

                            For Each activeKVP As KeyValuePair(Of String, clsProjekt) In activePortfolioProjects.liste

                                AlleProjekte.Add(activeKVP.Value)
                                If Not ShowProjekte.contains(activeKVP.Value.name) Then
                                    ShowProjekte.Add(activeKVP.Value)
                                End If

                            Next

                            ' get all projects referenced in a PortfolioVariantList , put them into AlleProjekte , ShowProjekte 
                            For Each tmpKey As String In kvp.Value
                                Dim hproj As clsProjekt = ImportProjekte.getProject(tmpKey)
                                If Not IsNothing(hproj) Then
                                    AlleProjekte.Add(hproj)
                                    ShowProjekte.AddAnyway(hproj)
                                End If
                            Next

                            ' create the Portfolio and store it 
                            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                         cName:=portfolioName, vName:=kvp.Key)

                            outPutCollection.Clear()
                            Call storeSingleConstellationToDB(outPutCollection, toStoreConstellation, Nothing)

                            msgTxt = toStoreConstellation.constellationName & " ( " & toStoreConstellation.variantName & " )"

                            'Console.WriteLine("Portfolio Variant stored: " & msgTxt)
                            Call logger(ptErrLevel.logInfo, "Portfolio Variant stored: ", msgTxt)


                        Catch ex As Exception
                            Call logger(ptErrLevel.logError, "Failure when preparing store of portfolio ", ex.Message)
                        End Try


                    Next

                    result = True

                End If


            Else
                result = True
                Call logger(ptErrLevel.logInfo, "no Portfolio Variants created because no Portfolio Name for results was provided", "")
            End If



        Catch ex As Exception

        End Try


        ' now empty the complete session  
        Call emptyRPASession()

        processCreateHedgedVariants = result

    End Function

    Private Function processAutoAllocatePortfolio() As Boolean

        Dim result As Boolean = True
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""

        Dim atleastOneError As Boolean = False

        Dim outputCollection As New Collection

        Dim heute As Date = Date.Now

        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight

        ' set it back to undefined
        showRangeLeft = 0
        showRangeRight = 0

        Try

            ' now get Portfolio Name and Variante-NAme 
            Dim params() As String = getPortfolioNames()

            Dim rankingList As SortedList(Of Integer, clsRankingParameters) = getRanking(PTRpa.visboSuggestResourceAllocation)


            Dim portfolioName As String = params(0)
            Dim variantName As String = params(1)

            Dim projectVariantName As String = "auto"
            If params(2) <> "" Then
                projectVariantName = params(2).Trim
            End If

            Dim autoAllocate As Boolean = True

            ShowProjekte.Clear()
            AlleProjekte.Clear()
            ImportProjekte.Clear()
            projectConstellations.Clear()

            ' now load the the portfolio and all projects of portfolio 
            ' hole Portfolio (pName,vName) aus der db
            Dim cTime As Date = heute
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(portfolioName,
                                                                                               "", cTime, Err, variantName:=variantName, storedAtOrBefore:=heute)


            If Not IsNothing(myConstellation) Then

                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & myActivePortfolio, " start of Operation ... ")
                ' tmpname in die Session-Liste wieder aufnehmen

                projectConstellations.Add(myConstellation)

                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, heute)

                    If Not IsNothing(hproj) Then

                        ' now Create a Variant from that , if it is not already the very same variant
                        If hproj.variantName <> projectVariantName Then
                            hproj = hproj.createVariant(projectVariantName, "auto-created variant")
                        End If

                        ImportProjekte.Add(hproj, updateCurrentConstellation:=False)

                    Else
                        Call logger(ptErrLevel.logWarning, "Loading " & kvp.Key & " failed ..", " Operation continued ...")
                        'Console.WriteLine("Loading " & kvp.Key & " failed ..", " Operation continued ...")
                        atleastOneError = True
                    End If

                Next

                ' now do the operation 

                For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In rankingList

                    Dim myProj As clsProjekt = ImportProjekte.getProjectbyName(kvp.Value.projectName)
                    Dim fmsg As String = ""

                    If Not IsNothing(myProj) Then

                        AlleProjekte.Add(myProj)
                        ShowProjekte.AddAnyway(myProj)

                        Call ShowProjekte.autoAllocate(myProj.name, "", False, fmsg, suggestedIDs:=kvp.Value.peopleSuggestions)

                        If fmsg = "" Then

                            outputCollection.Clear()
                            Dim storeProj As clsProjekt = ShowProjekte.getProject(myProj.name)

                            If Not IsNothing(storeProj) Then
                                If Not atleastOneError Then
                                    If storeSingleProjectToDB(storeProj, outputCollection) Then
                                        Call logger(ptErrLevel.logInfo, "project team allocated and stored: ", storeProj.getShapeText)
                                        'Console.WriteLine("project team allocated and stored: " & storeProj.getShapeText)

                                        'If Not setWriteProtection(storeProj, False) Then
                                        '    Call logger(ptErrLevel.logWarning, "Aufheben Write PRotection did not work ...  ", storeProj.getShapeText)
                                        'End If
                                    Else
                                        Call logger(ptErrLevel.logError, "store project failed : " & storeProj.getShapeText, outputCollection)
                                        'Console.WriteLine("!! ... store project team allocation failed : " & storeProj.getShapeText)
                                        atleastOneError = True
                                    End If
                                Else
                                    Call logger(ptErrLevel.logInfo, "because former Error occurred , no store happened : " & storeProj.getShapeText, outputCollection)
                                    'Console.WriteLine("!! ... because former Error occurred , no store happened: " & storeProj.getShapeText)
                                    atleastOneError = True
                                End If

                            End If

                        Else
                            ' failure 
                            atleastOneError = True
                            Call logger(ptErrLevel.logError, "Auto-Allocation failure: " & kvp.Key & " " & fmsg, " ... Operation continued ...")
                            'Console.WriteLine("!! ... Auto-Allocation failure : " & myProj.getShapeText)
                        End If
                    Else
                        ' failure 
                        atleastOneError = True
                        Call logger(ptErrLevel.logError, "Auto-Allocation failure: could not read " & kvp.Value.projectName & " " & kvp.Value.projectVariantName, " ... Operation continued ...")
                        'Console.WriteLine("Auto-Allocation failure: could not read " & kvp.Value.projectName & " " & kvp.Value.projectVariantName)

                    End If


                Next

                Call logger(ptErrLevel.logInfo, "Auto-Allocating Projects from Portfolio " & myActivePortfolio, " End of Operation ... ")

                ' now create the according Portfolio


            Else
                atleastOneError = True
                msgTxt = "Load Portfolio " & myActivePortfolio & " failed .."
                Call logger(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio, " failed ..")
                'Console.WriteLine(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio & " failed ..")
                Throw New ArgumentException(msgTxt)
            End If

            If Not atleastOneError Then

                Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:=projectVariantName)

                outputCollection.Clear()
                Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

                If outputCollection.Count > 0 Then
                    Call logger(ptErrLevel.logInfo, "Project List with Active Portfolio: ", outputCollection)
                    'Console.WriteLine("Portfolio created " & portfolioName & " ( " & projectVariantName & " )")
                End If

            Else
                Call logger(ptErrLevel.logInfo, "no store of portfolio because of former Error .. ", "")
                'Console.WriteLine("no store of portfolio because of former Error .. ")
            End If


        Catch ex As Exception

        End Try

        ' restore values 
        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        Call emptyRPASession()

        processAutoAllocatePortfolio = result

    End Function

    ''' <summary>
    ''' create a Portfolio Variant and according project variants in a way that there are no more any bottlenecks at people base
    ''' all projects, except those in the exception list will be handeld; i.e values distributed so that no person is being overloaded, 
    ''' values geing beyond that are assigned the according summary role and in a second step being auto-Allocated 
    ''' </summary>
    ''' <returns></returns>
    Private Function processAutoAdjustPortfolio() As Boolean

        Dim result As Boolean = True
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""



        Dim outputCollection As New Collection

        Dim heute As Date = Date.Now

        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight

        ' set it back to undefined
        showRangeLeft = 0
        showRangeRight = 0

        Try

            ' now get Portfolio Name and Variante-NAme 
            Dim params() As String = getPortfolioNames()
            Dim exceptionList As Collection = getNameList("Exception List")

            Dim portfolioName As String = params(0)
            Dim variantName As String = params(1)

            Dim projectVariantName As String = "auto"
            If params(2) <> "" Then
                projectVariantName = params(2).Trim
            End If

            Dim autoAllocate As Boolean = True

            ShowProjekte.Clear()
            AlleProjekte.Clear()
            projectConstellations.Clear()

            ' now load the the portfolio and all projects of portfolio 
            ' hole Portfolio (pName,vName) aus der db
            Dim cTime As Date = heute
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(portfolioName,
                                                                                               "", cTime, Err, variantName:=variantName, storedAtOrBefore:=heute)


            If Not IsNothing(myConstellation) Then

                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & myActivePortfolio, " start of Operation ... ")
                ' tmpname in die Session-Liste wieder aufnehmen

                projectConstellations.Add(myConstellation)

                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, heute)

                    If Not IsNothing(hproj) Then

                        ' now Create a Variant from that , if it is not in the exception list 
                        If Not exceptionList.Contains(hproj.name) Then
                            If hproj.variantName <> projectVariantName Then
                                hproj = hproj.createVariant(projectVariantName, "auto-created variant")
                            End If
                        End If

                        AlleProjekte.Add(hproj)

                        ' if it is already in ShowProjekte: remove it , then add this one 
                        ShowProjekte.AddAnyway(hproj)

                    Else
                        Call logger(ptErrLevel.logWarning, "Loading " & kvp.Key & " failed ..", " Operation continued ...")
                    End If

                Next

                ' now do the operation 
                For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                    If Not exceptionList.Contains(kvp.Value.name) And Not kvp.Value.hasActualValues Then
                        Dim fmsg As String = ""
                        Call ShowProjekte.autoDistribute(kvp.Value.name, "", fmsg)

                        If fmsg = "" Then
                            ' success
                            Call logger(ptErrLevel.logInfo, "Adjustment successful: " & kvp.Key, " ... Operation continued ...")

                            Call ShowProjekte.autoAllocate(kvp.Value.name, "", True, fmsg)

                            If fmsg = "" Then

                                outputCollection.Clear()

                                If storeSingleProjectToDB(kvp.Value, outputCollection) Then
                                    Call logger(ptErrLevel.logInfo, "project variant adjusted and stored: ", kvp.Value.getShapeText)
                                    'Console.WriteLine("project stored: " & kvp.Value.getShapeText)

                                    'If Not setWriteProtection(kvp.Value, False) Then
                                    '    Call logger(ptErrLevel.logWarning, "Aufheben Write PRotection did not work ...  ", kvp.Value.getShapeText)
                                    'End If
                                Else
                                    Call logger(ptErrLevel.logError, "project variant store failed: " & kvp.Value.getShapeText, outputCollection)
                                    'Console.WriteLine("!! ... project store failed: " & kvp.Value.getShapeText)
                                End If

                                Call logger(ptErrLevel.logInfo, "Auto-Allocation successful: " & kvp.Key, " ... Operation continued ...")
                            Else
                                ' failure 
                                Call logger(ptErrLevel.logError, "Auto-Allocation failure: " & kvp.Key & " " & fmsg, " ... Operation continued ...")
                            End If
                        Else
                            ' failure 
                            Call logger(ptErrLevel.logError, "Adjustment failure: " & kvp.Key & " " & fmsg, " ... Operation continued ...")
                        End If
                    Else
                        ' 
                        Call logger(ptErrLevel.logInfo, "not adjusted because it is in Exception List or is having actual values: " & kvp.Key, " Operation continued ...")
                    End If
                Next

                Call logger(ptErrLevel.logInfo, "Adjusting Projects from Portfolio " & myActivePortfolio, " End of Operation ... ")

                ' now create the according Portfolio


            Else
                msgTxt = "Load Portfolio " & myActivePortfolio & " failed .."
                Call logger(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio, " failed ..")
                Throw New ArgumentException(msgTxt)
            End If

            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:=projectVariantName)

            outputCollection.Clear()
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

            If outputCollection.Count > 0 Then
                Call logger(ptErrLevel.logInfo, "Project List with Active Portfolio: ", outputCollection)
            End If

        Catch ex As Exception

        End Try




        ' restore values 
        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        Call emptyRPASession()

        processAutoAdjustPortfolio = result
    End Function


    ''' <summary>
    ''' reads all projects of a given Portfolio into storage: all projects are then within clsProjekteAlle 'sessionListe'
    ''' updateCurrentConstellation is by default set to false, i.e a currentSessionConstellation is not defined by that in Default. 
    ''' </summary>
    ''' <param name="myPortfolioName"></param>
    ''' <param name="myPortfolioVName"></param>
    ''' <param name="sessionListe"></param>
    ''' <returns></returns>
    Private Function putPortfolioIntoSession(ByVal myPortfolioName As String, ByVal myPortfolioVName As String, ByRef sessionListe As clsProjekteAlle,
                                             Optional ByVal upDateCurrentConstellation As Boolean = False) As Boolean

        Dim allOk As Boolean = False
        Dim Err As New clsErrorCodeMsg


        Try
            sessionListe.Clear(upDateCurrentConstellation)

            If myPortfolioName = "" Then
                myPortfolioName = myActivePortfolio
                myPortfolioVName = ""
            End If

            Dim heute As Date = Date.Now
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(myPortfolioName,
                                                                                           "", heute, Err, variantName:=myPortfolioVName, storedAtOrBefore:=heute)


            If Not IsNothing(myConstellation) Then

                ' tmpname in die Session-Liste wieder aufnehmen
                projectConstellations.Add(myConstellation)
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, sessionListe, heute)
                    If Not IsNothing(hproj) Then

                        sessionListe.Add(hproj, upDateCurrentConstellation)

                    End If
                Next

                allOk = True
            Else
                Dim msgTxt As String = "Load Portfolio " & myPortfolioName & " failed .."
                Call logger(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio, " failed ..")
                allOk = False
            End If

        Catch ex As Exception
            allOk = False
        End Try

        putPortfolioIntoSession = allOk

    End Function





    ''' <summary>
    ''' in ImportProjekte sind alle aktuell eingelesenen Projekte 
    ''' </summary>
    ''' <returns></returns>
    Private Function processProjectListWithActivePortfolio(ByVal jobParameters As clsJobParameters, ByVal myKennung As PTRpa) As Boolean
        Dim result As Boolean = True
        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""

        Dim heute As Date = Date.Now

        ' cache löschen
        Dim result0 As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()

        Dim myActivePortfolio As String = jobParameters.portfolioName
        Dim portfolioVariantName As String = jobParameters.portfolioVariantName

        Dim projectVariantName As String = jobParameters.projectVariantName
        If projectVariantName = "" Then
            projectVariantName = "arb"
        End If

        Dim aggregationList As New List(Of String)
        Dim skillList As New List(Of String)

        ' for showing which projects could not be considered
        Dim missingList As New clsProjekteAlle

        If myKennung = PTRpa.visboFindProjectStart Then
            ' build aggregation List
            ' now get the aggregation Roles
            Dim aggregationRoles As SortedList(Of Integer, String) = RoleDefinitions.getAggregationRoles()
            Dim teamID As Integer = -1


            ' currently only Exclude of Roles & Skills is supported ..
            ' checkout aggregation Roles
            For Each ar As KeyValuePair(Of Integer, String) In aggregationRoles
                Dim tmpStrID As String = RoleDefinitions.bestimmeRoleNameID(ar.Key, teamID)
                If Not aggregationList.Contains(tmpStrID) Then
                    If jobParameters.donotConsiderRoleSkills.Count = 0 Then
                        aggregationList.Add(tmpStrID)
                    Else
                        If Not jobParameters.donotConsiderRoleSkills.Contains(ar.Value) Then
                            aggregationList.Add(tmpStrID)
                        End If
                    End If
                End If
            Next

            ' build Skill List 
            Dim skillIDs As Collection = ImportProjekte.getRoleSkillIDs()


            For Each si As String In skillIDs
                If Not skillList.Contains(si) Then
                    skillList.Add(si)
                End If

                ' new  
                If Not skillList.Contains(si) Then
                    If jobParameters.donotConsiderRoleSkills.Count = 0 Then
                        skillList.Add(si)
                    Else
                        If Not jobParameters.donotConsiderRoleSkills.Contains(si) Then
                            skillList.Add(si)
                        End If
                    End If
                End If
            Next
        End If



        Try
            ShowProjekte.Clear()
            AlleProjekte.Clear()

            ' now load the existing portfolio and all projects of portfolio 
            ' hole Portfolio (pName,vName) aus der db
            Dim cTime As Date = heute
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(myActivePortfolio,
                                                                                               "", cTime, Err, variantName:="", storedAtOrBefore:=heute)


            ' tk 1.5. 
            Dim nextLineNumber As Integer = myConstellation.getMaxRowNumber + 1

            If Not IsNothing(myConstellation) Then
                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & myActivePortfolio, " start of Operation ... ")
                ' tmpname in die Session-Liste wieder aufnehmen
                projectConstellations.Add(myConstellation)

                ' rowNr ist needed to use it definietion of Portfolio sorting / sequence

                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, heute)

                    Try
                        hproj.tfZeile = myConstellation.getBoardZeile(pName)
                        If hproj.tfZeile > nextLineNumber Then
                            nextLineNumber = hproj.tfZeile + 1
                        End If
                    Catch ex As Exception
                        hproj.tfZeile = 2
                    End Try

                    If Not IsNothing(hproj) Then

                        AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                        ' removes hproj from ShowProjekte, if already in there
                        ShowProjekte.AddAnyway(hproj)

                    Else
                        Call logger(ptErrLevel.logWarning, "Loading " & kvp.Key & " failed ..", " Operation continued ...")
                    End If

                Next

                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & myActivePortfolio, " End of Operation ... ")

            Else
                msgTxt = "Load Portfolio " & myActivePortfolio & " failed .."
                Call logger(ptErrLevel.logError, "Load Portfolio " & myActivePortfolio, " failed ..")
                Throw New ArgumentException(msgTxt)
            End If

            ' get the ranking list 
            'Dim rankingList As SortedList(Of Integer, String) = getRanking()
            Dim rankingList As SortedList(Of Integer, clsRankingParameters) = getRanking(myKennung)


            ' now create a Portfolio variant with unchanged new projects ...
            Dim removeSPList As New List(Of String)
            Dim removeAPList As New List(Of String)

            Dim first As Boolean = True
            Dim minMonthColumn As Integer = 0
            Dim maxMonthColumn As Integer = 0

            Dim outputCollection As New Collection

            Dim myRowNr As Integer = nextLineNumber
            For Each rankingPair As KeyValuePair(Of Integer, clsRankingParameters) In rankingList
                Dim key As String = calcProjektKey(rankingPair.Value.projectName, rankingPair.Value.projectVariantName)
                Dim hproj As clsProjekt = ImportProjekte.getProject(key)
                If Not IsNothing(hproj) Then

                    ' bestimme die Line Number 
                    hproj.tfZeile = myRowNr

                    If first Then
                        first = False
                        minMonthColumn = getColumnOfDate(hproj.startDate)
                        maxMonthColumn = getColumnOfDate(hproj.endeDate)
                    Else
                        Dim myMin As Integer = getColumnOfDate(hproj.startDate)
                        Dim myMax As Integer = getColumnOfDate(hproj.endeDate)
                        If myMin < minMonthColumn Then
                            minMonthColumn = myMin
                        End If
                        If myMax > maxMonthColumn Then
                            maxMonthColumn = myMax
                        End If
                    End If

                    ' check whether or not project is beginning after today ..
                    If DateDiff(DateInterval.Day, Date.Now, hproj.startDate) < 0 Then
                        outputCollection.Add(hproj.getShapeText)
                    End If

                    If Not AlleProjekte.Containskey(key) Then
                        AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                        removeAPList.Add(key)
                    Else
                        ' bring updated hproj into AlleProjekte
                        AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                    End If

                    If Not ShowProjekte.contains(hproj.name) Then
                        ShowProjekte.Add(hproj)
                        removeSPList.Add(hproj.name)
                    Else
                        ShowProjekte.AddAnyway(hproj)
                    End If


                End If

                myRowNr = myRowNr + 1
            Next

            ' now Check whether or not minMonthCol ist in Future, if not end it , because that is not allowed 
            If minMonthColumn < getColumnOfDate(Date.Now) + 1 Then
                Call logger(ptErrLevel.logError, "Find best Project Start - no projects allowed to start today or before today ", outputCollection)
                result = False
                ShowProjekte.Clear()
                AlleProjekte.Clear()
                ImportProjekte.Clear(False)
            End If

            If result = True Then


                Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                cName:=myActivePortfolio, vName:=portfolioVariantName)


                Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

                If outputCollection.Count > 0 Then
                    Call logger(ptErrLevel.logInfo, "Project List Import, Store Portfolio-Variant " & portfolioVariantName & " result:", outputCollection)
                End If

                ' now rest Showprojekte to formerStatus 
                For Each tmpName As String In removeAPList
                    AlleProjekte.Remove(tmpName)
                Next

                For Each tmpName As String In removeSPList
                    ShowProjekte.Remove(tmpName)
                Next

                ' nach dem Remove ist evtl wieder das SortCriteria zurückgesetzt
                ' die Sortierung wieder nach Zeile herstellen 
                currentSessionConstellation.sortCriteria = ptSortCriteria.customTF


                ' now check whether there are overutilizations 
                ' if so , move showRangeLeft and showrangeRight  1 by 1 , until there are no overutilizations any more 

                showRangeLeft = minMonthColumn
                showRangeRight = maxMonthColumn
                Dim stopValue As Integer = showRangeRight

                Dim overutilizationFound As Boolean = False
                Dim referenceMSValues As Double() = Nothing
                Dim referencePHValues As Double() = Nothing

                Dim sumIterations As Integer = 0

                If myKennung = PTRpa.visboFindProjectStart Then
                    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, jobParameters.allowedOverloadMonth, jobParameters.allowedOverloadTotal)
                End If


                If overutilizationFound And myKennung = PTRpa.visboFindProjectStart Then
                    msgTxt = "there are already resource bottlenecks in the starting portfolio " & myActivePortfolio
                    Call logger(ptErrLevel.logError, msgTxt, " please solve this first before considering new projects ... calculation stopped ..")
                    result = False
                End If
                '

                If result = True Then



                    ' create variant , if necessary
                    ' rankingList keeps the sequence within the Excel file. So user adds some fields important to him for prioritization , he add these fields , sorts it in th eExcel. 
                    ' It then represents the sequence: Row1 is the most important project 

                    myRowNr = nextLineNumber

                    For Each rankingPair As KeyValuePair(Of Integer, clsRankingParameters) In rankingList

                        sumIterations = 0
                        Dim key As String = calcProjektKey(rankingPair.Value.projectName, rankingPair.Value.projectVariantName)
                        Dim hproj As clsProjekt = ImportProjekte.getProject(key)

                        If Not IsNothing(hproj) Then

                            Try
                                hproj.tfZeile = myRowNr
                            Catch ex As Exception

                            End Try

                            Dim stdDuration As Integer = hproj.dauerInDays
                            Dim myDuration As Integer = stdDuration
                            'Dim minDuration As Integer = CInt(stdDuration * 0.7)
                            Dim minDuration As Integer = CInt(stdDuration * rankingPair.Value.shortestDuration)
                            Dim latestEndDate As Date = rankingPair.Value.latestEnd
                            Dim biggestOffsettoEnd As Integer = 0

                            If DateDiff(DateInterval.Day, hproj.endeDate, latestEndDate) > 0 Then
                                biggestOffsettoEnd = DateDiff(DateInterval.Day, hproj.endeDate, latestEndDate)
                            End If

                            Dim storeRequired As Boolean = False

                            Dim newStartDate As Date = hproj.startDate
                            Dim newEndDate As Date = hproj.endeDate


                            ' now define showrangeLeft and showrangeRight from hproj 
                            showRangeLeft = getColumnOfDate(hproj.startDate)
                            showRangeRight = getColumnOfDate(hproj.endeDate)

                            ' have to happen here because just before hproj is added to ShowProjekte, find out what the situation is before ...
                            If myKennung = PTRpa.visboFindProjectStartPM Then
                                ' now define the reference Values for Phases and Milestones 
                                referenceMSValues = ShowProjekte.getMilestonesFrequency(jobParameters.getMilestoneNames)
                                referencePHValues = ShowProjekte.getPhaseFrequency(jobParameters.getPhaseNames)
                            Else
                                ' now define skill-List, because it is good enough to only consider skills of the hproj under consideration
                                skillList.Clear()
                                Dim skillIDs As Collection = hproj.getSkillNameIds

                                For Each si As String In skillIDs
                                    If Not skillList.Contains(si) Then
                                        If jobParameters.donotConsiderRoleSkills.Count = 0 Then
                                            skillList.Add(si)
                                        Else
                                            If Not jobParameters.donotConsiderRoleSkills.Contains(si) Then
                                                skillList.Add(si)
                                            End If
                                        End If
                                    End If
                                Next
                            End If


                            ' check auf Exists is not necessary with AlleProjekte, because it will be replaced if it already exists 
                            AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                            ShowProjekte.AddAnyway(hproj)


                            If myKennung = PTRpa.visboFindProjectStart Then
                                overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, jobParameters.allowedOverloadMonth, jobParameters.allowedOverloadTotal)
                            Else
                                overutilizationFound = ShowProjekte.overLoadMSPhasesFound(jobParameters.getMilestoneNames, jobParameters.limitMilestones,
                                                                                          referenceMSValues,
                                                                                          jobParameters.getPhaseNames, jobParameters.limitPhases,
                                                                                          referencePHValues)
                            End If



                            If overutilizationFound Then

                                ' create variant if not already done
                                If hproj.variantName <> projectVariantName Then
                                    hproj = hproj.createVariant(projectVariantName, "variant to avoid resource bottlenecks")
                                    AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                                End If

                                Dim deltaInDays As Integer = jobParameters.defaultDeltaInDays
                                ' now modify this one ...

                                Dim endIterations As Integer = 0
                                Dim durationIterations As Integer = 0

                                Dim maxEndIterations As Integer = CInt(biggestOffsettoEnd / deltaInDays)
                                Dim maxDurationIterations As Integer = CInt((stdDuration - minDuration) / deltaInDays)

                                Dim rememberStartDate As Date = hproj.startDate
                                Dim rememberEndDate As Date = hproj.endeDate

                                Try
                                    Dim tmpProj As clsProjekt = Nothing
                                    Do While overutilizationFound And endIterations <= maxEndIterations
                                        ' move project by deltaIndays

                                        newStartDate = rememberStartDate.AddDays(deltaInDays)
                                        durationIterations = 1

                                        Do While overutilizationFound And durationIterations <= maxDurationIterations

                                            newEndDate = rememberEndDate
                                            tmpProj = moveProject(hproj, newStartDate, newEndDate)
                                            sumIterations = sumIterations + 1

                                            If Not IsNothing(tmpProj) Then

                                                hproj = tmpProj

                                                ' now define showrangeLeft and showrangeRight from hproj 
                                                showRangeLeft = getColumnOfDate(hproj.startDate)
                                                showRangeRight = getColumnOfDate(hproj.endeDate)

                                                If myKennung = PTRpa.visboFindProjectStartPM Then
                                                    ' aus ShowProjekte rausnehmen, um ReferenzValues zu bestimmen 
                                                    If ShowProjekte.contains(hproj.name) Then
                                                        ShowProjekte.Remove(hproj.name, False)
                                                    End If
                                                    referenceMSValues = ShowProjekte.getMilestonesFrequency(jobParameters.getMilestoneNames)
                                                    referencePHValues = ShowProjekte.getPhaseFrequency(jobParameters.getPhaseNames)
                                                End If

                                                ' now replace in AlleProjekte, ShowProjekte 
                                                AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                                                ShowProjekte.AddAnyway(hproj)

                                                Dim infomsg As String = "... trying out " & hproj.getShapeText & hproj.startDate.ToShortDateString & " - " & hproj.endeDate.ToShortDateString
                                                'Console.WriteLine(infomsg)
                                                Call logger(ptErrLevel.logInfo, "find best start ", infomsg)

                                                If myKennung = PTRpa.visboFindProjectStart Then
                                                    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, jobParameters.allowedOverloadMonth, jobParameters.allowedOverloadTotal)
                                                Else
                                                    overutilizationFound = ShowProjekte.overLoadMSPhasesFound(jobParameters.getMilestoneNames, jobParameters.limitMilestones,
                                                                                          referenceMSValues,
                                                                                          jobParameters.getPhaseNames, jobParameters.limitPhases,
                                                                                          referencePHValues)
                                                End If



                                            Else
                                                ' Error occurred 
                                                Throw New ArgumentException("tmpProj is Nothing")
                                            End If

                                            newStartDate = newStartDate.AddDays(deltaInDays)
                                            durationIterations = durationIterations + 1
                                        Loop

                                        If overutilizationFound Then

                                            rememberStartDate = rememberStartDate.AddDays(deltaInDays)
                                            rememberEndDate = rememberEndDate.AddDays(deltaInDays)

                                            tmpProj = moveProject(hproj, rememberStartDate, rememberEndDate)
                                            ' 

                                            sumIterations = sumIterations + 1

                                            If Not IsNothing(tmpProj) Then

                                                hproj = tmpProj

                                                ' now define showrangeLeft and showrangeRight from hproj 
                                                showRangeLeft = getColumnOfDate(hproj.startDate)
                                                showRangeRight = getColumnOfDate(hproj.endeDate)

                                                If myKennung = PTRpa.visboFindProjectStartPM Then
                                                    ' aus ShowProjekte rausnehmen, um ReferenzValues zu bestimmen 
                                                    If ShowProjekte.contains(hproj.name) Then
                                                        ShowProjekte.Remove(hproj.name, False)
                                                    End If
                                                    referenceMSValues = ShowProjekte.getMilestonesFrequency(jobParameters.getMilestoneNames)
                                                    referencePHValues = ShowProjekte.getPhaseFrequency(jobParameters.getPhaseNames)
                                                End If


                                                ' now replace in AlleProjekte, ShowProjekte 
                                                AlleProjekte.Add(hproj, sortkey:=hproj.tfZeile)
                                                ShowProjekte.AddAnyway(hproj)

                                                Dim infomsg As String = "... trying out " & hproj.getShapeText & hproj.startDate.ToShortDateString & " - " & hproj.endeDate.ToShortDateString
                                                Call logger(ptErrLevel.logInfo, "find best start ", infomsg)


                                                If myKennung = PTRpa.visboFindProjectStart Then
                                                    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, jobParameters.allowedOverloadMonth, jobParameters.allowedOverloadTotal)
                                                Else
                                                    overutilizationFound = ShowProjekte.overLoadMSPhasesFound(jobParameters.getMilestoneNames, jobParameters.limitMilestones,
                                                                                          referenceMSValues,
                                                                                          jobParameters.getPhaseNames, jobParameters.limitPhases,
                                                                                          referencePHValues)
                                                End If

                                            Else
                                                ' Error occurred 
                                                Throw New ArgumentException("tmpProj is Nothing")
                                            End If
                                        End If

                                        endIterations = endIterations + 1
                                    Loop

                                Catch ex As Exception
                                    Dim infomsg As String = "failure: could not create project-variant " & hproj.getShapeText & ex.Message
                                    Call logger(ptErrLevel.logError, "find best start ", infomsg)
                                    overutilizationFound = True
                                End Try

                                If Not overutilizationFound Then
                                    ' it is already in there ... but now needed to be stored
                                    storeRequired = True
                                Else
                                    ' take it out again , because there was no solution
                                    AlleProjekte.Remove(calcProjektKey(hproj))
                                    ShowProjekte.Remove(hproj.name)
                                End If

                            Else
                                ' all ok, just continue
                            End If

                            If storeRequired Then
                                Dim myMessages As New Collection
                                If storeSingleProjectToDB(hproj, myMessages) Then
                                    ' now for the sake of sequence in Constellation 
                                    myRowNr = myRowNr + 1

                                    Dim infomsg As String = "success: created " & sumIterations & " variants to avoid bottlenecks " & hproj.getShapeText
                                    Call logger(ptErrLevel.logInfo, "find best start ", infomsg)
                                Else
                                    ' take it out again , because there was no solution
                                    ShowProjekte.Remove(hproj.name)
                                    Dim infomsg As String = "... failed to store variant to avoid bottlenecks " & hproj.getShapeText
                                    Call logger(ptErrLevel.logError, "find best start ", infomsg)
                                End If
                            Else
                                If overutilizationFound Then
                                    ' for showing which projects could not be considered
                                    missingList.Add(hproj)
                                    Dim infomsg As String = "unsuccessful : tried out " & sumIterations & " variants for " & hproj.name
                                    Call logger(ptErrLevel.logWarning, "find best start ", infomsg)
                                Else
                                    ' now for the sake of sequence in Constellation 
                                    myRowNr = myRowNr + 1

                                    Dim infomsg As String = "success: could be added to portfolio variant as-is " & hproj.getShapeText
                                    Call logger(ptErrLevel.logInfo, "find best start ", infomsg)
                                End If

                            End If
                        Else
                            Dim infomsg As String = rankingPair.Value.projectName & " does not exist so far"
                            Call logger(ptErrLevel.logError, "find best start ", infomsg)
                        End If

                    Next

                    ' now consider to sort it again 
                    If myKennung = PTRpa.visboFindProjectStartPM Then
                        ' do the sorting according the very first phase , or if there is no phase according the first milestones 
                        ' start , then duration 
                        Dim isMilestone As Boolean = False
                        Dim sortItem As String = ""
                        If jobParameters.getPhaseNames.Count > 0 Then
                            sortItem = jobParameters.getPhaseNames.Item(0)
                        ElseIf jobParameters.getMilestoneNames.Count > 0 Then
                            sortItem = jobParameters.getMilestoneNames.Item(0)
                            isMilestone = True
                        End If

                        Dim myNewSortlist As New SortedList(Of Double, String)
                        Try

                            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                                Dim sortKrit As Double = 100000000.0

                                If Not isMilestone Then
                                    ' sort Kriterium is phase start and duration 
                                    Dim elemName As String = ""
                                    Dim breadCrumb As String = ""
                                    Dim lfdNr As Integer = 1
                                    Dim type As Integer = -1
                                    Dim pvName As String = ""
                                    Call splitHryFullnameTo2(sortItem, elemName, breadCrumb, type, pvName)
                                    Dim cPhase As clsPhase = kvp.Value.getPhase(elemName, lfdNr:=1)
                                    sortKrit = DateDiff(DateInterval.Day, StartofCalendar, cPhase.getStartDate) + 1 / cPhase.dauerInDays

                                    Dim kritDelta As Double = 0.00001
                                    Do While myNewSortlist.ContainsKey(sortKrit)
                                        sortKrit = sortKrit + kritDelta
                                    Loop

                                Else
                                    ' sort Kriterium is phase start and duration 
                                    Dim elemName As String = ""
                                    Dim breadCrumb As String = ""
                                    Dim lfdNr As Integer = 1
                                    Dim type As Integer = -1
                                    Dim pvName As String = ""
                                    Call splitHryFullnameTo2(sortItem, elemName, breadCrumb, type, pvName)
                                    Dim cMilestone As clsMeilenstein = kvp.Value.getMilestone(elemName, lfdNr:=1)
                                    sortKrit = DateDiff(DateInterval.Day, StartofCalendar, cMilestone.getDate)

                                    Dim kritDelta As Double = 0.00001
                                    Do While myNewSortlist.ContainsKey(sortKrit)
                                        sortKrit = sortKrit + kritDelta
                                    Loop
                                End If
                                ' jetzt enthält die Sortierte Liste den Eintrag nicht mehr 
                                myNewSortlist.Add(sortKrit, kvp.Value.name)

                            Next

                            ' now re-define the sorting in the constellation
                            Dim tmpSortListe As New SortedList(Of String, String)
                            Dim zeile As Integer = 2
                            Dim formatStr As String = "00000000"
                            For Each kvp2 As KeyValuePair(Of Double, String) In myNewSortlist
                                tmpSortListe.Add(zeile.ToString(formatStr), kvp2.Value)
                                zeile = zeile + 1
                            Next

                            ' jetzt das als die Sort-Liste der Konstellation übernehmen 
                            currentSessionConstellation.sortListe(ptSortCriteria.customTF) = tmpSortListe
                        Catch ex As Exception

                        End Try
                    End If

                    Dim pfVariantName As String = jobParameters.portfolioVariantName & " - " & projectVariantName
                    toStoreConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                                cName:=myActivePortfolio, vName:=pfVariantName)

                    outputCollection.Clear()
                    Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

                    If outputCollection.Count > 0 Then
                        Call logger(ptErrLevel.logInfo, "Project List Import, Store Portfolio-Variant: ", outputCollection)
                    End If

                End If


                ' now Store Constellation 
                If missingList.Count > 0 Then
                    Dim portfolioName As String = jobParameters.portfolioName
                    Dim pfVariantName As String = jobParameters.portfolioVariantName & " - " & jobParameters.projectVariantName
                    Dim ok As Boolean = storeConstellationFromProjectList(missingList, portfolioName, pfVariantName & " missing")
                End If
            End If




        Catch ex As Exception
            result = False
        End Try

        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        processProjectListWithActivePortfolio = result

    End Function


    ''' <summary>
    ''' all projects need to be read alrady in ImportProjekte
    ''' </summary>
    ''' <returns></returns>
    Private Function defineFeasiblePortfolio() As Boolean

        Dim result As Boolean = True
        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight

        Dim overloadAllowedinMonths As Double = 1.05
        Dim overloadAllowedTotal As Double = 1.0


        Try
            ' now get the aggregation Roles
            Dim aggregationRoles As SortedList(Of Integer, String) = RoleDefinitions.getAggregationRoles()
            Dim aggregationList As New List(Of String)

            Dim teamID As Integer = -1

            Dim exceptionList As Collection = getConsiderationList(excludedNames:=True)

            ' checkout aggregation Roles
            For Each ar As KeyValuePair(Of Integer, String) In aggregationRoles
                Dim tmpStrID As String = RoleDefinitions.bestimmeRoleNameID(ar.Key, teamID)
                If Not aggregationList.Contains(tmpStrID) And Not exceptionList.Contains(tmpStrID) Then
                    aggregationList.Add(tmpStrID)
                End If
            Next

            ' Get the Ranking out of Excel List , it is just the ordering of the rows 
            ' value holds the AllProjekte.Key, i.e name#variantName
            Dim rankingList As SortedList(Of Integer, String) = getRanking()

            ' takes all the projects which could not be considered first time ... 
            Dim rankingList2 As New SortedList(Of Integer, String)
            Dim missingList As New clsProjekteAlle

            Dim abbruchDate As Date = New Date(2023, 6, 30)

            Dim tmpValues As Double() = getOverloadParams()
            Dim tmpNames As String() = getPortfolioNames()


            Dim portfolioName As String = tmpNames(0)
            Dim variantName As String = tmpNames(1)


            Try
                abbruchDate = CDate(tmpNames(2))
            Catch ex As Exception
                abbruchDate = New Date(2023, 6, 30)
            End Try


            overloadAllowedinMonths = tmpValues(0)
            overloadAllowedTotal = tmpValues(1)


            AlleProjekte.Clear()
            ' now make sure all projects are in AlleProjekte
            For Each ppair As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste

                If Not AlleProjekte.Containskey(ppair.Key) Then
                    AlleProjekte.Add(ppair.Value)
                End If

            Next

            ' then empty ShowProjekte again 
            ShowProjekte.Clear()

            ' rankingList keeps the sequence within the Excel file. So user adds some fields important to him for prioritization , he add these fields , sorts it in th eExcel. 
            ' It then represents the sequence: Row1 is the most important project , Row2 the scond, and so forth
            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList

                ' check whether or not there is already such a project-variant in the DB

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Remove(hproj.name)
                End If

                ShowProjekte.Add(hproj)

                Dim skillIDs As Collection = hproj.getSkillNameIds
                Dim skillList As New List(Of String)

                For Each si As String In skillIDs
                    If Not skillList.Contains(si) And Not exceptionList.Contains(si) Then
                        skillList.Add(si)
                    End If
                Next


                ' now consider the tiemFrame this project occupies ..
                showRangeLeft = getColumnOfDate(hproj.startDate)
                showRangeRight = getColumnOfDate(hproj.endeDate)


                Dim overutilizationFound As Boolean = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedinMonths, overloadAllowedTotal)

                ' before checking for utilization , do the auto-distribution, that is according to freecapacity try to distribute it in the timeframe that all fits
                ' example: with a normalized distribution it may come to a bootleneck: 15 PT in JAn, 15 PT in Feb. But there is 12 PT in Jan, 18 PT in Feb free capacity , so this works out. 
                Dim myMessages As New Collection
                Dim storeRequired As Boolean = False

                ' tk 24.4. Auto-Distribute : take it out , because it changes too much in the monthly needs 
                'If overutilizationFound Then
                '    Dim infomsg As String = "try out variant with optimized distribution according to resource needs:  " & hproj.getShapeText
                '    Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", infomsg)

                '    hproj = hproj.createVariant(variantName, "consider requested sum - distribute according free capacity ")
                '    Dim errorMsg As String = ""

                '    ' put it into AlleProjekte
                '    AlleProjekte.Add(hproj)
                '    ShowProjekte.AddAnyway(hproj)

                '    Call ShowProjekte.autoDistribute(hproj.name, "", errorMsg)


                '    ' now calculate again ... 
                '    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedinMonths, overloadAllowedTotal)

                '    ' to trigger saving of the new variant , if there was first a bottleneck, then no more ...
                '    storeRequired = Not overutilizationFound

                'End If


                If overutilizationFound Then

                    ' take it out again, because there was no solution
                    ShowProjekte.Remove(hproj.name)
                    Dim infomsg As String = "with default start-Date & end-Date not considered because of bottlenecks, will be tried out later ... " & hproj.name
                    Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", infomsg)

                    'Console.WriteLine(infomsg)

                    rankingList2.Add(rankingPair.Key, rankingPair.Value)

                Else
                    ' all ok, just continue
                    Dim infomsg As String = " ... considered " & hproj.getShapeText
                    Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", infomsg)

                    'Console.WriteLine(infomsg)

                    ' now if there was created the variant
                    If storeRequired Then

                        Dim tmpMessages As New Collection
                        If storeSingleProjectToDB(hproj, tmpMessages) Then
                            Dim mymsg As String = "tried out new value distribution:  worked out to find solution for  " & hproj.getShapeText
                            Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", mymsg)

                            'Console.WriteLine(mymsg)

                        End If
                    Else
                        Dim mymsg As String = "could be considered unchanged: " & hproj.getShapeText
                        Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", mymsg)

                        'Console.WriteLine(mymsg)
                    End If
                End If



            Next

            ' now the second wave is going to come ... 


            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList2

                Dim anzLoops As Integer = 1

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If Not ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Add(hproj)
                End If

                Dim skillIDs As Collection = hproj.getSkillNameIds
                Dim skillList As New List(Of String)

                For Each si As String In skillIDs
                    If Not skillList.Contains(si) And Not exceptionList.Contains(si) Then
                        skillList.Add(si)
                    End If
                Next

                ' now consider the timeFrame this project occupies ..

                Dim key As String = calcProjektKey(hproj)
                ' create variant if not already done
                If hproj.variantName <> variantName Then
                    hproj = hproj.createVariant(variantName, "variant was created and moved to avoid resource bottlenecks")
                    ' bring that into AlleProjekte
                    key = calcProjektKey(hproj)
                    If AlleProjekte.Containskey(key) Then
                        AlleProjekte.Remove(key)
                    End If
                    AlleProjekte.Add(hproj)
                End If

                Dim deltaInDays As Integer = 3

                Dim iterations As Integer = 0

                Dim newStartDate As Date = hproj.startDate.AddDays(deltaInDays)
                Dim newEndDate As Date = hproj.endeDate.AddDays(deltaInDays)

                Dim overutilizationFound As Boolean = True

                Do While overutilizationFound And newEndDate <= abbruchDate
                    ' move project by deltaIndays

                    Dim tmpProj As clsProjekt = moveProject(hproj, newStartDate, newEndDate)

                    If Not IsNothing(tmpProj) Then

                        hproj = tmpProj

                        showRangeLeft = getColumnOfDate(tmpProj.startDate)
                        showRangeRight = getColumnOfDate(tmpProj.endeDate)


                        '' now replace in ShowProjekte 
                        'AlleProjekte.Remove(key)
                        'ShowProjekte.Remove(tmpProj.name)
                        ' add the new, altered version 
                        AlleProjekte.Add(hproj)
                        ShowProjekte.AddAnyway(hproj)

                        overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedinMonths, overloadAllowedTotal)

                        ' tk 24.4.22 takte it out again , autoDistribute changes too much
                        If overutilizationFound Then
                            Dim fmsg As String = ""
                            Call ShowProjekte.autoDistribute(hproj.name, "", fmsg)
                            overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedinMonths, overloadAllowedTotal)
                        End If

                        If overutilizationFound Then
                            newStartDate = newStartDate.AddDays(deltaInDays)
                            newEndDate = newEndDate.AddDays(deltaInDays)

                            anzLoops = anzLoops + 1
                        End If

                    Else
                        ' Error occurred 
                        Exit Do
                    End If

                Loop

                If Not overutilizationFound Then
                    ' it is already in there ... but now needed to be stored
                    Dim myMessages As New Collection
                    If storeSingleProjectToDB(hproj, myMessages) Then
                        Dim infomsg As String = "tried out " & anzLoops & " different start/ends to avoid bottlenecks, found solution for  " & hproj.getShapeText
                        Call logger(ptErrLevel.logInfo, "define feasible portfolio: ", infomsg)

                        'Console.WriteLine(infomsg)

                    Else
                        ' take it out again , because there was no solution
                        ShowProjekte.Remove(hproj.name)

                        Dim mlKEy As String = calcProjektKey(hproj.name, "")
                        hproj = ImportProjekte.getProject(mlKEy)
                        missingList.Add(hproj)

                        Dim infomsg As String = "... failure: could not store " & hproj.getShapeText
                        Call logger(ptErrLevel.logError, "define feasible portfolio: ", infomsg)

                        'Console.WriteLine(infomsg)
                    End If


                Else
                    ' take it out again , because there was no solution
                    AlleProjekte.Remove(key)
                    ShowProjekte.Remove(hproj.name)

                    ' take it out again , because there was no solution
                    Dim infomsg As String = "... could finally not be considered  " & hproj.name

                    Call logger(ptErrLevel.logWarning, "define feasible portfolio: ", infomsg)

                    'Console.WriteLine(infomsg)

                    Dim mlKEy As String = calcProjektKey(hproj.name, "")
                    hproj = ImportProjekte.getProject(mlKEy)
                    missingList.Add(hproj)

                End If

            Next

            ' ----------------


            ' now create the portfolio Variant from ShowProjekte 
            ' now create the Portfolio from ShowProjekte content 

            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:=variantName)
            Dim errMsg As New clsErrorCodeMsg
            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

            Dim outputCollection As New Collection
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)

            If outputCollection.Count > 0 Then
                Call logger(ptErrLevel.logInfo, "feasible Portfolio definition, ", outputCollection)
            End If

            ' now Store Constellation 
            If missingList.Count > 0 Then
                Dim ok As Boolean = storeConstellationFromProjectList(missingList, portfolioName, variantName & " missing")
            End If

        Catch ex As Exception
            Call logger(ptErrLevel.logError, "feasible Portfolio definition, unexpected failure:", ex.Message)
            result = False
        End Try

        defineFeasiblePortfolio = result

    End Function


    ''' <summary>
    ''' performs creation and optimization when no activePortfolio is defined or does exist
    ''' </summary>
    ''' <param name="portfolioName"></param>
    ''' <param name="overloadAllowedInMonths"></param>
    ''' <param name="overloadAllowedTotal"></param>
    ''' <returns></returns>
    Private Function processProjectListWithoutActivePortfolio(ByVal aggregationList As List(Of String),
                                                              ByVal skillList As List(Of String),
                                                              ByVal portfolioName As String,
                                                              ByVal overloadAllowedInMonths As Double,
                                                              ByVal overloadAllowedTotal As Double) As Boolean
        Dim result As Boolean = True
        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight

        ' cache löschen
        result = CType(databaseAcc, DBAccLayer.Request).clearCache()

        Try
            ' Get the Ranking out of Excel List , it is just the ordering of the rows 
            ' value holds the AllProjekte.Key, i.e name#variantName
            Dim rankingList As SortedList(Of Integer, String) = getRanking()


            AlleProjekte.Clear()
            ' now make sure all projects are in AlleProjekte
            For Each ppair As KeyValuePair(Of String, clsProjekt) In ImportProjekte.liste
                If Not AlleProjekte.Containskey(ppair.Key) Then
                    AlleProjekte.Add(ppair.Value)
                End If
            Next


            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If Not ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Add(hproj)
                End If
            Next


            ' currentSessionConstellation is build by alle the Showprojekte.add and AlleProjekte.add Commands ...
            ' create form that a portfolio, only containing the show-Elements 
            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:="")

            ' now store the Portfolio , with name portfolioName
            Dim errMsg As New clsErrorCodeMsg
            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

            Dim outputCollection As New Collection
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)

            ' define the range, necessary to check whether or not there are bottlenecks 
            showRangeLeft = ShowProjekte.getMinMonthColumn
            showRangeRight = ShowProjekte.getMaxMonthColumn


            ' then empty ShowProjekte again 
            ShowProjekte.Clear()


            ' 1. now start with the (next-)highest ranked project, 
            ' 2. If there are no bottlenecks, keep it in ShowProjekte; 
            '    if there are bottlenecks create a variant with name [arb], then move variant by 7 days until there is no bottleneck any more or until project has been moved by approx 6 months
            '    if bottleneck cannot be solved, take project out of potential portfolio 
            ' 3. Go to 1.


            ' rankingList keeps the sequence within the Excel file. So user adds some fields important to him for prioritization , he add these fields , sorts it in th eExcel. 
            ' It then represents the sequence: Row1 is the most important project 
            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If Not ShowProjekte.contains(hproj.name) Then
                    ShowProjekte.Add(hproj)
                End If

                Dim overutilizationFound As Boolean = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)


                If overutilizationFound Then
                    Dim key As String = calcProjektKey(hproj)
                    ' create variant if not already done
                    If hproj.variantName <> "arb" Then
                        hproj = hproj.createVariant("arb", "variant was created and moved to avoid resource bottlenecks")
                        ' bring that into AlleProjekte
                        key = calcProjektKey(hproj)
                        If AlleProjekte.Containskey(key) Then
                            AlleProjekte.Remove(key)
                        End If
                        AlleProjekte.Add(hproj)
                    End If

                    Dim deltaInDays As Integer = 7
                    Dim maxIterations As Integer = CInt(182 / deltaInDays)
                    Dim iterations As Integer = 0

                    Do While overutilizationFound And iterations <= maxIterations
                        ' move project by deltaIndays

                        Dim newStartDate As Date = hproj.startDate.AddDays(deltaInDays)
                        Dim newEndDate As Date = hproj.endeDate.AddDays(deltaInDays)

                        Dim tmpProj As clsProjekt = moveProject(hproj, newStartDate, newEndDate)

                        If Not IsNothing(tmpProj) Then

                            hproj = tmpProj

                            ' now replace in ShowProjekte 
                            AlleProjekte.Remove(key)
                            ShowProjekte.Remove(tmpProj.name)
                            ' add the new, altered version 
                            AlleProjekte.Add(tmpProj)
                            ShowProjekte.Add(tmpProj)

                            overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)

                            If overutilizationFound Then
                                iterations = iterations + 1
                            End If

                        Else
                            ' Error occurred 
                            Exit Do
                        End If

                    Loop

                    If Not overutilizationFound Then
                        ' it is already in there ... but now needed to be stored
                        Dim myMessages As New Collection
                        If storeSingleProjectToDB(hproj, myMessages) Then
                            Dim infomsg As String = "created variant to avoid bottlenecks " & hproj.getShapeText
                            Call logger(ptErrLevel.logInfo, infomsg, myMessages)
                            'Console.WriteLine(infomsg)
                        Else
                            ' take it out again , because there was no solution
                            ShowProjekte.Remove(hproj.name)
                            Dim infomsg As String = "... failed to create variant to avoid bottlenecks " & hproj.getShapeText
                            Call logger(ptErrLevel.logError, infomsg, myMessages)
                            'Console.WriteLine(infomsg)
                        End If


                    Else
                        ' take it out again , because there was no solution
                        AlleProjekte.Remove(key)
                        ShowProjekte.Remove(hproj.name)
                    End If

                Else
                    ' all ok, just continue
                End If

            Next

            ' now create the portfolio Variant arb from ShowProjekte 
            ' now create the Portfolio from ShowProjekte content 

            toStoreConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:="arb")


            outputCollection.Clear()
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)

            If outputCollection.Count > 0 Then
                Call logger(ptErrLevel.logError, "Project List Import, Store Portfolio-Variant arb failed:", outputCollection)
            End If

        Catch ex As Exception
            result = False
        End Try

        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        processProjectListWithoutActivePortfolio = result


    End Function

    Private Function processMppFile(ByVal fileName As String, ByVal importDate As Date) As Boolean

        Dim myName As String = My.Computer.FileSystem.GetName(fileName)
        Dim allOk As Boolean = False
        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboMPP.ToString, myName)

        ' cache löschen
        Dim result As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()

        Try

            Dim hproj As clsProjekt = New clsProjekt

            ' Definition für ein eventuelles Mapping
            Dim mapProj As clsProjekt = Nothing
            Call awinImportMSProject("RPA", fileName, hproj, mapProj, importDate)


            ' now protocol whether or not there are unknown cost and roles used in the MS projct file 
            If Not IsNothing(hproj) Then

                allOk = True

                If missingRoleDefinitions.Count > 0 Or missingCostDefinitions.Count > 0 Then

                    Dim outputCollection As New Collection
                    Dim outputLine As String = ""
                    For Each mrd As KeyValuePair(Of Integer, clsRollenDefinition) In missingRoleDefinitions.liste
                        If awinSettings.englishLanguage Then
                            outputLine = "unknown Role: " & mrd.Value.name
                        Else
                            outputLine = "unbekannte Rolle: " & mrd.Value.name
                        End If
                        outputCollection.Add(outputLine)
                    Next

                    For Each mcd As KeyValuePair(Of Integer, clsKostenartDefinition) In missingCostDefinitions.liste
                        If awinSettings.englishLanguage Then
                            outputLine = "unknown Cost: " & mcd.Value.name
                        Else
                            outputLine = "unbekannte Kostenart: " & mcd.Value.name
                        End If

                        outputCollection.Add(outputLine)
                    Next

                    If awinSettings.englishLanguage Then
                        outputLine = "unknown Elements: please modify organisation-file or input ..."
                    Else
                        outputLine = "Unbekannte Elemente: bitte in Organisations-Datei korrigieren"
                    End If

                    outputCollection.Add(outputLine)

                    Call logger(ptErrLevel.logWarning, "RPA Import MS Project " & myName, outputCollection)

                Else
                    ' everything ok
                End If


                Try
                    Dim keyStr As String = calcProjektKey(hproj)
                    ImportProjekte.Add(hproj, updateCurrentConstellation:=False)

                    If Not IsNothing(mapProj) Then
                        keyStr = calcProjektKey(mapProj)
                        ImportProjekte.Add(mapProj, updateCurrentConstellation:=False)

                    End If
                Catch ex2 As Exception
                    allOk = False
                    Call logger(ptErrLevel.logError, "RPA Error Importing MS project file: file already exists ", myName)
                End Try

                If allOk Then
                    Try
                        Call importProjekteEintragen(importDate, drawPlanTafel:=False, fileFrom3rdParty:=True, getSomeValuesFromOldProj:=True, calledFromActualDataImport:=False, calledFromRPA:=True)
                    Catch ex2 As Exception
                        allOk = False
                        Call logger(ptErrLevel.logError, "RPA Error Importing project brief " & PTRpa.visboMPP.ToString, myName)
                    End Try
                End If

                If allOk Then
                    ' store Project 
                    allOk = storeImportProjekte()
                    ' empty session 
                    Call emptyRPASession()
                    Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboMPP.ToString, myName)
                End If

            Else
                allOk = False
                Call logger(ptErrLevel.logError, "end Processing: " & PTRpa.visboMPP.ToString, myName)
            End If


        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "end Processing: " & PTRpa.visboMPP.ToString, myName)
        End Try

        processMppFile = allOk

    End Function

    Private Function processInitialOrga(ByVal myName As String) As Boolean

        Dim allOK As Boolean = False
        Dim msgTxt As String = ""
        Dim orgaImportConfig As New SortedList(Of String, clsConfigOrgaImport)
        Dim outputCollection As New Collection

        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboInitialOrga.ToString, myName)

        Try

            ' Dim importedOrga As clsOrganisation = ImportOrganisation(outputCollection)
            Dim importedOrga As clsOrganisation = ImportOrganisation(outputCollection, orgaImportConfig)


            If outputCollection.Count > 0 Then
                Dim errmsg As String = vbLf & " .. Exit .. Organisation not imported  "
                outputCollection.Add(errmsg)

                Call logger(ptErrLevel.logError, "RPA Error Importing Organisation ", outputCollection)

            ElseIf importedOrga.count > 0 Then

                ' TopNodes und OrgaTeamChilds bauen 
                Call importedOrga.allRoles.buildTopNodes()

                Dim err As New clsErrorCodeMsg
                Dim result As Boolean = False
                ' ute -> überprüfen bzw. fertigstellen ... 
                Dim orgaName As String = ptSettingTypes.organisation.ToString

                ' andere Rollen als Orga-Admin können Orga einlesen, aber eben nicht speichern ! 
                'result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(importedOrga,
                '                                        CStr(settingTypes(ptSettingTypes.organisation)),
                '                                        orgaName,
                '                                        importedOrga.validFrom,
                '                                        err)
                result = CType(databaseAcc, DBAccLayer.Request).storeTSOOrganisationToDB(importedOrga,
                                                                                  orgaName,
                                                                                  importedOrga.validFrom,
                                                                                  err)

                If result = True Then
                    allOK = True
                    msgTxt = "ok, Organisation, valid from " & importedOrga.validFrom.ToShortDateString & " stored ..."
                    'Console.WriteLine(msgTxt)
                    Call logger(ptErrLevel.logInfo, PTRpa.visboInitialOrga.ToString, msgTxt)
                Else
                    allOK = False
                    msgTxt = "Storing organisaiton failed "
                    Call logger(ptErrLevel.logError, PTRpa.visboInitialOrga.ToString, msgTxt)
                End If
            End If

            Call logger(ptErrLevel.logInfo, "endProcessing: " & PTRpa.visboInitialOrga.ToString, myName)
        Catch ex As Exception
            allOK = False
        End Try

        processInitialOrga = allOK

    End Function

    Private Function processRoundTripOrga(ByVal myName As String) As Boolean

        Dim allOK As Boolean = False
        Dim msgTxt As String = ""
        'Dim orgaImportConfig As New SortedList(Of String, clsConfigOrgaImport)
        Dim outputCollection As New Collection
        'Try

        '    Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboRoundtripOrga.ToString, myName)

        '    ' ===========================================================
        '    ' Konfigurationsdatei lesen und Validierung durchführen

        '    ' Ur: 21.02.2022: geändert auf configuration Orga im VC als Setting

        '    ' wenn es gibt - lesen der ControllingSheet und anderer, die durch configActualDataImport beschrieben sind
        '    'Dim configOrgaImport As String = My.Computer.FileSystem.CombinePath(configfilesOrdner, "configOrgaImport.xlsx")

        '    Dim orgaImportConfig As New SortedList(Of String, clsConfigOrgaImport)
        '    Dim lastrow As Integer = 0

        '    Call logger(ptErrLevel.logInfo, "start reading configuration Orga: " & PTRpa.visboRoundtripOrga.ToString, "VCSetting configuration Orga")

        '    ' check Config-File - zum Einlesen der Oragnistation gemäß Konfiguration
        '    ' hier werden Werte für die Konfiguration gelesen aus dem VCSetting "configuration Orga"
        '    Dim allesOK As Boolean = checkOrgaImportConfig("configuration Orga", myName, orgaImportConfig, lastrow, outputCollection)

        '    If Not allesOK Then
        '        Call logger(ptErrLevel.logError, "error reading configuration Orga: " & PTRpa.visboRoundtripOrga.ToString, "VCSetting configuration Orga does not exist")
        '        processRoundTripOrga = False
        '        Exit Function
        '    End If

        '    Try

        '        ' Dim importedOrga As clsOrganisation = ImportOrganisation(outputCollection)
        '        Dim importedOrga As clsOrganisation = ImportOrganisation(outputCollection, orgaImportConfig)


        '        If outputCollection.Count > 0 Then
        '            Dim errmsg As String = vbLf & " .. Exit .. Organisation not imported  "
        '            outputCollection.Add(errmsg)

        '            Call logger(ptErrLevel.logError, "RPA Error Importing Organisation ", outputCollection)

        '        ElseIf importedOrga.count > 0 Then

        '            ' TopNodes und OrgaTeamChilds bauen 
        '            Call importedOrga.allRoles.buildTopNodes()

        '            Dim err As New clsErrorCodeMsg
        '            Dim result As Boolean = False
        '            ' ute -> überprüfen bzw. fertigstellen ... 
        '            Dim orgaName As String = ptSettingTypes.organisation.ToString

        '            ' andere Rollen als Orga-Admin können Orga einlesen, aber eben nicht speichern ! 
        '            result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(importedOrga,
        '                                                CStr(settingTypes(ptSettingTypes.organisation)),
        '                                                orgaName,
        '                                                importedOrga.validFrom,
        '                                                err)

        '            If result = True Then
        '                allOK = True
        '                msgTxt = "ok, Organisation, valid from " & importedOrga.validFrom.ToShortDateString & " stored ..."
        '                Console.WriteLine(msgTxt)
        '                Call logger(ptErrLevel.logInfo, PTRpa.visboRoundtripOrga.ToString, msgTxt)
        '            Else
        '                allOK = False
        '                msgTxt = "Storing organisaiton failed "
        '                Call logger(ptErrLevel.logError, PTRpa.visboRoundtripOrga.ToString, msgTxt)
        '            End If
        '        End If

        '        Call logger(ptErrLevel.logInfo, "endProcessing: " & PTRpa.visboRoundtripOrga.ToString, myName)
        '    Catch ex As Exception
        '        allOK = False
        '    End Try

        'Catch ex As Exception
        '    allOK = False
        '    msgTxt = ""
        '    Call logger(ptErrLevel.logError, PTRpa.visboRoundtripOrga.ToString, ex.Message)
        'End Try

        msgTxt = "This Import will no longer be supported! " & " NOW: Download the Orga, change it and upload it via WebUI"
        Call logger(ptErrLevel.logError, "VisboRoundTripOrga", msgTxt)
        allOK = False

        processRoundTripOrga = allOK

    End Function


    Private Function processVisboJira(ByVal myName As String, ByVal importDate As Date) As Boolean

        Dim allOk As Boolean = True
        Dim aktDateTime As Date = Date.Now

        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboJira.ToString, myName)

        'check the pre-conditions
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If

        ' cache löschen
        Dim result As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()


        'read File with Jira-Projects and put it into ImportProjekte
        Try
            '' read the file and import into hproj
            'Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)
            Dim JIRAProjectsConfig As New SortedList(Of String, clsConfigProjectsImport)
            Dim projectsFile As String = ""
            Dim lastrow As Integer = 0
            Dim outputString As String = ""
            Dim dateiName As String = ""
            Dim listofArchivAllg As New List(Of String)
            Dim outPutCollection As New Collection
            Dim configJIRAProjects As String = ""


            Dim outputLine As String = ""

            Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

            ' Konfigurationsdatei lesen und Validierung durchführen

            ' wenn es gibt - lesen der Jira und anderer, die durch configCapaImport beschrieben sind
            ' no longer necessary
            ' Dim configJIRAProjects As String = My.Computer.FileSystem.CombinePath(configfilesOrdner, "configJIRAProjectImport.xlsx")

            ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
            Dim allesOK As Boolean = checkJIRAProjectImportConfig(configJIRAProjects, projectsFile, JIRAProjectsConfig, lastrow, outPutCollection)

            If allesOK Then

                Dim listofVorlagen As New Collection
                listofVorlagen.Add(myName)
                listofArchivAllg = readProjectsJIRA(listofVorlagen, JIRAProjectsConfig, outPutCollection)

                Try
                    ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                    ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 

                    Call importProjekteEintragen(importDate, drawPlanTafel:=False, fileFrom3rdParty:=True, getSomeValuesFromOldProj:=True, calledFromActualDataImport:=False, calledFromRPA:=True)

                Catch ex As Exception
                    If awinSettings.englishLanguage Then
                        Call MsgBox("Error at Import: " & vbLf & ex.Message)
                    Else
                        Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                    End If

                End Try
            Else
                Call logger(ptErrLevel.logError, "processVisboJira", outPutCollection)
                allOk = False
            End If

            ' store Projects
            If allOk Then
                allOk = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboJira.ToString, myName)

        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Jira Project file ", ex1.Message)
        End Try

        processVisboJira = allOk

    End Function


    Private Function processVisboUrlaubsplaner(ByVal myName As String, ByVal importDate As Date, ByRef errMessages As Collection) As Boolean

        Dim outPutline As String = ""
        Dim lastrow As Integer = 0
        Dim listofArchivUrlaub As New List(Of String)
        Dim listofArchivConfig As New List(Of String)
        Dim result As Boolean = False
        Dim outputCollection As New Collection
        Dim aktDateTime As Date = Date.Now

        'check the pre-conditions
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If


        Dim changedOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)

        ' Timer
        Dim sw As clsStopWatch
        sw = New clsStopWatch
        sw.StartTimer()

        If Not IsNothing(changedOrga) Then

            If changedOrga.allRoles.Count > 0 Then

                RoleDefinitions = changedOrga.allRoles
                CostDefinitions = changedOrga.allCosts


                Call logger(ptErrLevel.logInfo, "Einlesen Verfügbarkeiten " & myName, "processVisboUrlaubsplaner", anzFehler)
                result = readAvailabilityOfRole(myName, outputCollection)
                If result Then
                    ' hier: merken der erfolgreich importierten KapaFiles
                    listofArchivUrlaub.Add(myName)
                    Call logger(ptErrLevel.logInfo, "Einlesen Verfügbarkeiten " & myName, outputCollection)
                Else
                    Call logger(ptErrLevel.logError, "Einlesen Verfügbarkeiten " & myName, outputCollection)
                End If

                '' wenn es gibt - lesen der Urlaubslisten DateiName "Urlaubsplaner*.xlsx
                'listofArchivUrlaub = readInterneAnwesenheitslisten(outputCollection)

                If listofArchivUrlaub.Count > 0 Then

                    changedOrga.allRoles = RoleDefinitions

                    If outputCollection.Count = 0 Then
                        ' keine Fehler aufgetreten ... 
                        ' jetzt wird die Orga als Setting weggespeichert ... 
                        Dim err As New clsErrorCodeMsg

                        ' ute -> überprüfen bzw. fertigstellen ... 
                        ' Dim orgaName As String = ptSettingTypes.organisation.ToString

                        If (myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Or myCustomUserRole.customUserRole = ptCustomUserRoles.Alles) Or visboClient = "VISBO RPA / " Then


                            ' tk 13.4.22 wozu brauchen wir das hier ? 
                            'Dim orga As clsOrganisation = CType(databaseAcc, DBAccLayer.Request).retrieveTSOrgaFromDB("organisation", Date.Now, err, False, True, False)


                            result = storeCapasOfRoles()

                            If result = True Then
                                Call logger(ptErrLevel.logInfo, "ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " updated ...", "processUrlaubsplaner: ", -1)
                            Else
                                msgTxt = "Error when writing Capacities to Database..." & vbCrLf & err.errorMsg
                                Call logger(ptErrLevel.logError, msgTxt, "processUrlaubsplaner: ", -1)
                                outputCollection.Add(msgTxt)
                            End If
                        Else
                            msgTxt = "Error when writing Capacities to Database...- wrong customUserRole" & vbCrLf & myCustomUserRole.customUserRole
                            Call logger(ptErrLevel.logError, msgTxt, "processUrlaubsplaner: ", -1)
                            outputCollection.Add(msgTxt)
                            result = False
                        End If
                    Else
                        Call logger(ptErrLevel.logError, "processUrlaubsplaner: ", outputCollection)
                    End If
                Else
                    result = False
                    If outputCollection.Count > 0 Then
                        Call logger(ptErrLevel.logError, "processUrlaubsplaner: ", outputCollection)
                    Else
                        If awinSettings.englishLanguage Then
                            msgTxt = "there do not exists any 'Urlaubsplaner*'!"
                        Else
                            msgTxt = "Es existiert kein 'Urlaubsplaner*.*' !"
                        End If
                        Call logger(ptErrLevel.logError, "processUrlaubsplaner: ", msgTxt)
                        outputCollection.Add(msgTxt)
                    End If

                End If

            Else
                If awinSettings.englishLanguage Then
                    msgTxt = "No valid roles! Please import one first!"
                Else
                    msgTxt = "Die gültige Organisation beinhaltet keine Rollen! "
                End If
                Call logger(ptErrLevel.logError, "processUrlaubsplaner: ", msgTxt)
                outputCollection.Add(msgTxt)
            End If

        Else

            If awinSettings.englishLanguage Then
                msgTxt = "No valid organization! Please import one first!"
            Else
                msgTxt = "Es existiert keine gültige Organisation! Bitte zuerst Organisation importieren"
            End If
            outputCollection.Add(msgTxt)


            Dim errMsg As String
            If awinSettings.englishLanguage Then
                errMsg = "Capacities not updated - please first remove the errors in the importfiles ... "
                outputCollection.Add(errMsg)
            Else
                errMsg = "Kapazitäten wurden nicht aktualisiert - bitte erst die Import-Dateien korrigieren ... "
                outputCollection.Add(errMsg)
            End If
            Call logger(ptErrLevel.logError, "processUrlaubsplaner: ", outputCollection)

        End If


        Dim ti As Long = sw.EndTimer()
        errMessages = outputCollection

        processVisboUrlaubsplaner = result

    End Function

    ''' <summary>
    ''' standard import of actual data like Instart
    ''' </summary>
    ''' <param name="myName"></param>
    ''' <param name="importDate"></param>
    ''' <returns></returns>
    Private Function processVisboActualData1(ByVal myName As String, ByVal importDate As Date) As Boolean

        Dim allOk As Boolean = True
        Dim aktDateTime As Date = Date.Now

        'check the pre-conditions
        lastReadingOrganisation = readOrganisations()
        lastReadingCustomization = readCustomizations()
        If lastReadingCustomization <= Date.MinValue Then
            Call logger(ptErrLevel.logError, "processVisboActualData1", "the import of actual data requires the existence of a customization setting")
        End If
        'End If

        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboActualData1.ToString, myName)

        ' cache löschen
        Dim result0 As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()

        Dim weitermachen As Boolean = False
        Dim outPutCollection As New Collection
        Dim outPutline As String = ""
        Dim result As Boolean = False
        Dim listOfArchivFiles As New List(Of String)
        Dim dateiname As String = myName

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
        Dim msgTxt As String = "Actual Data Departments:  "
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

        Call logger(ptErrLevel.logInfo, msgTxt, "processVisboActualData1", anzFehler)

        weitermachen = True

        result = readActualData(dateiname)
        If result Then
            listOfArchivFiles.Add(dateiname)
        End If

        allOk = allOk And result

        Try
            ' store Projects
            If allOk Then
                allOk = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboActualData1.ToString, myName)

        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Actual Data (modus 1)", ex1.Message)
        End Try

        processVisboActualData1 = allOk

    End Function


    Public Function processVisboActualData2(ByVal myName As String, ByVal portfolioName As String, ByVal dirName As String, ByVal importDate As Date) As Boolean

        Dim allOk As Boolean = True
        Dim aktDateTime As Date = Date.Now

        logfileNamePath = createLogfileName(rpaFolder, myName)
        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboActualData2.ToString, myName)

        ' cache löschen
        Dim result0 As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()

        'check the pre-conditions
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If
        'If DateDiff(DateInterval.Hour, lastReadingCustomization, aktDateTime) > 24 Then
        lastReadingCustomization = readCustomizations()
            If lastReadingCustomization <= Date.MinValue Then
                Call logger(ptErrLevel.logError, "processVisboActualData2", "the import of actual data requires the existence of a customization setting")
            End If
        'End If

        Dim weitermachen As Boolean = False
        Dim outPutCollection As New Collection
        Dim outPutline As String = ""
        Dim result As Boolean = False
        Dim listOfArchivFilesAllg As New List(Of String)
        Dim dateiname As String = myName

        Dim selectedWB As String = ""
        Dim actualDataFile As String = ""
        Dim actualDataConfig As New SortedList(Of String, clsConfigActualDataImport)
        Dim lastrow As Integer
        Dim listOfErrorImportFilesAllg As New List(Of String)
        Dim anzFiles As Integer = 0

        Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)
        ' erstmal protokollieren, zu welchen Abteilungen Istdaten gelesen und substituiert werden 
        ' alle Planungen zu den Rollen, die in dieser Referatsliste aufgeführt sind, werden gelöscht 


        ' IstDaten - relevante Orga-Units aufsammeln für Import Istdaten

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


        ' Konfigurations-Dateien lesen 
        ' ===========================================================
        ' Konfigurationsdatei lesen und Validierung durchführen
        Dim configActualDataImport As String = My.Computer.FileSystem.CombinePath(configfilesOrdner, "configActualDataImport.xlsx")

        ' check Config-File - zum Einlesen der Istdaten gemäß Konfiguration
        ' hier werden Werte für actualDataFile, actualDataConfig gesetzt
        Dim checkConfigOK As Boolean = checkActualDataImportConfig(configActualDataImport, actualDataFile, actualDataConfig, lastrow, outPutCollection)

        ' read files with actualData 
        ' ==========================

        If checkConfigOK Then

            Dim listOfImportfilesAllg As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirName, FileIO.SearchOption.SearchTopLevelOnly, actualDataFile)
            anzFiles = listOfImportfilesAllg.Count

            If listOfImportfilesAllg.Count >= 1 Then
                ' Vorbereitungen für die Aufnahme der verschiedenen Excel-File Daten in die unterschiedlichen Projekte
                Dim editActualDataMonth As New frmInfoActualDataMonth
                Dim lastValidMonth As Integer = 0  ' angegeben in dem Dialog
                Dim IstdatenDate As Date
                Dim curMonth As Integer = 0
                Dim hrole As New clsRollenDefinition
                Dim cacheProjekte As New clsProjekteAlle


                ' Istdaten immer vom Vormonat einlesen
                IstdatenDate = CDate(importDate).AddMonths(-1)

                Dim referenzPortfolioName As String = myActivePortfolio

                Dim curTimeStamp As Date = Date.MinValue
                Dim err As New clsErrorCodeMsg
                Dim referenzPortfolio As clsConstellation = Nothing

                If referenzPortfolioName = "" Then

                    Dim txtMsg As String = "kein Portfolio gewählt - Abbruch!"
                    If awinSettings.englishLanguage Then
                        txtMsg = "no Portfolio selected - Cancelled ..."
                    End If
                    Call logger(ptErrLevel.logError, "processVisboActualData2", txtMsg)
                    'Console.WriteLine(txtMsg)

                    processVisboActualData2 = False

                    Exit Function

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
                    Call logger(ptErrLevel.logError, "processVisboActualData2", txtMsg)
                    'Console.WriteLine(txtMsg)

                    processVisboActualData2 = False

                    Exit Function

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

                    Call logger(ptErrLevel.logInfo, outPutline, "processVisboActualData2", anzFehler)

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

                    allOk = allOk And result
                Next

                If listOfImportfilesAllg.Count = listOfArchivFilesAllg.Count Then           ' dann sind alle korrekt durchgelaufen

                    ' jetzt kommt die zweite Bearbeitungs-Welle

                    ' jetzt wird hier überprüft 
                    ' gibt es Projekte im Referenz-Portfolio, die keine Ist-Daten erhalten haben - dann sollte jetzt ggf. hier ein Nuller Eintrag im array für diese Projekte erfolgen 

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

                        ' es fehlen timeSheets von manchen Mitarbeitern
                        For Each roleName As String In missingTimeSheets
                            ReDim logArray(5)
                            ' ins Protokoll eintragen 
                            logArray(0) = " Mitarbeiter ohne TimeSheet: "
                            If awinSettings.englishLanguage Then
                                logArray(0) = "Employee without TimeSheet: "
                            End If
                            logArray(1) = ""
                            logArray(2) = roleName
                            logArray(4) = ""

                            Call logger(ptErrLevel.logWarning, "processVisboActualData2", logArray)

                        Next
                    End If

                    ' Ende check : haben alle internen Mitarbeiter ein TimeSheet abgeliefert ... 

                    ' wenn auch externe Rollen Istdaten haben
                    ' welche externen Rollen haben keine Istdaten .. ? 

                    ' Projekte, die keine Istdaten erhalten, aber im referenzPortfolio sind, erhalten Istdaten = 0
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

                                Call logger(ptErrLevel.logWarning, "processVisboActualData2", logArray)

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

                                Call logger(ptErrLevel.logError, "processVisboActualData2", logArray)

                                allOk = allOk And False
                            End If

                            '' im Output anzeigen ... 
                            'logmessage = logArray(0) & hproj.name
                            'outPutCollection.Add(logmessage)

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

                            Call logger(ptErrLevel.logWarning, "processVisboActualData2", logArray)

                            '' im Output anzeigen ... 
                            'logmessage = logArray(0) & vPKvP.Key
                            'outPutCollection.Add(logmessage)

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

                        '' im Output anzeigen ... 
                        'logmessage = logArray(0) & substituteUnit
                        'outPutCollection.Add(logmessage)

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

                            'ur:211206: den Status zu ändern ist hier unglücklich, da dies im VP erledigt werden müsste Soll dies erfolgen???
                            ' wenn es noch nicht beauftragt ist ... dann beauftragen 
                            If hproj.vpStatus = VProjectStatus(PTVPStati.initialized) Or
                                hproj.vpStatus = VProjectStatus(PTVPStati.proposed) Or
                                hproj.vpStatus = VProjectStatus(PTVPStati.stopped) Or
                                hproj.vpStatus = VProjectStatus(PTVPStati.finished) Then
                                Try
                                    If awinSettings.englishLanguage Then
                                        msgTxt = "Attention! Your project " & hproj.name & "/" & hproj.variantName & "is actually - " & hproj.vpStatus & "!!"
                                    Else
                                        msgTxt = "Achtung! Das Projekt " & hproj.name & "/" & hproj.variantName & "hat aktuell den Status - " & hproj.vpStatus & "!!"
                                    End If
                                    Call logger(ptErrLevel.logWarning, "processVisboActualData2", msgTxt)
                                    'ur: 211206: hproj.vpStatus = VProjectStatus(PTVPStati.ordered)
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

                    Call logger(ptErrLevel.logWarning, "Import ActualData2", logmessage)

                    ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                    Try
                        Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=False, fileFrom3rdParty:=False,
                                             getSomeValuesFromOldProj:=False, calledFromActualDataImport:=True, calledFromRPA:=True)


                        ' ImportDatei ins archive-Directory schieben

                        If listOfArchivFilesAllg.Count > 0 Then
                            'Call moveFilesInArchiv(listOfArchivFilesAllg, importOrdnerNames(PTImpExp.actualData))
                            Call moveFilesInArchiv(listOfArchivFilesAllg, dirName)
                        End If

                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            Call MsgBox("Error at Import: " & vbLf & ex.Message)
                        Else
                            Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                        End If

                    End Try

                    allOk = allOk And result

                    Try
                        ' store Projects
                        If allOk Then
                            allOk = storeImportProjekte()
                        End If

                        ' empty session 
                        Call emptyRPASession()

                        Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboActualData2.ToString, myName)

                    Catch ex1 As Exception
                        allOk = False
                        Call logger(ptErrLevel.logError, "RPA Error Importing Actual Data (modus 2)", ex1.Message)
                    End Try

                Else


                    For Each errImp As String In listOfErrorImportFilesAllg
                        Dim errImpName As String = My.Computer.FileSystem.GetName(errImp)
                        Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, errImpName)
                        My.Computer.FileSystem.MoveFile(errImp, newDestination, True)
                        Call logger(ptErrLevel.logError, "failed: ", errImp)
                        Dim logfileName As String = My.Computer.FileSystem.GetName(logfileNamePath)
                        Dim newLog As String = My.Computer.FileSystem.CombinePath(failureFolder, logfileName)
                        My.Computer.FileSystem.MoveFile(logfileNamePath, newLog, True)

                    Next

                    errMsgCode = New clsErrorCodeMsg
                    result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                                                                                    & myName & ": with errors ..." & vbCrLf _
                                                                                    & "Look for more details in the Failure-Folder", errMsgCode)
                    ' Fehler erfolgt
                    ' Dateien müssen in failure-Directory verschoben werden
                    'Call MsgBox("TODO")

                End If


            Else
                If awinSettings.englishLanguage Then
                    outPutline = "No file to import actual data"
                Else
                    outPutline = "Es gibt keine Datei zum Importieren von Istdaten"
                End If

                Call logger(ptErrLevel.logWarning, outPutline, "processVisboActualData2", anzFehler)

            End If

        Else
            ' Fehlermeldung für Konfigurationsfile nicht vorhanden
            If awinSettings.englishLanguage Then
                outPutline = "Error: either no configuration file found or worng definitions ! " & configActualDataImport
            Else
                outPutline = "Fehler: entweder fehlt die Konfigurations-Datei oder sie enthält fehlerhafte Definitionen ! " & configActualDataImport
            End If
            Call logger(ptErrLevel.logError, outPutline, "processVisboActualData2", anzFehler)

            allOk = allOk And False

        End If    ' checkConfigOK

        processVisboActualData2 = allOk

    End Function

    Public Function processInstartProposal(ByVal myName As String, ByVal portfolioName As String, ByVal dirName As String, ByVal importDate As Date) As Boolean
        Dim allOk As Boolean = True
        Dim aktDateTime As Date = Date.Now
        Dim instartImportConfigOK As Boolean = False

        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboProposal.ToString, myName)

        'check the pre-conditions
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If


        ' cache löschen
        Dim result As Boolean = CType(databaseAcc, DBAccLayer.Request).clearCache()


        'read File with Proposals Instart and put it into ImportProjekte
        Try
            '' read the file and import into hproj
            'Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)
            Dim projectConfig As New SortedList(Of String, clsConfigProjectsImport)
            Dim projectsFile As String = ""
            Dim lastrow As Integer = 0
            Dim outputString As String = ""
            Dim dateiName As String = ""
            Dim listofArchivAllg As New List(Of String)
            Dim outPutCollection As New Collection
            Dim configProposalImport As String = ""


            Dim outputLine As String = ""

            Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

            ' Konfigurationsdatei lesen und Validierung durchführen

            ' wenn es gibt - lesen der Jira und anderer, die durch configCapaImport beschrieben sind
            ' no longer necessary
            ' Dim configJIRAProjects As String = My.Computer.FileSystem.CombinePath(configfilesOrdner, "configJIRAProjectImport.xlsx")

            ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
            Dim allesOK As Boolean = checkProjectImportConfig(configProposalImport, projectsFile, projectConfig, lastrow, outPutCollection)

            If allesOK Then


                Dim listofVorlagen As New Collection
                listofVorlagen.Add(myName)
                If projectsFile = projectConfig("DateiName").ProjectsFile Then
                    listofArchivAllg = readProjectsAllg(listofVorlagen, projectConfig, outPutCollection, ptImportTypen.instartCalcTemplateImport)
                End If
                'listofArchivAllg = readProjectsJIRA(listofVorlagen, JIRAProjectsConfig, outPutCollection)

                If listofArchivAllg.Count > 0 Then
                    Call moveFilesInArchiv(listofArchivAllg, importOrdnerNames(PTImpExp.projectWithConfig))
                End If

                If allesOK Then
                    ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                    Try
                        ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                        ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                        Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False, getSomeValuesFromOldProj:=False, calledFromActualDataImport:=False, calledFromRPA:=True)

                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            Call MsgBox("Error at Import: " & vbLf & ex.Message)
                        Else
                            Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                        End If

                    End Try
                Else

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

            Try
                ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 

                Call importProjekteEintragen(importDate, drawPlanTafel:=False, fileFrom3rdParty:=True, getSomeValuesFromOldProj:=True, calledFromActualDataImport:=False, calledFromRPA:=True)

            Catch ex As Exception
                If awinSettings.englishLanguage Then
                    Call MsgBox("Error at Import: " & vbLf & ex.Message)
                Else
                    Call MsgBox("Fehler bei Import: " & vbLf & ex.Message)
                End If

            End Try

            'Else
            '    Call logger(ptErrLevel.logError, "processInstartProposal", outPutCollection)
            '    allOk = False
            'End If

            ' store Projects
            If allOk Then
                allOk = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboInstartProposal.ToString, myName)

        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Jira Project file ", ex1.Message)
        End Try

        processInstartProposal = allOk
    End Function
    ''' <summary>
    ''' Gibt das jeweilige Ergebnis weiter fürs logfile und schiebt die jeweilige Datei in die entsprechenden Folder
    ''' </summary>
    ''' <param name="fullfileName"></param>
    ''' <param name="allOK"></param>
    Public Sub processResult(ByVal fullfileName As String, ByVal allOK As Boolean, ByVal meldungen As Collection)

        Dim myName As String = My.Computer.FileSystem.GetName(fullfileName)

        ' reading messages of logfile
        Dim errMessages As Collection = readlogger(ptErrLevel.logError)
        Dim warnMessages As Collection = readlogger(ptErrLevel.logWarning)

        Dim mailMessage As String = "[" & Format(Now, "dd.MM.yyyy hh:mm:ss") & "] " & vbCrLf

        For i As Integer = 1 To meldungen.Count
            Dim text As String = CStr(meldungen.Item(i))
            mailMessage = mailMessage & text & vbCrLf
        Next

        For i As Integer = 1 To warnMessages.Count
            Dim text As String = CStr(warnMessages.Item(i))
            mailMessage = mailMessage & text & vbCrLf
        Next

        'nur dann ist fehlerfrei importiert
        allOK = allOK And meldungen.Count < 1 And errMessages.Count < 1

        If allOK Then
            Dim newDestination As String = My.Computer.FileSystem.CombinePath(successFolder, myName)
            My.Computer.FileSystem.MoveFile(fullfileName, newDestination, True)
            Call logger(ptErrLevel.logInfo, "success: ", myName)

            errMsgCode = New clsErrorCodeMsg
            result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & myName & ": successful ..." & vbCrLf _
                                                                            & mailMessage, errMsgCode)

        Else
            Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
            If My.Computer.FileSystem.FileExists(fullfileName) Then
                My.Computer.FileSystem.MoveFile(fullfileName, newDestination, True)
                Call logger(ptErrLevel.logError, "failed: ", fullfileName)

                'Dim errMessages As Collection = readlogger(ptErrLevel.logError)

                Dim logfileName As String = My.Computer.FileSystem.GetName(logfileNamePath)
                Dim newLog As String = My.Computer.FileSystem.CombinePath(failureFolder, logfileName)
                My.Computer.FileSystem.MoveFile(logfileNamePath, newLog, True)

                'Dim mailMessage As String = "[" & Format(Now, "dd.MM.yyyy hh:mm:ss") & "] " & vbCrLf

                'For i As Integer = 1 To meldungen.Count
                '    Dim text As String = CStr(meldungen.Item(i))
                '    mailMessage = mailMessage & text & vbCrLf
                'Next

                'For i As Integer = 1 To errMessages.Count
                '    Dim text As String = CStr(errMessages.Item(i))
                '    mailMessage = mailMessage & text & vbCrLf
                'Next


                errMsgCode = New clsErrorCodeMsg
                'result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                '                                                            & myName & ": with errors ..." & vbCrLf _
                '                                                            & "Look for more details in the Failure-Folder: " & failureFolder, errMsgCode)
                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                                                                            & myName & ": with errors ..." & vbCrLf _
                                                                            & mailMessage, errMsgCode)
            End If
        End If


        ' wieder in das normale logfile schreiben
        logfileNamePath = createLogfileName(rpaFolder)

        If Not result Then
            If awinSettings.englishLanguage Then
                msgTxt = "Sending an Email to report the result failed !"
            Else
                msgTxt = "Beim Senden einer Email, um das Ergebnis zu melden, ging schief !"
            End If
            Call logger(ptErrLevel.logError, "processResult", msgTxt)
        End If
    End Sub

    Public Function processNewImportFile(ByVal fileName As String) As Boolean

        Dim fullFileName As String = fileName
        Dim myName As String = ""
        Dim rpaCategory As New PTRpa
        Dim result As Boolean = False

        ' Completion-File delivered?
        completedOK = LCase(fullFileName).Contains(LCase("Timesheet_completed"))
        If completedOK Then


            Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & fullFileName & "' was created at: " & Date.Now().ToLongDateString)

            'Einlesen der TimeSheets - Telair
            ' nachsehen ob collect vollständig
            myName = My.Computer.FileSystem.GetName(fullFileName)
            result = processVisboActualData2(myName, myActivePortfolio, collectFolder, Date.Now())
            ' TODO: löschen des Timesheet-compl
            If result Then
                Dim newDestination As String = My.Computer.FileSystem.CombinePath(successFolder, myName)
                My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                Call logger(ptErrLevel.logInfo, "success: ", myName)

                ' wieder in das normale logfile schreiben
                logfileNamePath = createLogfileName(rpaFolder)
                errMsgCode = New clsErrorCodeMsg
                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & myName & ": successful ...", errMsgCode)
            Else
                Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
                If My.Computer.FileSystem.FileExists(fullFileName) Then
                    My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                    'Call logger(ptErrLevel.logError, "failed: ", fullFileName)
                    'Dim logfileName As String = My.Computer.FileSystem.GetName(logfileNamePath)
                    'Dim newLog As String = My.Computer.FileSystem.CombinePath(failureFolder, logfileName)
                    'My.Computer.FileSystem.MoveFile(logfileNamePath, newLog, True)

                    ' wieder in das normale logfile schreiben
                    'logfileNamePath = createLogfileName(rpaFolder)
                    'errMsgCode = New clsErrorCodeMsg
                    'result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                    '                                                            & myName & ": with errors ..." & vbCrLf _
                    '                                                            & "Look for more details in the Failure-Folder", errMsgCode)
                End If
            End If
        Else
            If My.Computer.FileSystem.FileExists(fullFileName) And Not fullFileName.Contains("~$") Then

                Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & fullFileName & "' was created at: " & Date.Now().ToLongDateString)

                'FileExtension ansehen
                Dim fileExt As String = My.Computer.FileSystem.GetFileInfo(fullFileName).Extension
                Select Case fileExt
                    Case ".xlsx"

                        myName = My.Computer.FileSystem.GetName(fullFileName)

                        ' Bestimme den Import-Typ der zu importierenden Daten
                        rpaCategory = bestimmeRPACategory(fullFileName)

                        If rpaCategory = PTRpa.visboUnknown Then
                            ' move file to unknown Folder ... 
                            Dim newDestination As String = My.Computer.FileSystem.CombinePath(unknownFolder, myName)
                            My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                            Call logger(ptErrLevel.logInfo, "unknown file / category: ", myName)
                        Else
                            result = importOneProject(fullFileName, rpaCategory, Date.Now())
                            If result Then
                                Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & fullFileName & "' was imported successfully at: " & Date.Now().ToLongDateString)
                            End If
                        End If
                    Case ".mpp"

                        myName = My.Computer.FileSystem.GetName(fullFileName)

                        ' Import Typ ist Microsoft Project File
                        rpaCategory = PTRpa.visboMPP

                        ' Import wird durchgeführt
                        result = importOneProject(fullFileName, rpaCategory, Date.Now())
                        If result Then
                            Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & fullFileName & "' was imported successfully at: " & Date.Now().ToLongDateString)
                        End If

                    Case Else
                        myName = My.Computer.FileSystem.GetName(fullFileName)
                        rpaCategory = PTRpa.visboUnknown
                        ' move file to unknown Folder ... 
                        Dim newDestination As String = My.Computer.FileSystem.CombinePath(unknownFolder, myName)

                        Try
                            My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                        Catch ex As Exception
                            Call MsgBox("try catch watch.created" & ex.Message)
                        End Try

                        Call logger(ptErrLevel.logInfo, "unknown file / category: unknown", myName)

                        errMsgCode = New clsErrorCodeMsg
                        result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                                                                                    & myName & vbCrLf & " unknown file / category ...", errMsgCode)
                End Select
            Else
                Dim a As String = ""
            End If
        End If


        'If My.Computer.FileSystem.FileExists(fullFileName) And Not fullFileName.Contains("~$") Then

        '    Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & fullFileName & "' was created at: " & Date.Now().ToLongDateString)

        '    'FileExtension ansehen
        '    Dim fileExt As String = My.Computer.FileSystem.GetFileInfo(fullFileName).Extension
        '    Select Case fileExt
        '        Case ".xlsx"

        '            myName = My.Computer.FileSystem.GetName(fullFileName)

        '            ' Bestimme den Import-Typ der zu importierenden Daten
        '            rpaCategory = bestimmeRPACategory(fullFileName)

        '            If rpaCategory = PTRpa.visboUnknown Then
        '                ' move file to unknown Folder ... 
        '                Dim newDestination As String = My.Computer.FileSystem.CombinePath(unknownFolder, myName)
        '                My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
        '                Call logger(ptErrLevel.logInfo, "unknown file / category: ", myName)
        '            Else
        '                result = importOneProject(fullFileName, rpaCategory, Date.Now())
        '                If result Then
        '                    Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & fullFileName & "' was imported successfully at: " & Date.Now().ToLongDateString)
        '                End If
        '            End If
        '        Case ".mpp"

        '            myName = My.Computer.FileSystem.GetName(fullFileName)

        '            ' Import Typ ist Microsoft Project File
        '            rpaCategory = PTRpa.visboMPP

        '            ' Import wird durchgeführt
        '            result = importOneProject(fullFileName, rpaCategory, Date.Now())
        '            If result Then
        '                Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & fullFileName & "' was imported successfully at: " & Date.Now().ToLongDateString)
        '            End If

        '        Case Else
        '            myName = My.Computer.FileSystem.GetName(fullFileName)
        '            rpaCategory = PTRpa.visboUnknown
        '            ' move file to unknown Folder ... 
        '            Dim newDestination As String = My.Computer.FileSystem.CombinePath(unknownFolder, myName)

        '            Try
        '                My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
        '            Catch ex As Exception
        '                Call MsgBox("try catch watch.created" & ex.Message)
        '            End Try

        '            Call logger(ptErrLevel.logInfo, "unknown file / category: unknown", myName)

        '            errMsgCode = New clsErrorCodeMsg
        '            result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
        '                                                                        & myName & vbCrLf & " unknown file / category ...", errMsgCode)
        '    End Select
        'Else
        '    Dim a As String = ""
        'End If

        processNewImportFile = result

    End Function
End Module
