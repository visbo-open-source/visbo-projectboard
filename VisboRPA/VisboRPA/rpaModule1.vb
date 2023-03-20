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
    ' indicates if or if not successful Import will send an email
    Public noSucessEmails As Boolean

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


    Public Enum PTRpa
        ' represents the standard VISBO Projectbrief with Stammdaten, Ressources, Termine, Attribute 
        visboProject = 0

        ' represents the standard VISBO Excel project with just only name and Schedules, Appearances, and the like 
        visboExcelSchedules = 1

        ' represents the standard MS Project *.mpp File 
        visboMPP = 2

        ' represents the Jira File, as customized in JiraConfig customized  
        visboJira = 3

        ' represents the Instart AngebotsKalkulation Template 
        visboInstartProposal = 4

        ' represents the VISBO AngebotsKalkulation Template 
        visboProposal = 5

        ' represents the Telair Tagetik New Projects List
        visboNewTagetik = 6

        ' represents the Telair Update Project File
        visboUpdateTagetik = 7

        ' represents the standard VISBO Project Creation by BatchList 
        visboProjectList = 8

        ' represents the AllianzType Istdaten Import 
        visboActualData1 = 9

        ' represents the InstartType Istdaten Import 
        visboActualData2 = 10

        ' represents the Telair Istdaten Import 
        visboActualData3 = 11

        ' represents the initial VISBO Excel Organisation
        visboInitialOrga = 12

        ' represents the roundtrip VISBO Excel Organisation
        visboRoundtripOrga = 13

        ' represents the default Urlaubskalender from VISBO 
        visboDefaultCapacity = 14

        ' represents the Zeuss Urlaubskalender from VISBO 
        visboZeussCapacity = 15

        ' represents the Instart Type of Urlaubs-Information 
        visboEGeckoCapacity = 16

        ' represents the Allianz-Type Daten of Externe Rahmenverträge 
        visboExternalContracts = 17

        ' represents the classic modifier strcture 
        visboModifierCapacities = 18

        ' represents the unknown 
        visboUnknown = 19

        ' visbo Find Project Starts
        visboFindProjectStart = 20

        ' represents the CostAssertion of Telair
        visboCostAssertion = 21


        ' represents the Automatic Team Allocation
        visboSuggestResourceAllocation = 22

        ' represent the setitngs 
        visboJsonSetting = 23

        ' represents the Auto-Distribution
        visboAutoAdjust = 24

        ' create hedged variants 
        visboCreateHedgedVariant = 25

        ' visbo Find Project Starts with regard of frequency Phases, milestones
        visboFindProjectStartPM = 26

        ' find feasible Portfolio
        visboFindfeasiblePortfolio = 27

        ' represents the weser ressourcenplan
        visboWWWRessourcen = 28

        ' dataQuality Check 
        visboDataQualityCheck = 29

        ' rename projects in Batch
        visboRenameProjects = 30

        ' create baselines in Batch
        visboCreateBaselineProjects = 31

        ' assign attributes such as buinsessUnits in batchmode
        visboAssignAttributes = 32

    End Enum


    Public Sub Main()
        ' reads the VISBO RPA folder und treats each file it finds there appropriately
        ' in most cases new project and portfolio versions will be written 
        ' suggestions for Team Members will follow 
        ' automation in resource And team allocation will follow

        Dim actDir = My.Computer.FileSystem.CurrentDirectory

        ' success Emails will be sent
        noSucessEmails = False

        'Call MsgBox("TempFile:" & tempFile)

        ' name des aktuell laufenden Clients
        'visboClient = "VISBO RPA /"
        visboClient = divClients(client.VisboRPA)

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

        ' soll ein neuer Login gemacht werden - dann true, wenn VISBOMode = Demoe
        awinSettings.autoLogin = (My.Settings.VISBOMode <> "Demo")

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

        If Not awinSettings.autoLogin Then
            Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "Login for starting")
            loginErfolgreich = logInToMongoDB(True)
        Else

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
        Dim listOfArchivFiles As New List(Of String)
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
                    allOk = processModifierCapacities(fname, importDate, errMessages)
                    'Call logger(ptErrLevel.logError, "import Modifier Capacities", " not yet implemented !")

                Case CInt(PTRpa.visboExternalContracts)
                    allOk = True
                    allOk = processExternalContracts(fname, importDate, errMessages)
             '       Call logger(ptErrLevel.logError, "import external Contracts", " not yet implemented !")


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
                    allOk = processNewTagetik(fname, myActivePortfolio, collectFolder, importDate)
                    'Call logger(ptErrLevel.logError, "Import new Projects of Tagetik", " not yet integrated !")

                Case CInt(PTRpa.visboUpdateTagetik)
                    allOk = True
                    allOk = processUpdateTagetik(fname, myActivePortfolio, collectFolder, importDate)
                    'Call logger(ptErrLevel.logError, "Import Project-update of Tagetik", " not yet integrated !")

                Case CInt(PTRpa.visboEGeckoCapacity)
                    allOk = True
                    allOk = processEGeckoCapacity(fname, importDate, errMessages)
                    'Call logger(ptErrLevel.logError, "Import Capacities coming from eGecko", " not yet integrated !")

                Case CInt(PTRpa.visboInstartProposal)
                    allOk = processInstartProposal(fname, myActivePortfolio, collectFolder, importDate)
                    'Call logger(ptErrLevel.logError, "Import Calc-Sheet ", " not yet integrated !")

                Case CInt(PTRpa.visboWWWRessourcen)
                    allOk = processWWWRessourcen(fname, myActivePortfolio, collectFolder, importDate)
                    'Call logger(ptErrLevel.logError, "Import Calc-Sheet ", " not yet integrated !")

                Case CInt(PTRpa.visboProposal)
                    allOk = True
                    Call logger(ptErrLevel.logError, "Import visbo proposal ", " not yet integrated !")

                Case CInt(PTRpa.visboZeussCapacity)
                    allOk = True
                    currentWB.Close(SaveChanges:=False)
                    allOk = processZeussCapacity(fname, importDate, errMessages, listOfArchivFiles)
                    'Call logger(ptErrLevel.logError, "Import Zeuss-Capacities ", " not yet integrated !")

                Case CInt(PTRpa.visboFindProjectStart)

                    allOk = processFindProjectStart(myName)

                Case CInt(PTRpa.visboFindProjectStartPM)

                    allOk = processFindProjectStart(myName, PTRpa.visboFindProjectStartPM)

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


                Case CInt(PTRpa.visboCostAssertion)

                    allOk = processCostAssertion(fname, myActivePortfolio, collectFolder, importDate)


                Case CInt(PTRpa.visboDataQualityCheck)
                    Try
                        allOk = processDataQualityCheck()
                    Catch ex As Exception

                    End Try

                Case CInt(PTRpa.visboRenameProjects)
                    Try
                        allOk = processRenameProjects()
                    Catch ex As Exception

                    End Try

                Case CInt(PTRpa.visboCreateBaselineProjects)
                    Try
                        allOk = processCreateBaselines()
                    Catch ex As Exception

                    End Try

                Case CInt(PTRpa.visboAssignAttributes)
                    Try
                        allOk = processAssignAttributes()
                    Catch ex As Exception

                    End Try

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
                    If Not IsNothing(currentWB) Then
                        currentWB.Close(SaveChanges:=False)
                    End If
                End If
            Catch ex As Exception

            End Try

            ' here the logfiles and the importfiles will be moved to a folder depending on the result of the import
            If Not rpaCat = PTRpa.visboActualData2 Then
                If listOfArchivFiles.Count > 0 Then
                    For Each archivFile As String In listOfArchivFiles
                        Call processResult(archivFile, allOk, errMessages)
                    Next
                Else
                    Call processResult(fname, allOk, errMessages)
                End If
            Else
                If listOfArchivFiles.Count > 0 Then
                    For Each archivFile As String In listOfArchivFiles
                        Call processResult(archivFile, allOk, errMessages)
                    Next
                Else
                    Call processResult(fname, allOk, errMessages)
                End If
                'Call processResult(fname, allOk, errMessages)
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
    Public Sub emptyRPASession()
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

                        Call logger(ptErrLevel.logWarning, "baseline couldn't be created: ", outputCollection)
                        myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung

                        If storeSingleProjectToDB(kvp.Value, outputCollection) Then
                            ok = ok And True
                            Call logger(ptErrLevel.logInfo, "project stored: ", kvp.Value.getShapeText)
                            Console.WriteLine("project stored: " & kvp.Value.getShapeText)
                        Else
                            ok = ok And False
                            Call logger(ptErrLevel.logError, "project store failed: ", outputCollection)
                            Console.WriteLine("!! ... project store failed: " & kvp.Value.getShapeText)
                        End If

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

                        ' Check auf Auto Adjust Resource Bottlenecks
                        If result = PTRpa.visboUnknown Then
                            result = checkAutoAdjustPortfolio(currentWB)
                        End If

                        ' Check auf VISBO Project Template  
                        ' Check auf VISBO Project Brief and VISBO Project Template
                        ' Template has to contain the word "template" within the filename

                        If result = PTRpa.visboUnknown Then
                            result = checkProjectBrief(currentWB)
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

                        ' Check auf VISBO Urlaubskalender 
                        If result = PTRpa.visboUnknown Then
                            result = checkUrlaubsplaner(currentWB)
                        End If

                        ' Check auf Modifier Kapazitäten
                        If result = PTRpa.visboUnknown Then
                            result = checkModifierExternKapa(currentWB)
                        End If

                        ' Check auf externe Rahmenverträge 
                        If result = PTRpa.visboUnknown Then
                            result = checkExtRahmenvertr(currentWB)
                        End If

                        ' Check auf Instart eGecko Urlaube ...(Instart) 
                        If result = PTRpa.visboUnknown Then
                            result = checkInstartUrlaub(currentWB)
                        End If

                        ' Check auf Zeuss Kapazitäten... (Telair)
                        If result = PTRpa.visboUnknown Then
                            result = checkZeussCapacity(currentWB)
                        End If

                        ' Check auf Ist-Daten 
                        If result = PTRpa.visboUnknown Then
                            result = checkActualData1(currentWB)
                        End If

                        ' Check auf Telair TimeSheets
                        If result = PTRpa.visboUnknown Then
                            result = checkActualData2(currentWB)
                        End If

                        ' Check auf Tagetik new Project List 
                        If result = PTRpa.visboUnknown Then
                            result = checkTagetikProjectList(currentWB)
                        End If

                        ' Check auf Tagetik update projects 
                        If result = PTRpa.visboUnknown Then
                            result = checkTagetikUpdateProjectList(currentWB)
                        End If

                        ' Check auf Cost-Assertion Telair 
                        If result = PTRpa.visboUnknown Then
                            result = checkCostAssertion(currentWB)
                        End If

                        ' Check auf Instart Calculation Template 
                        If result = PTRpa.visboUnknown Then
                            result = checkInstartProposal(currentWB)
                        End If

                        ' Check auf Weser Ressourcenplanung 
                        If result = PTRpa.visboUnknown Then
                            result = checkWWWRessourcen(currentWB)
                        End If

                        ' Check auf Data Quality Check (Currently BHTC) 
                        If result = PTRpa.visboUnknown Then
                            result = checkDataQuality(currentWB)
                        End If

                        ' Check auf Rename Projects 
                        If result = PTRpa.visboUnknown Then
                            result = checkRename(currentWB)
                        End If

                        ' Check auf Create Baseline Projects 
                        If result = PTRpa.visboUnknown Then
                            result = checkBaselineCreation(currentWB)
                        End If

                        If result = PTRpa.visboUnknown Then
                            result = checkAssignAttributes(currentWB)
                        End If
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
    ''' checks whether or not it is a ModifierCapacities-File
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkModifierExternKapa(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim possibleTableNames() As String = {"Kapazität"}
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
                result = PTRpa.visboModifierCapacities
            End If


        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try


        checkModifierExternKapa = result
    End Function
    ''' <summary>
    ''' checks whether or not it is a File with external Contracts (like Allianz)
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkExtRahmenvertr(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim possibleTableNames() As String = {"externe Vertraege", "externe Rahmenvertraege", "Werte in Euro"}
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

    '''' <summary>
    '''' returns form Parameters the Portfolio-Name and Vname 
    '''' </summary>
    '''' <returns></returns>
    '''' 
    'Public Function getNameList(ByVal blattName As String) As Collection
    '    Dim result As New Collection


    '    Try

    '        Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
    '                                                        Global.Microsoft.Office.Interop.Excel.Workbook)

    '        Dim currentWS As xlns.Worksheet = CType(currentWB.Sheets.Item(blattName), Global.Microsoft.Office.Interop.Excel.Worksheet)

    '        Dim zeile As Integer = 2
    '        Dim spalte As Integer = 1



    '        If Not IsNothing(currentWS) Then
    '            With currentWS
    '                Dim lastRow As Integer = CType(.Cells(20000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

    '                While zeile <= lastRow
    '                    Dim pName As String = CStr(CType(.Cells(zeile, spalte), Global.Microsoft.Office.Interop.Excel.Range).Value)
    '                    If Not IsNothing(pName) Then
    '                        pName = pName.Trim

    '                        If pName <> "" Then
    '                            If Not result.Contains(pName) Then
    '                                result.Add(pName, pName)
    '                            End If
    '                        End If

    '                        zeile = zeile + 1
    '                    End If
    '                End While

    '            End With
    '        End If
    '    Catch ex As Exception

    '    End Try

    '    getNameList = result
    'End Function


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
    Private Function checkActualData2(ByVal currentWB As xlns.Workbook) As PTRpa
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

        checkActualData2 = result
    End Function

    Private Function checkDataQuality(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "Data Quality Check"

        Try
            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName), xlns.Worksheet)

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                result = PTRpa.visboDataQualityCheck

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkDataQuality = result

    End Function
    ''' <summary>
    '''  checks whether or not a file is a Weser Ressourcenplanung
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkWWWRessourcen(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "Kostencontrolling"

        Try
            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName), xlns.Worksheet)

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                Dim firstUsefullLine As Integer = CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlDown).Row
                Dim zweiteZeile As xlns.Range = CType(currentWS.Rows.Item(firstUsefullLine), xlns.Range)
                Try

                    verifiedStructure = CStr(zweiteZeile.Cells(1, 2).value).Trim.Contains("Projektressourcenplan")

                Catch ex As Exception
                    verifiedStructure = False
                End Try

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboWWWRessourcen
        Else
            result = PTRpa.visboUnknown
        End If

        checkWWWRessourcen = result
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

                Catch ex As Exception
                    verifiedStructure = False
                End Try

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboInstartProposal
        Else
            result = PTRpa.visboUnknown
        End If

        checkInstartProposal = result
    End Function



    ''' <summary>
    ''' checks whether or not a file is a Instart eGecko-Urlaubskalender
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkInstartUrlaub(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "Arbeitszeitauswertung"   '????

        Dim currentWS As xlns.Worksheet = Nothing
        Dim found As Boolean = False
        Dim wb As xlns.Workbook = currentWB

        Try
            For Each tmpSheet As xlns.Worksheet In currentWB.Worksheets

                If tmpSheet.Name.Contains(blattName) Then
                    found = True
                    currentWS = tmpSheet
                    Exit For
                End If
            Next

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else


                'Dim firstUsefullLine As Integer = CType(currentWS.Cells(1, 2), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlDown).Row
                'Dim zweiteZeile As xlns.Range = CType(currentWS.Rows.Item(firstUsefullLine), xlns.Range)
                'Try

                '    verifiedStructure = CStr(zweiteZeile.Cells(1, 2).value).Trim.Contains("Phase/Arbeitspaket")

                'Catch ex As Exception
                '    verifiedStructure = False
                'End Try

                verifiedStructure = True

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboEGeckoCapacity
        Else
            result = PTRpa.visboUnknown
        End If

        checkInstartUrlaub = result
    End Function




    ''' <summary>
    ''' checks whether or not a file is a Zeuss (Telair)-Urlaubskalender
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkZeussCapacity(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim fName As String = "Zeuss"   '????

        Dim currentWS As xlns.Worksheet = Nothing
        Dim found As Boolean = False
        Dim wb As xlns.Workbook = currentWB

        Try
            If currentWB.Name.Contains(fName) Then

                currentWS = currentWB.Worksheets.Item(1)

                If IsNothing(currentWS) Then
                    result = PTRpa.visboUnknown
                Else


                    Dim firstUsefullLine As Integer = CType(currentWS.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlDown).Row
                    Dim zweiteZeile As xlns.Range = CType(currentWS.Rows.Item(firstUsefullLine), xlns.Range)
                    Try

                        verifiedStructure = CStr(zweiteZeile.Cells(1, 1).value).Trim.Contains("Jahr:")

                    Catch ex As Exception
                        verifiedStructure = False
                    End Try

                    verifiedStructure = True

                End If
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboZeussCapacity
        Else
            result = PTRpa.visboUnknown
        End If

        checkZeussCapacity = result
    End Function


    ''' <summary>
    ''' checks whether or not a file is a Create Project List (Tagetik) Telair
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkTagetikProjectList(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim blattName As String = "BASE"

        Try
            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName), xlns.Worksheet)

            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                Dim firstLine As xlns.Range = CType(currentWS.Rows.Item(1), xlns.Range)

                Try
                    verifiedStructure = CStr(firstLine.Cells(1, 1).value).Trim.Contains("Budget")

                Catch ex As Exception
                    verifiedStructure = False
                End Try

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboNewTagetik
        Else
            result = PTRpa.visboUnknown
        End If

        checkTagetikProjectList = result
    End Function


    ''' <summary>
    ''' checks whether or not a file is a Update Project List Telair (Tagetik)
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkTagetikUpdateProjectList(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim verifiedStructure As Boolean = False
        Dim possibleTableNames() As String = {"Telair"}

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


            If IsNothing(currentWS) Then
                result = PTRpa.visboUnknown
            Else
                Dim tmpRange As xlns.Range = CType(currentWS.Range(currentWS.Cells(1, 1), currentWS.Cells(30, 40)), xlns.Range)
                Dim xxx As Object = tmpRange.Find("Forecast")
                If Not IsNothing(xxx) Then
                    result = PTRpa.visboUpdateTagetik
                Else
                    result = PTRpa.visboUnknown
                End If
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try



        checkTagetikUpdateProjectList = result
    End Function


    ''' <summary>
    ''' checks whether or not it is a cost Assertion of Telair
    ''' </summary>
    ''' <param name="currentWB"></param>
    ''' <returns></returns>
    Private Function checkCostAssertion(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim requiredTableName As String = "Stammdaten"
        Dim possibleTableNames() As String = {"Assertion", "to-do", "Kalkulation", "Calculation"}
        Dim verifiedStructure As Boolean = False
        Try

            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets, xlns.Sheets).Item(requiredTableName)
            Dim found As Boolean = False


            If IsNothing(currentWB) Then
                result = PTRpa.visboUnknown
            Else
                For Each tmpSheet As xlns.Worksheet In CType(currentWB.Worksheets, xlns.Sheets)
                    For Each tblname As String In possibleTableNames
                        If tmpSheet.Name.ToLower.Contains(tblname.ToLower) Then
                            found = True
                            currentWS = tmpSheet
                            Exit For
                        End If
                    Next
                Next
            End If

            If found Then
                result = PTRpa.visboCostAssertion
            End If

        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        checkCostAssertion = result
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

            If isTemplate Then
                ' read the file and import into vproj
                Call awinImportProjectmitHrchy(Nothing, vproj, True, importDate)

                Dim template As New clsProjekt
                ' mache aus clsprojektVorlage ein 'clsProjekt'
                Dim startDate As Date = StartofCalendar
                Dim endDate As Date = startDate.AddDays(vproj.dauerInDays - 1)
                Dim myProject As clsProjekt = Nothing
                ' zunächst die eingelesene Vorlage in die Liste der Projektvorlagen hinzufügen
                Projektvorlagen.Add(vproj)
                template = erstelleProjektAusVorlage(myProject, vproj.VorlagenName, vproj.VorlagenName, startDate, endDate, vproj.Erloes, 0, 5.0, 5.0, "0", vproj.VorlagenName, "", "", True)
                If Not IsNothing(template) Then
                    template.name = vproj.VorlagenName
                    template.projectType = ptPRPFType.projectTemplate
                End If
                hproj = template

            Else

                ' read the file and import into hproj
                Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)

            End If


            allOK = Not IsNothing(hproj)

            If allOK Then
                Try

                    ImportProjekte.Add(hproj, updateCurrentConstellation:=False)

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

                ' actualize Template-List 
                If isTemplate And allOK Then
                    lastReadingProjectTemplates = readProjectTemplates()
                End If
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

    ''' <summary>
    ''' Liest eine in Excel beschriebene Liste mit Projekten  (VISBO defined)
    ''' </summary>
    ''' <param name="myName"></param>
    ''' <param name="myActivePortfolio"></param>
    ''' <returns></returns>
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

    ''' <summary>
    ''' modifies capacities of all person resources , be it interns or externs 
    ''' </summary>
    ''' <param name="myName"></param>
    ''' <param name="importDate"></param>
    ''' <param name="errMessages"></param>
    ''' <returns></returns>
    Private Function processModifierCapacities(ByVal myName As String, ByVal importDate As Date, ByRef errMessages As Collection) As Boolean

        Dim listOfArchivFiles As New List(Of String)
        Dim result As Boolean = False

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Dim outputCollection As New Collection

        lastReadingOrganisation = readOrganisations()

        Dim changedOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)

        If Not IsNothing(changedOrga) Then

            If changedOrga.allRoles.Count > 0 Then

                RoleDefinitions = changedOrga.allRoles
                CostDefinitions = changedOrga.allCosts


                ' Liste enthält die Datei-Namen der erfolgreich eingelesenen externen Kapazitäts-Files 
                Dim listOfArchivExtern As New List(Of String)

                ' wenn es gibt - lesen der Externen Verträge 
                result = readKapaModifier(myName, listOfArchivExtern, errMessages)

                If result Then
                    Call logger(ptErrLevel.logInfo, "Import capacities from file " & myName & " successful", "processModifierCapacities", anzFehler)
                Else
                    Call logger(ptErrLevel.logError, "Import capacities from file " & myName & " NOT successful", "processModifierCapacities", anzFehler)
                    For Each singleMsg As String In errMessages
                        Call logger(ptErrLevel.logError, singleMsg, "processModifierCapacities", anzFehler)
                    Next
                End If

                If listOfArchivExtern.Count > 0 Then

                    changedOrga.allRoles = RoleDefinitions

                    If outputCollection.Count = 0 Then
                        ' keine Fehler aufgetreten ... 
                        ' jetzt wird die Orga als Setting weggespeichert ... 
                        Dim err As New clsErrorCodeMsg
                        Dim resultSum As Boolean = True
                        Dim capas As clsCapas = Nothing

                        ' ute -> überprüfen bzw. fertigstellen ... 
                        Dim orgaName As String = ptSettingTypes.organisation.ToString

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Or (visboClient = divClients(client.VisboRPA)) Then

                            ' now stores everything from RoleDefinitions what needs to be stored ... 
                            resultSum = storeCapasOfRoles()

                            If resultSum = True Then
                                Call logger(ptErrLevel.logInfo, "ok, capacities " & myName & " successfully updated ...", "", -1)
                                listOfArchivFiles = listOfArchivExtern

                            Else
                                Call logger(ptErrLevel.logError, "Error when writing capacities " & myName & "to Database..." & vbCrLf & err.errorMsg, "", -1)
                                listOfArchivFiles = listOfArchivExtern
                            End If

                            result = resultSum

                        Else
                            Call logger(ptErrLevel.logError, "ok, capacities " & myName & " temporarily updated ...", "", -1)

                        End If

                    Else

                        Call showOutPut(outputCollection, "Importing capacities", "... mit Fehlern abgebrochen ...")
                        Call logger(ptErrLevel.logError, "processModifierCapacities: ", outputCollection)

                    End If
                Else
                    If outputCollection.Count > 0 Then

                        Call showOutPut(outputCollection, "Importing Capacities", "... mit Fehlern abgebrochen ...")
                        Call logger(ptErrLevel.logError, "processModifierCapacities: ", outputCollection)
                    Else

                        If awinSettings.englishLanguage Then
                            Call logger(ptErrLevel.logError, "no Files to import ...", "processModifierCapacities: ", anzFehler)
                        Else
                            Call logger(ptErrLevel.logError, "es gab keine Dateien zum Einlesen ... ", "processModifierCapacities: ", anzFehler)
                        End If
                    End If

                End If
            Else
                If awinSettings.englishLanguage Then
                    Call logger(ptErrLevel.logError, "No valid roles! Please import one first!", "processModifierExternContracts: ", anzFehler)
                Else
                    Call logger(ptErrLevel.logError, "Die gültige Organisation beinhaltet keine Rollen! ", "processModifierExternContracts: ", anzFehler)

                End If
            End If

        Else
            If awinSettings.englishLanguage Then
                Call logger(ptErrLevel.logError, "No valid organization! Please import one first!", "processModifierCapacities: ", anzFehler)
            Else
                Call logger(ptErrLevel.logError, "Es existiert keine gültige Organisation! Bitte zuerst Organisation importieren", "processModifierCapacities: ", anzFehler)
            End If


            Dim errMsg As String = "Kapazitäten wurden nicht aktualisiert - bitte erst die Import-Dateien korrigieren ... "
            outputCollection.Add(errMsg)
            Call showOutPut(outputCollection, "Importing Capacities", "")
            Call logger(ptErrLevel.logError, "processModifierCapacities: ", outputCollection)

        End If

        processModifierCapacities = result
    End Function


    Private Function processExternalContracts(ByVal myName As String, ByVal importDate As Date, ByRef errMessages As Collection) As Boolean


        Dim actualDataFile As String = ""
        Dim actualDataConfig As New SortedList(Of String, clsConfigActualDataImport)
        Dim outPutline As String = ""
        Dim lastrow As Integer = 0
        Dim listOfArchivFiles As New List(Of String)
        Dim result As Boolean = False

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False

        Dim outputCollection As New Collection

        lastReadingOrganisation = readOrganisations()

        Dim changedOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)

        If Not IsNothing(changedOrga) Then

            If changedOrga.allRoles.Count > 0 Then

                RoleDefinitions = changedOrga.allRoles
                CostDefinitions = changedOrga.allCosts


                ' Liste enthält die Datei-Namen der erfolgreich eingelesenen externen Kapazitäts-Files 
                Dim listOfArchivExtern As New List(Of String)

                ' wenn es gibt - lesen der Externen Verträge 
                result = readKapaExtern(myName, listOfArchivExtern, errMessages)

                If result Then
                    Call logger(ptErrLevel.logInfo, "Import external contracts from file " & myName & " successful", "readMonthlyExternKapasEV", anzFehler)
                End If

                If listOfArchivExtern.Count > 0 Then

                    changedOrga.allRoles = RoleDefinitions

                    If outputCollection.Count = 0 Then
                        ' keine Fehler aufgetreten ... 
                        ' jetzt wird die Orga als Setting weggespeichert ... 
                        Dim err As New clsErrorCodeMsg
                        Dim resultSum As Boolean = True
                        Dim capas As clsCapas = Nothing

                        ' ute -> überprüfen bzw. fertigstellen ... 
                        Dim orgaName As String = ptSettingTypes.organisation.ToString

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Or (visboClient = divClients(client.VisboRPA)) Then

                            ' now stores everything from RoleDefinitions what needs to be stored ... 
                            resultSum = storeCapasOfRoles()

                            If resultSum = True Then
                                Call logger(ptErrLevel.logInfo, "ok, external Contracts " & myName & " successfully updated ...", "", -1)
                                listOfArchivFiles = listOfArchivExtern

                                '' verschieben der Kapa-Dateien Kapazität* Modifier  in den ArchivOrdner
                                'Call moveFilesInArchiv(listOfArchivExtern, importOrdnerNames(PTImpExp.Kapas))
                                '' verschieben der Kapa-Dateien Urlaubsplaner*.xlsx in den ArchivOrdner
                                'Call moveFilesInArchiv(listofArchivUrlaub, importOrdnerNames(PTImpExp.Kapas))
                                '' verschieben der Kapa-Dateien,die durch configCapaImport.xlsx beschrieben sind, in den ArchivOrdner
                                'Call moveFilesInArchiv(listofArchivConfig, importOrdnerNames(PTImpExp.Kapas))

                            Else
                                Call logger(ptErrLevel.logError, "Error when writing Capacities of external contract " & myName & "to Database..." & vbCrLf & err.errorMsg, "", -1)
                                listOfArchivFiles = listOfArchivExtern
                            End If

                            result = resultSum

                        Else
                            Call logger(ptErrLevel.logError, "ok, external Contracts " & myName & " temporarily updated ...", "", -1)

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

        processExternalContracts = result
    End Function


    Private Function processEGeckoCapacity(ByVal myName As String, ByVal importDate As Date, ByRef errMessages As Collection) As Boolean

        Dim result As Boolean = False
        Dim actualDataFile As String = ""
        Dim actualDataConfig As New SortedList(Of String, clsConfigActualDataImport)
        Dim outPutline As String = ""
        Dim lastrow As Integer = 0
        Dim listofArchivConfig As New List(Of String)

        appInstance.EnableEvents = False
        appInstance.ScreenUpdating = False
        enableOnUpdate = False


        Dim outputCollection As New Collection

        Dim changedOrga As clsOrganisation = validOrganisations.getOrganisationValidAt(Date.Now)

        If Not IsNothing(changedOrga) Then

            If changedOrga.allRoles.Count > 0 Then

                RoleDefinitions = changedOrga.allRoles
                CostDefinitions = changedOrga.allCosts

                ' Liste enthält die Datei-Namen der erfolgreich eingelesenen externen Kapazitäts-Files 
                Dim listOfArchivExtern As New List(Of String)

                ' wenn es gibt - lesen der EGecko-Files o.ä., die durch configCapaImport beschrieben sind
                Dim configCapaImport As String = configfilesOrdner & "\" & "configCapaImport.xlsx"
                If My.Computer.FileSystem.FileExists(configCapaImport) Then

                    listofArchivConfig = readInterneAnwesenheitslistenAllg(configCapaImport, actualDataConfig, outputCollection, myName)
                Else
                    outPutline = "There is no Config-File for the capacities!"
                    Call logger(ptErrLevel.logWarning, outPutline, "PTImportKapas", anzFehler)
                End If

                If listofArchivConfig.Count > 0 Then

                    changedOrga.allRoles = RoleDefinitions

                    If outputCollection.Count = 0 Then
                        ' keine Fehler aufgetreten ... 
                        ' jetzt wird die Orga als Setting weggespeichert ... 
                        Dim err As New clsErrorCodeMsg
                        Dim resultSum As Boolean = True
                        Dim capas As clsCapas = Nothing

                        ' Dim orga As clsOrganisation = Nothing

                        ' ute -> überprüfen bzw. fertigstellen ... 
                        Dim orgaName As String = ptSettingTypes.organisation.ToString

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Or (visboClient = divClients(client.VisboRPA)) Then

                            ' tk wozu brauche ich das hier ? 
                            ' orga = CType(databaseAcc, DBAccLayer.Request).retrieveTSOrgaFromDB("organisation", Date.Now, err, False, True, False)

                            ' now stores everything from RoleDefinitions what needs to be stored ... 
                            resultSum = storeCapasOfRoles()

                            If resultSum = True Then
                                Call logger(ptErrLevel.logInfo, "ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " updated ...", "", -1)

                            Else
                                Call logger(ptErrLevel.logError, "Error when writing Capacities to Database..." & vbCrLf & err.errorMsg, "", -1)
                            End If

                        Else
                            'Call MsgBox("ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " temporarily updated ...")
                            Call logger(ptErrLevel.logInfo, "ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " temporarily updated ...", "", -1)
                        End If

                        result = resultSum

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

        enableOnUpdate = True
        appInstance.EnableEvents = True
        appInstance.ScreenUpdating = True

        processEGeckoCapacity = result

    End Function

    Private Function processZeussCapacity(ByVal myName As String, ByVal importDate As Date, ByRef errMessages As Collection, ByRef listOfArchivFiles As List(Of String)) As Boolean


        Dim actualDataFile As String = ""
        Dim actualDataConfig As New SortedList(Of String, clsConfigActualDataImport)
        Dim outPutline As String = ""
        Dim lastrow As Integer = 0
        Dim listofArchivUrlaub As New List(Of String)
        Dim listofArchivConfig As New List(Of String)
        Dim configActualDataImport As String = "configActualDataImport.xlsx"
        Dim result As Boolean = False

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
                ' Call readMonthlyModifierKapas(outputCollection, listOfArchivExtern)

                ' wenn es gibt - lesen der Externen Verträge 
                ' Call readMonthlyExternKapasEV(outputCollection, listOfArchivExtern)

                '' wenn es gibt - lesen der Urlaubslisten DateiName "Urlaubsplaner*.xlsx
                ' listofArchivUrlaub = readInterneAnwesenheitslisten(outputCollection)

                ''  check Config-File - zum Einlesen der Istdaten gemäß Konfiguration -
                ''  - hier benötigt um den Kalender von IstDaten und Urlaubsdaten aufeinander abzustimmen
                configfilesOrdner = My.Computer.FileSystem.CombinePath(awinPath, configfilesOrdner)
                configActualDataImport = configfilesOrdner & "\" & "configActualDataImport.xlsx"
                Dim allesOK As Boolean = checkActualDataImportConfig(configActualDataImport, actualDataFile, actualDataConfig, lastrow, outputCollection)

                ' wenn es gibt - lesen der Zeuss- listen und anderer, die durch configCapaImport beschrieben sind
                Dim configCapaImport As String = configfilesOrdner & "\" & "configCapaImport.xlsx"
                If My.Computer.FileSystem.FileExists(configCapaImport) Then

                    listofArchivConfig = readInterneAnwesenheitslistenAllg(configCapaImport, actualDataConfig, outputCollection, myName)
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
                        Dim resultSum As Boolean = True
                        Dim capas As clsCapas = Nothing

                        ' Dim orga As clsOrganisation = Nothing

                        ' ute -> überprüfen bzw. fertigstellen ... 
                        Dim orgaName As String = ptSettingTypes.organisation.ToString

                        If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Or (visboClient = divClients(client.VisboRPA)) Then

                            ' tk wozu brauche ich das hier ? 
                            ' orga = CType(databaseAcc, DBAccLayer.Request).retrieveTSOrgaFromDB("organisation", Date.Now, err, False, True, False)

                            ' now stores everything from RoleDefinitions what needs to be stored ... 
                            resultSum = storeCapasOfRoles()

                            If resultSum = True Then
                                Call logger(ptErrLevel.logInfo, "ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " updated ...", "", -1)
                                listOfArchivFiles = listofArchivConfig

                                '' verschieben der Kapa-Dateien Kapazität* Modifier  in den ArchivOrdner
                                'Call moveFilesInArchiv(listOfArchivExtern, importOrdnerNames(PTImpExp.Kapas))
                                '' verschieben der Kapa-Dateien Urlaubsplaner*.xlsx in den ArchivOrdner
                                'Call moveFilesInArchiv(listofArchivUrlaub, importOrdnerNames(PTImpExp.Kapas))
                                '' verschieben der Kapa-Dateien,die durch configCapaImport.xlsx beschrieben sind, in den ArchivOrdner
                                'Call moveFilesInArchiv(listofArchivConfig, importOrdnerNames(PTImpExp.Kapas))

                            Else
                                Call logger(ptErrLevel.logError, "Error when writing Capacities to Database..." & vbCrLf & err.errorMsg, "", -1)
                                listOfArchivFiles = listofArchivConfig
                            End If

                            result = resultSum

                        Else
                            Call logger(ptErrLevel.logError, "ok, Capacities in organisation, valid from " & changedOrga.validFrom.ToString & " temporarily updated ...", "", -1)
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

        enableOnUpdate = True
        appInstance.EnableEvents = True

        appInstance.ScreenUpdating = True
        processZeussCapacity = result
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

                        If (myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Or myCustomUserRole.customUserRole = ptCustomUserRoles.Alles) Or (visboClient = divClients(client.VisboRPA)) Then


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

                            Call logger(ptErrLevel.logInfo, "PTImportIstDaten", logArray, logDblArray)
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
            Dim configProposalImportName As String = "configCalcTemplateImport.xlsx"
            Dim configProposalImport As String = My.Computer.FileSystem.CombinePath(configfilesOrdner, configProposalImportName)

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

                'If listofArchivAllg.Count > 0 Then
                '    Call moveFilesInArchiv(listofArchivAllg, importOrdnerNames(PTImpExp.projectWithConfig))
                'End If

                allesOK = (listofArchivAllg.Count > 0 And outPutCollection.Count = 0)

                If allesOK Then
                    ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                    Try
                        ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                        ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                        Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False, getSomeValuesFromOldProj:=False, calledFromActualDataImport:=False, calledFromRPA:=True)


                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            outputString = "Error at Import: " & vbLf & ex.Message
                        Else
                            outputString = "Fehler bei Import: " & vbLf & ex.Message
                        End If
                        outPutCollection.Add(outputString)

                    End Try

                Else

                    Call logger(ptErrLevel.logError, "checkProjectImportConfig", outPutCollection)

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

            End If


            allOk = allOk And allesOK

        Catch ex2 As Exception
            allOk = False
        End Try

        Try
            ' store Projects
            If allOk Then
                allOk = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboProposal.ToString, myName)

        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Projects Proposal", ex1.Message)
        End Try



        processInstartProposal = allOk
    End Function



    Public Function processWWWRessourcen(ByVal myName As String, ByVal portfolioName As String, ByVal dirName As String, ByVal importDate As Date) As Boolean
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
            Dim configWWWRessourcenImport As String = "configCalcTemplateImport.xlsx"


            Dim outputLine As String = ""

            Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

            ' Konfigurationsdatei lesen und Validierung durchführen

            ' wenn es gibt - lesen der Jira und anderer, die durch configCapaImport beschrieben sind
            ' no longer necessary
            ' Dim configJIRAProjects As String = My.Computer.FileSystem.CombinePath(configfilesOrdner, "configJIRAProjectImport.xlsx")

            ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
            configWWWRessourcenImport = configfilesOrdner & "\" & configWWWRessourcenImport
            Dim allesOK As Boolean = checkProjectImportConfig(configWWWRessourcenImport, projectsFile, projectConfig, lastrow, outPutCollection)

            If allesOK Then

                Dim listofVorlagen As New Collection
                listofVorlagen.Add(myName)
                If projectsFile = projectConfig("DateiName").ProjectsFile Then
                    listofArchivAllg = readProjectsAllg(listofVorlagen, projectConfig, outPutCollection, ptImportTypen.instartCalcTemplateImport)
                End If


                allesOK = (listofArchivAllg.Count > 0 And outPutCollection.Count = 0)

                If allesOK Then
                    ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                    Try
                        ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                        ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                        Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False, getSomeValuesFromOldProj:=False, calledFromActualDataImport:=False, calledFromRPA:=True)


                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            outputString = "Error at Import: " & vbLf & ex.Message
                        Else
                            outputString = "Fehler bei Import: " & vbLf & ex.Message
                        End If
                        outPutCollection.Add(outputString)

                    End Try

                Else

                    Call logger(ptErrLevel.logError, "checkProjectImportConfig", outPutCollection)

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

            End If


            allOk = allOk And allesOK

        Catch ex2 As Exception
            allOk = False
        End Try

        Try
            ' store Projects
            If allOk Then
                allOk = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboProposal.ToString, myName)

        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Projects Proposal", ex1.Message)
        End Try



        processWWWRessourcen = allOk
    End Function



    Public Function processNewTagetik(ByVal myName As String, ByVal portfolioName As String, ByVal dirName As String, ByVal importDate As Date) As Boolean
        Dim allOk As Boolean = True
        Dim aktDateTime As Date = Date.Now
        Dim telairImportConfigOK As Boolean = False


        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboNewTagetik.ToString, myName)

        'check the pre-conditions
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If
        If DateDiff(DateInterval.Hour, lastReadingProjectTemplates, aktDateTime) > 2 Then
            lastReadingProjectTemplates = readProjectTemplates()
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
            Dim configProjectsImport As String = "configProjectImport.xlsx"


            Dim outputLine As String = ""

            Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

            ' Konfigurationsdatei lesen und Validierung durchführen


            ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
            configProjectsImport = configfilesOrdner & "\" & configProjectsImport
            telairImportConfigOK = checkProjectImportConfig(configProjectsImport, projectsFile, projectConfig, lastrow, outPutCollection)

            If outPutCollection.Count > 0 Then
                Call logger(ptErrLevel.logError, "processNewTagetik", outPutCollection)
            End If


            If telairImportConfigOK Then

                Dim listofVorlagen As New Collection
                listofVorlagen.Add(myName)
                If projectsFile = projectConfig("DateiName").ProjectsFile Then
                    listofArchivAllg = readProjectsAllg(listofVorlagen, projectConfig, outPutCollection, ptImportTypen.telairTagetikImport)
                End If
                'listofArchivAllg = readProjectsJIRA(listofVorlagen, JIRAProjectsConfig, outPutCollection)

                'If listofArchivAllg.Count > 0 Then
                '    Call moveFilesInArchiv(listofArchivAllg, importOrdnerNames(PTImpExp.projectWithConfig))
                'End If

                allOk = (listofArchivAllg.Count > 0 And outPutCollection.Count = 0)

                If allOk Then
                    ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                    Try
                        ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                        ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                        Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False, getSomeValuesFromOldProj:=False, calledFromActualDataImport:=False, calledFromRPA:=True)


                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            outputString = "Error at Import: " & vbLf & ex.Message
                        Else
                            outputString = "Fehler bei Import: " & vbLf & ex.Message
                        End If
                        outPutCollection.Add(outputString)

                    End Try

                Else

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

            End If

            allOk = allOk And telairImportConfigOK

        Catch ex2 As Exception
            allOk = False
        End Try

        Try
            ' store Projects
            If allOk Then
                allOk = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboNewTagetik.ToString, myName)

        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Cost Assertion Projects", ex1.Message)
        End Try



        processNewTagetik = allOk
    End Function





    Public Function processUpdateTagetik(ByVal myName As String, ByVal portfolioName As String, ByVal dirName As String, ByVal importDate As Date) As Boolean
        Dim allOk As Boolean = True
        Dim aktDateTime As Date = Date.Now
        Dim telairUpdateConfigOK As Boolean = False


        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboUpdateTagetik.ToString, myName)

        'check the pre-conditions
        If DateDiff(DateInterval.Hour, lastReadingOrganisation, aktDateTime) > 2 Then
            lastReadingOrganisation = readOrganisations()
        End If
        If DateDiff(DateInterval.Hour, lastReadingProjectTemplates, aktDateTime) > 2 Then
            lastReadingProjectTemplates = readProjectTemplates()
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
            Dim configProjectsUpdates As String = "configProjectUpdates.xlsx"


            Dim outputLine As String = ""

            Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

            ' Konfigurationsdatei lesen und Validierung durchführen


            ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
            ' Konfigurationsdatei lesen und Validierung durchführen
            configfilesOrdner = My.Computer.FileSystem.CombinePath(awinPath, configfilesOrdner)
            configProjectsUpdates = configfilesOrdner & "\" & configProjectsUpdates
            telairUpdateConfigOK = checkProjectImportConfig(configProjectsUpdates, projectsFile, projectConfig, lastrow, outPutCollection)

            If outPutCollection.Count > 0 Then
                Call logger(ptErrLevel.logError, "processUpdateTagetik", outPutCollection)
            End If


            If telairUpdateConfigOK Then

                Dim listofVorlagen As New Collection
                listofVorlagen.Add(myName)
                If projectsFile = projectConfig("DateiName").ProjectsFile Then
                    listofArchivAllg = readProjectsAllg(listofVorlagen, projectConfig, outPutCollection, ptImportTypen.telairTagetikUpdate)
                End If
                'listofArchivAllg = readProjectsJIRA(listofVorlagen, JIRAProjectsConfig, outPutCollection)

                'If listofArchivAllg.Count > 0 Then
                '    Call moveFilesInArchiv(listofArchivAllg, importOrdnerNames(PTImpExp.projectWithConfig))
                'End If

                allOk = (listofArchivAllg.Count > 0 And outPutCollection.Count = 0)

                If allOk Then
                    ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                    Try
                        ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                        ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                        Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False, getSomeValuesFromOldProj:=False, calledFromActualDataImport:=False, calledFromRPA:=True)


                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            outputString = "Error at Import: " & vbLf & ex.Message
                        Else
                            outputString = "Fehler bei Import: " & vbLf & ex.Message
                        End If
                        outPutCollection.Add(outputString)

                    End Try

                Else

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

            End If

            allOk = allOk And telairUpdateConfigOK

        Catch ex2 As Exception
            allOk = False
        End Try

        Try
            ' store Projects
            If allOk Then
                allOk = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboUpdateTagetik.ToString, myName)

        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Cost Assertion Projects", ex1.Message)
        End Try



        processUpdateTagetik = allOk
    End Function

    ''' <summary>
    ''' does create retrospectively baselines for given project. 
    ''' All versions later than the desired date for baseline are deleted, 
    ''' then a baseline is created, the the versions are restored
    ''' and so are their key metric values to the just created baseline
    ''' </summary>
    ''' <param name="blattName"></param>
    ''' <returns></returns>
    Public Function processCreateBaselines(Optional blattName As String = "Baseline Creation")
        Dim atleastOneError As Boolean = False

        Dim zeile As Integer = 2

        Dim currentProjectName As String = ""
        Dim currentVariantName As String = ""
        Dim baselineDate As Date = Nothing


        Dim err As New clsErrorCodeMsg


        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattName = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattName),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    Dim lastRow As Integer = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    While zeile <= lastRow

                        Try

                            currentProjectName = CStr(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                            Try
                                If Not IsNothing(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                    currentVariantName = CStr(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                                Else
                                    currentVariantName = ""
                                End If
                            Catch ex As Exception
                                currentVariantName = ""
                            End Try


                            Try
                                If Not IsNothing(CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Value) Then
                                    baselineDate = CDate(CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Value).Date.AddHours(23).AddMinutes(59)
                                Else
                                    baselineDate = Date.MinValue
                                End If
                            Catch ex As Exception
                                baselineDate = Date.MinValue
                            End Try





                            ' check1: does project exist at all 
                            ' check2: does the timestamp exist at all

                            ' versionsToRestore as Date()
                            Dim versionsToRestore As New Collection
                            Dim heute As Date = Date.Now
                            Dim vpID As String = ""

                            Dim check1 As Boolean = CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(currentProjectName, currentVariantName, Date.Now, err)
                            ' only valid if there does not yet exist any baseline
                            Dim check2 As Boolean = Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(currentProjectName, "pfv", Date.Now, err)

                            If check1 And check2 Then
                                ' add check: if a baseline already exists: the baselineDate has to be later than the last baseline
                                Dim history As clsProjektHistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(currentProjectName, currentVariantName, StartofCalendar, Date.Now, err)



                                If history.Count > 0 Then

                                    ' in case Date.MinValue: look for timestamp later than 3 months after project start 
                                    If IsNothing(baselineDate) Then
                                        baselineDate = history.Last.startDate.AddMonths(3)
                                    End If

                                    If baselineDate = Date.MinValue Then
                                        baselineDate = history.Last.startDate.AddMonths(3)
                                    End If

                                    Dim useAsBaselineProject As clsProjekt = history.getProjectbefore(baselineDate)
                                    If IsNothing(useAsBaselineProject) Then
                                        useAsBaselineProject = history.getProjectAfter(baselineDate)
                                    End If

                                    If IsNothing(useAsBaselineProject) Then
                                        useAsBaselineProject = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(currentProjectName, currentVariantName, vpID, baselineDate.AddSeconds(10), err)
                                    End If

                                    Dim ok As Boolean
                                    If Not IsNothing(useAsBaselineProject) Then
                                        ' get all timestamps after the baselineDate, because they need to be deleted and the restored again ..
                                        Dim myfollowingprojects As SortedList(Of Date, clsProjekt) = history.getProjectsAfter(baselineDate)
                                        Dim myCopiedProjects As New SortedList(Of Date, clsProjekt)

                                        Dim noErrors As Boolean = True
                                        For Each ckvp As KeyValuePair(Of Date, clsProjekt) In myfollowingprojects
                                            Dim tmpProj As clsProjekt = ckvp.Value.createVariant("tmp", "")
                                            tmpProj.variantName = ""

                                            If ckvp.Value.isIdenticalTo(tmpProj) Then
                                                tmpProj.timeStamp = ckvp.Key
                                                myCopiedProjects.Add(ckvp.Key, tmpProj)
                                            Else
                                                noErrors = False
                                            End If
                                        Next

                                        If noErrors Then
                                            '
                                            ' now delete all vpv timestamps after baselineDate
                                            '
                                            For Each kvp As KeyValuePair(Of Date, clsProjekt) In myfollowingprojects
                                                If kvp.Key = kvp.Value.timeStamp Then
                                                    Try
                                                        If CType(databaseAcc, DBAccLayer.Request).deleteProjectTimestampFromDB(kvp.Value.name, kvp.Value.variantName, kvp.Key, dbUsername, err) Then
                                                            Call logger(ptErrLevel.logInfo, "Project Version deleted ", kvp.Value.name & " " & kvp.Value.variantName)
                                                        Else
                                                            Call logger(ptErrLevel.logError, "Project Version could not be deleted ", kvp.Value.name & " " & kvp.Value.variantName & " " & err.errorMsg)
                                                        End If
                                                    Catch ex As Exception
                                                        Call logger(ptErrLevel.logError, "Project Version could not be deleted ", kvp.Value.name & " " & kvp.Value.variantName & " " & ex.Message)
                                                    End Try

                                                Else
                                                    ' timestamps are not identical 
                                                End If

                                            Next

                                            '
                                            ' now write the baseline
                                            '
                                            Dim myBaselineProject As clsProjekt = useAsBaselineProject.createVariant("pfv", "Plan Version used: " & baselineDate.ToString)
                                            myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager
                                            myBaselineProject.timeStamp = baselineDate


                                            Dim mergedProj As clsProjekt = Nothing
                                            Try
                                                ok = CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(myBaselineProject, dbUsername, mergedProj, err)

                                                If ok Then
                                                    Call logger(ptErrLevel.logInfo, "Baseline stored ", myBaselineProject.name & " " & myBaselineProject.timeStamp.ToString)
                                                Else
                                                    Call logger(ptErrLevel.logError, "Baseline not stored ", myBaselineProject.name & " " & err.errorMsg)
                                                End If

                                                ' now delete the latest , auto-created plan-version with current time-stamp
                                                ' was created because a baseline was written
                                                vpID = ""
                                                Dim mylatestProj As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(currentProjectName, "", vpID, Date.Now, err)
                                                ok = CType(databaseAcc, DBAccLayer.Request).deleteProjectTimestampFromDB(currentProjectName, "", mylatestProj.timeStamp, dbUsername, err)
                                                If ok Then
                                                    Call logger(ptErrLevel.logInfo, "Version deleted ", currentProjectName & " " & mylatestProj.timeStamp.ToString)
                                                Else
                                                    Call logger(ptErrLevel.logError, "Version delete failed ", currentProjectName & " " & mylatestProj.timeStamp.ToString & " " & err.errorMsg)
                                                End If

                                                ' now create tmp Variant 
                                                Dim savemyLatestPRoj As clsProjekt = mylatestProj.createVariant("tmp", "")
                                                savemyLatestPRoj.variantName = ""
                                                savemyLatestPRoj.timeStamp = baselineDate.AddSeconds(5)

                                                myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung
                                                ok = CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(savemyLatestPRoj, dbUsername, mergedProj, err)
                                                If ok Then
                                                    Call logger(ptErrLevel.logInfo, "Version stored ", savemyLatestPRoj.name & " " & savemyLatestPRoj.timeStamp.ToString)
                                                Else
                                                    Call logger(ptErrLevel.logError, "Version not stored ", savemyLatestPRoj.name & " " & err.errorMsg)
                                                End If
                                            Catch ex As Exception
                                                Call logger(ptErrLevel.logError, "Baseline Version could not be stored ", myBaselineProject.name & " " & myBaselineProject.timeStamp.ToString & " " & ex.Message)
                                            End Try


                                            ' now write all the deleted vpv versions 
                                            myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung
                                            For Each kvp As KeyValuePair(Of Date, clsProjekt) In myCopiedProjects

                                                Try
                                                    If CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(kvp.Value, dbUsername, mergedProj, err) Then
                                                        Call logger(ptErrLevel.logInfo, "Project Version stored", kvp.Value.name & " " & kvp.Value.timeStamp.ToString)
                                                    Else
                                                        Call logger(ptErrLevel.logError, "Project Version NOT stored", kvp.Value.name & " " & kvp.Value.timeStamp.ToString & " " & err.errorMsg)
                                                    End If

                                                Catch ex As Exception
                                                    Call logger(ptErrLevel.logError, "Project Version could not be stored ", kvp.Value.name & " " & kvp.Value.timeStamp.ToString & " " & ex.Message)
                                                End Try


                                            Next

                                        Else
                                            Call logger(ptErrLevel.logError, "Project Copy did not produce identical versions, no actions taken", currentProjectName)
                                        End If

                                    Else
                                        Call logger(ptErrLevel.logError, "There is no version at or before  ", baselineDate.ToString)
                                    End If

                                Else
                                    ' now it seems to exist exactly one version, use that and create a baseline from It 
                                End If

                            Else
                                If Not check1 Then
                                    Call logger(ptErrLevel.logError, "Project does not exist:  ", currentProjectName)
                                End If

                                If Not check2 Then
                                    Call logger(ptErrLevel.logError, "Project Baseline already exists ", currentProjectName)
                                End If
                            End If



                        Catch ex As Exception
                            atleastOneError = True
                            Call logger(ptErrLevel.logError, "Exception in renaming, line ", zeile.ToString & ex.Message)
                        End Try

                        zeile = zeile + 1

                    End While


                End With
            End If

        Catch ex As Exception
            atleastOneError = True
            Throw New Exception("Fehler In Process Rename Projects" & ex.Message)
        End Try


        processCreateBaselines = Not atleastOneError
    End Function

    ''' <summary>
    ''' assigns attributes in batch
    ''' </summary>
    ''' <param name="blattName"></param>
    ''' <returns></returns>
    Public Function processAssignAttributes(Optional ByVal blattName As String = "Assign Attributes") As Boolean

        Dim atleastOneError As Boolean = False

        Dim zeile As Integer = 2

        Dim myProjectName As String = ""

        Dim mybusinessUnit As String = ""
        Dim myResponsible As String = ""
        Dim responsible As String = ""
        Dim strategyFit As Double = 0.0
        Dim riskKPI As Double = 0.0

        Dim err As New clsErrorCodeMsg


        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattName = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattName),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    Dim lastRow As Integer = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    While zeile <= lastRow

                        Try

                            myProjectName = CStr(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                            mybusinessUnit = CStr(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                            myResponsible = CStr(CType(.Cells(zeile, 3), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                            ' set myCustom User Role 
                            myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung
                            If myProjectName <> "" Then
                                ' check1: does current Project exist? 

                                ' check1: does oldName exist at all?

                                Dim vpID As String = ""
                                Dim myProject As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectfromDB(myProjectName, " ", vpID, Date.Now, err)
                                Dim check1 As Boolean = Not IsNothing(myProject)

                                Dim check2 As Boolean = False
                                Dim i As Integer = 1
                                While i <= businessUnitDefinitions.Count And Not check2

                                    If businessUnitDefinitions.ElementAt(i - 1).Value.name = mybusinessUnit Then
                                        check2 = True
                                    Else
                                        i = i + 1
                                    End If

                                End While


                                If check1 And check2 Then

                                    Dim storeRequired As Boolean = (myProject.businessUnit <> mybusinessUnit) Or ((myProject.leadPerson <> myResponsible) And (myResponsible <> ""))
                                    Dim outputCollection As New Collection
                                    msgTxt = ""
                                    If storeRequired Then
                                        If myProject.businessUnit <> mybusinessUnit And mybusinessUnit <> "" Then
                                            myProject.businessUnit = mybusinessUnit
                                            msgTxt = mybusinessUnit & " "
                                        End If

                                        If myProject.leadPerson <> myResponsible And myResponsible <> "" Then
                                            myProject.leadPerson = myResponsible
                                            msgTxt = msgTxt & myResponsible
                                        End If

                                        Try
                                            Dim mergedProj As clsProjekt = Nothing
                                            'If storeSingleProjectToDB(myProject, outputCollection) Then
                                            myProject.timeStamp = Date.Now
                                            If CType(databaseAcc, DBAccLayer.Request).storeProjectToDB(myProject, dbUsername, mergedProj, err, True) Then
                                                msgTxt = msgTxt & " was assigned to " & myProject.name
                                                Call logger(ptErrLevel.logInfo, "project stored ", msgTxt)
                                            Else
                                                Call logger(ptErrLevel.logError, "project store with new business Unit failed: " & msgTxt, outputCollection)

                                            End If
                                        Catch ex As Exception
                                            Call logger(ptErrLevel.logError, "project store with new business Unit failed: " & msgTxt, outputCollection)
                                        End Try

                                    Else
                                        Call logger(ptErrLevel.logInfo, "project has already  ", mybusinessUnit)
                                    End If

                                Else
                                    If Not check1 Then
                                        ' Logging
                                        atleastOneError = True
                                        Call logger(ptErrLevel.logError, "Project does not exist: ", myProjectName)
                                    End If
                                    If Not check2 Then
                                        ' Logging
                                        atleastOneError = True
                                        Call logger(ptErrLevel.logError, "Business Unit not known: ", myProjectName & " " & mybusinessUnit)
                                    End If
                                End If
                            Else
                                ' logging: no valid rename Parameters
                                atleastOneError = True
                                Call logger(ptErrLevel.logError, "No project name Given in Zeile : ", zeile.ToString)
                            End If

                        Catch ex As Exception
                            atleastOneError = True
                            Call logger(ptErrLevel.logError, "Exception in renaming, line ", zeile.ToString & ex.Message)
                        End Try

                        zeile = zeile + 1

                    End While


                End With
            End If

        Catch ex As Exception
            atleastOneError = True
            Throw New Exception("Fehler In Process Rename Projects" & ex.Message)
        End Try

        processAssignAttributes = Not atleastOneError
    End Function


    ''' <summary>
    ''' does rename projects in batch. 
    ''' in the table there need to be oldname , newName
    ''' oldname need to exist, newname must not exist
    ''' </summary>
    ''' <param name="blattName"></param>
    ''' <returns></returns>
    Public Function processRenameProjects(Optional ByVal blattName As String = "Rename Projects") As Boolean

        Dim atleastOneError As Boolean = False

        Dim zeile As Integer = 2

        Dim currentProjectName As String = ""
        Dim newProjectName As String = ""
        Dim err As New clsErrorCodeMsg


        Try
            Dim activeWSListe As xlns.Worksheet = Nothing
            If blattName = "" Then
                activeWSListe = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            Else
                activeWSListe = CType(appInstance.ActiveWorkbook.Worksheets.Item(blattName),
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)
            End If

            If Not IsNothing(activeWSListe) Then

                With activeWSListe

                    Dim lastRow As Integer = CType(.Cells(2000, 1), Global.Microsoft.Office.Interop.Excel.Range).End(xlns.XlDirection.xlUp).Row

                    While zeile <= lastRow

                        Try

                            currentProjectName = CStr(CType(.Cells(zeile, 1), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim
                            newProjectName = CStr(CType(.Cells(zeile, 2), Global.Microsoft.Office.Interop.Excel.Range).Value).Trim

                            If currentProjectName <> "" And newProjectName <> "" And currentProjectName <> newProjectName Then
                                ' check1: does current Project exist? 

                                ' check1: does oldName exist at all?
                                Dim check1 As Boolean = CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(currentProjectName, "", Date.Now, err)
                                ' check2: is newName not yet existent? 
                                Dim check2 As Boolean = Not CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(newProjectName, "", Date.Now, err)

                                If check1 And check2 Then

                                    If CType(databaseAcc, DBAccLayer.Request).renameProjectsInDB(currentProjectName, newProjectName, dbUsername, err) Then
                                        Call logger(ptErrLevel.logInfo, "Rename success: ", currentProjectName & " -> " & newProjectName)
                                    Else
                                        atleastOneError = True
                                        Call logger(ptErrLevel.logError, "Rename Failed: " & currentProjectName & " -> " & newProjectName, err.errorMsg)
                                    End If

                                Else
                                    If Not check1 Then
                                        ' Logging
                                        atleastOneError = True
                                        Call logger(ptErrLevel.logError, "Project to rename does not exist: ", currentProjectName)
                                    End If
                                    If Not check2 Then
                                        ' Logging
                                        atleastOneError = True
                                        Call logger(ptErrLevel.logError, "Project with new name does already exist: ", newProjectName)
                                    End If
                                End If
                            Else
                                ' logging: no valid rename Parameters
                                atleastOneError = True
                                Call logger(ptErrLevel.logError, "no valid renaming parameters : ", currentProjectName & " -> " & newProjectName)
                            End If

                        Catch ex As Exception
                            atleastOneError = True
                            Call logger(ptErrLevel.logError, "Exception in renaming, line ", zeile.ToString & ex.Message)
                        End Try

                        zeile = zeile + 1

                    End While


                End With
            End If

        Catch ex As Exception
            atleastOneError = True
            Throw New Exception("Fehler In Process Rename Projects" & ex.Message)
        End Try

        processRenameProjects = Not atleastOneError

    End Function

    Public Function processCostAssertion(ByVal myName As String, ByVal portfolioName As String, ByVal dirName As String, ByVal importDate As Date) As Boolean
        Dim allOk As Boolean = True
        Dim aktDateTime As Date = Date.Now
        Dim telairCostAssertionImportConfigOK As Boolean = False


        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboCostAssertion.ToString, myName)

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
            Dim projectCostAssertConfig As New SortedList(Of String, clsConfigProjectsImport)
            Dim projectsFile As String = ""
            Dim lastrow As Integer = 0
            Dim outputString As String = ""
            Dim dateiName As String = ""
            Dim listofArchivAllg As New List(Of String)
            Dim outPutCollection As New Collection
            Dim configCostAssertionImport As String = "configCostAssertionImport.xlsx"


            Dim outputLine As String = ""

            Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

            ' Konfigurationsdatei lesen und Validierung durchführen
            configfilesOrdner = My.Computer.FileSystem.CombinePath(awinPath, configfilesOrdner)
            configCostAssertionImport = configfilesOrdner & "\" & configCostAssertionImport
            ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
            telairCostAssertionImportConfigOK = checkProjectImportConfig(configCostAssertionImport, projectsFile, projectCostAssertConfig, lastrow, outPutCollection)

            If telairCostAssertionImportConfigOK Then

                Dim listofVorlagen As New Collection
                listofVorlagen.Add(myName)
                If projectsFile = projectCostAssertConfig("DateiName").ProjectsFile Then
                    listofArchivAllg = readProjectsAllg(listofVorlagen, projectCostAssertConfig, outPutCollection, ptImportTypen.telairCostAssertionImport)
                End If
                'listofArchivAllg = readProjectsJIRA(listofVorlagen, JIRAProjectsConfig, outPutCollection)

                'If listofArchivAllg.Count > 0 Then
                '    Call moveFilesInArchiv(listofArchivAllg, importOrdnerNames(PTImpExp.projectWithConfig))
                'End If

                allOk = (listofArchivAllg.Count > 0 And outPutCollection.Count = 0)

                If allOk Then
                    ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                    Try
                        ' es muss der Parameter FileFrom3RdParty auf False gesetzt sein
                        ' dieser Parameter bewirkt, dass die alten Ressourcen-Zuordnungen aus der Datenbank übernommen werden, wenn das eingelesene File eine Ressourcen Summe von 0 hat. 
                        Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=True, fileFrom3rdParty:=False, getSomeValuesFromOldProj:=False, calledFromActualDataImport:=False, calledFromRPA:=True)


                    Catch ex As Exception
                        If awinSettings.englishLanguage Then
                            outputString = "Error at Import: " & vbLf & ex.Message
                        Else
                            outputString = "Fehler bei Import: " & vbLf & ex.Message
                        End If
                        outPutCollection.Add(outputString)

                    End Try

                Else

                    Call logger(ptErrLevel.logError, "processCostAssertion", outPutCollection)

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

            End If


            allOk = allOk And telairCostAssertionImportConfigOK

        Catch ex2 As Exception
            allOk = False
        End Try

        Try
            ' store Projects
            If allOk Then
                allOk = storeImportProjekte()
            End If

            ' empty session 
            Call emptyRPASession()

            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboCostAssertion.ToString, myName)

        Catch ex1 As Exception
            allOk = False
            Call logger(ptErrLevel.logError, "RPA Error Importing Cost Assertion Projects", ex1.Message)
        End Try



        processCostAssertion = allOk
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

            If noSucessEmails Then

                ' do nothing
            Else

                errMsgCode = New clsErrorCodeMsg
                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & myName & ": successful ..." & vbCrLf _
                                                                                & mailMessage, errMsgCode)
            End If


        Else
            Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
            If My.Computer.FileSystem.FileExists(fullfileName) Then
                My.Computer.FileSystem.MoveFile(fullfileName, newDestination, True)
                Call logger(ptErrLevel.logError, "failed: ", fullfileName)

                'Dim errMessages As Collection = readlogger(ptErrLevel.logError)

                Dim logfileName As String = My.Computer.FileSystem.GetName(logfileNamePath)
                Dim newLog As String = My.Computer.FileSystem.CombinePath(failureFolder, logfileName)
                My.Computer.FileSystem.MoveFile(logfileNamePath, newLog, True)

                errMsgCode = New clsErrorCodeMsg
                'result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                '                                                            & myName & ": with errors ..." & vbCrLf _
                '                                                            & "Look for more details in the Failure-Folder: " & failureFolder, errMsgCode)
                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                                                                            & myName & ": with errors ..." & vbCrLf _
                                                                            & mailMessage, errMsgCode)
            End If
        End If

        Try
            Dim ok As Boolean = cancelLocksMyProjects(dbUsername)
        Catch ex As Exception
            ' evt. keine locks vorhanden
        End Try

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

    Public Function cancelLocksMyProjects(ByVal user As String) As Boolean
        Dim err As New clsErrorCodeMsg
        Dim msgTxt As String = ""
        result = False

        ' all locks of my projects will be deleted
        If CType(databaseAcc, DBAccLayer.Request).cancelWriteProtections(user, err, False) Then
            If awinSettings.englishLanguage Then
                msgTxt = "Your temporary write locks have been lifted"
            Else
                msgTxt = "Ihre vorübergehenden Schreibsperren wurden aufgehoben"
            End If
            If awinSettings.visboDebug Then
                Call MsgBox(msgTxt)
            End If
            Call logger(ptErrLevel.logInfo, "cancelLocksMyProjects", msgTxt)
            result = True
        Else
            If awinSettings.englishLanguage Then
                msgTxt = "Your temporary write locks could not be lifted"
            Else
                msgTxt = "Ihre vorübergehenden Schreibsperren konnten nicht aufgehoben werden"
            End If
            If awinSettings.visboDebug Then
                Call MsgBox(msgTxt)
            End If
            Call logger(ptErrLevel.logInfo, "cancelLocksMyProjects", msgTxt)
            result = False
        End If

        cancelLocksMyProjects = result
    End Function

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

                If noSucessEmails Then

                    ' do nothing
                Else

                    errMsgCode = New clsErrorCodeMsg
                    result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & myName & ": successful ...", errMsgCode)
                End If

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
                Select Case LCase(fileExt)
                    Case ".xlsx", ".xlsm"

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

                    Case ".xlsm"

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
