Imports xlns = Microsoft.Office.Interop.Excel
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Newtonsoft.Json
Imports System.IO
Imports DBAccLayer
Imports WebServerAcc
Imports System.Security.Principal


Module rpaModule1


    Public Sub Main()
        ' reads the VISBO RPA folder und treats each file it finds there appropriately
        ' in most cases new project and portfolio versions will be written 
        ' suggestions for Team Members will follow 
        ' automation in resource And team allocation will follow
        Dim msgTxt As String = ""

        Dim anzFiles As Integer = 0

        Dim rpaPath As String = My.Settings.rpaPath
        Dim swPath As String = My.Settings.swPath

        Dim rpaFolder As String = My.Computer.FileSystem.CombinePath(rpaPath, "RPA")
        Dim successFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "success")
        Dim failureFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "failure")
        Dim logfileFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "logfiles")
        Dim unknownFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "unknown")
        Dim settingsFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "settings")
        Dim settingJsonFile As String = My.Computer.FileSystem.CombinePath(settingsFolder, "rpa_setting.json")


        Dim myActivePortfolio As String = ""

        Dim listToProcess As New SortedList(Of String, Integer)

        Try
            If My.Computer.FileSystem.FileExists(settingJsonFile) Then
                Dim jsonSetting As String = File.ReadAllText(settingJsonFile)
                Dim inputvalues As clsRPASetting = JsonConvert.DeserializeObject(Of clsRPASetting)(jsonSetting)

                ' is there a activePortfolio
                myActivePortfolio = inputvalues.activePortfolio

                ' now check whether or not the folder are existings , if not create them 
                If Not My.Computer.FileSystem.DirectoryExists(successFolder) Then
                    My.Computer.FileSystem.CreateDirectory(successFolder)
                End If

                If Not My.Computer.FileSystem.DirectoryExists(failureFolder) Then
                    My.Computer.FileSystem.CreateDirectory(failureFolder)
                End If

                If Not My.Computer.FileSystem.DirectoryExists(logfileFolder) Then
                    My.Computer.FileSystem.CreateDirectory(logfileFolder)
                End If

                If Not My.Computer.FileSystem.DirectoryExists(unknownFolder) Then
                    My.Computer.FileSystem.CreateDirectory(unknownFolder)
                End If

                ' read all files, categorize and verify them  
                msgTxt = "Starting ..."
                Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)


                ' 
                ' startUpRPA setzt awinSettings, liest orga, appearances und andere Settings - analog awinSetTypen , allerdings nie mit Versuch, etwas von Platte zu lesen ... 
                If startUpRPA(inputvalues.VisboCenter, inputvalues.VisboUrl, swPath) Then

                    ' read all Excel based files 
                    Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(rpaFolder, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsx")


                    For Each fullFileName As String In listOfImportfiles

                        Dim myName As String = My.Computer.FileSystem.GetName(fullFileName)
                        Dim rpaCategory As PTRpa = bestimmeRPACategory(fullFileName)

                        If rpaCategory = PTRpa.visboUnknown Then
                            ' move file to unknown Folder ... 
                            Dim newDestination As String = My.Computer.FileSystem.CombinePath(unknownFolder, myName)
                            My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                            Call logger(ptErrLevel.logInfo, "unknown file / category: ", myName)
                        Else

                            If Not listToProcess.ContainsKey(myName) Then
                                listToProcess.Add(fullFileName, CInt(rpaCategory))
                            End If
                        End If



                    Next

                    ' read all Microsoft Project Files 
                    listOfImportfiles = My.Computer.FileSystem.GetFiles(rpaFolder, FileIO.SearchOption.SearchTopLevelOnly, "*.mpp")
                    For Each fullFileName As String In listOfImportfiles

                        Dim myName As String = My.Computer.FileSystem.GetName(fullFileName)
                        Dim rpaCategory As PTRpa = PTRpa.visboMPP

                        If Not listToProcess.ContainsKey(myName) Then
                            listToProcess.Add(fullFileName, CInt(rpaCategory))
                        End If

                    Next


                    ImportProjekte.Clear()
                    Dim importOrganisations As New clsOrganisations
                    Dim importCustomization As New clsCustomization
                    Dim importAppearances As New clsAppearances
                    Dim importDate As Date = Date.Now()


                    For Each kvp As KeyValuePair(Of String, Integer) In listToProcess

                        Dim myName As String = My.Computer.FileSystem.GetName(kvp.Key)
                        Dim currentWB As xlns.Workbook = Nothing
                        Dim allOk As Boolean = False

                        Try

                            If Not kvp.Value = PTRpa.visboMPP Then
                                currentWB = appInstance.Workbooks.Open(kvp.Key)
                            End If


                            Select Case kvp.Value
                                Case CInt(PTRpa.visboProjectList)

                                    Dim portfolioName As String = myName.Substring(0, myName.IndexOf(".xls"))

                                    Dim overloadAllowedinMonths As Double = 1.0
                                    Dim overloadAllowedTotal As Double = 1.03


                                    Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboProjectList.ToString, myName)
                                    Dim readProjects As Integer = 0
                                    Dim createdProjects As Integer = 0
                                    Dim importedProjects As Integer = ImportProjekte.Count
                                    allOk = awinImportProjektInventur(readProjects, createdProjects)
                                    If allOk Then
                                        Call logger(ptErrLevel.logInfo, "Project List imported: " & myName, readProjects & " read; " & createdProjects & " created")
                                        allOk = storeImportProjekte()
                                    Else
                                        Call logger(ptErrLevel.logError, "failure in Processing: " & myName, PTRpa.visboProjectList.ToString)
                                    End If

                                    If allOk Then


                                        ' Get the Ranking out of Excel List , it is just the ordering of the rows 
                                        ' value holds the AllProjekte.Key, i.e name#variantName
                                        Dim rankingList As SortedList(Of Integer, String) = getRanking()

                                        ' if Portfolio with active Projects is given and exists:  
                                        ' then we probably do have a brownfield
                                        If myActivePortfolio <> "" Then
                                            ' load portfolio projects 
                                        Else
                                            myActivePortfolio = "active projects"
                                        End If

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

                                        showRangeLeft = ShowProjekte.getMinMonthColumn
                                        showRangeRight = ShowProjekte.getMaxMonthColumn


                                        Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=portfolioName, vName:="")

                                        Dim errMsg As New clsErrorCodeMsg
                                        Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)

                                        Dim outputCollection As New Collection
                                        Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, dbPortfolioNames)


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

                                        Dim skillIDs As Collection = ShowProjekte.getRoleSkillIDs()

                                        If skillIDs.Count > 0 Then
                                            For Each tmpStrID As String In skillIDs
                                                If Not skillList.Contains(tmpStrID) Then
                                                    skillList.Add(tmpStrID)
                                                End If
                                            Next
                                        End If

                                        ' then empty ShowProjekte again 
                                        ShowProjekte.Clear()

                                        ' now check for each project , whether there are 
                                        For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList

                                            Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                                            If Not ShowProjekte.contains(hproj.name) Then
                                                ShowProjekte.Add(hproj)
                                            End If

                                            Dim overutilizationFound As Boolean = ShowProjekte.overLoadFound(aggregationList, False, overloadAllowedinMonths, overloadAllowedTotal)

                                            If Not overutilizationFound Then
                                                overutilizationFound = ShowProjekte.overLoadFound(skillList, False, overloadAllowedinMonths, overloadAllowedTotal)
                                            End If

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

                                                        overutilizationFound = ShowProjekte.overLoadFound(aggregationList, False, overloadAllowedinMonths, overloadAllowedTotal)

                                                        If Not overutilizationFound Then
                                                            overutilizationFound = ShowProjekte.overLoadFound(skillList, False, overloadAllowedinMonths, overloadAllowedTotal)
                                                        End If

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
                                                        Console.WriteLine(infomsg)
                                                    Else
                                                        ' take it out again , because there was no solution
                                                        ShowProjekte.Remove(hproj.name)
                                                        Dim infomsg As String = "... failed to create variant to avoid bottlenecks " & hproj.getShapeText
                                                        Call logger(ptErrLevel.logError, infomsg, myMessages)
                                                        Console.WriteLine(infomsg)
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

                                    Else
                                        ' no additional logger necessary - is done in storeImportProjekte
                                    End If


                                    ' now empty the complete session  
                                    Call emptyRPASession()
                                    Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboProjectList.ToString, myName)


                                Case CInt(PTRpa.visboMPP)

                                    Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboMPP.ToString, myName)

                                    Try

                                        Dim hproj As clsProjekt = New clsProjekt

                                        ' Definition für ein eventuelles Mapping
                                        Dim mapProj As clsProjekt = Nothing
                                        Call awinImportMSProject("RPA", kvp.Key, hproj, mapProj, importDate)

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

                                Case CInt(PTRpa.visboProject)

                                    Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboProject.ToString, myName)

                                    'read Project Brief and put it into ImportProjekte
                                    Try
                                        Dim hproj As clsProjekt = Nothing

                                        ' read the file and import into hproj
                                        Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)

                                        allOk = Not IsNothing(hproj)
                                        If allOk Then
                                            Try
                                                Dim keyStr As String = calcProjektKey(hproj)
                                                ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                                                'AlleProjekte.Add(hproj, updateCurrentConstellation:=False)

                                                Call importProjekteEintragen(importDate, drawPlanTafel:=False, fileFrom3rdParty:=True, getSomeValuesFromOldProj:=True, calledFromActualDataImport:=False, calledFromRPA:=True)
                                            Catch ex2 As Exception
                                                allOk = False
                                                Call logger(ptErrLevel.logError, "RPA Error Importing MS Project file " & PTRpa.visboProject.ToString, ex2.Message)
                                            End Try
                                        Else
                                            Call logger(ptErrLevel.logError, "RPA Error Importing MS Project file " & PTRpa.visboProject.ToString, myName)
                                        End If

                                        ' store Project 
                                        If allOk Then
                                            allOk = storeImportProjekte()
                                        End If

                                        ' empty session 
                                        Call emptyRPASession()

                                        Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboProject.ToString, myName)

                                    Catch ex1 As Exception
                                        allOk = False
                                        Call logger(ptErrLevel.logError, "RPA Error Importing MS Project file ", ex1.Message)
                                    End Try

                                Case CInt(PTRpa.visboJira)
                                    allOk = True

                                    Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboJira.ToString, myName)

                                    'read File with Jira-Projects and put it into ImportProjekte
                                    Try

                                        'Dim hproj As clsProjekt = Nothing

                                        '' read the file and import into hproj
                                        'Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)
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

                                            Dim listofVorlagen As New Collection
                                            listofVorlagen.Add(kvp.Key)
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

                                Case CInt(PTRpa.visboDefaultCapacity)
                                    allOk = True

                                Case CInt(PTRpa.visboInitialOrga)

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
                                            result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(importedOrga,
                                                                                    CStr(settingTypes(ptSettingTypes.organisation)),
                                                                                    orgaName,
                                                                                    importedOrga.validFrom,
                                                                                    err)

                                            If result = True Then
                                                allOk = True
                                                msgTxt = "ok, Organisation, valid from " & importedOrga.validFrom.ToShortDateString & " stored ..."
                                                Console.WriteLine(msgTxt)
                                                Call logger(ptErrLevel.logInfo, PTRpa.visboInitialOrga.ToString, msgTxt)
                                            Else
                                                allOk = False
                                                msgTxt = "Storing organisaiton failed "
                                                Call logger(ptErrLevel.logError, PTRpa.visboInitialOrga.ToString, msgTxt)
                                            End If
                                        End If

                                        Call logger(ptErrLevel.logInfo, "endProcessing: " & PTRpa.visboInitialOrga.ToString, myName)
                                    Catch ex As Exception
                                        allOk = False
                                    End Try

                                Case CInt(PTRpa.visboModifierCapacities)
                                    allOk = True
                                Case CInt(PTRpa.visboExternalContracts)
                                    allOk = True
                                Case CInt(PTRpa.visboActualData1)
                                    allOk = True
                                Case Else

                            End Select

                            If Not (kvp.Value = PTRpa.visboMPP Or kvp.Value = PTRpa.visboJira) Then

                                If allOk Then
                                    CType(currentWB.Worksheets(1), xlns.Worksheet).Cells(1, 1).interior.color = visboFarbeGreen
                                Else
                                    CType(currentWB.Worksheets(1), xlns.Worksheet).Cells(1, 1).interior.color = visboFarbeRed
                                End If

                                currentWB.Close(SaveChanges:=True)

                            End If

                            If allOk Then
                                Dim newDestination As String = My.Computer.FileSystem.CombinePath(successFolder, myName)
                                My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)

                                Call logger(ptErrLevel.logInfo, "success: ", myName)
                            Else
                                Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
                                My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                                Call logger(ptErrLevel.logError, "failed: ", myName)
                            End If

                        Catch ex As Exception
                            Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
                            My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                            Call logger(ptErrLevel.logError, "failed: ", ex.Message)
                            If Not kvp.Value = PTRpa.visboMPP Then
                                currentWB.Close(SaveChanges:=True)
                            End If
                        End Try

                        If allOk Then
                            Console.WriteLine(myName & ": successful ...")
                        Else
                            Console.WriteLine(myName & ": failed ...")
                        End If

                    Next

                    Console.WriteLine("end of jobs!")
                    msgTxt = "End of RPA ..."
                    Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)
                Else
                    msgTxt = "wrong settings - exited without performing jobs ...."
                    Call MsgBox(msgTxt)
                    Console.WriteLine(msgTxt)
                    Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)
                End If

                ' now store User Login Data
                My.Settings.userNamePWD = awinSettings.userNamePWD

                ' speichern 
                My.Settings.Save()

            Else
                ' Exit ! 
                ' read all files, categorize and verify them  
                msgTxt = "Exit - there is no File " & settingJsonFile
                Call logger(ptErrLevel.logError, "VISBO Robotic Process automation", msgTxt)
                Console.WriteLine(msgTxt)
            End If

            ' now release all writeProtections ...
            Dim errMsgCode As New clsErrorCodeMsg
            If CType(databaseAcc, DBAccLayer.Request).cancelWriteProtections(dbUsername, errMsgCode) Then
                If awinSettings.visboDebug Then
                    Call MsgBox("Ihre vorübergehenden Schreibsperren wurden aufgehoben")
                End If
            Else
                msgTxt = "Write Protections could not be released ! Please do so in Web-UI ..."
                Call logger(ptErrLevel.logError, "VISBO Robotic Process automation End", msgTxt)
                Console.WriteLine(msgTxt)
            End If

        Catch ex As Exception
            msgTxt = "Exit - Failure in rpa Main: " & ex.Message
            Call logger(ptErrLevel.logError, "VISBO Robotic Process automation", msgTxt)
            Console.WriteLine(msgTxt)
        End Try




    End Sub

    Private Sub emptyRPASession()
        ImportProjekte.Clear()
        ShowProjekte.Clear(False)
        AlleProjekte.Clear(False)
        AlleProjektSummaries.Clear(False)
    End Sub



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
                        Console.WriteLine("project stored: " & kvp.Value.getShapeText)
                    Else
                        ok = ok And False
                        Call logger(ptErrLevel.logError, "project store failed: ", outputCollection)
                        Console.WriteLine("!! ... project store failed: " & kvp.Value.getShapeText)
                    End If

                Else
                    myCustomUserRole.customUserRole = ptCustomUserRoles.ProjektLeitung

                    If storeSingleProjectToDB(kvp.Value, outputCollection) Then
                        ok = ok And True
                        Call logger(ptErrLevel.logInfo, "project updated: ", kvp.Value.getShapeText)
                        Console.WriteLine("project updated: " & kvp.Value.getShapeText)
                    Else
                        ok = ok And False
                        Call logger(ptErrLevel.logInfo, "project update failed: ", outputCollection)
                        Console.WriteLine("!! ... project update failed: " & kvp.Value.getShapeText)
                    End If

                End If

            Next

        Catch ex As Exception
            ok = False
            Call logger(ptErrLevel.logError, "Store Projects from List failed", ex.Message)
            Console.WriteLine("!!!! Store Projects from List failed" & ex.Message)
        End Try

        storeImportProjekte = ok
    End Function
    Private Function startUpRPA(ByVal mongoName As String, ByVal url As String, ByVal path As String) As Boolean

        Dim result As Boolean = False

        ' ggf hier noch die appInstance setzen ... 
        appInstance = New xlns.Application

        Try

            If readawinSettings(path) Then

                result = True
                ' independent of what is given in projectboardConfig.xml
                awinSettings.databaseName = mongoName
                awinSettings.databaseURL = url
                ' gespeichertes (verschlüsselt) Username und Pwd aus den Settings holen 
                awinSettings.rememberUserPwd = True
                awinSettings.userNamePWD = My.Settings.userNamePWD

                ' returns false if anything goes wrong .. 
                result = rpaSetTypen()

            End If


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
            Call logger(ptErrLevel.logInfo, "startUpRPA", "localPath:" & awinPath)
            Call logger(ptErrLevel.logInfo, "startUpRPA", "GlobalPath:" & globalPath)


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


            ' Erzeugen des Report Ordners, wenn er nicht schon existiert ..

            reportOrdnerName = awinPath & "Reports\"
            Try
                My.Computer.FileSystem.CreateDirectory(reportOrdnerName)
            Catch ex As Exception

            End Try

            ' ------------------
            ' tk 10.10.21
            ' normally not necessary

            'importOrdnerNames(PTImpExp.visbo) = awinPath & "Import\VISBO Steckbriefe"
            'importOrdnerNames(PTImpExp.rplan) = awinPath & "Import\RPLAN-Excel"
            'importOrdnerNames(PTImpExp.msproject) = awinPath & "Import\MSProject"
            'importOrdnerNames(PTImpExp.batchlists) = awinPath & "Import\Batch Projektlisten"
            'importOrdnerNames(PTImpExp.modulScen) = awinPath & "Import\Modulare Szenarien"
            'importOrdnerNames(PTImpExp.addElements) = awinPath & "Import\AddOn Regeln"
            'importOrdnerNames(PTImpExp.rplanrxf) = awinPath & "Import\RXF Files"
            'importOrdnerNames(PTImpExp.massenEdit) = awinPath & "Import\MassEdit"
            'importOrdnerNames(PTImpExp.offlineData) = awinPath & "Import\OfflineData"
            'importOrdnerNames(PTImpExp.scenariodefs) = awinPath & "Import\Scenario Definitions"
            'importOrdnerNames(PTImpExp.Orga) = awinPath & "Import\Organisation"
            'importOrdnerNames(PTImpExp.customUserRoles) = awinPath & "Import\CustomUserRoles"
            'importOrdnerNames(PTImpExp.actualData) = awinPath & "Import\ActualData"
            'importOrdnerNames(PTImpExp.Kapas) = awinPath & "Import\Capacities"
            'importOrdnerNames(PTImpExp.projectWithConfig) = awinPath & "Import\Projects With Config"
            'importOrdnerNames(PTImpExp.rpa) = awinPath & "Import\RPA"

            'exportOrdnerNames(PTImpExp.visbo) = awinPath & "Export\VISBO Steckbriefe"
            'exportOrdnerNames(PTImpExp.rplan) = awinPath & "Export\RPLAN-Excel"
            'exportOrdnerNames(PTImpExp.msproject) = awinPath & "Export\MSProject"
            'exportOrdnerNames(PTImpExp.batchlists) = awinPath & "Export\Scenario Definitions"
            'exportOrdnerNames(PTImpExp.modulScen) = awinPath & "Export\Modulare Szenarien"
            'exportOrdnerNames(PTImpExp.massenEdit) = awinPath & "Export\MassEdit"
            'exportOrdnerNames(PTImpExp.scenariodefs) = awinPath & "Export\Scenario Definitions"

            '' jetzt werden die Directories alle angelegt, sofern Sie nicht schon existieren ... 
            'For di As Integer = 0 To importOrdnerNames.Length - 1
            '    Try

            '        If Not IsNothing(importOrdnerNames(di)) Then
            '            My.Computer.FileSystem.CreateDirectory(importOrdnerNames(di))
            '        Else
            '            importOrdnerNames(di) = "-"
            '        End If

            '    Catch ex As Exception

            '    End Try
            'Next

            'For di As Integer = 0 To exportOrdnerNames.Length - 1
            '    Try
            '        If Not IsNothing(exportOrdnerNames(di)) Then
            '            My.Computer.FileSystem.CreateDirectory(exportOrdnerNames(di))
            '        Else
            '            exportOrdnerNames(di) = "-"
            '        End If

            '    Catch ex As Exception

            '    End Try
            'Next
            ' end changes tl 10.10.21
            ' --------------------------------------------------------

            StartofCalendar = StartofCalendar.Date

            DiagrammTypen(0) = "Phase"
            DiagrammTypen(1) = "Rolle"
            DiagrammTypen(2) = "Kostenart"
            DiagrammTypen(3) = "Portfolio"
            DiagrammTypen(4) = "Ergebnis"
            DiagrammTypen(5) = "Meilenstein"
            DiagrammTypen(6) = "Meilenstein Trendanalyse"
            DiagrammTypen(7) = "Phasen-Kategorie"
            DiagrammTypen(8) = "Meilenstein-Kategorie"
            DiagrammTypen(9) = "Cash-Flow"


            Try
                repMessages = XMLImportReportMsg(repMsgFileName, awinSettings.ReportLanguage)
                Call setLanguageMessages()
            Catch ex As Exception

            End Try

            autoSzenarioNamen(0) = "before Optimization"
            autoSzenarioNamen(1) = "1. Optimum"
            autoSzenarioNamen(2) = "2. Optimum"
            autoSzenarioNamen(3) = "3. Optimum"

            '
            ' die Namen der Worksheets Ressourcen und Portfolio verfügbar machen
            ' die Zahlen müssen korrespondieren mit der globalen Enumeration ptTables 
            arrWsNames(1) = "repCharts" ' Tabellenblatt zur Aufnahme der Charts für Reports 
            arrWsNames(2) = "Vorlage" ' depr
            ' arrWsNames(3) = 
            arrWsNames(ptTables.MPT) = "MPT"                          ' Multiprojekt-Tafel 
            arrWsNames(4) = "Einstellungen"                ' in Customization File 
            ' arrWsNames(5) = 
            arrWsNames(ptTables.meRC) = "meRC"                          ' Edit Ressourcen
            arrWsNames(6) = "meTE"                          ' Edit Termine
            arrWsNames(7) = "Darstellungsklassen"           ' wird in awinsettypen hinter MPT kopiert; nimmt für die Laufzeit die Darstellungsklassen auf 
            arrWsNames(8) = "Phasen-Mappings"               ' in Customization
            arrWsNames(9) = "meAT"                          ' Edit Attribute 
            arrWsNames(10) = "Meilenstein-Mappings"         ' in Customization
            ' arrWsNames(11) = 
            arrWsNames(ptTables.meCharts) = "meCharts"                     ' Massen-Edit Charts 
            arrWsNames(ptTables.mptPfCharts) = "mptPfCharts"                     ' vorbereitet: Portfolio Charts 
            arrWsNames(ptTables.mptPrCharts) = "mptPrCharts"                     ' vorbereitet: Projekt Charts 
            arrWsNames(14) = "Objekte" ' depr
            arrWsNames(15) = "missing Definitions"          ' in Customization File 


            awinSettings.applyFilter = False

            showRangeLeft = 0
            showRangeRight = 0

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
                Try
                    loginErfolgreich = CType(databaseAcc, DBAccLayer.Request).login(awinSettings.databaseURL, awinSettings.databaseName, awinSettings.VCid, dbUsername, dbPasswort, err)
                Catch ex As Exception
                    loginErfolgreich = False
                End Try


                If Not loginErfolgreich Then
                    loginErfolgreich = logInToMongoDB(True)
                End If

            Else
                loginErfolgreich = logInToMongoDB(True)
            End If


            ' das folgende darf nur gemacht werden, wenn auch awinsetting.visboserver gilt ... 


            If loginErfolgreich Then

                ' jetzt muss geprüft werden, ob es mehr als ein zugelassenes VISBO Center gibt , ist dann der Fall wenn es ein # im awinsettings.databaseNAme gibt 
                Dim listOfVCs As List(Of String) = CType(databaseAcc, DBAccLayer.Request).retrieveVCsForUser(err)

                If listOfVCs.Count = 1 Then
                    ' alles ok, nimm dieses  VC
                    If awinSettings.databaseName <> "" Then
                        If awinSettings.databaseName <> listOfVCs.Item(0).ToUpper Then
                            Throw New ArgumentException("No access to this VISBO Center " & awinSettings.databaseName)
                        Else
                            ' make sure it is exactly the name , consideruing lower and upper case
                            awinSettings.databaseName = listOfVCs.Item(0)
                        End If
                    Else
                        awinSettings.databaseName = listOfVCs.Item(0)
                    End If
                    Dim changeOK As Boolean = CType(databaseAcc, DBAccLayer.Request).updateActualVC(awinSettings.databaseName, awinSettings.VCid, err)
                    If Not changeOK Then
                        Throw New ArgumentException("No access to this VISBO Center ... program ends  ..." & vbCrLf & err.errorMsg)
                    End If

                ElseIf listOfVCs.Count > 1 Then
                    ' now choose what is  das gewünschte VC aus
                    If Not listOfVCs.Contains(awinSettings.databaseName) Then
                        Throw New ArgumentException("No access to this VISBO Center " & awinSettings.databaseName)
                    End If

                Else
                    ' user has no access to any VISBO Center 
                    Throw New ArgumentException("No access to a VISBO Center ")
                End If

            Else
                ' no valid Login
                Throw New ArgumentException("No valid Login")
            End If

            '
            ' Read appearance Definitions
            appearanceDefinitions.liste = CType(databaseAcc, DBAccLayer.Request).retrieveAppearancesFromDB("", Date.Now, False, err)
            If IsNothing(appearanceDefinitions.liste) Then
                ' user has no access to any VISBO Center 
                Throw New ArgumentException("No appearance Definitions in VISBO")
            End If

            '
            ' Read Customizations 
            Dim customizations As clsCustomization = CType(databaseAcc, DBAccLayer.Request).retrieveCustomizationFromDB("", Date.Now, False, err)

            If Not IsNothing(customizations) Then
                StartofCalendar = customizations.kalenderStart
                Call logger(ptErrLevel.logInfo, "rpaSetTypen", " StartOfCalendar: " & StartofCalendar.ToString)

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
            Else
                Throw New ArgumentException("No customization in VISBO")
            End If

            '
            ' now read Organisation 
            Dim currentOrga As clsOrganisation = CType(databaseAcc, DBAccLayer.Request).retrieveOrganisationFromDB("", Date.Now, False, err)

            If Not IsNothing(currentOrga) Then
                If currentOrga.count > 0 Then

                    If currentOrga.count > 0 Then
                        validOrganisations.addOrga(currentOrga)
                    End If

                    CostDefinitions = currentOrga.allCosts
                    RoleDefinitions = currentOrga.allRoles

                Else
                    Throw New ArgumentException("No organisation in VISBO")
                End If
            Else
                Throw New ArgumentException("No organisation in VISBO")
            End If

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
            Dim projectTemplates As clsProjekteAlle = CType(databaseAcc, DBAccLayer.Request).retrieveProjectTemplatesFromDB(err)

            If Not IsNothing(projectTemplates) Then
                Dim projVorlage As clsProjektvorlage
                For Each kvp As KeyValuePair(Of String, clsProjekt) In projectTemplates.liste

                    projVorlage = createTemplateOfProject(kvp.Value)
                    ' hiermit wird die _Dauer gesetzt
                    Dim vorlagenDauer = projVorlage.dauerInDays

                    Projektvorlagen.Add(projVorlage)
                Next
            End If

            If awinSettings.englishLanguage Then
                windowNames(PTwindows.mpt) = "VISBO Multiproject-Board"
                windowNames(PTwindows.massEdit) = "edit projects: "
                windowNames(PTwindows.meChart) = "project and portfolio Charts: "
                windowNames(PTwindows.mptpf) = "Portfolio Charts: "
                windowNames(PTwindows.mptpr) = "Project Charts"
            Else
                windowNames(PTwindows.mpt) = "VISBO Multiprojekt-Tafel"
                windowNames(PTwindows.massEdit) = "Projekte editieren: "
                windowNames(PTwindows.meChart) = "Projekt und Portfolio Charts: "
                windowNames(PTwindows.mptpf) = "Portfolio Charts: "
                windowNames(PTwindows.mptpr) = "Projekt Charts"
            End If


            projectboardViews(PTview.mpt) = Nothing
            projectboardViews(PTview.mptpr) = Nothing
            projectboardViews(PTview.mptprpf) = Nothing
            projectboardViews(PTview.meOnly) = Nothing
            projectboardViews(PTview.meChart) = Nothing

            projectboardWindows(PTwindows.mpt) = Nothing
            projectboardWindows(PTwindows.mptpr) = Nothing
            projectboardWindows(PTwindows.mptpf) = Nothing
            projectboardWindows(PTwindows.massEdit) = Nothing
            projectboardWindows(PTwindows.meChart) = Nothing

            result = True

        Catch ex As Exception

            result = False
            Call logger(ptErrLevel.logError, "rpaSetTypen", ex.Message)

        End Try

        rpaSetTypen = result

    End Function

    Private Function bestimmeRPACategory(ByVal fileName As String) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown

        ' Open fileName 
        If Not IsNothing(fileName) Then

            If My.Computer.FileSystem.FileExists(fileName) Then

                Try
                    Dim currentWB As xlns.Workbook = appInstance.Workbooks.Open(fileName)

                    Try
                        ' Check auf Project Batch-List
                        If result = PTRpa.visboUnknown Then
                            result = checkProjectBatchList(currentWB)
                        End If

                        ' Check auf VISBO Project Brief
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

                        ' Check auf VISBO Project Template  

                        ' Check auf Urlaubskalender 

                        ' Check auf Modifier Kapazitäten

                        ' Check auf externe Rahmenverträge 
                        If result = PTRpa.visboUnknown Then
                            result = checkExtRahmenvertr(currentWB)
                        End If

                        ' Check auf Instart eGecko Urlaube ... 

                        ' Check auf Zeuss Kapazitäten

                        ' Check auf Ist-Daten 

                        ' Check auf Telair TimeSheets

                        ' Check auf Tagetik new Project List 

                        ' Check auf Tagetik update projects 

                        ' Check auf Instart Calculation Template 

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
                        ersteZeile.Cells(1, 13).value.trim = "Business Unit" And
                        ersteZeile.Cells(1, 14).value.trim = "Description"

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

    Private Function checkVCOrganisation(ByVal currentWB As xlns.Workbook) As PTRpa
        Dim result As PTRpa = PTRpa.visboUnknown
        Dim possibleTableNames() As String = {"VisboCenterOrganisation"}
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

                If found Then
                    verifiedStructure = CStr(currentWS.Cells(1, 1).value).Trim = "name" And
                                        CStr(currentWS.Cells(1, 2).value).Trim = "uid"
                    Exit For
                End If

            Next


        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try


        checkVCOrganisation = result
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
    ''' returns the sequence of the project-names 
    ''' there is only one project-variant per ranking allowed
    ''' </summary>
    ''' <returns></returns>
    Public Function getRanking() As SortedList(Of Integer, String)

        Dim rankingList As New SortedList(Of Integer, String)
        Dim nameList As New SortedList(Of String, String)
        Dim key As String

        Dim zeile As Integer, spalte As Integer


        Dim tfZeile As Integer = 2
        Dim listOfpNames As New SortedList(Of String, String)
        Dim pName As String = ""
        Dim variantName As String = ""

        Dim lastRow As Integer


        Dim geleseneProjekte As Integer


        Dim firstZeile As xlns.Range


        zeile = 2
        spalte = 1
        geleseneProjekte = 0




        Try
            Dim activeWSListe As xlns.Worksheet = CType(appInstance.ActiveWorkbook.ActiveSheet,
                                                            Global.Microsoft.Office.Interop.Excel.Worksheet)

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
        Catch ex As Exception

            Throw New Exception("Fehler In Portfolio-Datei" & ex.Message)
        End Try

        getRanking = rankingList
    End Function

End Module
