Imports xlns = Microsoft.Office.Interop.Excel
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Newtonsoft.Json
Imports System.IO
Imports DBAccLayer
Imports WebServerAcc
Imports System.Security.Principal


Module rpaModule1

    Public myActivePortfolio As String = ""
    Public inputvalues As clsRPASetting = Nothing


    Public Sub Main()
        ' reads the VISBO RPA folder und treats each file it finds there appropriately
        ' in most cases new project and portfolio versions will be written 
        ' suggestions for Team Members will follow 
        ' automation in resource And team allocation will follow

        Dim nonStop As Boolean = True
        Dim errMsgCode As New clsErrorCodeMsg
        Dim msgTxt As String = ""
        Dim result As Boolean = False

        Dim rpaPath As String = My.Settings.rpaPath
        Dim swPath As String = My.Settings.swPath

        Dim rpaFolder As String = My.Computer.FileSystem.CombinePath(rpaPath, "RPA")
        Dim successFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "success")
        Dim failureFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "failure")
        Dim collectFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "collect")
        Dim logfileFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "logfiles")
        Dim unknownFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "unknown")
        Dim settingsFolder As String = My.Computer.FileSystem.CombinePath(rpaFolder, "settings")
        Dim settingJsonFile As String = My.Computer.FileSystem.CombinePath(settingsFolder, "rpa_setting.json")

        'Dim myActivePortfolio As String = ""
        'Dim inputvalues As clsRPASetting = Nothing

        ' Read the Setting-file of RPA
        If My.Computer.FileSystem.FileExists(settingJsonFile) Then
            Dim jsonSetting As String = File.ReadAllText(settingJsonFile)
            inputvalues = JsonConvert.DeserializeObject(Of clsRPASetting)(jsonSetting)
            ' is there a activePortfolio
            myActivePortfolio = inputvalues.activePortfolio
        Else
            ' Exit ! 
            ' read all files, categorize and verify them  
            msgTxt = "Exit - there is no File " & settingJsonFile
            Call logger(ptErrLevel.logError, "VISBO Robotic Process automation", msgTxt)
            Console.WriteLine(msgTxt)

            ' break the RPA - Service
            nonStop = False
        End If



        Try
            Dim anzFiles As Integer = 0

            '' FileNamen für logging zusammenbauen
            logfileNamePath = createLogfileName(rpaFolder, "")

            Dim listToProcess As New SortedList(Of String, Integer)
            Dim listToProcess2 As New SortedList(Of String, Integer)
            Dim listActualDataFiles As New SortedList(Of String, Integer)

            '    If My.Computer.FileSystem.FileExists(settingJsonFile) Then
            '        Dim jsonSetting As String = File.ReadAllText(settingJsonFile)
            '        Dim inputvalues As clsRPASetting = JsonConvert.DeserializeObject(Of clsRPASetting)(jsonSetting)

            '        ' is there a activePortfolio
            '        myActivePortfolio = inputvalues.activePortfolio

            ' now check whether or not the folder are existings , if not create them 
            If Not My.Computer.FileSystem.DirectoryExists(successFolder) Then
                My.Computer.FileSystem.CreateDirectory(successFolder)
            End If

            If Not My.Computer.FileSystem.DirectoryExists(failureFolder) Then
                My.Computer.FileSystem.CreateDirectory(failureFolder)
            End If

            If Not My.Computer.FileSystem.DirectoryExists(collectFolder) Then
                My.Computer.FileSystem.CreateDirectory(collectFolder)
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

            visboClient = "VISBO RPA / "

            ' 
            ' startUpRPA  liest orga, appearances und andere Settings - analog awinSetTypen , allerdings nie mit Versuch, etwas von Platte zu lesen ... 
            nonStop = startUpRPA(inputvalues.VisboCenter, inputvalues.VisboUrl, swPath)


            If nonStop Then
                ' Sendet eine Email an den User
                errMsgCode = New clsErrorCodeMsg
                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & "correct start of the RPA", errMsgCode)

            Else
                msgTxt = "wrong settings - exited without performing jobs ...."
                'Call MsgBox(msgTxt)
                Console.WriteLine(msgTxt)
                Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)
                errMsgCode = New clsErrorCodeMsg
                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & msgTxt, errMsgCode)

                ' Stoppt den Service aufgrund von ungültigen Settings
                nonStop = False
            End If



            ' never ending loop for importing the different files - RPA
            Do While nonStop

                Dim myName As String = ""
                Dim rpaCategory As New PTRpa
                listToProcess = New SortedList(Of String, Integer)
                listToProcess2 = New SortedList(Of String, Integer)


                Try

                    ' Completion-File delivered?
                    Dim completionFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(rpaFolder, FileIO.SearchOption.SearchTopLevelOnly, "Timesheet_completed*.*")
                    Dim completedOK As Boolean = (completionFiles.Count > 0)


                    ' read all Excel based files 
                    Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(rpaFolder, FileIO.SearchOption.SearchTopLevelOnly, "*.xlsx")



                    For Each fullFileName As String In listOfImportfiles

                        myName = My.Computer.FileSystem.GetName(fullFileName)

                        ' Bestimme den Import-Typ der zu importierenden Daten
                        rpaCategory = bestimmeRPACategory(fullFileName)

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

                        myName = My.Computer.FileSystem.GetName(fullFileName)
                        rpaCategory = PTRpa.visboMPP

                        If Not listToProcess.ContainsKey(myName) Then
                            listToProcess.Add(fullFileName, CInt(rpaCategory))
                        End If

                    Next

                    listOfImportfiles = Nothing

                    ImportProjekte.Clear()
                    Dim importOrganisations As New clsOrganisations
                    Dim importCustomization As New clsCustomization
                    Dim importAppearances As New clsAppearances
                    Dim importDate As Date = Date.Now()
                    Dim allOk As Boolean = False


                    If completedOK Then

                        logfileNamePath = createLogfileName(rpaFolder, myActivePortfolio)

                        ' that means, all timesheets are in the RPA folder
                        For Each kvp As KeyValuePair(Of String, Integer) In listToProcess

                            'collect the Timesheets for actualData in one separate list and dir 'collect'
                            If kvp.Value = PTRpa.visboActualData2 Then
                                myName = My.Computer.FileSystem.GetName(kvp.Key)
                                Dim newDestination As String = My.Computer.FileSystem.CombinePath(collectFolder, myName)
                                My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                                Call logger(ptErrLevel.logInfo, "collect: ", myName)
                                listActualDataFiles.Add(newDestination, kvp.Value)
                            Else
                                ' all other files to import
                                listToProcess2.Add(kvp.Key, kvp.Value)
                            End If
                        Next

                        listToProcess = listToProcess2

                        ' import actualData like Timesheets from collectFolder
                        allOk = processVisboActualData2("Timesheets", myActivePortfolio, collectFolder, importDate)

                        For Each kvp As KeyValuePair(Of String, Integer) In listActualDataFiles
                            myName = My.Computer.FileSystem.GetName(kvp.Key)
                            If allOk Then
                                Dim newDestination As String = My.Computer.FileSystem.CombinePath(successFolder, myName)
                                My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                                Call logger(ptErrLevel.logInfo, "success: ", myName)
                                Console.WriteLine(myName & ": successful ...")
                                errMsgCode = New clsErrorCodeMsg
                                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & myName & ": successful ...", errMsgCode)
                            Else
                                Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
                                My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                                Call logger(ptErrLevel.logError, "failed: ", myName)
                                Console.WriteLine(myName & ": with errors ...")

                                errMsgCode = New clsErrorCodeMsg
                                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                                                                                            & myName & ": with errors ..." & vbCrLf _
                                                                                            & "Look for more details in the Failure-Folder", errMsgCode)

                            End If

                        Next

                        ' logfile in entsprechenden folder verschieben
                        Dim logfileName As String = My.Computer.FileSystem.GetName(logfileNamePath)
                        If Not allOk Then
                            Dim newLog As String = My.Computer.FileSystem.CombinePath(failureFolder, logfileName)
                            My.Computer.FileSystem.MoveFile(logfileNamePath, newLog, True)
                        End If

                    End If

                    For Each kvp As KeyValuePair(Of String, Integer) In listToProcess

                        myName = My.Computer.FileSystem.GetName(kvp.Key)
                        Dim currentWB As xlns.Workbook = Nothing


                        Try

                            If Not kvp.Value = PTRpa.visboMPP _
                                And Not kvp.Value = PTRpa.visboActualData1 _
                                And Not kvp.Value = PTRpa.visboActualData2 Then

                                appInstance.DisplayAlerts = False
                                currentWB = appInstance.Workbooks.Open(kvp.Key)
                            End If

                            logfileNamePath = createLogfileName(rpaFolder, myName)
                            Select Case kvp.Value
                                Case CInt(PTRpa.visboProjectList)

                                    allOk = processProjectList(myName, myActivePortfolio)

                                Case CInt(PTRpa.visboFindProjectStart)

                                    allOk = processFindProjectStart(myName, myActivePortfolio)

                                Case CInt(PTRpa.visboMPP)

                                    allOk = processMppFile(kvp.Key, importDate)

                                Case CInt(PTRpa.visboProject)

                                    allOk = processVisboBrief(myName, importDate)

                                Case CInt(PTRpa.visboJira)

                                    allOk = processVisboJira(kvp.Key, importDate)

                                Case CInt(PTRpa.visboDefaultCapacity)
                                    allOk = True

                                Case CInt(PTRpa.visboInitialOrga)

                                    allOk = processInitialOrga(myName)

                                Case CInt(PTRpa.visboRoundtripOrga)

                                    allOk = processRoundTripOrga(myName)

                                Case CInt(PTRpa.visboModifierCapacities)

                                    allOk = True

                                Case CInt(PTRpa.visboExternalContracts)

                                    allOk = True

                                Case CInt(PTRpa.visboActualData1)

                                    allOk = processVisboActualData1(kvp.Key, importDate)

                                Case CInt(PTRpa.visboActualData2)

                                    'Dim completionFiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(rpaFolder, FileIO.SearchOption.SearchTopLevelOnly, "Timesheet_completed*.*")
                                    ' in collectFolder verschieben
                                    Dim newDestination As String = My.Computer.FileSystem.CombinePath(collectFolder, myName)
                                    My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                                    Call logger(ptErrLevel.logInfo, "collect: ", myName)
                                    ' nachsehen ob collect vollständig
                                    If completionFiles.Count > 0 Then
                                        allOk = processVisboActualData2(kvp.Key, myActivePortfolio, collectFolder, importDate)
                                    End If


                                Case Else

                            End Select

                            ' Sendet eine Email an den User

                            'Dim result_sendEmail As Boolean = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("files abgearbeitet", errMsgCode)

                            If Not (kvp.Value = PTRpa.visboMPP Or
                                        kvp.Value = PTRpa.visboJira Or
                                        kvp.Value = PTRpa.visboActualData1 Or
                                        kvp.Value = PTRpa.visboActualData2) Then

                                If allOk Then
                                    CType(currentWB.Worksheets(1), xlns.Worksheet).Cells(1, 1).interior.color = visboFarbeGreen
                                Else
                                    CType(currentWB.Worksheets(1), xlns.Worksheet).Cells(1, 1).interior.color = visboFarbeRed
                                End If
                                currentWB.Close(SaveChanges:=True)
                            End If

                            'If Not IsNothing(currentWB) Then
                            '    currentWB.Close(SaveChanges:=True)
                            'End If

                            If Not kvp.Value = PTRpa.visboActualData2 Then

                                If allOk Then
                                    Dim newDestination As String = My.Computer.FileSystem.CombinePath(successFolder, myName)
                                    My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                                    Call logger(ptErrLevel.logInfo, "success: ", myName)
                                    'Dim logfileName As String = My.Computer.FileSystem.GetName(logfileNamePath)
                                    'Dim newLog As String = My.Computer.FileSystem.CombinePath(successFolder, logFileName)
                                    'My.Computer.FileSystem.MoveFile(logfileNamePath, newLog, True)
                                    Console.WriteLine(myName & ": successful ...")
                                    errMsgCode = New clsErrorCodeMsg
                                    result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & myName & ": successful ...", errMsgCode)


                                Else
                                    Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
                                    My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                                    Call logger(ptErrLevel.logError, "failed: ", myName)
                                    Dim logfileName As String = My.Computer.FileSystem.GetName(logfileNamePath)
                                    Dim newLog As String = My.Computer.FileSystem.CombinePath(failureFolder, logfileName)
                                    My.Computer.FileSystem.MoveFile(logfileNamePath, newLog, True)

                                    errMsgCode = New clsErrorCodeMsg
                                    result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                                                                                                & myName & ": with errors ..." & vbCrLf _
                                                                                                & "Look for more details in the Failure-Folder", errMsgCode)
                                End If

                            End If



                        Catch ex As Exception
                            Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
                            My.Computer.FileSystem.MoveFile(kvp.Key, newDestination, True)
                            Call logger(ptErrLevel.logError, "failed: ", ex.Message)
                            If Not kvp.Value = PTRpa.visboMPP Then
                                currentWB.Close(SaveChanges:=True)
                            End If
                            Console.WriteLine(myName & ": failed ...")
                        End Try

                    Next

                    Console.WriteLine("looking for next jobs!")
                    'msgTxt = "looking for next jobs!"
                    'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)

                Catch ex As Exception

                    If awinSettings.englishLanguage Then
                        msgTxt = "Error importing: "
                    Else
                        msgTxt = "Fehler beim Import von: "
                    End If
                    Call logger(ptErrLevel.logsevereError, msgTxt, myName & "/" & rpaCategory.ToString)
                End Try

            Loop



            ' now store User Login Data
            My.Settings.userNamePWD = awinSettings.userNamePWD

                ' speichern 
                My.Settings.Save()

                '' now release all writeProtections ...
            'Dim errMsgCode As New clsErrorCodeMsg
            'If CType(databaseAcc, DBAccLayer.Request).cancelWriteProtections(dbUsername, errMsgCode) Then
            '    If awinSettings.visboDebug Then
            '        Call MsgBox("Ihre vorübergehenden Schreibsperren wurden aufgehoben")
            '    End If
            'Else
            '    msgTxt = "Write Protections could not be released ! Please do so in Web-UI ..."
            '    Call logger(ptErrLevel.logError, "VISBO Robotic Process automation End", msgTxt)
            '    Console.WriteLine(msgTxt)
            'End If


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

                awinSettings.visboServer = True

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

                    Dim tmpActDataString As String = currentOrga.allRoles.getActualdataOrgaUnits
                    If tmpActDataString = "" And awinSettings.ActualdataOrgaUnits <> "" Then
                        ' do nothing, leave it as is 
                    Else
                        awinSettings.ActualdataOrgaUnits = tmpActDataString
                    End If

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
                    appInstance.DisplayAlerts = False
                    Dim currentWB As xlns.Workbook = appInstance.Workbooks.Open(fileName, UpdateLinks:=0)
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
        Dim blattName2 As String = "Parameters"


        Try

            Dim currentWS As xlns.Worksheet = CType(currentWB.Worksheets.Item(blattName1), xlns.Worksheet)
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

            End If
        Catch ex As Exception
            result = PTRpa.visboUnknown
        End Try

        If verifiedStructure Then
            result = PTRpa.visboFindProjectStart
        Else
            result = PTRpa.visboUnknown
        End If

        checkFindBestStarts = result
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

    Public Function getOverloadParams() As Double()

        Dim blattName As String = "Parameters"
        Dim result As Double()
        ReDim result(1)
        result(0) = 1.05
        result(1) = 1.0

        Try
            Dim currentWB As xlns.Workbook = CType(appInstance.ActiveWorkbook,
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

    Private Function processVisboBrief(ByVal myName As String, ByVal importDate As Date) As Boolean

        Dim allOK As Boolean = False
        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboProject.ToString, myName)

        'read Project Brief and put it into ImportProjekte
        Try
            Dim hproj As clsProjekt = Nothing

            ' read the file and import into hproj
            Call awinImportProjectmitHrchy(hproj, Nothing, False, importDate)

            allOK = Not IsNothing(hproj)
            If allOK Then
                Try
                    Dim keyStr As String = calcProjektKey(hproj)
                    ImportProjekte.Add(hproj, updateCurrentConstellation:=False)
                    'AlleProjekte.Add(hproj, updateCurrentConstellation:=False)

                    Call importProjekteEintragen(importDate, drawPlanTafel:=False, fileFrom3rdParty:=True, getSomeValuesFromOldProj:=True, calledFromActualDataImport:=False, calledFromRPA:=True)
                Catch ex2 As Exception
                    allOK = False
                    Call logger(ptErrLevel.logError, "RPA Error Importing MS Project file " & PTRpa.visboProject.ToString, ex2.Message)
                End Try
            Else
                Call logger(ptErrLevel.logError, "RPA Error Importing MS Project file " & PTRpa.visboProject.ToString, myName)
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

    Private Function processFindProjectStart(ByVal myName As String, ByVal myActivePortfolio As String) As Boolean

        Dim allOk As Boolean = False

        Try
            Dim portfolioName As String = myName.Substring(0, myName.IndexOf(".xls"))
            Dim overloadAllowedinMonths As Double = 1.05
            Dim overloadAllowedTotal As Double = 1.0

            Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboFindProjectStart.ToString, myName)
            Dim readProjects As Integer = 0
            Dim createdProjects As Integer = 0
            'Dim importedProjects As Integer = ImportProjekte.Count

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


            ' jetzt alle Projekte aus der Liste holen und die OverloadParams holen 
            Try
                Dim listOfProjs As SortedList(Of Integer, String) = getRanking()
                Dim tmpValues As Double() = getOverloadParams()

                overloadAllowedinMonths = tmpValues(0)
                overloadAllowedTotal = tmpValues(1)

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


            If allOk Then
                Call logger(ptErrLevel.logInfo, "Project List imported: " & myName, ImportProjekte.Count & " read; ")
            Else
                Call logger(ptErrLevel.logError, "failure in Processing: " & myName, PTRpa.visboFindProjectStart.ToString)
            End If

            If allOk Then

                Dim skillIDs As Collection = ImportProjekte.getRoleSkillIDs()

                For Each si As String In skillIDs
                    If Not skillList.Contains(si) Then
                        skillList.Add(si)
                    End If
                Next

                Dim noActivePortfolio As Boolean = True
                Dim dbPortfolioNames As New SortedList(Of String, String)

                ' if Portfolio with active Projects is given and exists:  
                ' then we probably do have a brownfield
                If myActivePortfolio <> "" Then

                    Dim errMsg As New clsErrorCodeMsg
                    dbPortfolioNames = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)
                    noActivePortfolio = Not dbPortfolioNames.ContainsKey(myActivePortfolio)
                End If

                If noActivePortfolio Then
                    Call logger(ptErrLevel.logError, "no active Portfolio: " & myActivePortfolio, PTRpa.visboFindProjectStart.ToString)
                Else
                    ' check whether and how projects are fitting to the already existing Portfolio 
                    allOk = processProjectListWithActivePortfolio(aggregationList,
                                                                     skillList,
                                                                     myActivePortfolio, dbPortfolioNames(myActivePortfolio), portfolioName, overloadAllowedinMonths, overloadAllowedTotal)
                End If

            Else
                ' no additional logger necessary - is done in storeImportProjekte
            End If


            ' now empty the complete session  
            Call emptyRPASession()
            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboProjectList.ToString, myName)

        Catch ex As Exception
            Call logger(ptErrLevel.logError, "errors occurred when processing: " & PTRpa.visboProjectList.ToString, myName & ": " & ex.Message)
        End Try

        processFindProjectStart = allOk

    End Function

    Private Function processProjectList(ByVal myName As String, ByVal myActivePortfolio As String) As Boolean

        Dim allOk As Boolean = False

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



            allOk = awinImportProjektInventur(readProjects, createdProjects)
            If allOk Then
                Call logger(ptErrLevel.logInfo, "Project List imported: " & myName, readProjects & " read; " & createdProjects & " created")
                allOk = storeImportProjekte()
            Else
                Call logger(ptErrLevel.logError, "failure in Processing: " & myName, PTRpa.visboProjectList.ToString)
            End If

            If allOk Then

                Dim skillIDs As Collection = ImportProjekte.getRoleSkillIDs()

                For Each si As String In skillIDs
                    If Not skillList.Contains(si) Then
                        skillList.Add(si)
                    End If
                Next

                Dim doTheInitialJob As Boolean = True
                Dim dbPortfolioNames As New SortedList(Of String, String)

                ' if Portfolio with active Projects is given and exists:  
                ' then we probably do have a brownfield
                If myActivePortfolio <> "" Then
                    ' load portfolio projects 
                    ' now store the Portfolio , with name portfolioName
                    Dim errMsg As New clsErrorCodeMsg
                    dbPortfolioNames = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errMsg)
                    doTheInitialJob = Not dbPortfolioNames.ContainsKey(myActivePortfolio)
                End If

                If doTheInitialJob Then
                    allOk = processProjectListWithoutActivePortfolio(aggregationList,
                                                                     skillList,
                                                                     portfolioName, overloadAllowedinMonths, overloadAllowedTotal)
                Else
                    ' check whether and how projects are fitting to the already existing Portfolio 
                    allOk = processProjectListWithActivePortfolio(aggregationList,
                                                                     skillList,
                                                                     myActivePortfolio, dbPortfolioNames(myActivePortfolio), portfolioName, overloadAllowedinMonths, overloadAllowedTotal)
                End If

            Else
                ' no additional logger necessary - is done in storeImportProjekte
            End If


            ' now empty the complete session  
            Call emptyRPASession()
            Call logger(ptErrLevel.logInfo, "end Processing: " & PTRpa.visboProjectList.ToString, myName)

        Catch ex As Exception
            Call logger(ptErrLevel.logError, "errors occurred when processing: " & PTRpa.visboProjectList.ToString, myName & ": " & ex.Message)
        End Try

        processProjectList = allOk

    End Function

    ''' <summary>
    ''' in ImportProjekte sind alle aktuell eingelesenen Projekte 
    ''' </summary>
    ''' <param name="myActivePortfolio"></param>
    ''' <param name="listName"></param>
    ''' <param name="overloadAllowedInMonths"></param>
    ''' <param name="overloadAllowedTotal"></param>
    ''' <returns></returns>
    Private Function processProjectListWithActivePortfolio(ByVal aggregationList As List(Of String),
                                                           ByVal skillList As List(Of String),
                                                           ByVal myActivePortfolio As String,
                                                           ByVal myPortfolioVPID As String,
                                                           ByVal listName As String,
                                                           ByVal overloadAllowedInMonths As Double,
                                                           ByVal overloadAllowedTotal As Double) As Boolean
        Dim result As Boolean = True
        Dim saveShowRangeLeft As Integer = showRangeLeft
        Dim saveShowRangeRight As Integer = showRangeRight
        Dim Err As New clsErrorCodeMsg
        Dim msgTxt As String = ""

        Dim heute As Date = Date.Now

        Try
            ShowProjekte.Clear()
            AlleProjekte.Clear()

            ' now load the the portfolio and all projects of portfolio 
            ' hole Portfolio (pName,vName) aus der db
            Dim cTime As Date = heute
            Dim myConstellation As clsConstellation = CType(databaseAcc, DBAccLayer.Request).retrieveOneConstellationFromDB(myActivePortfolio,
                                                                                               "", cTime, Err, variantName:="", storedAtOrBefore:=heute)

            If Not IsNothing(myConstellation) Then
                Call logger(ptErrLevel.logInfo, "Loading Projects from Portfolio " & myActivePortfolio, " start of Operation ... ")
                ' tmpname in die Session-Liste wieder aufnehmen
                projectConstellations.Add(myConstellation)
                For Each kvp As KeyValuePair(Of String, clsConstellationItem) In myConstellation.Liste

                    Dim pName As String = getPnameFromKey(kvp.Key)
                    Dim vName As String = getVariantnameFromKey(kvp.Key)
                    Dim hproj As clsProjekt = getProjektFromSessionOrDB(pName, vName, AlleProjekte, heute)
                    If Not IsNothing(hproj) Then

                        AlleProjekte.Add(hproj)
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
            Dim rankingList As SortedList(Of Integer, String) = getRanking()
            Dim deltaInDays As Integer

            ' now create a Portfolio variant with unchanged new projects ...
            Dim removeSPList As New List(Of String)
            Dim removeAPList As New List(Of String)

            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList
                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                If Not IsNothing(hproj) Then
                    If Not AlleProjekte.Containskey(rankingPair.Value) Then
                        AlleProjekte.Add(hproj)
                        removeAPList.Add(rankingPair.Value)
                    Else
                        ' bring updated hproj into AlleProjekte
                        AlleProjekte.Add(hproj)
                    End If

                    If Not ShowProjekte.contains(hproj.name) Then
                        ShowProjekte.Add(hproj)
                        removeSPList.Add(hproj.name)
                    Else
                        ShowProjekte.AddAnyway(hproj)
                    End If


                End If
            Next

            Dim toStoreConstellation As clsConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=myActivePortfolio, vName:=listName)

            Dim outputCollection As New Collection
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

            If outputCollection.Count > 0 Then
                Call logger(ptErrLevel.logInfo, "Project List Import, Store Portfolio-Variant " & listName & " result:", outputCollection)
            End If

            ' now rest Showprojekte to formerStatus 
            For Each tmpName As String In removeAPList
                AlleProjekte.Remove(tmpName)
            Next

            For Each tmpName As String In removeSPList
                ShowProjekte.Remove(tmpName)
            Next


            ' now check whether there are overutilizations 
            ' if so , move showRangeLeft and showrangeRight  1 by 1 , until there are no overutilizations any more 

            showRangeLeft = ShowProjekte.getMinMonthColumn
            showRangeRight = ShowProjekte.getMaxMonthColumn
            Dim stopValue As Integer = showRangeRight

            Dim overutilizationFound As Boolean = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)

            ' now move the timeframe step by step until there are no overutilizations any more
            Do While overutilizationFound And showRangeLeft <= stopValue

                showRangeLeft = showRangeLeft + 1
                showRangeRight = showRangeRight + 1
                overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)

            Loop

            If overutilizationFound Then
                msgTxt = "no timeframe to be found to start new projects " & myActivePortfolio
                Call logger(ptErrLevel.logError, msgTxt, " calculation failed ..")
                Throw New ArgumentException(msgTxt)
            End If
            '



            ' create variant , if necessary
            ' rankingList keeps the sequence within the Excel file. So user adds some fields important to him for prioritization , he add these fields , sorts it in th eExcel. 
            ' It then represents the sequence: Row1 is the most important project 
            For Each rankingPair As KeyValuePair(Of Integer, String) In rankingList

                Dim hproj As clsProjekt = ImportProjekte.getProject(rankingPair.Value)
                Dim stdDuration As Integer = hproj.dauerInDays
                Dim myDuration As Integer = stdDuration
                Dim minDuration As Integer = CInt(stdDuration * 0.8)



                Dim storeRequired As Boolean = False

                Dim newStartDate As Date
                Dim newEndDate As Date
                Dim key As String = calcProjektKey(hproj)

                If getColumnOfDate(hproj.startDate) < showRangeLeft Then

                    ' create variant if not already done
                    If hproj.variantName <> "arb" Then
                        hproj = hproj.createVariant("arb", "variant was created and moved to avoid resource bottlenecks")
                        ' bring that into AlleProjekte
                        key = calcProjektKey(hproj)
                        deltaInDays = DateDiff(DateInterval.Day, hproj.startDate, getDateofColumn(showRangeLeft, False))

                        newStartDate = hproj.startDate.AddDays(deltaInDays)
                        newEndDate = hproj.endeDate.AddDays(deltaInDays)

                        Dim tmpProj As clsProjekt = moveProject(hproj, newStartDate, newEndDate)

                        If Not IsNothing(tmpProj) Then
                            hproj = tmpProj
                            storeRequired = True
                        Else
                            msgTxt = "project could be moved"
                        End If

                    End If

                End If


                ' check auf Exists is not necessary with AlleProjekte, because it will be replaced if it already exists 
                AlleProjekte.Add(hproj)
                ShowProjekte.AddAnyway(hproj)

                ' now define skill-List, because it is good enough to only consider skills of the hproj under consideration 
                skillList.Clear()
                Dim skillIDs As Collection = hproj.getSkillNameIds

                For Each si As String In skillIDs
                    If Not skillList.Contains(si) Then
                        skillList.Add(si)
                    End If
                Next

                ' now define showrangeLeft and showrangeRight from hproj 
                showRangeLeft = getColumnOfDate(hproj.startDate)
                showRangeRight = getColumnOfDate(hproj.endeDate)

                overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)
                Dim sumIterations As Integer = 0
                Dim endIterations As Integer = 0
                Dim durationIterations As Integer = 0

                If overutilizationFound Then

                    ' create variant if not already done
                    If hproj.variantName <> "arb" Then
                        hproj = hproj.createVariant("arb", "variant to avoid resource bottlenecks")

                        key = calcProjektKey(hproj)
                        AlleProjekte.Add(hproj)
                    End If

                    deltaInDays = 7
                    Dim maxEndIterations As Integer = CInt(182 / deltaInDays)
                    Dim maxDurationIterations As Integer = CInt((stdDuration - minDuration) / deltaInDays) + 1

                    Dim rememberStartDate As Date = hproj.startDate
                    Dim rememberEndDate As Date = hproj.endeDate

                    Try
                        Dim tmpProj As clsProjekt = Nothing
                        Do While overutilizationFound And endIterations <= maxEndIterations
                            ' move project by deltaIndays

                            newStartDate = rememberStartDate.AddDays(deltaInDays)

                            Do While overutilizationFound And durationIterations <= maxDurationIterations

                                newEndDate = rememberEndDate
                                tmpProj = moveProject(hproj, newStartDate, newEndDate)


                                If Not IsNothing(tmpProj) Then

                                    hproj = tmpProj

                                    ' now replace in AlleProjekte, ShowProjekte 
                                    AlleProjekte.Add(tmpProj)
                                    ShowProjekte.AddAnyway(tmpProj)

                                    ' now define showrangeLeft and showrangeRight from hproj 
                                    showRangeLeft = getColumnOfDate(hproj.startDate)
                                    showRangeRight = getColumnOfDate(hproj.endeDate)

                                    Dim infomsg As String = "... trying out " & hproj.getShapeText & hproj.startDate.ToShortDateString & " - " & hproj.endeDate.ToShortDateString
                                    Console.WriteLine(infomsg)

                                    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)

                                    If overutilizationFound Then
                                        durationIterations = durationIterations + 1
                                    End If

                                Else
                                    ' Error occurred 
                                    Throw New ArgumentException("tmpProj is Nothing")
                                End If

                                newStartDate = newStartDate.AddDays(deltaInDays)
                            Loop

                            If overutilizationFound Then

                                rememberStartDate = rememberStartDate.AddDays(deltaInDays)
                                rememberEndDate = rememberEndDate.AddDays(deltaInDays)

                                tmpProj = moveProject(hproj, rememberStartDate, rememberEndDate)

                                If Not IsNothing(tmpProj) Then

                                    hproj = tmpProj

                                    ' now replace in AlleProjekte, ShowProjekte 
                                    AlleProjekte.Add(tmpProj)
                                    ShowProjekte.AddAnyway(tmpProj)

                                    ' now define showrangeLeft and showrangeRight from hproj 
                                    showRangeLeft = getColumnOfDate(hproj.startDate)
                                    showRangeRight = getColumnOfDate(hproj.endeDate)

                                    Dim infomsg As String = "... trying out " & hproj.getShapeText & hproj.startDate.ToShortDateString & " - " & hproj.endeDate.ToShortDateString
                                    Console.WriteLine(infomsg)

                                    overutilizationFound = ShowProjekte.overLoadFound(aggregationList, skillList, False, overloadAllowedInMonths, overloadAllowedTotal)

                                    If overutilizationFound Then
                                        endIterations = endIterations + 1
                                    End If

                                Else
                                    ' Error occurred 
                                    Throw New ArgumentException("tmpProj is Nothing")
                                End If
                            End If


                        Loop

                    Catch ex As Exception
                        Dim infomsg As String = "failure: could not create project-variant " & hproj.getShapeText
                        Dim myMessages As New Collection
                        Call logger(ptErrLevel.logError, infomsg, myMessages)
                        overutilizationFound = True
                    End Try

                    If Not overutilizationFound Then
                        ' it is already in there ... but now needed to be stored
                        storeRequired = True
                    Else
                        ' take it out again , because there was no solution
                        AlleProjekte.Remove(key)
                        ShowProjekte.Remove(hproj.name)
                    End If

                Else
                    ' all ok, just continue
                End If

                If storeRequired Then
                    Dim myMessages As New Collection
                    If storeSingleProjectToDB(hproj, myMessages) Then
                        Dim infomsg As String = "success: created " & endIterations & " variants to avoid bottlenecks " & hproj.getShapeText
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
                    Dim infomsg As String = "success: could be added to portfolio variant as-is " & hproj.getShapeText
                    Dim myMessages As New Collection
                    Call logger(ptErrLevel.logInfo, infomsg, myMessages)
                    Console.WriteLine(infomsg)
                End If

            Next

            toStoreConstellation = currentSessionConstellation.copy(dontConsiderNoShows:=True,
                                                                                            cName:=myActivePortfolio, vName:=listName & "-arb")

            outputCollection.Clear()
            Call storeSingleConstellationToDB(outputCollection, toStoreConstellation, Nothing)

            If outputCollection.Count > 0 Then
                Call logger(ptErrLevel.logError, "Project List Import, Store Portfolio-Variant arb failed:", outputCollection)
            End If


        Catch ex As Exception
            result = False
        End Try

        showRangeLeft = saveShowRangeLeft
        showRangeRight = saveShowRangeRight

        processProjectListWithActivePortfolio = result

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
                result = CType(databaseAcc, DBAccLayer.Request).storeVCSettingsToDB(importedOrga,
                                                        CStr(settingTypes(ptSettingTypes.organisation)),
                                                        orgaName,
                                                        importedOrga.validFrom,
                                                        err)

                If result = True Then
                    allOK = True
                    msgTxt = "ok, Organisation, valid from " & importedOrga.validFrom.ToShortDateString & " stored ..."
                    Console.WriteLine(msgTxt)
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
        Try

            Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboRoundtripOrga.ToString, myName)

            ' ===========================================================
            ' Konfigurationsdatei lesen und Validierung durchführen

            ' wenn es gibt - lesen der ControllingSheet und anderer, die durch configActualDataImport beschrieben sind
            Dim configOrgaImport As String = awinPath & configfilesOrdner & "configOrgaImport.xlsx"
            Dim orgaImportConfig As New SortedList(Of String, clsConfigOrgaImport)
            Dim lastrow As Integer = 0

            Call logger(ptErrLevel.logInfo, "start reading configuration: " & PTRpa.visboRoundtripOrga.ToString, configOrgaImport)

            ' check Config-File - zum Einlesen der Istdaten gemäß Konfiguration
            ' hier werden Werte für actualDataFile, actualDataConfig gesetzt
            Dim allesOK As Boolean = checkOrgaImportConfig(configOrgaImport, myName, orgaImportConfig, lastrow, outputCollection)

            If Not allesOK Then
                Call logger(ptErrLevel.logError, "error reading configuration: " & PTRpa.visboRoundtripOrga.ToString, configOrgaImport)
                processRoundTripOrga = False
                Exit Function
            End If

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
                        allOK = True
                        msgTxt = "ok, Organisation, valid from " & importedOrga.validFrom.ToShortDateString & " stored ..."
                        Console.WriteLine(msgTxt)
                        Call logger(ptErrLevel.logInfo, PTRpa.visboRoundtripOrga.ToString, msgTxt)
                    Else
                        allOK = False
                        msgTxt = "Storing organisaiton failed "
                        Call logger(ptErrLevel.logError, PTRpa.visboRoundtripOrga.ToString, msgTxt)
                    End If
                End If

                Call logger(ptErrLevel.logInfo, "endProcessing: " & PTRpa.visboRoundtripOrga.ToString, myName)
            Catch ex As Exception
                allOK = False
            End Try

        Catch ex As Exception
            allOK = False
            msgTxt = ""
            Call logger(ptErrLevel.logError, PTRpa.visboRoundtripOrga.ToString, ex.Message)
        End Try

        processRoundTripOrga = allOK

    End Function


    Private Function processVisboJira(ByVal myName As String, ByVal importDate As Date) As Boolean

        Dim allOk As Boolean = True

        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboJira.ToString, myName)

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

            Dim outputLine As String = ""

            Dim boardWasEmpty As Boolean = (ShowProjekte.Count > 0)

            ' Konfigurationsdatei lesen und Validierung durchführen

            ' wenn es gibt - lesen der Zeuss- listen und anderer, die durch configCapaImport beschrieben sind
            Dim configJIRAProjects As String = awinPath & configfilesOrdner & "configJIRAProjectImport.xlsx"

            ' Read & check Config-File - ist evt.  in my.settings.xlsConfig festgehalten
            Dim allesOK As Boolean = checkProjectImportConfig(configJIRAProjects, projectsFile, JIRAProjectsConfig, lastrow, outPutCollection)

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
    ''' standard import of actual data like Instart
    ''' </summary>
    ''' <param name="myName"></param>
    ''' <param name="importDate"></param>
    ''' <returns></returns>
    Private Function processVisboActualData1(ByVal myName As String, ByVal importDate As Date) As Boolean

        Dim allOk As Boolean = True

        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboActualData1.ToString, myName)

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


    Private Function processVisboActualData2(ByVal myName As String, ByVal portfolioName As String, ByVal dirName As String, ByVal importDate As Date) As Boolean

        Dim allOk As Boolean = True

        Call logger(ptErrLevel.logInfo, "start Processing: " & PTRpa.visboActualData2.ToString, myName)

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
        Dim configActualDataImport As String = awinPath & configfilesOrdner & "configActualDataImport.xlsx"

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
                    Console.WriteLine(txtMsg)

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
                    Console.WriteLine(txtMsg)

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

                        ' es fehlen timeSheets von manchen Mitarbeitern
                        allOk = allOk And False
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

                    Call logger(ptErrLevel.logWarning, "Import ActualData2", logmessage)

                    ' Auch wenn unbekannte Rollen und Kosten drin waren - die Projekte enthalten die ja dann nicht und können deshalb aufgenommen werden ..
                    Try
                        Call importProjekteEintragen(importDate:=importDate, drawPlanTafel:=False, fileFrom3rdParty:=False,
                                             getSomeValuesFromOldProj:=False, calledFromActualDataImport:=True)


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

                    Next
                    ' Fehler erfolgt
                    ' Dateien müssen in failure-Directory verschoben werden
                    Call MsgBox("TODO")

                End If


            Else
                If awinSettings.englishLanguage Then
                    outPutline = "No file to import actual data"
                Else
                    outPutline = "Es gibt keine Datei zum Importieren von Istdaten"
                End If

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

            allOk = allOk And False

        End If    ' checkConfigOK

        processVisboActualData2 = allOk

    End Function


    ''''' <summary>
    ''''' composition of the FileName of the different Logfiles for the RPA Import as well
    ''''' </summary>
    ''''' <param name="rpaFolder"></param>
    ''''' <param name="rpaImportfile"></param>
    ''''' <returns></returns>
    ''Public Function createLogfileName(Optional ByVal rpaFolder As String = "", Optional ByVal rpaImportfile As String = "") As String
    ''    ' FileNamen zusammenbauen
    ''    Dim logfileOrdner As String = "logfiles"


    ''    If IsNothing(awinPath) Then
    ''        Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    ''        awinPath = My.Computer.FileSystem.CombinePath(curUserDir, "VISBO")
    ''    End If
    ''    Dim logfilePath As String = My.Computer.FileSystem.CombinePath(awinPath, logfileOrdner)

    ''    ' write logfiles in rpaFolder, if RPA was started
    ''    If Not IsNothing(rpaFolder) And rpaFolder <> "" Then
    ''        logfilePath = My.Computer.FileSystem.CombinePath(rpaFolder, logfileOrdner)
    ''    End If

    ''    Dim logfileName As String = "logfile" & "_" & rpaImportfile & "_" & logDate.Year.ToString & logDate.Month.ToString("0#") & logDate.Day.ToString("0#") & "_" & logDate.TimeOfDay.ToString.Replace(":", "-") & ".txt"
    ''    Dim logfileNamePath As String = My.Computer.FileSystem.CombinePath(logfilePath, logfileName)
    ''    ' Fragen, ob bereits existiert - eventuell nicht nötig
    ''    If Not My.Computer.FileSystem.DirectoryExists(logfilePath) Then
    ''        My.Computer.FileSystem.CreateDirectory(logfilePath)
    ''    End If
    ''    createLogfileName = logfileNamePath
    ''End Function

    ''''' <summary>
    '''''  schreibt in das logfile mit Errorlevel
    ''''' </summary>
    ''''' <param name="errLevel"></param>
    ''''' <param name="text"></param>
    ''''' <param name="addOn"></param>
    ''''' <param name="anzFehler"></param>
    ''Public Sub logger(ByVal logfileNamePath As String, ByVal errLevel As Integer, ByVal text As String, ByVal addOn As String, ByVal anzFehler As Long)

    ''    Try
    ''        Dim strMeld As String
    ''        ' tk 15.8. in Order to avoid warning statements in Visual Studio 
    ''        ' once this is needed in the code, then uncomment it accordingly 
    ''        'Const ForReading = 1, ForWriting = 2, ForAppending = 8

    ''        Const ForAppending = 8
    ''        Const logTrennz As String = " , "
    ''        ' logfile-stream erzeugen
    ''        Dim fs = CreateObject("Scripting.FileSystemObject")

    ''        '' FileNamen zusammenbauen
    ''        'Dim logfileNamePath As String = createLogfileName(rpaFolder, rpaImportfile)

    ''        'Dim logfileOrdner As String = "logfiles"
    ''        'If IsNothing(awinPath) Then
    ''        '    Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    ''        '    awinPath = My.Computer.FileSystem.CombinePath(curUserDir, "VISBO")
    ''        'End If
    ''        'Dim logfilePath As String = My.Computer.FileSystem.CombinePath(awinPath, logfileOrdner)

    ''        '' write logfiles in rpaFolder, if RPA was started
    ''        'If Not IsNothing(rpaFolder) Then
    ''        '    logfilePath = rpaFolder
    ''        'End If

    ''        'Dim logfileName As String = "logfile" & "_" & rpaImportfile & "_" & logDate.Year.ToString & logDate.Month.ToString("0#") & logDate.Day.ToString("0#") & "_" & logDate.TimeOfDay.ToString.Replace(":", "-") & ".txt"
    ''        'Dim logfileNamePath As String = My.Computer.FileSystem.CombinePath(logfilePath, logfileName)
    ''        '' Fragen, ob bereits existiert - eventuell nicht nötig
    ''        'If Not My.Computer.FileSystem.DirectoryExists(logfilePath) Then
    ''        '    My.Computer.FileSystem.CreateDirectory(logfilePath)
    ''        'End If


    ''        ' logfile öffnen
    ''        Dim logf = fs.OpenTextFile(logfileNamePath, ForAppending, True, 0)
    ''        strMeld = "[" & Format(Now, "dd.MM.yyyy hh:mm:ss") & "] " & logTrennz & errorLevel(errLevel) & logTrennz & addOn & logTrennz & text
    ''        logf.writeline(strMeld)
    ''        logf.close()


    ''    Catch ex As Exception
    ''        If awinSettings.englishLanguage Then
    ''            Call MsgBox("ERROR while Writing to logfile" & ex.Message)
    ''        Else
    ''            Call MsgBox("Fehler beim Schreiben ins logfile" & ex.Message)
    ''        End If

    ''    End Try


    ''End Sub



    ''''' <summary>
    ''''' schreibt die Inhalte der Collection als String in das Logfile
    ''''' </summary>
    ''''' <param name="meldungen"></param>
    ''Public Sub logger(ByVal logfileNamePath As String, ByVal errLevel As Integer, ByVal addOn As String, ByVal meldungen As Collection)

    ''    Try
    ''        Dim anzZeilen As Integer = meldungen.Count

    ''        Dim strMeld As String

    ''        ' tk 15.8. in Order to avoid warning statements in Visual Studio 
    ''        ' once this is needed in the code, then uncomment it accordingly 
    ''        'Const ForReading = 1, ForWriting = 2, ForAppending = 8
    ''        Const ForAppending = 8
    ''        Const logTrennz As String = " , "
    ''        ' logfile-stream erzeugen
    ''        Dim fs = CreateObject("Scripting.FileSystemObject")


    ''        '' FileNamen zusammenbauen
    ''        'logfileNamePath = createLogfileName(rpaFolder, rpaImportfile)

    ''        '' FileNamen zusammenbauen
    ''        'Dim logfileOrdner As String = "logfiles"
    ''        'If IsNothing(awinPath) Then
    ''        '    Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    ''        '    awinPath = My.Computer.FileSystem.CombinePath(curUserDir, "VISBO")
    ''        'End If
    ''        'Dim logfilePath As String = My.Computer.FileSystem.CombinePath(awinPath, logfileOrdner)
    ''        'Dim logfileName As String = "logfile" & "_" & logDate.Year.ToString & logDate.Month.ToString("0#") & logDate.Day.ToString("0#") & "_" & logDate.TimeOfDay.ToString.Replace(":", "-") & ".txt"
    ''        'Dim logfileNamePath As String = My.Computer.FileSystem.CombinePath(logfilePath, logfileName)
    ''        '' Fragen, ob bereits existiert - eventuell nicht nötig
    ''        'If Not My.Computer.FileSystem.DirectoryExists(logfilePath) Then
    ''        '    My.Computer.FileSystem.CreateDirectory(logfilePath)
    ''        'End If
    ''        ' logfile öffnen
    ''        Dim logf = fs.OpenTextFile(logfileNamePath, ForAppending, True, 0)

    ''        For i As Integer = 1 To anzZeilen

    ''            Dim text As String = CStr(meldungen.Item(i))
    ''            strMeld = "[" & Format(Now, "dd.MM.yyyy hh:mm:ss") & "] " & logTrennz & errorLevel(errLevel) & logTrennz & addOn & logTrennz & text
    ''            logf.writeline(strMeld)

    ''        Next
    ''        logf.close()
    ''    Catch ex As Exception

    ''    End Try
    ''End Sub

    ''''' <summary>
    ''''' ganz aanlog zu dem anderen logfile Schrieben, nur dass jetzt ein Array von String Werten übergeben wird, der in die einzelnen Spalten kommt 
    ''''' </summary>
    ''''' <param name="text"></param>
    ''Public Sub logger(ByVal logfileNamePath As String, ByVal errLevel As Integer, ByVal addOn As String, ByVal text() As String)

    ''    Try
    ''        Dim anzSpalten As Integer = text.Length

    ''        Dim strMeld As String

    ''        ' tk 15.8. in Order to avoid warning statements in Visual Studio 
    ''        ' once this is needed in the code, then uncomment it accordingly 
    ''        'Const ForReading = 1, ForWriting = 2, ForAppending = 8

    ''        Const ForAppending = 8
    ''        Const logTrennz As String = " , "
    ''        ' logfile-stream erzeugen
    ''        Dim fs = CreateObject("Scripting.FileSystemObject")


    ''        '' FileNamen zusammenbauen
    ''        'Dim logfileOrdner As String = "logfiles"
    ''        'If IsNothing(awinPath) Then
    ''        '    Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    ''        '    awinPath = My.Computer.FileSystem.CombinePath(curUserDir, "VISBO")
    ''        'End If
    ''        'Dim logfilePath As String = My.Computer.FileSystem.CombinePath(awinPath, logfileOrdner)
    ''        'Dim logfileName As String = "logfile" & "_" & logDate.Year.ToString & logDate.Month.ToString("0#") & logDate.Day.ToString("0#") & "_" & logDate.TimeOfDay.ToString.Replace(":", "-") & ".txt"
    ''        'Dim logfileNamePath As String = My.Computer.FileSystem.CombinePath(logfilePath, logfileName)
    ''        '' Fragen, ob bereits existiert - eventuell nicht nötig
    ''        'If Not My.Computer.FileSystem.DirectoryExists(logfilePath) Then
    ''        '    My.Computer.FileSystem.CreateDirectory(logfilePath)
    ''        'End If

    ''        ' Meldungstext zusammensetzen aus dem text-array
    ''        strMeld = "[" & Format(Now, "dd.MM.yyyy hh:mm:ss") & "] " & logTrennz & errorLevel(errLevel) & logTrennz & addOn
    ''        For i As Integer = 0 To anzSpalten - 1
    ''            strMeld = strMeld & logTrennz & CStr(text(i))
    ''        Next

    ''        ' logfile öffnen
    ''        Dim logf = fs.OpenTextFile(logfileNamePath, ForAppending, True, 0)
    ''        logf.writeline(strMeld)
    ''        logf.close()

    ''    Catch ex As Exception

    ''    End Try

    ''End Sub

    ''Public Sub logger(ByVal logfileNamePath As String, ByVal errLevel As Integer, ByVal addOn As String, ByVal text() As String, ByVal values() As Double)


    ''    Try
    ''        Dim anzSpalten As Integer = text.Length
    ''        Dim anzSpaltenValues As Integer = values.Length

    ''        Dim strMeld As String

    ''        ' tk 15.8. in Order to avoid warning statements in Visual Studio 
    ''        ' once this is needed in the code, then uncomment it accordingly 
    ''        'Const ForReading = 1, ForWriting = 2, ForAppending = 8

    ''        Const ForAppending = 8
    ''        Const logTrennz As String = " , "
    ''        ' logfile-stream erzeugen
    ''        Dim fs = CreateObject("Scripting.FileSystemObject")


    ''        '' FileNamen zusammenbauen
    ''        'Dim logfileOrdner As String = "logfiles"
    ''        'If IsNothing(awinPath) Then
    ''        '    Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    ''        '    awinPath = My.Computer.FileSystem.CombinePath(curUserDir, "VISBO")
    ''        'End If
    ''        'Dim logfilePath As String = My.Computer.FileSystem.CombinePath(awinPath, logfileOrdner)
    ''        'Dim logfileName As String = "logfile" & "_" & logDate.Year.ToString & logDate.Month.ToString("0#") & logDate.Day.ToString("0#") & "_" & logDate.TimeOfDay.ToString.Replace(":", "-") & ".txt"
    ''        'Dim logfileNamePath As String = My.Computer.FileSystem.CombinePath(logfilePath, logfileName)
    ''        '' Fragen, ob bereits existiert - eventuell nicht nötig
    ''        'If Not My.Computer.FileSystem.DirectoryExists(logfilePath) Then
    ''        '    My.Computer.FileSystem.CreateDirectory(logfilePath)
    ''        'End If

    ''        ' Meldungstext zusammensetzen aus dem text-array
    ''        strMeld = "[" & Format(Now, "dd.MM.yyyy hh:mm:ss") & "] " & logTrennz & errorLevel(errLevel) & logTrennz & addOn
    ''        For i As Integer = 0 To anzSpalten - 1
    ''            strMeld = strMeld & logTrennz & CStr(text(i))
    ''        Next
    ''        For k As Integer = 0 To anzSpaltenValues - 1
    ''            strMeld = strMeld & logTrennz & Format(values(k), "#,##0.##")
    ''        Next

    ''        ' logfile öffnen
    ''        Dim logf = fs.OpenTextFile(logfileNamePath, ForAppending, True, 0)
    ''        logf.writeline(strMeld)
    ''        logf.close()

    ''    Catch ex As Exception

    ''    End Try

    ''End Sub


    ''Public Sub logger(ByVal logfileNamePath As String, ByVal errLevel As Integer, ByVal addOn As String, ByVal strLog As String)
    ''    Try

    ''        Dim strMeld As String

    ''        ' tk 15.8. in Order to avoid warning statements in Visual Studio 
    ''        ' once this is needed in the code, then uncomment it accordingly 
    ''        'Const ForReading = 1, ForWriting = 2, ForAppending = 8

    ''        Const ForAppending = 8
    ''        Const logTrennz As String = " , "
    ''        ' logfile-stream erzeugen
    ''        Dim fs = CreateObject("Scripting.FileSystemObject")


    ''        '' FileNamen zusammenbauen
    ''        ' Dim logfileNamePath As String = createLogfileName(rpaFolder, rpaImportfile)

    ''        '' FileNamen zusammenbauen
    ''        'Dim logfileOrdner As String = "logfiles"
    ''        'If IsNothing(awinPath) Then
    ''        '    Dim curUserDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    ''        '    awinPath = My.Computer.FileSystem.CombinePath(curUserDir, "VISBO")
    ''        'End If
    ''        'Dim logfilePath As String = My.Computer.FileSystem.CombinePath(awinPath, logfileOrdner)
    ''        'Dim logfileName As String = "logfile" & "_" & logDate.Year.ToString & logDate.Month.ToString("0#") & logDate.Day.ToString("0#") & "_" & logDate.TimeOfDay.ToString.Replace(":", "-") & ".txt"
    ''        'Dim logfileNamePath As String = My.Computer.FileSystem.CombinePath(logfilePath, logfileName)
    ''        '' Fragen, ob bereits existiert - eventuell nicht nötig
    ''        'If Not My.Computer.FileSystem.DirectoryExists(logfilePath) Then
    ''        '    My.Computer.FileSystem.CreateDirectory(logfilePath)
    ''        'End If
    ''        ' logfile öffnen
    ''        Dim logf = fs.OpenTextFile(logfileNamePath, ForAppending, True, 0)
    ''        strMeld = "[" & Format(Now, "dd.MM.yyyy hh:mm:ss") & "] " & logTrennz & errorLevel(errLevel) & logTrennz & addOn & " , " & strLog
    ''        logf.writeline(strMeld)
    ''        logf.close()

    ''    Catch ex As Exception

    ''    End Try
    ''End Sub

End Module
