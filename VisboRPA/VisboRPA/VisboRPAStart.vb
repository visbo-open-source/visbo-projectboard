Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Newtonsoft.Json
Imports System.IO
Imports DBAccLayer
Imports WebServerAcc
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Security.Principal
Imports System.Diagnostics
Public Class VisboRPAStart

    Private Sub watchFolder_Changed(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Changed

        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Changed")
        ''code here for newly changed file Or directory

        logfileNamePath = createLogfileName(rpaModule1.rpaFolder)

        Call logger(ptErrLevel.logInfo, "WatchFolder_changed", "File '" & e.FullPath & "' was changed at: " & Date.Now().ToLongDateString)

        Dim fullFileName As String = e.FullPath
        Dim myName As String = ""
        Dim rpaCategory As New PTRpa
        Dim result As Boolean = False

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
                        Call logger(ptErrLevel.logInfo, "WatchFolder_changed", "File '" & e.FullPath & "' was imported successfully at: " & Date.Now().ToLongDateString)
                    End If
                End If
            Case ".mpp"

                myName = My.Computer.FileSystem.GetName(fullFileName)

                ' Import Typ ist Microsoft Project File
                rpaCategory = PTRpa.visboMPP

                ' Import wird durchgeführt
                result = importOneProject(fullFileName, rpaCategory, Date.Now())
                If result Then
                    Call logger(ptErrLevel.logInfo, "WatchFolder_changed", "File '" & e.FullPath & "' was imported successfully at: " & Date.Now().ToLongDateString)
                End If
            Case Else

        End Select
    End Sub

    Private Sub watchFolder_Created(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Created

        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Created")
        ''code here for newly changed file Or directory

        logfileNamePath = createLogfileName(rpaModule1.rpaFolder)


        Dim fullFileName As String = e.FullPath
        Dim myName As String = ""
        Dim rpaCategory As New PTRpa
        Dim result As Boolean = False

        ' Completion-File delivered?
        completedOK = LCase(fullFileName).Contains(LCase("Timesheet_completed"))
        If completedOK Then


            Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & e.FullPath & "' was created at: " & Date.Now().ToLongDateString)

            'Einlesen der TimeSheets - Telair
            ' nachsehen ob collect vollständig
            myName = My.Computer.FileSystem.GetName(fullFileName)
            result = processVisboActualData2(myName, myActivePortfolio, collectFolder, Date.Now())
            ' TODO: löschen des Timesheet-compl
            If result Then
                Dim newDestination As String = My.Computer.FileSystem.CombinePath(successFolder, myName)
                My.Computer.FileSystem.MoveFile(myName, newDestination, True)
                Call logger(ptErrLevel.logInfo, "success: ", myName)

                ' wieder in das normale logfile schreiben
                logfileNamePath = createLogfileName(rpaFolder)
                errMsgCode = New clsErrorCodeMsg
                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & myName & ": successful ...", errMsgCode)
            Else
                Dim newDestination As String = My.Computer.FileSystem.CombinePath(failureFolder, myName)
                If My.Computer.FileSystem.FileExists(fullFileName) Then
                    My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                    Call logger(ptErrLevel.logError, "failed: ", fullFileName)
                    Dim logfileName As String = My.Computer.FileSystem.GetName(logfileNamePath)
                    Dim newLog As String = My.Computer.FileSystem.CombinePath(failureFolder, logfileName)
                    My.Computer.FileSystem.MoveFile(logfileNamePath, newLog, True)

                    ' wieder in das normale logfile schreiben
                    logfileNamePath = createLogfileName(rpaFolder)
                    errMsgCode = New clsErrorCodeMsg
                    result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                                                                                & myName & ": with errors ..." & vbCrLf _
                                                                                & "Look for more details in the Failure-Folder", errMsgCode)
                End If
            End If
        End If


        If My.Computer.FileSystem.FileExists(fullFileName) And Not fullFileName.Contains("~$") Then

            Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & e.FullPath & "' was created at: " & Date.Now().ToLongDateString)

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
                            Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & e.FullPath & "' was imported successfully at: " & Date.Now().ToLongDateString)
                        End If
                    End If
                Case ".mpp"

                    myName = My.Computer.FileSystem.GetName(fullFileName)

                    ' Import Typ ist Microsoft Project File
                    rpaCategory = PTRpa.visboMPP

                    ' Import wird durchgeführt
                    result = importOneProject(fullFileName, rpaCategory, Date.Now())
                    If result Then
                        Call logger(ptErrLevel.logInfo, "watchFolder_Created", "File '" & e.FullPath & "' was imported successfully at: " & Date.Now().ToLongDateString)
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


    End Sub

    Private Sub watchFolder_Deleted(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Deleted

        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Deleted")
        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "File '" & e.FullPath & "' was deleted at: " & Date.Now().ToLongDateString)
    End Sub

    Private Sub watchFolder_Renamed(sender As Object, e As IO.RenamedEventArgs) Handles watchFolder.Renamed
        Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Renamed")
        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "File '" & e.FullPath & "' was renamed at: " & Date.Now().ToLongDateString)
    End Sub

    Private Sub VisboRPAStart_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' das folgende darf nur gemacht werden, wenn auch awinsetting.visboserver gilt ... 
        Dim err As New clsErrorCodeMsg

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
                Else
                    myVC = awinSettings.databaseName
                    VCSelection.Text = myVC
                End If

            ElseIf listOfVCs.Count > 1 Then
                VCSelection.Items.Clear()
                For Each vc In listOfVCs
                    VCSelection.Items.Add(vc)
                Next
                If listOfVCs.Contains(myVC) Then
                    VCSelection.Text = myVC
                    awinSettings.databaseName = myVC
                    Dim changeOK As Boolean = CType(databaseAcc, DBAccLayer.Request).updateActualVC(awinSettings.databaseName, awinSettings.VCid, err)
                    If Not changeOK Then
                        Throw New ArgumentException("No access to this VISBO Center ... program ends  ..." & vbCrLf & err.errorMsg)
                    Else
                        awinSettings.databaseName = myVC
                        VCSelection.Text = myVC
                    End If
                Else
                    Call logger(ptErrLevel.logInfo, "Load of Formular", "No access to this VISBO Center '" & myVC & "'... ")
                    awinSettings.databaseName = ""
                    VCSelection.Text = ""
                End If

            Else
                ' user has no access to any VISBO Center 
                Call logger(ptErrLevel.logInfo, "Load of Formular", "User has no access to any VISBO Center ... ")
                Throw New ArgumentException("No access to a VISBO Center ")
            End If

        Else
            ' no valid Login
            Call logger(ptErrLevel.logInfo, "Load of Formular", "No valid Login ... ")
            Throw New ArgumentException("No valid Login")
        End If

        If awinSettings.databaseName <> "" Then
            ' alle möglichen Portfolios anbieten
            Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, err)
            activePortfolioSel.Items.Clear()

            For Each vpf In dbPortfolioNames
                activePortfolioSel.Items.Add(vpf.Key)
            Next
            If dbPortfolioNames.ContainsKey(myActivePortfolio) Then
                activePortfolioSel.Text = myActivePortfolio
            End If
            If dbPortfolioNames.Count = 1 Then
                    activePortfolioSel.Text = dbPortfolioNames.ElementAt(0).Key
                End If
            End If


            'this is the path we want to monitor
            If rpaPath <> "" Then
            If My.Computer.FileSystem.DirectoryExists(rpaPath) Then
                rpaDir.Text = rpaPath
                watchFolder.Path = rpaPath
            End If
        End If

        'Add a list of Filter we want to specify
        'make sure you use OR for each Filter as we need to
        'all of those 


        'Set this property to true to start watching
        watchFolder.EnableRaisingEvents = True
    End Sub

    Private Sub VisboRPAStart_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        'Set this property to true to start watching
        watchFolder.EnableRaisingEvents = False

        Call logger(ptErrLevel.logInfo, "VisboRPA", "Process was stopped!")

        If loginErfolgreich Then

            My.Settings.rpaPath = rpaFolder
            My.Settings.VisboCenter = myVC
            My.Settings.activePortfolio = myActivePortfolio

            My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
            If My.Settings.rememberUserPWD Then
                My.Settings.userNamePWD = awinSettings.userNamePWD
            Else
                'Cancel User Login-Data
                My.Settings.userNamePWD = ""
            End If

            ' speichern 
            My.Settings.Save()
        Else
            ' wenn kein erfolgreicher Login stattgefunden hat
            My.Settings.Reset()
        End If


        Dim err As New clsErrorCodeMsg

        Dim logoutErfolgreich As Boolean = CType(databaseAcc, DBAccLayer.Request).logout(err)

        If logoutErfolgreich Then
            If awinSettings.visboDebug Then
                If awinSettings.englishLanguage Then
                    Call MsgBox(err.errorMsg & vbCrLf & "User don't have access to a VisboCenter any longer!")
                Else
                    Call MsgBox(err.errorMsg & vbCrLf & "User hat keinen Zugriff mehr zu einem VisboCenter!")
                End If
            End If
        End If

    End Sub

    Private Sub btn_start_Click(sender As Object, e As EventArgs) Handles btn_start.Click

        If Not IsNothing(rpaDir.Text) Then
            rpaFolder = rpaDir.Text
        End If

        If My.Computer.FileSystem.DirectoryExists(rpaFolder) Then
                'this is the path we want to monitor
                watchFolder.Path = rpaFolder

                'Set this property to true to start watching
                watchFolder.EnableRaisingEvents = True

            Call startWatching(rpaFolder)
        Else
            statusMessage.Text = "Please choose the RPA-Folder !"
        End If

    End Sub

    Private Sub btn_stop_Click(sender As Object, e As EventArgs) Handles btn_stop.Click


        MyBase.Close()


    End Sub

    Private Sub durchsuchen_Click(sender As Object, e As EventArgs) Handles durchsuchen.Click

        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            rpaDir.Text = FolderBrowserDialog1.SelectedPath
            rpaPath = rpaDir.Text

            rpaFolder = rpaPath

            Call startWatching(rpaFolder)
        End If

    End Sub


    Public Sub startWatching(ByVal rpaFolder As String)

        successFolder = My.Computer.FileSystem.CombinePath(rpaFolder, "success")
        failureFolder = My.Computer.FileSystem.CombinePath(rpaFolder, "failure")
        collectFolder = My.Computer.FileSystem.CombinePath(rpaFolder, "collect")
        logfileFolder = My.Computer.FileSystem.CombinePath(rpaFolder, "logfiles")
        unknownFolder = My.Computer.FileSystem.CombinePath(rpaFolder, "unknown")
        settingsFolder = My.Computer.FileSystem.CombinePath(rpaFolder, "settings")
        configfilesOrdner = settingsFolder
        settingJsonFile = My.Computer.FileSystem.CombinePath(settingsFolder, "rpa_setting.json")


        ' FileNamen für logging zusammenbauen
        logfileNamePath = createLogfileName(rpaFolder, "")


        Try

            Dim anzFiles As Integer = 0

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

            If Not My.Computer.FileSystem.DirectoryExists(settingsFolder) Then
                My.Computer.FileSystem.CreateDirectory(settingsFolder)
            End If


            Dim startup As Boolean = False

            ' Read the Setting-file of RPA
            If My.Computer.FileSystem.FileExists(settingJsonFile) Then
                Dim jsonSetting As String = File.ReadAllText(settingJsonFile)
                inputvalues = JsonConvert.DeserializeObject(Of clsRPASetting)(jsonSetting)
                If Not IsNothing(inputvalues) Then
                    awinSettings.proxyURL = inputvalues.proxyURL
                End If
                ' if there are already definitions from User, take these
                ' else take the defaults
                Dim settingConfigfilesOrdner As String = inputvalues.VisboConfigFiles
                settingConfigfilesOrdner = settingConfigfilesOrdner.Replace("\\", "\")

                If Not IsNothing(settingConfigfilesOrdner) And My.Computer.FileSystem.DirectoryExists(settingConfigfilesOrdner) Then
                    If configfilesOrdner <> settingConfigfilesOrdner Then
                        configfilesOrdner = settingConfigfilesOrdner
                    End If
                Else
                    'ConfigFolder is the settingFolder of rpaPath
                    configfilesOrdner = settingsFolder
                End If


                ' if there are already definitions from User, take these
                ' else take the defaults
                Dim settingPortfolioName As String = inputvalues.activePortfolio

                If IsNothing(myActivePortfolio) Then
                    myActivePortfolio = settingPortfolioName
                Else
                    myActivePortfolio = myActivePortfolio
                End If


                ' read all files, categorize and verify them  
                msgTxt = "Starting with ..."
                Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)

                visboClient = "VISBO RPA / "
                ' 
                ' startUpRPA  liest orga, appearances und andere Settings - analog awinSetTypen , allerdings nie mit Versuch, etwas von Platte zu lesen ... 
                startup = startUpRPA(awinSettings.databaseName, awinSettings.databaseURL, settingsFolder, awinSettings.proxyURL)

            Else
                ' read all files, categorize and verify them  
                msgTxt = "Starting with ..."
                Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)


                If Not (IsNothing(awinSettings.databaseName) Or IsNothing(awinSettings.databaseURL) Or IsNothing(settingsFolder)) Then
                    ' kein file rpa_setting.json vorhanden
                    msgTxt = "default settings will  be used. For more details have a look at the logfiles ...." & vbCrLf & rpaFolder & "\logfiles"
                    'Call MsgBox(msgTxt)
                    ' Console.WriteLine(msgTxt)
                    Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)

                    visboClient = "VISBO RPA / "
                    ' 
                    ' startUpRPA  liest orga, appearances und andere Settings - analog awinSetTypen , allerdings nie mit Versuch, etwas von Platte zu lesen ... 
                    startup = startUpRPA(awinSettings.databaseName, awinSettings.databaseURL, settingsFolder, awinSettings.proxyURL)
                End If


            End If


            If Not startup Then

                ' Exit ! 
                ' read all files, categorize and verify them  
                msgTxt = "Exit - Error starting the VisboRPA"
                Call logger(ptErrLevel.logError, "VISBO Robotic Process automation", msgTxt)

                ' break the RPA - Service

            End If

            If startup Then

                Call logger(ptErrLevel.logInfo, "VisboRPA: proxyURL", awinSettings.proxyURL)
                Call logger(ptErrLevel.logInfo, "VisboRPA: Visbo Plattform", awinSettings.databaseURL)
                Call logger(ptErrLevel.logInfo, "VisboRPA: Visbo Center", awinSettings.databaseName)
                Call logger(ptErrLevel.logInfo, "VisboRPA: active Portfolio", myActivePortfolio)
                Call logger(ptErrLevel.logInfo, "VisboRPA: Config Files Folder", configfilesOrdner)
                Call logger(ptErrLevel.logInfo, "VisboRPA: RPA Folder", rpaFolder)

                ' Email soll nicht an den User gesendet werden bei erfolgreichem Start
                ' Sendet eine Email an den User
                'errMsgCode = New clsErrorCodeMsg
                'result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & "correct start of the RPA", errMsgCode)
                'If Not result Then
                '    Call logger(ptErrLevel.logError, "RPA Service - On Start", errMsgCode.errorMsg)
                'Else
                'this is the path we want to monitor
                watchFolder.Path = rpaFolder

                'Set this property to true to start watching
                watchFolder.EnableRaisingEvents = True

                MyBase.WindowState = FormWindowState.Minimized

                'verwendete Definitionen nochmals eintragen
                rpaDir.Text = rpaFolder
                VCSelection.Text = myVC
                activePortfolioSel.Text = myActivePortfolio

                My.Settings.rpaPath = My.Computer.FileSystem.GetParentPath(rpaFolder)

                statusMessage.Text = "VisboRPA started successfully..."

                ' außer stop, darf kein button aktiv sein
                btn_start.Enabled = False
                    durchsuchen.Enabled = False
                'End If


            Else
                msgTxt = "default settings will  be used. For more details have a look at the logfiles ...." & vbCrLf & rpaFolder
                Call MsgBox(msgTxt)
                ' Console.WriteLine(msgTxt)
                Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)


            End If


        Catch ex As Exception
            Call logger(ptErrLevel.logError, "VISBO Robotic Process Automation", ex.Message)
        End Try
    End Sub

    Private Sub VCSelection_SelectedIndexChanged(sender As Object, e As EventArgs) Handles VCSelection.SelectedIndexChanged
        Dim errmsg As New clsErrorCodeMsg

        myVC = VCSelection.Text
        awinSettings.databaseName = myVC
        Dim changeOK As Boolean = CType(databaseAcc, DBAccLayer.Request).updateActualVC(awinSettings.databaseName, awinSettings.VCid, errmsg)
        If Not changeOK Then
            Call logger(ptErrLevel.logError, "VCSelection", "No access to this VISBO Center ... program ends  ..." & vbCrLf & errmsg.errorMsg)
        Else
            awinSettings.databaseName = myVC
            VCSelection.Text = myVC
        End If

        ' alle möglichen Portfolios anbieten
        Dim dbPortfolioNames As SortedList(Of String, String) = CType(databaseAcc, DBAccLayer.Request).retrievePortfolioNamesFromDB(Date.Now, errmsg)
        activePortfolioSel.Items.Clear()
        For Each vpf In dbPortfolioNames
            activePortfolioSel.Items.Add(vpf.Key)
        Next
        If dbPortfolioNames.Count = 1 Then
            activePortfolioSel.Text = dbPortfolioNames.ElementAt(0).Key
        Else
            If dbPortfolioNames.ContainsKey(myVC) Then
                activePortfolioSel.Text = myVC
            Else
                activePortfolioSel.Text = ""
            End If
        End If
    End Sub

    Private Sub activePortfolioSel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles activePortfolioSel.SelectedIndexChanged
        myActivePortfolio = activePortfolioSel.Text

    End Sub
End Class