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
                    My.Computer.FileSystem.MoveFile(fullFileName, newDestination, True)
                    Call logger(ptErrLevel.logInfo, "unknown file / category: ", myName)

                    errMsgCode = New clsErrorCodeMsg
                    result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf _
                                                                                & myName & vbCrLf & " unknown file / category ...", errMsgCode)
            End Select
        Else
            Dim a As String = ""
        End If


    End Sub

    Private Sub watchFolder_Deleted(sender As Object, e As IO.FileSystemEventArgs) Handles watchFolder.Deleted

        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Deleted")
        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "File '" & e.FullPath & "' was deleted at: " & Date.Now().ToLongDateString)
    End Sub

    Private Sub watchFolder_Renamed(sender As Object, e As IO.RenamedEventArgs) Handles watchFolder.Renamed
        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "watchFolder_Renamed")
        'Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", "File '" & e.FullPath & "' was renamed at: " & Date.Now().ToLongDateString)
    End Sub

    Private Sub VisboRPAStart_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        ' now store User Login Data
        My.Settings.userNamePWD = awinSettings.userNamePWD
        ' now cancel User Login Data
        'My.Settings.userNamePWD = ""
        My.Settings.rpaPath = rpaFolder
        ' speichern 
        My.Settings.Save()

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

        ''Set this property to true to start watching
        'watchFolder.EnableRaisingEvents = False

        'Call logger(ptErrLevel.logInfo, "VisboRPA", "Process was stopped!")
        '' now store User Login Data
        'My.Settings.userNamePWD = awinSettings.userNamePWD

        '' now delete User Login Data
        'My.Settings.userNamePWD = ""

        ''now cancel RPAFolder
        'My.Settings.rpaPath = rpaFolder

        '' speichern 
        'My.Settings.Save()
        'Dim err As New clsErrorCodeMsg

        'Dim logoutErfolgreich As Boolean = CType(databaseAcc, DBAccLayer.Request).logout(err)

        'If logoutErfolgreich Then
        '    If awinSettings.visboDebug Then
        '        If awinSettings.englishLanguage Then
        '            Call MsgBox(err.errorMsg & vbCrLf & "User don't have access to a VisboCenter any longer!")
        '        Else
        '            Call MsgBox(err.errorMsg & vbCrLf & "User hat keinen Zugriff mehr zu einem VisboCenter!")
        '        End If
        '    End If
        'End If

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
                ' is there a activePortfolio
                myActivePortfolio = inputvalues.activePortfolio
                configfilesOrdner = inputvalues.VisboConfigFiles
                configfilesOrdner = configfilesOrdner.Replace("\\", "\")


                ' read all files, categorize and verify them  
                msgTxt = "Starting with ..."
                Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)



                visboClient = "VISBO RPA / "
                ' 
                ' startUpRPA  liest orga, appearances und andere Settings - analog awinSetTypen , allerdings nie mit Versuch, etwas von Platte zu lesen ... 
                startup = startUpRPA(inputvalues.VisboCenter, inputvalues.VisboUrl, settingsFolder, inputvalues.proxyURL)

            Else
                startup = False
                ' Exit ! 
                ' read all files, categorize and verify them  
                msgTxt = "Exit - there is no File " & settingJsonFile
                Call logger(ptErrLevel.logError, "VISBO Robotic Process automation", msgTxt)

                ' break the RPA - Service

            End If

            If startup Then

                Call logger(ptErrLevel.logInfo, "VisboRPA: proxyURL", inputvalues.proxyURL)
                Call logger(ptErrLevel.logInfo, "VisboRPA: Visbo Plattform", inputvalues.VisboUrl)
                Call logger(ptErrLevel.logInfo, "VisboRPA: Visbo Center", inputvalues.VisboCenter)
                Call logger(ptErrLevel.logInfo, "VisboRPA: active Portfolio", myActivePortfolio)
                Call logger(ptErrLevel.logInfo, "VisboRPA: Config Files Folder", configfilesOrdner)
                Call logger(ptErrLevel.logInfo, "VisboRPA: RPA Folder", rpaFolder)

                ' Sendet eine Email an den User
                errMsgCode = New clsErrorCodeMsg
                result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & "correct start of the RPA", errMsgCode)
                If Not result Then
                    Call logger(ptErrLevel.logError, "RPA Service- On Start", errMsgCode.errorMsg)
                Else
                    'this is the path we want to monitor
                    watchFolder.Path = rpaFolder

                    'Set this property to true to start watching
                    watchFolder.EnableRaisingEvents = True

                    MyBase.WindowState = FormWindowState.Minimized

                    rpaDir.Text = rpaFolder
                    My.Settings.rpaPath = My.Computer.FileSystem.GetParentPath(rpaFolder)

                    statusMessage.Text = "VisboRPA started successfully..."

                    ' außer stop, darf kein button aktiv sein
                    btn_start.Enabled = False
                    durchsuchen.Enabled = False
                End If


            Else
                msgTxt = "wrong settings - exited without performing jobs ...."
                Call MsgBox(msgTxt)
                ' Console.WriteLine(msgTxt)
                Call logger(ptErrLevel.logInfo, "VISBO Robotic Process automation", msgTxt)
                'errMsgCode = New clsErrorCodeMsg
                'result = CType(databaseAcc, DBAccLayer.Request).sendEmailToUser("VISBO Robotic Process automation" & vbCrLf & msgTxt, errMsgCode)
                'If Not result Then
                '    Call logger(ptErrLevel.logError, "RPA Service- On Start", errMsgCode.errorMsg)
                'End If

            End If


        Catch ex As Exception
            Call logger(ptErrLevel.logError, "VISBO Robotic Process Automation", ex.Message)
        End Try
    End Sub

End Class