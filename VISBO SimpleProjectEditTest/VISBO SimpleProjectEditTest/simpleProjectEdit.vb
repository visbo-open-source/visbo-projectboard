Imports xlns = Microsoft.Office.Interop.Excel
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Newtonsoft.Json
Imports System.IO
Imports DBAccLayer
Imports WebServerAcc
Imports System.Security.Principal

Imports System.Diagnostics


Module simpleProjectEdit

    ''' <summary>
    ''' when called, all awinSetting Variables are set .. 
    ''' </summary>
    ''' <returns></returns>
    Public Function speSetTypen() As Boolean

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


            If My.Settings.awinPath = "" Then
                ' tk 12.12.18 damit wird sichergestellt, dass bei einer Installation die Demo Daten einfach im selben Directory liegen können
                ' im ProjectBoardConfig kann demnach entweder der leere String stehen oder aber ein relativer Pfad, der vom User/Home Directory ausgeht ... 
                'Dim locationOfProjectBoard = My.Computer.FileSystem.GetParentPath(appInstance.ActiveWorkbook.FullName)
                Dim locationOfSPE As String = My.Computer.FileSystem.CurrentDirectory
                locationOfSPE = "C:\VISBO"
                Dim stdConfigDataName As String = "VISBO Config Data"

                awinPath = My.Computer.FileSystem.CombinePath(locationOfSPE, stdConfigDataName)

                If My.Computer.FileSystem.DirectoryExists(awinPath) Then
                    ' alles ok
                Else
                    awinPath = My.Computer.FileSystem.CombinePath(curUserDir, stdConfigDataName)
                    If My.Computer.FileSystem.DirectoryExists(awinPath) Then
                        ' alles ok 
                    End If
                End If
            ElseIf My.Computer.FileSystem.DirectoryExists(My.Settings.awinPath) Then
                awinPath = My.Settings.awinPath
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
                    Call logger(ptErrLevel.logInfo, "speSetTypen", "Synchronized localPath with globalPath")

                Else

                    Call logger(ptErrLevel.logInfo, "speSetTypen", "no Synchronization between localPath and globalPath")

                End If

            End If

            StartofCalendar = StartofCalendar.Date


            'Try
            '    repMessages = XMLImportReportMsg(repMsgFileName, awinSettings.ReportLanguage)
            '    Call setLanguageMessages()
            'Catch ex As Exception

            'End Try



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
            Try
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

                If loginErfolgreich Then

                    ' jetzt muss geprüft werden, ob es mehr als ein zugelassenes VISBO Center gibt , ist dann der Fall wenn es ein # im awinsettings.databaseNAme gibt 
                    Dim listOfVCs As List(Of String) = CType(databaseAcc, DBAccLayer.Request).retrieveVCsForUser(err)

                    If listOfVCs.Count = 1 Then
                        ' alles ok, nimm dieses  VC
                        awinSettings.databaseName = listOfVCs.Item(0)

                        Dim changeOK As Boolean = CType(databaseAcc, DBAccLayer.Request).updateActualVC(awinSettings.databaseName, awinSettings.VCid, err)
                        If Not changeOK Then
                            Call logger(ptErrLevel.logError, "VISBO SPE load", "No access to this VISBO Center ... program ends  ..." & vbCrLf & err.errorMsg)
                            Throw New ArgumentException("No access to this VISBO Center ... program ends  ..." & vbCrLf & err.errorMsg)
                        Else
                            Dim myVC As String = awinSettings.databaseName

                        End If

                    ElseIf listOfVCs.Count > 1 Then
                        ' wähle das gewünschte VC aus
                        Dim chooseVC As New frmSelectOneItem
                        chooseVC.itemsCollection = listOfVCs
                        If chooseVC.ShowDialog = Windows.Forms.DialogResult.OK Then
                            ' alles ok 
                            awinSettings.databaseName = chooseVC.itemList.SelectedItem.ToString
                            Dim changeOK As Boolean = CType(databaseAcc, DBAccLayer.Request).updateActualVC(awinSettings.databaseName, awinSettings.VCid, err)
                            If Not changeOK Then
                                Call logger(ptErrLevel.logError, "VISBO SPE load", "No access to this VISBO Center ... program ends  ..." & vbCrLf & err.errorMsg)
                                Throw New ArgumentException("No access to this VISBO Center ... program ends  ..." & vbCrLf & err.errorMsg)
                            End If
                        Else
                            Throw New ArgumentException("no Selection of VISBO Center ... program ends  ..." & vbCrLf & err.errorMsg)
                        End If

                    Else
                        ' user has no access to any VISBO Center 
                        Call logger(ptErrLevel.logInfo, "Load of Formular", "User has no access to any VISBO Center ... ")
                        Throw New ArgumentException("No access to a VISBO Center ")
                    End If

                Else
                    ' no valid Login
                    Call logger(ptErrLevel.logInfo, "Load of Formular", "No valid Login ... ")
                    'Throw New ArgumentException("No valid Login")
                End If

                If Not loginErfolgreich Then

                    Call logger(ptErrLevel.logInfo, "LOGIN cancelled ...", "", -1)

                    If awinSettings.englishLanguage Then
                        Throw New ArgumentException("LOGIN cancelled ...")
                    Else
                        Throw New ArgumentException("LOGIN abgebrochen ...")
                    End If

                End If

            Catch ex As Exception

            End Try


            '' ur: 10032022: not needed for RPA
            '' Read appearance Definitions
            'appearanceDefinitions.liste = CType(databaseAcc, DBAccLayer.Request).retrieveAppearancesFromDB("", Date.Now, False, err)
            'If IsNothing(appearanceDefinitions.liste) Or appearanceDefinitions.liste.Count > 0 Then
            '    ' user has no access to any VISBO Center 
            '    msgTxt = "No appearance Definitions in VISBO"
            '    Call logger(ptErrLevel.logInfo, "rpaSetTypen", "")
            '    'Throw New ArgumentException(msgTxt)
            'End If
            '
            ' now read Customizations
            ''
            ''
            '' Read Customizations 
            Dim lastReadingCustomization As Date = readCustomizations()

            '
            ' now read Organisation 
            ''
            '' Read Customizations 
            ' muss später erfolgen: Dim lastReadingOrganisation As Date = readOrganisations()


            '
            ' now read customFieldDefinitions; is allowed to be empty
            customFieldDefinitions = CType(databaseAcc, DBAccLayer.Request).retrieveCustomFieldsFromDB(err)

            If IsNothing(customFieldDefinitions) Then
                customFieldDefinitions = New clsCustomFieldDefinitions
                Call logger(ptErrLevel.logInfo, "speSetTypen", "no CustomFieldDefinitions found")
            End If

            '
            ' myCustomUserRole wird by Default auf <Alles> gesetzt 
            '
            '' ur:5.5.22: dies soll durch ServerRechte ersetzt werden
            '
            ' TODO: RestCall-aufsetzen für Abfragen der Rechte zum aktuellen User

            myCustomUserRole = New clsCustomUserRole

            With myCustomUserRole
                .customUserRole = ptCustomUserRoles.Alles
                .specifics = ""
                .userName = dbUsername
            End With

            '' ur: here not necessary
            '' now read Vorlagen - maybe Empty
            'lastReadingProjectTemplates = readProjectTemplates()

            result = True

        Catch ex As Exception

            result = False
            Call logger(ptErrLevel.logError, "speSetTypen", ex.Message)
            Dim msg As String = ""

            If ex.Message.StartsWith("LOGIN cancelled") Or ex.Message.Contains("User") Then
                msg = ex.Message
            Else

            End If

            '??? Throw New ArgumentException(msg)

        End Try

        speSetTypen = result

    End Function

End Module
