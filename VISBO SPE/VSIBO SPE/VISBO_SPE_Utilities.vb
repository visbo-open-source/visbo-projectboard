Imports xlns = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic
Imports Newtonsoft.Json
Imports System.IO
Imports DBAccLayer
Imports WebServerAcc
Imports System.Security.Principal

Imports System.Diagnostics


Module VISBO_SPE_Utilities

    Public spe_vpid As String = ""
    Public spe_vpvid As String = ""
    Public spe_ott As String = ""

    ''' <summary>
    ''' when called, all awinSetting Variables are set .. 
    ''' </summary>
    ''' <returns></returns>
    Public Function speSetTypen(ByVal oneTimeToken As String) As Boolean

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

            'arrWsNames(4) = "Einstellungen"                ' in Customization File 
            '' arrWsNames(5) = 
            arrWsNames(ptTables.meRC) = "meRC"                          ' Edit Ressourcen
            arrWsNames(6) = "meTE"                          ' Edit Termine
            'arrWsNames(7) = "Darstellungsklassen"           ' wird in awinsettypen hinter MPT kopiert; nimmt für die Laufzeit die Darstellungsklassen auf 
            'arrWsNames(8) = "Phasen-Mappings"               ' in Customization
            arrWsNames(9) = "meAT"                          ' Edit Attribute 
            'arrWsNames(10) = "Meilenstein-Mappings"         ' in Customization
            '' arrWsNames(11) = 
            arrWsNames(ptTables.meCharts) = "meCharts"                     ' Massen-Edit Charts 
            arrWsNames(ptTables.mptPfCharts) = "mptPfCharts"                     ' vorbereitet: Portfolio Charts 
            arrWsNames(ptTables.mptPrCharts) = "mptPrCharts"                     ' vorbereitet: Projekt Charts 
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
                        loginErfolgreich = logInToMongoDB(True, oneTimeToken)
                    End If

                Else
                    If Not loginErfolgreich Then
                        loginErfolgreich = logInToMongoDB(True, oneTimeToken)
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
                        If chooseVC.ShowDialog = DialogResult.OK Then
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
                Throw New ArgumentException(ex.Message)
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

            Dim lastReadingOrganisation As Date = readOrganisations()


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

            ' ur: 20220511: eventuell nicht benötigt

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



    ''' <summary>
    ''' es werden nur Projekte an MassEdit übergeben ... sollten Summary Projekte in der Selection sein, werden die erst durch ihre Projekte, die im Show sind, ersetzt 
    ''' </summary>
    ''' <param name="meModus"></param>
    Public Sub massEditRcTeAt(ByVal meModus As ptModus)
        Dim todoListe As New Collection
        Dim projektTodoliste As New Collection
        Dim outputFenster As New frmOutputWindow
        Dim outputCollection As New Collection
        Dim outPutLine As String = ""


        '' now set visbozustaende
        '' necessary to know whether roles or cost need to be shown in building the forms to select roles , skills and costs 
        visboZustaende.projectBoardMode = meModus


        enableOnUpdate = False

        If ShowProjekte.Count >= 0 Then

            Call logger(ptErrLevel.logInfo, "massEditRcTeAt", "Projekte: " & ShowProjekte.Count)

            ' neue Methode 
            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                If Not todoListe.Contains(kvp.Key) Then
                    todoListe.Add(kvp.Key, kvp.Key)
                End If
            Next


            ' check, ob wirklich alle Projekte editiert werden sollen ... 
            If todoListe.Count = ShowProjekte.Count And todoListe.Count > 30 Then
                Dim yesNo As Integer
                yesNo = MsgBox("Wollen Sie wirklich alle Projekte editieren?", MsgBoxStyle.YesNo)
                If yesNo = MsgBoxResult.No Then
                    enableOnUpdate = True
                    Exit Sub
                End If
            End If



            If todoListe.Count >= 0 Then

                ' jetzt muss ggf noch showrangeLeft und showrangeRight gesetzt werden  

                ' Call enableControls(meModus)

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

                            '' ur:220506:  Call awinShowtimezone(showRangeLeft, showRangeRight, True)
                        Else
                            ' beim alten ShowRangeLeft lassen, wenn es Überlappungen gibt ..
                            Dim newLeft As Integer = ShowProjekte.getMinMonthColumn(todoListe)
                            Dim newRight As Integer = ShowProjekte.getMaxMonthColumn(todoListe)

                            If newLeft >= showRangeRight Or newRight <= showRangeLeft Then
                                ' neu bestimmen 
                                '' ur:220506:  Call awinShowtimezone(showRangeLeft, showRangeRight, False)

                                showRangeLeft = ShowProjekte.getMinMonthColumn(todoListe)
                                showRangeRight = ShowProjekte.getMaxMonthColumn(todoListe)

                                '' ur:220506:  Call awinShowtimezone(showRangeLeft, showRangeRight, True)

                            End If
                        End If

                        '' tk 15.2.19 Portfolio Manager darf Summary-Projekte bearbeiten , um sie dann als Vorgaben speichern zu können 
                        '' das wird in der Funktion substituteListeByPVnameIDs geregelt .. 
                        projektTodoliste = substituteListeByPVNameIDs(todoListe)

                        ' jetzt aufbauen der dbCacheProjekte, names are pvnames
                        Call buildCacheProjekte(projektTodoliste, namesArePvNames:=True)

                        Call logger(ptErrLevel.logInfo, "massEditRcTeAt", "before writeOnlineMassEditRessCost: " & showRangeLeft & " ,  " & showRangeRight & " ,  " & meModus)

                        Call writeOnlineMassEditRessCost(projektTodoliste, showRangeLeft, showRangeRight, meModus)


                    ElseIf meModus = ptModus.massEditTermine Then
                        ' tk 15.2.19 Portfolio Manager darf Summary-Projekte bearbeiten , um sie dann als Vorgaben speichern zu können 
                        ' das wird in der Funktion substituteListeByPVnameIDs geregelt .. 
                        projektTodoliste = substituteListeByPVNameIDs(todoListe)

                        ' jetzt aufbauen der dbCacheProjekte, names are pvnames
                        Call buildCacheProjekte(projektTodoliste, namesArePvNames:=True)

                        Call logger(ptErrLevel.logInfo, "massEditRcTeAt", "before writeOnlineMassEditRessCost: " & showRangeLeft & " ,  " & showRangeRight & " ,  " & meModus)

                        Call writeOnlineMassEditTermineSPE(projektTodoliste)

                    ElseIf meModus = ptModus.massEditAttribute Then
                        ' tk 15.2.19 Portfolio Manager darf Summary-Projekte bearbeiten , um sie dann als Vorgaben speichern zu können 
                        ' das wird in der Funktion substituteListeByPVnameIDs geregelt .. 
                        projektTodoliste = substituteListeByPVNameIDs(todoListe)

                        ' jetzt aufbauen der dbCacheProjekte, names are pNames
                        Call buildCacheProjekte(todoListe, namesArePvNames:=False)

                        Call logger(ptErrLevel.logInfo, "massEditRcTeAt", "before writeOnlineMassEditRessCost: " & showRangeLeft & " ,  " & showRangeRight & " ,  " & meModus)

                        Call writeOnlineMassEditAttribute(projektTodoliste)
                    Else
                        Exit Sub
                    End If

                    'appInstance.EnableEvents = True



                    'Try

                    '    If Not IsNothing(projectboardWindows(PTwindows.mpt)) Then
                    '        projectboardWindows(PTwindows.massEdit) = projectboardWindows(PTwindows.mpt).NewWindow
                    '    Else
                    '        projectboardWindows(PTwindows.massEdit) = appInstance.ActiveWindow.NewWindow
                    '    End If

                    'Catch ex As Exception
                    '    projectboardWindows(PTwindows.massEdit) = appInstance.ActiveWindow.NewWindow
                    'End Try

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


                    'With projectboardWindows(PTwindows.massEdit)
                    With appInstance.ActiveWindow

                        Try
                            .FreezePanes = False
                            .Split = False

                            If (meModus = ptModus.massEditRessSkills Or meModus = ptModus.massEditCosts) Then

                                If awinSettings.meExtendedColumnsView = True Then
                                    If ShowProjekte.Count = 1 Then
                                        .SplitRow = 1
                                        .SplitColumn = 5
                                        .FreezePanes = True
                                    Else
                                        .SplitRow = 1
                                        .SplitColumn = 8
                                        .FreezePanes = True
                                    End If
                                Else
                                    If ShowProjekte.Count = 1 Then
                                        .SplitRow = 1
                                        .SplitColumn = 4
                                        .FreezePanes = True
                                    Else
                                        .SplitRow = 1
                                        .SplitColumn = 7
                                        .FreezePanes = True
                                    End If
                                End If
                                .DisplayHeadings = False

                            ElseIf meModus = ptModus.massEditTermine Then
                                If ShowProjekte.Count = 1 Then
                                    .SplitRow = 1
                                    .SplitColumn = 3
                                    .FreezePanes = True
                                Else
                                    .SplitRow = 1
                                    .SplitColumn = 6
                                    .FreezePanes = True
                                End If

                            ElseIf meModus = ptModus.massEditAttribute Then
                                If ShowProjekte.Count = 1 Then
                                    .SplitRow = 1
                                    .SplitColumn = 2
                                    .FreezePanes = True
                                Else
                                    .SplitRow = 1
                                    .SplitColumn = 5
                                    .FreezePanes = True
                                End If
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
                'If awinSettings.englishLanguage Then
                '    Call MsgBox("no projects apply to criterias ...")
                'Else
                '    Call MsgBox("Es gibt keine Projekte, die zu der Auswahl passen ...")
                'End If
            End If


        Else

            enableOnUpdate = True
            If appInstance.EnableEvents = False Then
                appInstance.EnableEvents = True
            End If
            'If awinSettings.englishLanguage Then
            '    Call MsgBox("no projects loaded ...")
            'Else
            '    Call MsgBox("Es sind keine Projekte geladen ...")
            'End If

        End If


        appInstance.ScreenUpdating = True
        If appInstance.ScreenUpdating = False Then
            appInstance.ScreenUpdating = True
        End If


    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="todoListe"></param>
    Public Sub writeOnlineMassEditTermineSPE(ByVal todoListe As Collection)

        Dim err As New clsErrorCodeMsg

        ' wieviele Spalten werden hier angezeigt ... 
        Dim anzSpalten As Integer = 12
        If awinSettings.enableInvoices Then
            anzSpalten = 16
        End If

        'If todoListe.Count = 0 Then


        '    If awinSettings.englishLanguage Then
        '        Call MsgBox("no projects for mass-edit available ..")
        '    Else
        '        Call MsgBox("keine Projekte für den Massen-Edit vorhanden ..")
        '    End If

        '    Exit Sub
        'End If

        Try

            appInstance.EnableEvents = False

            ' jetzt die selectedProjekte Liste zurücksetzen ... ohne die currentConstellation zu verändern ...
            selectedProjekte.Clear()

            Dim currentWS As Excel.Worksheet = Nothing
            Dim currentWB As Excel.Workbook
            Dim startDateColumn As Integer = 5
            Dim tmpName As String


            ' hier muss jetzt das entsprechende File aufgemacht werden ...
            ' das File 
            Try
                currentWB = CType(appInstance.Workbooks.Item(myProjektTafel), Excel.Workbook)
                currentWS = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.meTE)), Excel.Worksheet)

                Try
                    ' off setzen des AutoFilter Modus ... 
                    If CType(currentWS, Excel.Worksheet).AutoFilterMode = True Then
                        'CType(CType(currentWS, Excel.Worksheet).Cells(1, 1), Excel.Range).Select()
                        CType(currentWS, Excel.Worksheet).Cells(1, 1).AutoFilter()
                    End If
                Catch ex As Exception

                End Try

                ' braucht man eigentlich nicht mehr, aber sicher ist sicher ...
                Try
                    With currentWS
                        Dim zRange As Excel.Range = CType(.Range(.Cells(2, 1), .Cells(visboZustaende.meMaxZeile, visboZustaende.meColED)), Excel.Range)
                        zRange.Clear()
                    End With
                Catch ex As Exception

                End Try


            Catch ex As Exception
                Call MsgBox("es gibt Probleme mit dem Mass-Edit Termine Worksheet ...")
                appInstance.EnableEvents = True
                Exit Sub
            End Try


            ' jetzt schreiben der ersten Zeile 
            Dim zeile As Integer = 1
            Dim spalte As Integer = 1

            Dim startSpalteDaten As Integer = 4
            'Dim roleCostNames As Excel.Range = Nothing
            Dim datesInput As Excel.Range = Nothing

            tmpName = ""

            ' Schreiben der Überschriften
            With CType(currentWS, Excel.Worksheet)

                If .ProtectContents Then
                    .Unprotect(Password:="x")
                End If


                If awinSettings.englishLanguage Then
                    CType(.Cells(1, 1), Excel.Range).Value = "Project-Nr"
                    CType(.Cells(1, 2), Excel.Range).Value = "Project-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Variant-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Element-Name"
                    CType(.Cells(1, 5), Excel.Range).Value = "Start-Date"
                    CType(.Cells(1, 6), Excel.Range).Value = "End-Date"
                    CType(.Cells(1, 7), Excel.Range).Value = "Trafficlight"
                    CType(.Cells(1, 8), Excel.Range).Value = "Explanation"
                    CType(.Cells(1, 9), Excel.Range).Value = "Deliverables"
                    CType(.Cells(1, 10), Excel.Range).Value = "Responsible"
                    CType(.Cells(1, 11), Excel.Range).Value = "% Done"
                    CType(.Cells(1, 12), Excel.Range).Value = "folder/document Link"
                    If awinSettings.enableInvoices Then
                        CType(.Cells(1, 13), Excel.Range).Value = "Invoice Value"
                        CType(.Cells(1, 14), Excel.Range).Value = "Term of payment"
                        CType(.Cells(1, 15), Excel.Range).Value = "Penalty Value"
                        CType(.Cells(1, 16), Excel.Range).Value = "Penalty Date"
                    End If



                Else
                    CType(.Cells(1, 1), Excel.Range).Value = "Projekt-Nummer"
                    CType(.Cells(1, 2), Excel.Range).Value = "Projekt-Name"
                    CType(.Cells(1, 3), Excel.Range).Value = "Varianten-Name"
                    CType(.Cells(1, 4), Excel.Range).Value = "Element-Name"
                    CType(.Cells(1, 5), Excel.Range).Value = "Start-Datum"
                    CType(.Cells(1, 6), Excel.Range).Value = "End-Datum"
                    CType(.Cells(1, 7), Excel.Range).Value = "Ampel"
                    CType(.Cells(1, 8), Excel.Range).Value = "Erläuterung"
                    CType(.Cells(1, 9), Excel.Range).Value = "Lieferumfänge"
                    CType(.Cells(1, 10), Excel.Range).Value = "Verantwortlich"
                    CType(.Cells(1, 11), Excel.Range).Value = "% abgeschlossen"
                    CType(.Cells(1, 12), Excel.Range).Value = "Link zum Dokument/Ordner"
                    If awinSettings.enableInvoices Then
                        CType(.Cells(1, 13), Excel.Range).Value = "Rechnungs-Betrag"
                        CType(.Cells(1, 14), Excel.Range).Value = "Zahlungsziel"
                        CType(.Cells(1, 15), Excel.Range).Value = "Vertrags-Strafe"
                        CType(.Cells(1, 16), Excel.Range).Value = "Datum Vertrags-Strafe"
                    End If
                End If

                ' das Erscheinungsbild der Zeile 1 bestimmen  
                Call massEditZeile1Appearance(ptTables.meTE)


            End With


            zeile = 2


            For Each pvName As String In todoListe

                Dim hproj As clsProjekt = Nothing
                If AlleProjekte.Containskey(pvName) Then
                    hproj = AlleProjekte.getProject(pvName)
                End If

                If Not IsNothing(hproj) Then

                    ' ist das Projekt geschützt ? 
                    ' wenn nein, dann temporär schützen 
                    Dim protectionText As String = ""
                    'Dim wpItem As clsWriteProtectionItem
                    Dim isProtectedbyOthers As Boolean
                    If myCustomUserRole.customUserRole = ptCustomUserRoles.OrgaAdmin Then
                        isProtectedbyOthers = True

                        protectionText = "Orga-Admin kann Daten nur sehen, nicht ändern ...  "
                        If awinSettings.englishLanguage Then
                            protectionText = "Orga-Admin may only view data ..."
                        End If
                    Else
                        ''isProtectedbyOthers = Not tryToprotectProjectforMe(hproj.name, hproj.variantName)
                        'If isProtectedbyOthers Then

                        '    protectionText = writeProtections.getProtectionText(calcProjektKey(hproj.name, hproj.variantName))

                        'End If
                    End If


                    ' jetzt wird für jedes Element in der Hierarchy eine Zeile rausgeschrieben 
                    ' das ist jetzt die rootphase-NameID
                    Dim curElemID As String = rootPhaseName
                    Dim indentLevel As Integer = 0
                    ' abbruchlevel gibt, wo die Funktion getNextIdOfId aufhört: erst an der Rootphase(=0) oder beim Element 
                    Dim abbruchLevel As Integer = 0
                    Dim indentOffset As Integer = 1

                    ' jetzt wird die Hierarchy abgeklappert .. beginnend mit dem ersten Element, der RootPhase
                    Do While curElemID <> ""

                        Dim cPhase As clsPhase = Nothing
                        Dim cMilestone As clsMeilenstein = Nothing
                        Dim isMilestone As Boolean = elemIDIstMeilenstein(curElemID)


                        ' unabhängig davon, ob es sich um einen Meilenstein oder eine Phase handelt .. 
                        ' Business-Unit , Projekt-Name, Varianten-Name müssen geschrieben werden 
                        ' Business-Unit
                        CType(currentWS.Cells(zeile, 1), Excel.Range).Value = hproj.kundenNummer
                        CType(currentWS.Cells(zeile, 1), Excel.Range).Locked = True
                        'CType(currentWS.Cells(zeile, 1), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray

                        ' 
                        ' Projekt-Name
                        CType(currentWS.Cells(zeile, 2), Excel.Range).Value = hproj.name
                        CType(currentWS.Cells(zeile, 2), Excel.Range).Locked = True
                        'CType(currentWS.Cells(zeile, 2), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray

                        ' geschützt oder nicht geschützt ? 
                        Dim currentCell As Excel.Range = CType(currentWS.Cells(zeile, 2), Excel.Range)
                        If isProtectedbyOthers Then

                            If isProtectedbyOthers Then
                                currentCell.Font.Color = awinSettings.protectedByOtherColor
                            End If

                            ' Kommentare löschen
                            currentCell.ClearComments()

                            currentCell.AddComment(Text:=protectionText)
                            currentCell.Comment.Visible = False

                        End If

                        ' Varianten-Name
                        CType(currentWS.Cells(zeile, 3), Excel.Range).Value = hproj.variantName
                        CType(currentWS.Cells(zeile, 3), Excel.Range).Locked = True
                        'CType(currentWS.Cells(zeile, 3), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray

                        ' 
                        ' jetzt kommen die Milestone bzw Phase-abhängigen Elemente  
                        If isMilestone Then


                            cMilestone = hproj.getMilestoneByID(curElemID)
                            Dim msName As String = cMilestone.name
                            Dim msNameID As String = cMilestone.nameID

                            ' schreibe den Meilenstein

                            ' Element-Name Meilenstein bzw. Phase inkl Indentlevel schreiben 
                            CType(currentWS.Cells(zeile, 4), Excel.Range).Value = cMilestone.name
                            CType(currentWS.Cells(zeile, 4), Excel.Range).IndentLevel = indentLevel
                            CType(currentWS.Cells(zeile, 4), Excel.Range).Locked = True
                            'CType(currentWS.Cells(zeile, 4), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray

                            ' jetzt die Kommentare schreiben 
                            CType(currentWS.Cells(zeile, 4), Excel.Range).ClearComments()

                            ' jetzt muss die genaue ID reingeschrieben werden
                            CType(currentWS.Cells(zeile, 4), Excel.Range).AddComment(Text:=msNameID)
                            CType(currentWS.Cells(zeile, 4), Excel.Range).Comment.Visible = False


                            ' Startdatum, gibt es bei Meilensteinen nicht, deswegen sperren  
                            CType(currentWS.Cells(zeile, 5), Excel.Range).Value = ""
                            CType(currentWS.Cells(zeile, 5), Excel.Range).Locked = True
                            'CType(currentWS.Cells(zeile, 5), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray

                            Dim isPastElement As Boolean = (DateDiff(DateInterval.Day, hproj.actualDataUntil, cMilestone.getDate) <= 0) And (cMilestone.percentDone = 1)

                            ' Ende-Datum 
                            CType(currentWS.Cells(zeile, 6), Excel.Range).Value = cMilestone.getDate.ToShortDateString
                            If isPastElement Then
                                ' Sperren ...
                                CType(currentWS.Cells(zeile, 5), Excel.Range).Interior.Color = XlRgbColor.rgbLightGrey
                                CType(currentWS.Cells(zeile, 6), Excel.Range).Interior.Color = XlRgbColor.rgbLightGrey
                                CType(currentWS.Cells(zeile, 6), Excel.Range).Font.Color = XlRgbColor.rgbBlack

                                CType(currentWS.Cells(zeile, 6), Excel.Range).Locked = True
                                'CType(currentWS.Cells(zeile, 6), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                            Else
                                If isProtectedbyOthers Then
                                    CType(currentWS.Cells(zeile, 6), Excel.Range).Locked = True
                                    'CType(currentWS.Cells(zeile, 6), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                Else
                                    CType(currentWS.Cells(zeile, 6), Excel.Range).Locked = False
                                End If

                            End If

                            ' Ampel-Farbe
                            CType(currentWS.Cells(zeile, 7), Excel.Range).Value = cMilestone.ampelStatus
                            If isProtectedbyOthers Then
                                CType(currentWS.Cells(zeile, 7), Excel.Range).Locked = True
                                'CType(currentWS.Cells(zeile, 7), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                            Else
                                CType(currentWS.Cells(zeile, 7), Excel.Range).Locked = False
                            End If



                            If cMilestone.ampelStatus = 1 Then
                                CType(currentWS.Cells(zeile, 7), Excel.Range).Interior.Color = visboFarbeGreen
                            ElseIf cMilestone.ampelStatus = 2 Then
                                CType(currentWS.Cells(zeile, 7), Excel.Range).Interior.Color = visboFarbeYellow
                            ElseIf cMilestone.ampelStatus = 3 Then
                                CType(currentWS.Cells(zeile, 7), Excel.Range).Interior.Color = visboFarbeRed
                            Else
                                ' keine Farbe 
                                'CType(currentWS.Cells(zeile, 7), Excel.Range).Interior.Color = visboFarbeNone
                            End If


                            ' Ampel-Erläuterung
                            CType(currentWS.Cells(zeile, 8), Excel.Range).Value = cMilestone.ampelErlaeuterung
                            If isProtectedbyOthers Then
                                CType(currentWS.Cells(zeile, 8), Excel.Range).Locked = True
                                'CType(currentWS.Cells(zeile, 8), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                            Else
                                CType(currentWS.Cells(zeile, 8), Excel.Range).Locked = False
                            End If


                            ' Lieferumfänge
                            CType(currentWS.Cells(zeile, 9), Excel.Range).Value = cMilestone.getAllDeliverables(vbLf)
                            If isProtectedbyOthers Then
                                CType(currentWS.Cells(zeile, 9), Excel.Range).Locked = True
                                'CType(currentWS.Cells(zeile, 9), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                            Else
                                CType(currentWS.Cells(zeile, 9), Excel.Range).Locked = False
                            End If



                            ' wer ist verantwortlich
                            CType(currentWS.Cells(zeile, 10), Excel.Range).Value = cMilestone.verantwortlich
                            If isProtectedbyOthers Then
                                CType(currentWS.Cells(zeile, 10), Excel.Range).Locked = True
                                'CType(currentWS.Cells(zeile, 10), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                            Else
                                CType(currentWS.Cells(zeile, 10), Excel.Range).Locked = False
                            End If


                            ' wieviel ist erledigt ? 
                            CType(currentWS.Cells(zeile, 11), Excel.Range).Value = cMilestone.percentDone.ToString("0#%")
                            If isProtectedbyOthers Then
                                CType(currentWS.Cells(zeile, 11), Excel.Range).Locked = True
                                'CType(currentWS.Cells(zeile, 11), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                            Else
                                CType(currentWS.Cells(zeile, 11), Excel.Range).Locked = False
                            End If


                            ' der Dokumenten Link 
                            CType(currentWS.Cells(zeile, 12), Excel.Range).Value = cMilestone.DocURL
                            If isProtectedbyOthers Then
                                CType(currentWS.Cells(zeile, 12), Excel.Range).Locked = True
                                'CType(currentWS.Cells(zeile, 12), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                            Else
                                CType(currentWS.Cells(zeile, 12), Excel.Range).Locked = False
                            End If

                            ' wenn Meilensteine Invoices / Penalties haben können 
                            If awinSettings.enableInvoices Then
                                ' der Rechnungsbetrag und Zahlungsziel 
                                If Not IsNothing(cMilestone.invoice) Then
                                    If cMilestone.invoice.Key > 0 Then
                                        CType(currentWS.Cells(zeile, 13), Excel.Range).Value = cMilestone.invoice.Key
                                        CType(currentWS.Cells(zeile, 14), Excel.Range).Value = cMilestone.invoice.Value
                                    End If
                                End If

                                ' die Penalty und das Penalty Date
                                If Not IsNothing(cMilestone.penalty) Then
                                    If cMilestone.penalty.Value > 0 Then
                                        CType(currentWS.Cells(zeile, 15), Excel.Range).Value = cMilestone.penalty.Value
                                        CType(currentWS.Cells(zeile, 16), Excel.Range).Value = cMilestone.penalty.Key
                                    End If
                                End If

                                If isProtectedbyOthers Then
                                    CType(currentWS.Cells(zeile, 13), Excel.Range).Locked = True
                                    CType(currentWS.Cells(zeile, 14), Excel.Range).Locked = True
                                    CType(currentWS.Cells(zeile, 15), Excel.Range).Locked = True
                                    CType(currentWS.Cells(zeile, 16), Excel.Range).Locked = True
                                Else
                                    CType(currentWS.Cells(zeile, 13), Excel.Range).Locked = False
                                    CType(currentWS.Cells(zeile, 14), Excel.Range).Locked = False
                                    CType(currentWS.Cells(zeile, 15), Excel.Range).Locked = False
                                    CType(currentWS.Cells(zeile, 16), Excel.Range).Locked = False
                                End If


                            End If



                        Else

                            cPhase = hproj.getPhaseByID(curElemID)
                            Dim phName As String = cPhase.name
                            Dim phNameID As String = cPhase.nameID

                            ' schreibe die Phase
                            With CType(currentWS, Excel.Worksheet)

                                ' Element-Name Meilenstein bzw. Phase
                                CType(.Cells(zeile, 4), Excel.Range).Value = cPhase.name
                                CType(.Cells(zeile, 4), Excel.Range).IndentLevel = indentLevel

                                CType(currentWS.Cells(zeile, 4), Excel.Range).Locked = True
                                'CType(.Cells(zeile, 4), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray

                                ' jetzt die Kommentare schreiben 
                                CType(currentWS.Cells(zeile, 4), Excel.Range).ClearComments()

                                ' jetzt muss die genaue ID reingeschrieben werden
                                CType(currentWS.Cells(zeile, 4), Excel.Range).AddComment(Text:=phNameID)
                                CType(currentWS.Cells(zeile, 4), Excel.Range).Comment.Visible = False


                                ' Startdatum 
                                CType(.Cells(zeile, 5), Excel.Range).Value = cPhase.getStartDate.ToShortDateString
                                If DateDiff(DateInterval.Day, hproj.actualDataUntil, cPhase.getStartDate) <= 0 Then
                                    ' Sperren ...
                                    CType(currentWS.Cells(zeile, 5), Excel.Range).Interior.Color = XlRgbColor.rgbLightGrey
                                    CType(currentWS.Cells(zeile, 5), Excel.Range).Font.Color = XlRgbColor.rgbBlack
                                    CType(currentWS.Cells(zeile, 5), Excel.Range).Locked = True
                                Else
                                    If isProtectedbyOthers Then
                                        CType(currentWS.Cells(zeile, 5), Excel.Range).Locked = True
                                        'CType(currentWS.Cells(zeile, 5), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                    Else
                                        CType(currentWS.Cells(zeile, 5), Excel.Range).Locked = False
                                    End If
                                End If


                                ' Ende-Datum 
                                CType(.Cells(zeile, 6), Excel.Range).Value = cPhase.getEndDate.ToShortDateString
                                If DateDiff(DateInterval.Day, hproj.actualDataUntil, cPhase.getEndDate) <= 0 Then
                                    ' Sperren ...
                                    CType(currentWS.Cells(zeile, 6), Excel.Range).Locked = True
                                    CType(currentWS.Cells(zeile, 6), Excel.Range).Interior.Color = XlRgbColor.rgbLightGrey
                                    CType(currentWS.Cells(zeile, 6), Excel.Range).Font.Color = XlRgbColor.rgbBlack
                                Else
                                    If isProtectedbyOthers Then
                                        CType(currentWS.Cells(zeile, 6), Excel.Range).Locked = True
                                        'CType(currentWS.Cells(zeile, 6), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                    Else
                                        CType(currentWS.Cells(zeile, 6), Excel.Range).Locked = False
                                    End If

                                End If

                                ' Ampel-Farbe
                                CType(.Cells(zeile, 7), Excel.Range).Value = cPhase.ampelStatus
                                If isProtectedbyOthers Then
                                    CType(currentWS.Cells(zeile, 7), Excel.Range).Locked = True
                                    'CType(currentWS.Cells(zeile, 7), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                Else
                                    CType(.Cells(zeile, 7), Excel.Range).Locked = False
                                End If


                                If cPhase.ampelStatus = 1 Then
                                    CType(.Cells(zeile, 7), Excel.Range).Interior.Color = visboFarbeGreen
                                ElseIf cPhase.ampelStatus = 2 Then
                                    CType(.Cells(zeile, 7), Excel.Range).Interior.Color = visboFarbeYellow
                                ElseIf cPhase.ampelStatus = 3 Then
                                    CType(.Cells(zeile, 7), Excel.Range).Interior.Color = visboFarbeRed
                                Else
                                    ' tk keine Farbe mehr 
                                    'CType(.Cells(zeile, 7), Excel.Range).Interior.Color = visboFarbeNone
                                End If


                                ' Ampel-Erläuterung
                                CType(.Cells(zeile, 8), Excel.Range).Value = cPhase.ampelErlaeuterung
                                If isProtectedbyOthers Then
                                    CType(currentWS.Cells(zeile, 8), Excel.Range).Locked = True
                                    'CType(currentWS.Cells(zeile, 8), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                Else
                                    CType(.Cells(zeile, 8), Excel.Range).Locked = False
                                End If


                                ' Lieferumfänge
                                CType(.Cells(zeile, 9), Excel.Range).Value = cPhase.getAllDeliverables(vbLf)
                                If isProtectedbyOthers Then
                                    CType(currentWS.Cells(zeile, 9), Excel.Range).Locked = True
                                    'CType(currentWS.Cells(zeile, 9), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                Else
                                    CType(.Cells(zeile, 9), Excel.Range).Locked = False
                                End If


                                ' wer ist verantwortlich
                                CType(.Cells(zeile, 10), Excel.Range).Value = cPhase.verantwortlich
                                If isProtectedbyOthers Then
                                    CType(currentWS.Cells(zeile, 10), Excel.Range).Locked = True
                                    'CType(currentWS.Cells(zeile, 10), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                Else
                                    CType(.Cells(zeile, 10), Excel.Range).Locked = False
                                End If


                                ' wieviel ist erledigt ? 
                                CType(.Cells(zeile, 11), Excel.Range).Value = cPhase.percentDone.ToString("0#%")
                                If isProtectedbyOthers Then
                                    CType(currentWS.Cells(zeile, 11), Excel.Range).Locked = True
                                    'CType(currentWS.Cells(zeile, 11), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                Else
                                    CType(.Cells(zeile, 11), Excel.Range).Locked = False
                                End If


                                ' der Dokumenten Link 
                                CType(currentWS.Cells(zeile, 12), Excel.Range).Value = cPhase.DocURL
                                If isProtectedbyOthers Then
                                    CType(currentWS.Cells(zeile, 12), Excel.Range).Locked = True
                                    'CType(currentWS.Cells(zeile, 12), Excel.Range).Interior.Color = XlRgbColor.rgbLightGray
                                Else
                                    CType(currentWS.Cells(zeile, 12), Excel.Range).Locked = False
                                End If


                                ' wenn Phasen Invoices / Penalties haben können 
                                If awinSettings.enableInvoices Then
                                    ' der Rechnungsbetrag und Zahlungsziel 
                                    If Not IsNothing(cPhase.invoice) Then
                                        If cPhase.invoice.Key > 0 Then
                                            CType(currentWS.Cells(zeile, 13), Excel.Range).Value = cPhase.invoice.Key
                                            CType(currentWS.Cells(zeile, 14), Excel.Range).Value = cPhase.invoice.Value
                                        End If
                                    End If

                                    ' die Penalty und das Penalty Date
                                    If Not IsNothing(cPhase.penalty) Then
                                        If cPhase.penalty.Value > 0 Then
                                            CType(currentWS.Cells(zeile, 15), Excel.Range).Value = cPhase.penalty.Value
                                            CType(currentWS.Cells(zeile, 16), Excel.Range).Value = cPhase.penalty.Key.Date
                                        End If
                                    End If

                                    If isProtectedbyOthers Then
                                        CType(currentWS.Cells(zeile, 13), Excel.Range).Locked = True
                                        CType(currentWS.Cells(zeile, 14), Excel.Range).Locked = True
                                        CType(currentWS.Cells(zeile, 15), Excel.Range).Locked = True
                                        CType(currentWS.Cells(zeile, 16), Excel.Range).Locked = True
                                    Else
                                        CType(currentWS.Cells(zeile, 13), Excel.Range).Locked = False
                                        CType(currentWS.Cells(zeile, 14), Excel.Range).Locked = False
                                        CType(currentWS.Cells(zeile, 15), Excel.Range).Locked = False
                                        CType(currentWS.Cells(zeile, 16), Excel.Range).Locked = False
                                    End If


                                End If

                            End With
                        End If


                        ' Zeile eins weiter ... 
                        zeile = zeile + 1
                        curElemID = hproj.hierarchy.getNextIdOfId(curElemID, indentLevel, abbruchLevel)

                    Loop

                End If

            Next


            ' jetzt die Größe der Spalten für BU, pName, vName, Phasen-Name, RC-Name anpassen 
            Dim infoBlock As Excel.Range
            Dim infoDataBlock As Excel.Range


            Dim firstHundredColumns As Excel.Range = Nothing

            With CType(currentWS, Excel.Worksheet)
                infoBlock = CType(.Range(.Columns(1), .Columns(anzSpalten)), Excel.Range)
                infoDataBlock = CType(.Range(.Cells(2, 1), .Cells(zeile + 100, anzSpalten)), Excel.Range)

                firstHundredColumns = CType(.Range(.Columns(1), .Columns(100)), Excel.Range)

                infoBlock.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                infoBlock.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

                ' die Besonderheiten abbilden 

                ' Phasen bzw. Meilenstein Name
                With CType(infoDataBlock.Columns(4), Excel.Range)
                    .WrapText = True
                End With

                ' Erläuterung
                With CType(infoDataBlock.Columns(8), Excel.Range)
                    .WrapText = True
                End With
                ' Lieferumfänge 
                With CType(infoDataBlock.Columns(9), Excel.Range)
                    .WrapText = True
                End With

                ' percent Done 
                With CType(infoDataBlock.Columns(11), Excel.Range)
                    .NumberFormat = "0#%"
                End With

                ' hier prüfen, ob es bereits Werte für massColValues gibt ..
                If massColFontValues(1, 2) > 0 Then
                    ' es wurden bereits mal Spaltenbreiten gesetzt 

                    For ik As Integer = 1 To 100
                        CType(firstHundredColumns.Columns(ik), Excel.Range).ColumnWidth = massColFontValues(1, ik)
                    Next

                Else
                    ' hier jetzt prüfen, ob nicht zu viel Platz eingenommen wird
                    Try
                        firstHundredColumns.AutoFit()
                    Catch ex As Exception

                    End Try



                End If

            End With

            ' löschen des ganzen Blattes
            If todoListe.Count = 0 Then
                infoDataBlock.Clear()
            End If

            appInstance.EnableEvents = True



        Catch ex As Exception
            Call MsgBox("Fehler in Aufbereitung Termine" & vbLf & ex.Message)
            appInstance.EnableEvents = True
        End Try


    End Sub


    ''' <summary>
    ''' bestimmt das Erscheinungsbild der ersten Zeile in einem Mass-Edit Fenster Ressourcen, Termine, Attribute
    ''' </summary>
    ''' <param name="tableTyp"></param>
    Private Sub massEditZeile1Appearance(ByVal tableTyp As Integer)

        Dim currentWS As Excel.Worksheet = CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(tableTyp)), Excel.Worksheet)
        Dim ersteZeile As Excel.Range = CType(currentWS.Rows(1), Excel.Range)

        With ersteZeile
            .RowHeight = awinSettings.zeilenhoehe1 + 5
            .Interior.Color = visboFarbeOrange
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = XlRgbColor.rgbWhite
            .VerticalAlignment = Excel.XlVAlign.xlVAlignTop
        End With

    End Sub


    Public Sub FollowHyperlinkToWebsite(ByVal hproj As clsProjekt, Optional ByVal type As String = "Capacity")

        Dim vpid As String = hproj.vpID
        Dim varID As String = CType(databaseAcc, DBAccLayer.Request).findVariantID(vpid, hproj.variantName)
        'Dim varID As String = "624dcfc6e89109508af0f7e2"

        Dim refDate As Date = Date.Now
        Dim StartOfProject As String = DateTimeToISODate(hproj.startDate)
        Dim EndOfProject As String = DateTimeToISODate(hproj.endeDate)
        Dim vonbis As String = "&from=" & StartOfProject & "&to=" & EndOfProject

        Dim variante As String = "&variantID=" & varID
        Dim serverURL As String = awinSettings.databaseURL.Replace("/api", "")

        appInstance.ActiveWorkbook.FollowHyperlink(Address:=serverURL & "/vpKeyMetrics/" & vpid & "?view=" & type & "&unit=PD" & vonbis & variante, NewWindow:=True)

    End Sub

    Public Sub loadGivenProject()
        ' Laden des übergebenen Projektes

        Dim err As New clsErrorCodeMsg

        If spe_vpid <> "" And spe_vpvid <> "" Then
            'TODO: hier holen des Projekte mit vpid... und vpvid...
            Dim hproj As clsProjekt = CType(databaseAcc, DBAccLayer.Request).retrieveOneProjectVersionfromDB(spe_vpid, spe_vpvid, err)
            If Not IsNothing(hproj) Then
                ShowProjekte.Add(hproj, False)
                AlleProjekte.Add(hproj, False)
            End If
            spe_vpid = ""
            spe_vpvid = ""
        End If

        appInstance.EnableEvents = True

        If AlleProjekte.Count > 0 Then
            ' Termine edit aufschalten
            'all MsgBox(currentProjektTafelModus)
            Call massEditRcTeAt(currentProjektTafelModus)
        End If

    End Sub
    ''' <summary>
    ''' Umwandlung einen Datum des Typs Date in einen ISO-Datums-String
    ''' </summary>
    ''' <param name="datumUhrzeit"></param>
    ''' <returns></returns>
    Public Function DateTimeToISODate(ByVal datumUhrzeit As Date) As String

        Dim ISODateandTime As String = Nothing
        Dim ISODate As String = ""
        Dim ISOTime As String = ""

        If datumUhrzeit >= Date.MinValue And datumUhrzeit <= Date.MaxValue Then
            ' DatumUhrzeit wird um 1 Sekunde erhöht, dass die 1000-stel keine Rolle spielen
            Dim hours As Integer = datumUhrzeit.Hour
            Dim minutes As Integer = datumUhrzeit.Minute
            Dim seconds As Integer = datumUhrzeit.Second
            Dim milliseconds As Integer = datumUhrzeit.Millisecond
            datumUhrzeit = datumUhrzeit.Date
            datumUhrzeit = datumUhrzeit.AddHours(hours).AddMinutes(minutes).AddSeconds(seconds).AddMilliseconds(0)
            ISODateandTime = datumUhrzeit.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        End If

        DateTimeToISODate = ISODateandTime

    End Function
End Module
