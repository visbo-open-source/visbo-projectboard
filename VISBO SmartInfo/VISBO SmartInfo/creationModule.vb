Imports ProjectBoardDefinitions
Imports System.Windows.Forms
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports System.Collections
Imports ProjectBoardBasic
Module creationModule

    ' defines the keyWords fpr Powerpoint Reporting Compoents, so that report component can be generated
    Friend projectComponentNames As String() = {"Projekt-Name", "Custom-Field", "selectedItems", "Einzelprojektsicht",
                                                "AllePlanElemente", "Swimlanes", "Swimlanes2", "TableBudgetCostAPVCV",
                                                "TableMilestoneAPVCV", "ProjektBedarfsChart", "Ampel-Farbe", "Ampel-Text",
                                                "Beschreibung", "Business-Unit", "SymTrafficLight", "SymRisks", "SymGoals",
                                                "SymTeam", "SymFinance", "SymSchedules", "SymPrPf", "Stand:", "Laufzeit:", "Verantwortlich:"}

    Friend multiprojectComponentNames As String() = {"Multiprojektsicht"}

    Friend portfolioComponentNames As String() = {}

    ' hier ist  projectboardCustomization.xlsx zu finden
    Friend customizationPath As String = ""

    '
    ' wird benötigt für ReportCreation
    Friend currentSldHasProjectTemplates As Boolean = False
    Friend currentSldHasMultiProjectTemplates As Boolean = False
    Friend currentSldHasPortfolioTemplates As Boolean = False

    Friend appearancesWereRead As Boolean = False
    ' Ende ReportCreation Spezifika
    '

    Public Sub readSettings(ByVal dbNameIsKnown As Boolean)
        With awinSettings


            ' ur:2020.12.1: Einstellungen für direkt MongoDB oder ReST-Server Zugriff
            .databaseURL = My.Settings.mongoDBURL
            .visboServer = My.Settings.VISBOServer
            .proxyURL = My.Settings.proxyServerURL
            .DBWithSSL = My.Settings.mongoDBWithSSL

            If Not dbNameIsKnown Then
                .databaseName = My.Settings.mongoDBname
            End If

            .awinPath = My.Settings.awinPath

            .rememberUserPwd = My.Settings.rememberUserPWD
            .userNamePWD = My.Settings.userNamePWD


            .mppShowProjectLine = My.Settings.showProjectLine
            .mppShowAllIfOne = My.Settings.showAllIfOne
            .mppShowAmpel = My.Settings.showAmpel
            .mppUseOriginalNames = My.Settings.useOriginalNames
            .mppShowPhName = My.Settings.showPhName
            .mppShowPhDate = My.Settings.showPhDate
            .mppUseAbbreviation = My.Settings.useAbbrev
            .mppShowMsName = My.Settings.showMsName
            .mppShowMsDate = My.Settings.showMsDate
            .mppKwInMilestone = My.Settings.kwInMilestone
            .mppVertikalesRaster = My.Settings.showVerticals
            .mppShowLegend = My.Settings.showLegend
            .mppSortiertDauer = My.Settings.sortiertDauer
            .mppShowHorizontals = My.Settings.showHorizontals
            .mppOnePage = My.Settings.allOnePage
            .mppExtendedMode = My.Settings.extendedMode
            .mppProjectsWithNoMPmayPass = My.Settings.projectswithNoPhMsmayPass

            If .mppSortiertDauer Then
                .mppShowAllIfOne = True
            End If

            ' now get path where projectboardCustomization.xlsx is to find 
            ' 
            customizationPath = My.Settings.customizationPath

            ' now define showLeftrange
            If My.Settings.calLeftDate <> Date.MinValue Then
                showRangeLeft = getColumnOfDate(My.Settings.calLeftDate)
            End If

            If My.Settings.calRightDate <> Date.MinValue Then
                showRangeRight = getColumnOfDate(My.Settings.calRightDate)
            End If

        End With



    End Sub

    Public Sub writeSettings()
        With awinSettings

            ' auskommentierte Settings bleiben unverändert während der Ausführung dieses Programms
            '' ur:2020.12.1: Einstellungen für direkt MongoDB oder ReST-Server Zugriff
            'My.Settings.mongoDBURL = .databaseURL
            'My.Settings.VISBOServer = .visboServer
            'My.Settings.proxyServerURL = .proxyURL
            'My.Settings.mongoDBWithSSL = .DBWithSSL
            'My.Settings.mongoDBname = .databaseName
            'My.Settings.awinPath = .awinPath

            ' folgende Settings werden im Link Settings vor dem Erzeugen einen Reports evt. modifiziert
            My.Settings.showProjectLine = .mppShowProjectLine
            My.Settings.showAllIfOne = .mppShowAllIfOne
            My.Settings.showAmpel = .mppShowAmpel
            My.Settings.useOriginalNames = .mppUseOriginalNames
            My.Settings.showPhName = .mppShowPhName
            My.Settings.showPhDate = .mppShowPhDate
            My.Settings.useAbbrev = .mppUseAbbreviation
            My.Settings.showMsName = .mppShowMsName
            My.Settings.showMsDate = .mppShowMsDate
            My.Settings.kwInMilestone = .mppKwInMilestone
            My.Settings.showVerticals = .mppVertikalesRaster
            My.Settings.showLegend = .mppShowLegend
            My.Settings.sortiertDauer = .mppSortiertDauer

            My.Settings.showHorizontals = .mppShowHorizontals
            My.Settings.allOnePage = .mppOnePage
            My.Settings.extendedMode = .mppExtendedMode
            My.Settings.projectswithNoPhMsmayPass = .mppProjectsWithNoMPmayPass


            ' now define showLeftrange
            If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
                My.Settings.calLeftDate = getDateofColumn(showRangeLeft, False)
                My.Settings.calRightDate = getDateofColumn(showRangeRight, True)
            End If
        End With

        ' Settings sichern für den nächsten Programm-Durchlauf
        My.Settings.Save()
    End Sub

    ''' <summary>
    ''' erzeugt den Bericht auf Grundlage des aktuell geladenen Powerpoints  
    ''' bei Aufruf ist sichergestellt, daß in Projekthistorie die Historie des Projektes steht 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub fillReportingComponentWithinPPT(ByRef hproj As clsProjekt,
                                          ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection,
                                          ByVal selectedRoles As Collection, ByVal selectedCosts As Collection,
                                          ByVal selectedBUs As Collection, ByVal selectedTyps As Collection,
                                          ByRef zeilenhoehe_sav As Double,
                                          ByRef legendFontSize As Single)

        Dim err As New clsErrorCodeMsg

        ' 4.10.19 tk wird doch gar nicht mehr gebraucht ? ist ja die currentPResentation
        'Dim pptCurrentPresentation As PowerPoint.Presentation = Nothing

        ' 4.10.19 tk ist die currentSlide
        'Dim pptSlide As PowerPoint.Slide = Nothing
        Dim shapeRange As PowerPoint.ShapeRange = Nothing

        Dim pptShape As PowerPoint.Shape
        Dim pname As String = hproj.name
        Dim fullName As String = hproj.getShapeText
        Dim top As Double, left As Double, width As Double, height As Double
        Dim htop As Double, hleft As Double, hwidth As Double, hheight As Double
        Dim pptSize As Single = 18
        ' ur: 16.04.2015: wird nun übergeben: Dim zeilenhoehe As Double = 0.0

        ' tk 5.10.19 wird nicht mehr referenziert
        'Dim auswahl As Integer
        'Dim compareToID As Integer

        Dim qualifier As String = " ", qualifier2 As String = " "
        'Dim notYetDone As Boolean = False
        Dim ze As String = " (" & awinSettings.kapaEinheit & ")"
        Dim ke As String = " (T€)"
        Dim heute As Date = Date.Now
        Dim istWerteexistieren As Boolean
        Dim boxName As String
        Dim listofShapes As New Collection

        Dim lproj As clsProjekt = Nothing
        Dim bproj As clsProjekt = Nothing
        Dim lastproj As clsProjekt = Nothing
        Dim lastElem As Integer
        ' das sind Formen , die zur in der Tabelle Vergleich Anzeige der Tendenz verwendet werden 
        Dim gleichShape As PowerPoint.Shape = Nothing
        Dim steigendShape As PowerPoint.Shape = Nothing
        Dim fallendShape As PowerPoint.Shape = Nothing
        Dim ampelShape As PowerPoint.Shape = Nothing
        Dim sternShape As PowerPoint.Shape = Nothing

        Dim reportCreationDate As Date = Date.Now

        Dim bigType As Integer = -1
        Dim compID As Integer = -1

        Dim msgTxt As String = ""


        Try   ' Projekthistorie aufbauen, sofern sie für das aktuelle hproj nicht schon aufgebaut ist
            ' die Projekthistorie wird immer (zumindest erstmal hier ...) nur von der Basis-Variante betrachtet ...


            Dim aktprojekthist As Boolean = False

            If projekthistorie.Count > 0 Then
                'aktprojekthist = (hproj.name = projekthistorie.First.name)
                aktprojekthist = (hproj.name = projekthistorie.First.name)
            End If



            If Not noDBAccessInPPT Then

                If CType(databaseAcc, DBAccLayer.Request).pingMongoDb() Then
                    Try

                        If Not aktprojekthist Then
                            projekthistorie = CType(databaseAcc, DBAccLayer.Request).retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:="",
                                                                        storedEarliest:=Date.MinValue, storedLatest:=Date.Now, err:=err)
                        End If


                        ' bei Projekten, egal ob standard Projekt oder Portfolio Projekt wird immer mit der Vorlagen-Variante verglichen
                        Dim vorgabeVariantName As String = ptVariantFixNames.pfv.ToString

                        ' ur: alt: bproj = CType(databaseAcc, DBAccLayer.Request).retrieveFirstContractedPFromDB(hproj.name, vorgabeVariantName, err)
                        bproj = projekthistorie.beauftragung

                        ' tk 19.1.19 das darf hier nicht mehr gemacht werden. Eine letzte Vorgabe kann später gemacht sein als der Planungsstand ... 
                        'Dim lDate As Date = hproj.timeStamp.AddMinutes(-1)
                        ' ur: alt: lproj = CType(databaseAcc, DBAccLayer.Request).retrieveLastContractedPFromDB(hproj.name, vorgabeVariantName, storedAtOrBefore:=Date.Now, err:=err)
                        lproj = projekthistorie.lastBeauftragung(Date.Now)


                    Catch ex As Exception
                        projekthistorie.clear()
                    End Try
                Else
                    Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
                End If
            Else
                Call MsgBox("Datenbank-Anbindung ist nicht aktiviert. Historie enthält nur das aktuelle Projekt " & hproj.name)
                Exit Sub
            End If

        Catch ex As Exception
            Call MsgBox("Fehler in Create: " & ex.Message)
            Exit Sub
        End Try

        Try
            lastElem = projekthistorie.Count - 1
            lastproj = projekthistorie.ElementAt(lastElem - 1)
        Catch ex As Exception
            lastElem = -1
            lastproj = Nothing
        End Try



        If DateDiff(DateInterval.Month, hproj.startDate, heute) > 0 Then
            istWerteexistieren = True
        Else
            istWerteexistieren = False
        End If

        ' tk 4.10.19 aktuell wird nur eine Slide behandelt ... 
        'Dim anzSlidesToAdd As Integer = 1
        'Dim anzahlCurrentSlides As Integer
        'Dim currentInsert As Integer = 1

        ' jetzt wird das CurrentPresentation File unter einem Dummy Namen gespeichert ..

        Dim kennzeichnung As String = ""
        Dim anzShapes As Integer

        Dim folieIX As Integer = 1
        Dim objectsToDo As Integer = 0
        Dim objectsDone As Integer = 0

        ' tk 4.10 aktuell macht er das einfach nur für die aktuelle Slide auf der man sitzt 
        While folieIX <= 1

            Call addSmartPPTSlideBaseInfo(currentSlide, reportCreationDate, ptPRPFType.project)

            ' jetzt werden die Charts gezeichnet 
            anzShapes = currentSlide.Shapes.Count
            Dim newShapeRange As PowerPoint.ShapeRange = Nothing
            Dim newShapeRange2 As PowerPoint.ShapeRange = Nothing
            Dim newShape As PowerPoint.Shape = Nothing

            ' müssen Phasen, Meilensteine gewählt werden ?
            ' (0) = true : es wird eine Selection benötigt 
            ' (1) = true : die Selection hat bereits stattgefunden
            Dim phMSSelNeeded(1) As Boolean
            phMSSelNeeded(0) = False
            phMSSelNeeded(1) = False

            ' müssen Rollen, Kostenarten gewählt werden ?
            ' (0) = true : es wird eine Selection benötigt 
            ' (1) = true : die Selection hat bereits stattgefunden
            Dim roleCostSelNeeded(1) As Boolean
            roleCostSelNeeded(0) = False
            roleCostSelNeeded(1) = False

            ' jetzt wird die listofShapes aufgebaut - das sind alle Shapes, die ersetzt werden müssen ...
            For i = 1 To anzShapes
                pptShape = currentSlide.Shapes(i)
                qualifier = ""
                With pptShape

                    Dim tmpStr(3) As String
                    Try

                        If .Title <> "" Then
                            tmpStr = .Title.Trim.Split(New Char() {CChar("("), CChar(")")}, 3)
                            kennzeichnung = tmpStr(0).Trim
                        Else
                            If CBool(.HasTextFrame) Then
                                Dim dummyStr As String = .TextFrame2.TextRange.Text
                                tmpStr = dummyStr.Trim.Split(New Char() {CChar("("), CChar(")")}, 3)
                                kennzeichnung = tmpStr(0).Trim
                            Else
                                kennzeichnung = "nicht identifizierbar"
                            End If
                        End If

                    Catch ex As Exception
                        kennzeichnung = "nicht identifizierbar"
                    End Try

                    If kennzeichnung = "Projekt-Name" Or
                        kennzeichnung = "Custom-Field" Or
                        kennzeichnung = "selectedItems" Or
                        kennzeichnung = "Einzelprojektsicht" Or
                        kennzeichnung = "Multiprojektsicht" Or
                        kennzeichnung = "AllePlanElemente" Or
                        kennzeichnung = "Swimlanes" Or
                        kennzeichnung = "Swimlanes2" Or
                        kennzeichnung = "TableBudgetCostAPVCV" Or
                        kennzeichnung = "TableMilestoneAPVCV" Or
                        kennzeichnung = "ProjektBedarfsChart" Or
                        kennzeichnung = "Ampel-Farbe" Or
                        kennzeichnung = "Ampel-Text" Or
                        kennzeichnung = "Beschreibung" Or
                        kennzeichnung = "Business-Unit" Or
                        kennzeichnung = "SymTrafficLight" Or
                        kennzeichnung = "SymRisks" Or
                        kennzeichnung = "SymGoals" Or
                        kennzeichnung = "SymTeam" Or
                        kennzeichnung = "SymFinance" Or
                        kennzeichnung = "SymSchedules" Or
                        kennzeichnung = "SymPrPf" Or
                        kennzeichnung = "Stand:" Or
                        kennzeichnung = "Laufzeit:" Or
                        kennzeichnung = "Verantwortlich:" Then

                        listofShapes.Add(pptShape)

                    ElseIf kennzeichnung = "gleich" Then
                        gleichShape = pptShape

                    ElseIf kennzeichnung = "steigend" Then
                        steigendShape = pptShape

                    ElseIf kennzeichnung = "fallend" Then
                        fallendShape = pptShape

                    ElseIf kennzeichnung = "ampel" Then
                        ampelShape = pptShape

                    ElseIf kennzeichnung = "stern" Then
                        sternShape = pptShape

                    End If


                End With

                If kennzeichnung = "Einzelprojektsicht" Or
                        kennzeichnung = "Swimlanes" Or
                        kennzeichnung = "Swimlanes2" Or
                        kennzeichnung = "Multiprojektsicht" Or
                        kennzeichnung = "TableMilestoneAPVCV" Then

                    phMSSelNeeded(0) = True

                ElseIf kennzeichnung = "TableBudgetCostAPVCV" Or
                    kennzeichnung = "ProjektBedarfsChart" Then

                    roleCostSelNeeded(0) = True
                End If


            Next

            ' je nachdem, welche Komponenten jetzt erstellt werden sollen 
            ' muss hier noch die Auswahl der selectedPhases passieren 

            If phMSSelNeeded(0) = True And Not phMSSelNeeded(1) = True Then

                Dim listOfFormerSelectedProjects As String() = Nothing
                If selectedProjekte.Count > 0 Then

                    listOfFormerSelectedProjects = selectedProjekte.Liste.Keys.ToArray

                End If


                showRangeLeft = ShowProjekte.getMinMonthColumn
                showRangeRight = ShowProjekte.getMaxMonthColumn + 3

                ' ur:2020.12.04: löschen der evt. zuvor ausgewählten Phasen und Meilensteine
                selectedPhases.Clear()
                selectedMilestones.Clear()

                ' jetzt die selectedProjekte auf ein Projekt setzen, das wird nämlich dann verwendet , um im TreeView bei 
                ' die Struktur Auswahl zu machen 
                selectedProjekte.Clear(False)
                If ShowProjekte.Count > 0 Then
                    selectedProjekte.Add(ShowProjekte.getProject(1), False)
                End If

                Dim frmSelectionPhMs As New frmSelectPhasesMilestones
                If frmSelectionPhMs.ShowDialog = Windows.Forms.DialogResult.OK Then

                    If Not IsNothing(frmSelectionPhMs.selectedPhases) Then
                        selectedPhases = frmSelectionPhMs.selectedPhases
                    Else
                        selectedPhases = New Collection
                    End If

                    If Not IsNothing(frmSelectionPhMs.selectedMilestones) Then
                        selectedMilestones = frmSelectionPhMs.selectedMilestones
                    Else
                        selectedMilestones = New Collection
                    End If


                Else
                    Exit Sub
                End If

                selectedProjekte.Clear(False)

                showRangeLeft = getColumnOfDate(frmSelectionPhMs.vonDate.Value)
                showRangeRight = getColumnOfDate(frmSelectionPhMs.bisDate.Value)

                phMSSelNeeded(1) = True
                If Not IsNothing(listOfFormerSelectedProjects) Then
                    selectedProjekte.Clear(False)

                    For Each tmpName As String In listOfFormerSelectedProjects
                        If ShowProjekte.contains(tmpName) Then
                            selectedProjekte.Add(ShowProjekte.getProject(tmpName), False)
                        End If
                    Next
                End If

                '' jetzt muss für den Multiprojekt Report noch showrangeLeft und Right gesetzt werden 
                'showRangeLeft = ShowProjekte.getMinMonthColumn - 1
                'showRangeRight = ShowProjekte.getMaxMonthColumn + 3

            End If



            For Each tmpShape As PowerPoint.Shape In listofShapes


                Try
                    pptShape = tmpShape
                    qualifier = ""
                    qualifier2 = ""
                    kennzeichnung = ""

                    With pptShape
                        .Name = "Shape" & .Id.ToString
                        Dim tst As String = .AlternativeText
                        If .Title <> "" Then

                            Call title2kennzQualifier(.Title, kennzeichnung, qualifier, qualifier2)
                            boxName = kennzeichnung

                        Else
                            ' Start neu

                            Call title2kennzQualifier(.TextFrame2.TextRange.Text, kennzeichnung, qualifier, qualifier2)
                            boxName = kennzeichnung

                        End If

                        ' wenn .AlternativeText was enthält ; das wird z.Bsp in Tabelle PRojektziele benötigt ...
                        If .AlternativeText <> "" Then
                            qualifier2 = .AlternativeText
                        End If


                        top = .Top
                        left = .Left
                        height = .Height
                        width = .Width

                        ' ur:27.04.2016
                        ' ''Try
                        ' ''    boxName = .TextFrame2.TextRange.Text
                        ' ''Catch ex As Exception
                        ' ''    boxName = " "
                        ' ''End Try

                        If CBool(.TextFrame2.HasText) Then
                            boxName = .TextFrame2.TextRange.Text
                        Else
                            boxName = ""
                        End If


                        htop = 100
                        hleft = 100
                        hwidth = 300
                        hheight = 400

                        Select Case kennzeichnung

                            Case "Projekt-Name"

                                fullName = hproj.getShapeText

                                If qualifier.Length > 0 Then
                                    If qualifier.Trim <> "Enlarge13" Then
                                        .TextFrame2.TextRange.Text = fullName & ": " & qualifier
                                    Else
                                        .TextFrame2.TextRange.Text = fullName
                                    End If
                                Else
                                    .TextFrame2.TextRange.Text = fullName
                                End If

                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                           ptReportBigTypes.components, ptReportComponents.prName)

                            Case "selectedItems"

                                Dim selTxt As String = ""

                                If selectedRoles.Count > 0 Then
                                    For Each tmpRoleID As String In selectedRoles
                                        Dim teamID As Integer = -1
                                        Dim tmpRoleName As String = RoleDefinitions.getRoleDefByIDKennung(tmpRoleID, teamID).name
                                        If selTxt = "" Then
                                            selTxt = tmpRoleName
                                        Else
                                            selTxt = selTxt & "; " & tmpRoleName
                                        End If
                                    Next
                                End If

                                If selectedCosts.Count > 0 Then
                                    Dim firstTime As Boolean = True
                                    For Each tmpCostName As String In selectedCosts
                                        If selTxt = "" Then
                                            selTxt = tmpCostName
                                        Else
                                            If firstTime Then
                                                selTxt = selTxt & vbLf & tmpCostName
                                            Else
                                                selTxt = selTxt & "; " & tmpCostName
                                            End If
                                        End If
                                        firstTime = False
                                    Next
                                End If

                                ' wenn nichts in selTxtx drin steht , ist es auch gut. Dann "verschwindet" dieses Feld ...
                                .TextFrame2.TextRange.Text = selTxt



                            Case "Custom-Field"
                                If qualifier.Length > 0 Then
                                    ' existiert der überhaupt 
                                    Dim uid As Integer = customFieldDefinitions.getUid(qualifier)

                                    If uid <> -1 Then
                                        Dim cftype As Integer = customFieldDefinitions.getTyp(uid)

                                        Select Case cftype
                                            Case ptCustomFields.Str
                                                Dim wert As String = hproj.getCustomSField(uid)
                                                If Not IsNothing(wert) Then
                                                    .TextFrame2.TextRange.Text = qualifier & ": " & wert
                                                Else
                                                    .TextFrame2.TextRange.Text = qualifier & " : n.a"
                                                End If

                                            Case ptCustomFields.Dbl
                                                Dim wert As Double = hproj.getCustomDField(uid)
                                                If Not IsNothing(wert) Then
                                                    .TextFrame2.TextRange.Text = qualifier & ": " & wert.ToString("#0.##")
                                                Else
                                                    .TextFrame2.TextRange.Text = qualifier & " : n.a"
                                                End If

                                            Case ptCustomFields.bool
                                                Dim wert As Boolean = hproj.getCustomBField(uid)

                                                If Not IsNothing(wert) Then
                                                    If wert Then
                                                        ' Sprache !
                                                        .TextFrame2.TextRange.Text = qualifier & ": Yes"
                                                    Else
                                                        ' Sprache !
                                                        .TextFrame2.TextRange.Text = qualifier & ": No"
                                                    End If

                                                Else
                                                    .TextFrame2.TextRange.Text = qualifier & " : n.a"
                                                End If

                                        End Select

                                        Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                           ptReportBigTypes.components, ptReportComponents.prCustomField)
                                    Else
                                        .TextFrame2.TextRange.Text = "Custom-Field " & qualifier &
                                            " existiert nicht !"
                                    End If

                                Else
                                    ' n.a"
                                    .TextFrame2.TextRange.Text = "Custom-Field ohne Namen.."
                                End If


                            Case "AllePlanElemente"

                                Try

                                    Dim i As Integer = 0
                                    Dim tmpphases As New Collection
                                    Dim tmpMilestones As New Collection
                                    Dim minCal As Boolean = False
                                    If qualifier2.Length > 0 Then
                                        minCal = (qualifier2.Trim = "minCal")
                                    End If

                                    ' alle Phasennamen des Projektes hproj in die Collection tmpphases bringen
                                    For Each cphase In hproj.AllPhases

                                        Dim tmpstr As String = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                                        If tmpstr <> "" Then
                                            tmpstr = tmpstr & "#" & cphase.name
                                            If Not tmpphases.Contains(tmpstr) Then
                                                tmpphases.Add(tmpstr, tmpstr)
                                            End If

                                        End If


                                    Next



                                    ' alle Meilensteine-Namen des Projektes hproj in die collection tmpMilestones bringen
                                    Dim mSList As SortedList(Of Date, String)

                                    mSList = hproj.getMilestones        ' holt alle Meilensteine in Form ihrer nameID sortiert nach Datum

                                    If mSList.Count > 0 Then
                                        For Each kvp As KeyValuePair(Of Date, String) In mSList

                                            Dim tmpstr = hproj.hierarchy.getBreadCrumb(kvp.Value) & "#" & hproj.getMilestoneByID(kvp.Value).name
                                            If Not tmpMilestones.Contains(tmpstr) Then
                                                tmpMilestones.Add(tmpstr, tmpstr)
                                            End If

                                        Next
                                    End If


                                    ' die Slide mit Tag kennzeichnen ... 
                                    Dim pptFirstTime As Boolean = True
                                    Call drawMultiprojectViewinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
                                                                  tmpphases, tmpMilestones,
                                                                  translateToRoleNames(selectedRoles), selectedCosts,
                                                                  selectedBUs, selectedTyps,
                                                                  False, False, hproj, kennzeichnung, minCal)
                                    .TextFrame2.TextRange.Text = ""
                                    '.ZOrder(MsoZOrderCmd.msoSendToBack)
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message
                                    objectsDone = objectsToDo

                                End Try

                            Case "Multiprojektsicht"

                                Try
                                    Dim tmpProjekt As New clsProjekt

                                    Dim minCal As Boolean = False
                                    If pptShape.AlternativeText.Length > 0 Then
                                        minCal = (pptShape.AlternativeText.Trim = "minCal")
                                    End If

                                    Dim pptFirstTime As Boolean = True
                                    Call drawMultiprojectViewinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
                                                              selectedPhases, selectedMilestones,
                                                              translateToRoleNames(selectedRoles), selectedCosts,
                                                              selectedBUs, selectedTyps,
                                                              True, False, tmpProjekt, kennzeichnung, minCal)
                                    .TextFrame2.TextRange.Text = ""
                                    '.ZOrder(MsoZOrderCmd.msoSendToBack)
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message
                                    objectsDone = objectsToDo
                                End Try


                            Case "Einzelprojektsicht"


                                Try
                                    Dim minCal As Boolean = False
                                    If qualifier2.Length > 0 Then
                                        minCal = (qualifier2.Trim = "minCal")
                                    End If
                                    Dim pptFirstTime As Boolean = True
                                    Call drawMultiprojectViewinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
                                                                      selectedPhases, selectedMilestones,
                                                                      translateToRoleNames(selectedRoles), selectedCosts,
                                                                      selectedBUs, selectedTyps,
                                                                      True, False, hproj, kennzeichnung, minCal)
                                    .TextFrame2.TextRange.Text = ""
                                    '.ZOrder(MsoZOrderCmd.msoSendToBack)
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message
                                    objectsDone = objectsToDo
                                End Try



                            Case "Swimlanes"

                                Try

                                    Dim minCal As Boolean = False
                                    If qualifier2.Length > 0 Then
                                        minCal = (qualifier2.Trim = "minCal")
                                    End If

                                    Dim pptFirstTime As Boolean = True
                                    Call zeichneSwimlane2SichtinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
                                                                      selectedPhases, selectedMilestones,
                                                                      translateToRoleNames(selectedRoles), selectedCosts,
                                                                      selectedBUs, selectedTyps,
                                                                      False, hproj, kennzeichnung, minCal)

                                    .TextFrame2.TextRange.Text = ""
                                    '.ZOrder(MsoZOrderCmd.msoSendToBack)

                                    ' sonst wird pptLasttime benötigt, um bei mehreren PRojekten 
                                    ' swimlaneMode wird erst nach Ende der While Schleife ausgewertet - in diesem Fall wird die tmpSav Folie gelöscht 
                                    'swimlaneMode = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message & ": iDkey = " & iDkey
                                    objectsDone = objectsToDo
                                End Try


                            Case "Swimlanes2"

                                Dim formerSetting As Boolean = awinSettings.mppExtendedMode


                                Try

                                    Dim minCal As Boolean = False
                                    If qualifier2.Length > 0 Then
                                        minCal = (qualifier2.Trim = "minCal")
                                    End If

                                    Dim pptFirstTime As Boolean = True
                                    Call zeichneSwimlane2SichtinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
                                                                      selectedPhases, selectedMilestones,
                                                                      translateToRoleNames(selectedRoles), selectedCosts,
                                                                      selectedBUs, selectedTyps,
                                                                      False, hproj, kennzeichnung, minCal)
                                    awinSettings.mppExtendedMode = formerSetting
                                    .TextFrame2.TextRange.Text = ""
                                    '.ZOrder(MsoZOrderCmd.msoSendToBack)

                                    ' sonst wird pptLasttime benötigt, um bei mehreren Projekten 
                                    ' swimlaneMode wird erst nach Ende der While Schleife ausgewertet - in diesem Fall wird die tmpSav Folie gelöscht 
                                    'swimlaneMode = True
                                Catch ex As Exception
                                    awinSettings.mppExtendedMode = formerSetting
                                    .TextFrame2.TextRange.Text = ex.Message & ": iDkey = " & iDkey
                                    objectsDone = objectsToDo
                                End Try






                            Case "TableMilestoneAPVCV"

                                Try
                                    ' wenn es im Qualifier angegebene Rollen und Kostenarten gibt, dann haben die Prio vor der interaktiven Auswahl 
                                    ' erstmal werden nur Meilensteinen betrachtet ...
                                    Dim sMilestones As Collection = selectedMilestones

                                    If Not IsNothing(qualifier2) Then
                                        If qualifier2.Length > 0 Then
                                            sMilestones = New Collection
                                            Dim tmpStr() As String = qualifier2.Split(New Char() {CChar(vbLf), CChar(vbCr)})

                                            For Each tmpPMName As String In tmpStr

                                                sMilestones.Add(tmpPMName)
                                            Next

                                        End If
                                    End If

                                    Dim q1 As String = "0"
                                    Dim q2 As String = "0"


                                    ' in Q2 steht die Anzahl der Meilensteine , in q1 könnte später die Anzahl der Phasen stehen  
                                    q2 = sMilestones.Count.ToString




                                    ' die smart Powerpoint Table Info wird in dieser MEthode gesetzt ...
                                    ' tk 24.6.18 damit man unabhängig von selectedMilestones in der PPT-Vorlage feste Meilensteine angeben kann 
                                    Call zeichneTableMilestoneAPVCV(pptShape, hproj, bproj, lproj, sMilestones, q1, q2)
                                    'Call zeichneProjektTabelleZiele(pptShape, hproj, selectedMilestones, qualifier, qualifier2)


                                Catch ex As Exception

                                End Try

                            Case "TableBudgetCostAPVCV"

                                Try
                                    ' es können hier keine interaktiven Qualifier angegeben werden 
                                    ' 
                                    Dim q1 As String = qualifier ' gibt ggf an, ob PT ausgegeben werden soll 
                                    Dim q2 As String = qualifier2

                                    ' es werden drei Fälle unterschieden
                                    ' 1. qualifier2 = "" ->  die Budget, PK, SK, Ergebnis Übersicht  :todoCollection leer, q1= 0 , q2=0
                                    ' 2. qualifier2 = %used% -> es wird die gemeinsame Liste ermittelt ; todoCollection leer oder mit Inhalt, q1=-1, q2= -1



                                    ' die smart Powerpoint Table Info wird in dieser MEthode gesetzt ...
                                    ' tk 24.6.18 damit man unabhängig von selectedMilestones in der PPT-Vorlage feste Werte / Identifier angeben kann 
                                    Call zeichneTableBudgetCostAPVCV(pptShape, hproj, bproj, lproj, q1, q2)


                                Catch ex As Exception

                                End Try


                            Case "ProjektBedarfsChart"

                                Try
                                    Dim smartChartInfo As clsSmartPPTChartInfo = getChartParametersFromQ1(qualifier)

                                    ' Text im ShapeContainer / Platzhalter zurücksetzen 
                                    .TextFrame2.TextRange.Text = ""


                                    With smartChartInfo
                                        .q2 = bestimmeRoleQ2(qualifier2, selectedRoles)
                                        .bigType = ptReportBigTypes.charts

                                        ' muss mit dem ersten oder letzten verglichen werden ? 
                                        .hproj = hproj
                                        .vpid = hproj.vpID
                                        If .vergleichsTyp = PTVergleichsTyp.erster Then
                                            .vglProj = bproj
                                        ElseIf .vergleichsTyp = PTVergleichsTyp.letzter Then
                                            .vglProj = lproj
                                        End If

                                    End With

                                    Call createProjektChartInPPTNew(smartChartInfo, pptShape)

                                    boxName = ""
                                    'notYetDone = False
                                    'End If
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message
                                End Try



                            Case "Ampel-Farbe"

                                If boxName = kennzeichnung Then
                                    If englishLanguage Then
                                        boxName = "Ampel-Farbe"
                                    Else
                                        boxName = "Traffic Light"
                                    End If
                                    'boxName = repMessages.getmsg(230)
                                End If

                                Select Case hproj.ampelStatus
                                    Case 0
                                        .Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                                    Case 1
                                        .Fill.ForeColor.RGB = CInt(awinSettings.AmpelGruen)
                                    Case 2
                                        .Fill.ForeColor.RGB = CInt(awinSettings.AmpelGelb)
                                    Case 3
                                        .Fill.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                                    Case Else
                                End Select

                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prAmpel
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)


                            Case "Ampel-Text"

                                If boxName = kennzeichnung Then
                                    If englishLanguage Then
                                        boxName = "Ampel-Text"
                                    Else
                                        boxName = "Traffic Light Explanation"
                                    End If
                                    'boxName = repMessages.getmsg(225)
                                End If
                                '.TextFrame2.TextRange.Text = boxName & ": " & hproj.ampelErlaeuterung
                                ' keine String Ampel-Text mehr rein-machen
                                .TextFrame2.TextRange.Text = hproj.ampelErlaeuterung

                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prAmpelText
                                qualifier2 = boxName
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "Business-Unit"

                                If boxName = kennzeichnung Then
                                    If englishLanguage Then
                                        boxName = "Business-Unit:"
                                    Else
                                        boxName = "Business-Unit:"
                                    End If
                                    'boxName = repMessages.getmsg(226)
                                End If
                                .TextFrame2.TextRange.Text = boxName & " " & hproj.businessUnit

                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prBusinessUnit
                                qualifier2 = boxName
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "Beschreibung"

                                If boxName = kennzeichnung Then
                                    If englishLanguage Then
                                        boxName = "Beschreibung"
                                    Else
                                        boxName = "Description"
                                    End If
                                    'boxName = repMessages.getmsg(227)
                                End If
                                '.TextFrame2.TextRange.Text = boxName & ": " & hproj.description
                                ' jetzt ohne boxName ...
                                '.TextFrame2.TextRange.Text = boxName & ": " & hproj.description
                                .TextFrame2.TextRange.Text = hproj.description

                                Try
                                    If hproj.variantDescription.Length > 0 Then
                                        ' jetzt ohne boxName
                                        '.TextFrame2.TextRange.Text = boxName & ": " & hproj.description & vbLf & vbLf &
                                        '    "Varianten-Beschreibung: " & hproj.variantDescription

                                        .TextFrame2.TextRange.Text = hproj.description & vbLf & vbLf &
                                            "Varianten-Beschreibung: " & hproj.variantDescription
                                    End If
                                Catch ex As Exception

                                End Try

                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prDescription
                                qualifier2 = boxName
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "SymTrafficLight"

                                ' hier wird das entsprechende Licht gesetzt ...
                                Call switchOnTrafficLightColor(pptShape, hproj.ampelStatus)
                                ' hier wird das Symbol aufgeladen mit der entsprechenden Smart-Info 
                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prSymTrafficLight
                                qualifier2 = ""
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "SymRisks"
                                ' hier wird das Symbol aufgeladen mit der entsprechenden Smart-Info 
                                ' hier wird das Symbol aufgeladen mit der entsprechenden Smart-Info 
                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prSymRisks
                                qualifier2 = ""
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "SymGoals"
                                ' hier wird das Symbol aufgeladen mit der entsprechenden Smart-Info 
                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prSymDescription
                                qualifier2 = ""
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)


                            Case "SymFinance"
                                ' hier wird das Symbol aufgeladen mit der entsprechenden Smart-Info 
                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prSymFinance
                                qualifier2 = ""
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "SymSchedules"
                                ' hier wird das Symbol aufgeladen mit der entsprechenden Smart-Info 
                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prSymSchedules
                                qualifier2 = ""
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "SymTeam"
                                ' hier wird das Symbol aufgeladen mit der entsprechenden Smart-Info 
                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prSymTeam
                                qualifier2 = ""
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "SymProject"
                                ' hier wird das Symbol aufgeladen mit der entsprechenden Smart-Info 
                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prSymProject
                                qualifier2 = ""
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "Stand:"

                                If boxName = kennzeichnung Then
                                    If englishLanguage Then
                                        boxName = "Version:"
                                    Else
                                        boxName = "Stand:"
                                    End If
                                    'boxName = repMessages.getmsg(223)
                                End If

                                .TextFrame2.TextRange.Text = boxName & " " & Date.Now.ToString("d", repCult) & " (DB: " & hproj.timeStamp.ToString("d", repCult) & ")"
                                '.TextFrame2.TextRange.Text = boxName & " " & hproj.timeStamp.ToString("d", repCult)
                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prStand
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "Laufzeit:"

                                If boxName = kennzeichnung Then
                                    If englishLanguage Then
                                        boxName = "Project Time:"
                                    Else
                                        boxName = "Laufzeit:"
                                    End If
                                    'boxName = repMessages.getmsg(228)
                                End If
                                .TextFrame2.TextRange.Text = boxName & " " & textZeitraum(hproj.startDate, hproj.endeDate)

                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prLaufzeit
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)

                            Case "Verantwortlich:"

                                If boxName = kennzeichnung Then
                                    If englishLanguage Then
                                        boxName = "Verantwortlich:"
                                    Else
                                        boxName = "Responsible:"
                                    End If
                                    'boxName = repMessages.getmsg(229)
                                End If
                                .TextFrame2.TextRange.Text = boxName & " " & hproj.leadPerson

                                bigType = ptReportBigTypes.components
                                compID = ptReportComponents.prVerantwortlich
                                qualifier2 = boxName
                                Call addSmartPPTCompInfo(pptShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                                          bigType, compID)
                            Case Else
                        End Select



                    End With

                Catch ex As Exception

                    tmpShape.TextFrame2.TextRange.Text = ex.Message & vbLf & tmpShape.Title & ": Fehler ..."

                End Try

            Next





            ' jetzt muss die ListofShapes wieder geleert werden 

            listofShapes.Clear()

            ' jetzt müssen die Hilfs-Shapes, die evtl auf der Folie sind, gelöscht werden 
            If Not IsNothing(gleichShape) Then
                gleichShape.Delete()
                gleichShape = Nothing
            End If

            If Not IsNothing(steigendShape) Then
                steigendShape.Delete()
                steigendShape = Nothing
            End If

            If Not IsNothing(fallendShape) Then
                fallendShape.Delete()
                fallendShape = Nothing
            End If

            If Not IsNothing(ampelShape) Then
                ampelShape.Delete()
                ampelShape = Nothing
            End If

            'Next

            If objectsDone >= objectsToDo Or awinSettings.mppOnePage Then
                folieIX = folieIX + 1
                'pptFirstTime = True  ' damit die Folie für die Legende geholt wird
                'Try
                '    If Not IsNothing(pptCurrentPresentation.Slides("tmpSav")) Then
                '        pptCurrentPresentation.Slides("tmpSav").Delete()   ' Vorlage in passender Größe wird nun nicht mehr benötigt
                '    End If
                'Catch ex As Exception

                'End Try
                objectsToDo = 0
                objectsDone = 0
            End If

        End While ' folieIX <= anzSlidestoAdd


    End Sub


    '''' <summary>
    '''' erstellt Balken und Curve Projekt-Diagramme , Soll-Ist 
    '''' </summary>
    '''' <param name="sCInfo"></param>
    '''' <param name="pptAppl"></param>
    '''' <param name="presentationName"></param>
    '''' <param name="currentSlideName"></param>
    '''' <param name="chartContainer"></param>
    Public Sub createProjektChartInPPTNew(ByVal sCInfo As clsSmartPPTChartInfo,
                                      ByVal chartContainer As PowerPoint.Shape)

        ' Festlegen der Titel Schriftgrösse
        Dim titleFontSize As Single = 14
        If chartContainer.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
            titleFontSize = chartContainer.TextFrame2.TextRange.Font.Size
        End If



        ' Parameter Definitionen
        Dim top As Single = chartContainer.Top
        Dim left As Single = chartContainer.Left
        Dim height As Single = chartContainer.Height
        Dim width As Single = chartContainer.Width

        ' tk 5.10.19 hier nicht notwendig , weil in ppt
        'Dim currentPresentation As PowerPoint.Presentation = pptAppl.Presentations.Item(presentationName)
        'Dim currentSlide As PowerPoint.Slide = currentPresentation.Slides.Item(currentSlideName)

        Dim diagramTitle As String = " "
        Dim IstCharttype As Microsoft.Office.Core.XlChartType
        Dim PlanChartType As Microsoft.Office.Core.XlChartType
        Dim vglChartType As Microsoft.Office.Core.XlChartType

        Dim considerIstDaten As Boolean = False

        ' tk 19.4.19 wenn es sich um ein Portfolio handelt, dann muss rausgefunden werden, was der kleinste Ist-Daten-Value ist 
        If sCInfo.prPF = ptPRPFType.portfolio Then
            considerIstDaten = (ShowProjekte.actualDataUntil > StartofCalendar.AddMonths(showRangeLeft - 1))
        ElseIf sCInfo.prPF = ptPRPFType.project Then
            considerIstDaten = sCInfo.hproj.actualDataUntil > sCInfo.hproj.startDate
        End If



        If sCInfo.chartTyp = PTChartTypen.CurveCumul Then
            IstCharttype = Microsoft.Office.Core.XlChartType.xlArea

            If considerIstDaten Then
                PlanChartType = Microsoft.Office.Core.XlChartType.xlArea
            Else
                PlanChartType = Microsoft.Office.Core.XlChartType.xlLine
            End If

            vglChartType = Microsoft.Office.Core.XlChartType.xlLine
        Else
            IstCharttype = Microsoft.Office.Core.XlChartType.xlColumnStacked
            PlanChartType = Microsoft.Office.Core.XlChartType.xlColumnStacked
            vglChartType = Microsoft.Office.Core.XlChartType.xlLine
        End If

        Dim plen As Integer
        Dim pstart As Integer

        Dim Xdatenreihe() As String = Nothing
        Dim tdatenreihe() As Double = Nothing
        Dim istDatenReihe() As Double = Nothing

        Dim prognoseDatenReihe() As Double = Nothing
        Dim vdatenreihe() As Double = Nothing
        Dim internKapaDatenreihe() As Double = Nothing
        Dim vDatensumme As Double = 0.0
        Dim tDatenSumme As Double


        Dim pkIndex As Integer = CostDefinitions.Count


        Dim titelTeile(1) As String
        Dim titelTeilLaengen(1) As Integer
        Dim tmpcollection As New Collection

        Dim found As Boolean = False

        Dim pname As String = sCInfo.pName



        '
        ' hole die Projektdauer; berücksichtigen: die können unterschiedlich starten und unterschiedlich lang sein
        ' deshalb muss die Zeitspanne bestimmt werden, die beides umfasst  
        '

        Call bestimmePstartPlen(sCInfo, pstart, plen)


        ' hier werden die Istdaten, die Prognosedaten, die Vergleichsdaten sowie die XDaten bestimmt
        Dim errMsg As String = ""
        Call bestimmeXtipvDatenreihen(pstart, plen, sCInfo,
                                       Xdatenreihe, tdatenreihe, vdatenreihe, istDatenReihe, prognoseDatenReihe, internKapaDatenreihe, errMsg)

        If errMsg <> "" Then
            ' es ist ein Fehler aufgetreten
            If chartContainer.HasTextFrame = Microsoft.Office.Core.MsoTriState.msoTrue Then
                chartContainer.TextFrame2.TextRange.Text = errMsg
            End If
            Exit Sub
        End If

        ' jetzt die Farbe bestimme
        Dim balkenFarbe As Integer = bestimmeBalkenFarbe(sCInfo)


        Dim vProjDoesExist As Boolean = Not IsNothing(sCInfo.vglProj)

        If sCInfo.chartTyp = PTChartTypen.CurveCumul Then
            tDatenSumme = tdatenreihe(tdatenreihe.Length - 1)
            vDatensumme = vdatenreihe(vdatenreihe.Length - 1)
        Else
            tDatenSumme = tdatenreihe.Sum
            vDatensumme = vdatenreihe.Sum

        End If

        Dim startRed As Integer = 0
        Dim lengthRed As Integer = 0
        diagramTitle = bestimmeChartDiagramTitle(sCInfo, tDatenSumme, vDatensumme, startRed, lengthRed)

        ' jetzt wird das Diagramm in Powerpoint erzeugt ...
        Dim newPPTChart As PowerPoint.Shape = currentSlide.Shapes.AddChart(Left:=left, Top:=top, Width:=width, Height:=height)
        'Dim newPPTChart As PowerPoint.Shape = currentSlide.Shapes.AddChart(Type:=Microsoft.Office.Core.XlChartType.xlColumnStacked, Left:=left, Top:=top,
        '                                                           Width:=width, Height:=height)
        ' 
        ' tk brauchen wir das ?  
        'Dim tmpWB As Excel.Workbook = CType(newPPTChart.Chart.ChartData.Workbook, Excel.Workbook)


        ' jetzt kommt das Löschen der alten SeriesCollections . . 
        With newPPTChart.Chart
            Try
                Dim anz As Integer = CInt(CType(.SeriesCollection, PowerPoint.SeriesCollection).Count)
                Do While anz > 0
                    .SeriesCollection(1).Delete()
                    anz = anz - 1
                Loop
            Catch ex As Exception

            End Try
        End With

        ' Start
        Try

            If Not IsNothing(newPPTChart.Chart.ChartData) Then


                With newPPTChart.Chart.ChartData

                    .Workbook.Application.Visible = False
                    .Workbook.Application.Width = 50
                    .Workbook.Application.Height = 15
                    .Workbook.Application.Top = 10
                    .Workbook.Application.Left = -120
                    .Workbook.Application.WindowState = -4140 '## Minimize Excel
                End With


            End If

        Catch ex As Exception

        End Try

        ' Ende 


        ' jetzt werden die Collections in dem Chart aufgebaut ...
        With newPPTChart.Chart

            Dim dontShowPlanung As Boolean = False

            If sCInfo.prPF = ptPRPFType.portfolio Then
                If ShowProjekte.actualDataUntil >= StartofCalendar Then
                    dontShowPlanung = getColumnOfDate(ShowProjekte.actualDataUntil) >= showRangeRight
                End If

            Else
                If sCInfo.hproj.hasActualValues Then
                    dontShowPlanung = getColumnOfDate(sCInfo.hproj.actualDataUntil) >= getColumnOfDate(sCInfo.hproj.endeDate)
                End If
            End If


            If Not dontShowPlanung Then
                With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                    If sCInfo.prPF = ptPRPFType.portfolio Then
                        .Name = bestimmeLegendNameIPB("PS") & Date.Now.ToShortDateString
                        .Interior.Color = balkenFarbe
                    Else
                        .Name = bestimmeLegendNameIPB("P") & sCInfo.hproj.timeStamp.ToShortDateString
                        .Interior.Color = visboFarbeBlau
                    End If

                    .Values = prognoseDatenReihe
                    .XValues = Xdatenreihe
                    .ChartType = PlanChartType

                    If sCInfo.chartTyp = PTChartTypen.CurveCumul And Not considerIstDaten Then
                        ' es handelt sich um eine Line
                        .Format.Line.Weight = 4
                        .Format.Line.ForeColor.RGB = visboFarbeBlau
                        .Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
                    End If

                End With
            End If

            ' Beauftragung bzw. Vergleichsdaten
            If sCInfo.prPF = ptPRPFType.portfolio Then
                'series
                With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                    .Name = bestimmeLegendNameIPB("C")
                    .Values = vdatenreihe
                    .XValues = Xdatenreihe

                    .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                    With .Format.Line
                        .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid
                        .ForeColor.RGB = visboFarbeRed
                        .Weight = 2
                    End With


                End With

                Dim tmpSum As Double = internKapaDatenreihe.Sum
                If vdatenreihe.Sum > tmpSum And tmpSum > 0 Then
                    ' es gibt geplante externe Ressourcen ... 
                    With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
                        .HasDataLabels = False
                        '.name = "Kapazität incl. Externe"
                        .Name = bestimmeLegendNameIPB("CI")
                        '.Name = repMessages.getmsg(118)

                        .Values = internKapaDatenreihe
                        .XValues = Xdatenreihe
                        .ChartType = Microsoft.Office.Core.XlChartType.xlLine
                        With .Format.Line
                            .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineSysDot
                            .ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbFuchsia
                            .Weight = 2
                        End With

                    End With
                End If

            Else
                If Not IsNothing(sCInfo.vglProj) Then

                    'series
                    With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)

                        .Name = bestimmeLegendNameIPB("B") & sCInfo.vglProj.timeStamp.ToShortDateString
                        .Values = vdatenreihe
                        .XValues = Xdatenreihe

                        .ChartType = vglChartType

                        If vglChartType = Microsoft.Office.Core.XlChartType.xlLine Then
                            With .Format.Line
                                .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
                                .ForeColor.RGB = visboFarbeOrange
                                .Weight = 4
                            End With
                        Else
                            ' ggf noch was definieren ..
                        End If

                    End With

                End If
            End If


            ' jetzt kommt der Neu-Aufbau der Series-Collections
            If considerIstDaten Then

                ' jetzt die Istdaten zeichnen 
                With CType(CType(.SeriesCollection, PowerPoint.SeriesCollection).NewSeries, PowerPoint.Series)
                    If sCInfo.prPF = ptPRPFType.portfolio Then
                        .Name = bestimmeLegendNameIPB("IS")
                    Else
                        .Name = bestimmeLegendNameIPB("I")
                    End If
                    .Interior.Color = awinSettings.SollIstFarbeArea
                    .Values = istDatenReihe
                    .XValues = Xdatenreihe
                    .ChartType = IstCharttype
                End With

            End If


        End With

        ' ---- ab hier Achsen und Überschrift setzen 

        With CType(newPPTChart.Chart, PowerPoint.Chart)
            '
            .HasAxis(PowerPoint.XlAxisType.xlCategory) = True
            .HasAxis(PowerPoint.XlAxisType.xlValue) = True

            .SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementPrimaryValueAxisShow)

            Try
                With CType(.Axes(PowerPoint.XlAxisType.xlCategory), PowerPoint.Axis)

                    .HasTitle = False
                    If titleFontSize - 4 >= 6 Then
                        .TickLabels.Font.Size = titleFontSize - 4
                    Else
                        .TickLabels.Font.Size = 6
                    End If


                    ' tk 9.7.19 führt zu Fehler
                    'If .Format.TextFrame2.HasText = MsoTriState.msoCTrue Then
                    '    If titleFontSize - 4 >= 6 Then
                    '        .Format.TextFrame2.TextRange.Font.Size = titleFontSize - 4
                    '    Else
                    '        .Format.TextFrame2.TextRange.Font.Size = 6
                    '    End If
                    'End If

                End With
            Catch ex As Exception

            End Try

            Try
                With CType(.Axes(PowerPoint.XlAxisType.xlValue), PowerPoint.Axis)

                    .HasTitle = False
                    .MinimumScale = 0

                    If titleFontSize - 4 >= 6 Then
                        .TickLabels.Font.Size = titleFontSize - 4
                    Else
                        .TickLabels.Font.Size = 6
                    End If

                    ' führt immer zu Fehler 
                    'If .Format.TextFrame2.HasText = MsoTriState.msoCTrue Then
                    '    If titleFontSize - 4 >= 6 Then
                    '        .Format.TextFrame2.TextRange.Font.Size = titleFontSize - 4
                    '    Else
                    '        .Format.TextFrame2.TextRange.Font.Size = 6
                    '    End If
                    'End If

                End With
            Catch ex As Exception

            End Try

            Try
                .HasLegend = True
                With .Legend
                    .Position = PowerPoint.XlLegendPosition.xlLegendPositionTop

                    If titleFontSize - 4 >= 6 Then
                        .Font.Size = titleFontSize - 4
                    Else
                        .Font.Size = 6
                    End If

                End With
            Catch ex As Exception

            End Try

            .HasTitle = True
            .ChartTitle.Text = " " ' Platzhalter 

        End With

        ' 
        ' ---- hier dann final den Titel setzen 
        With newPPTChart.Chart
            .HasTitle = True
            .ChartTitle.Text = diagramTitle
            .ChartTitle.Font.Size = titleFontSize
            .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbBlack

            If startRed > 0 And lengthRed > 0 Then
                ' die aktuelle Summe muss rot eingefärbt werden 
                .ChartTitle.Format.TextFrame2.TextRange.Characters(startRed,
                    lengthRed).Font.Fill.ForeColor.RGB = Microsoft.Office.Interop.PowerPoint.XlRgbColor.rgbRed
            End If

        End With

        newPPTChart.Chart.Refresh()

        ' jetzt das Excel wieder schliessen 
        ' tk brauchen wir das ? 8.10.19
        'tmpWB.Close(SaveChanges:=False)
        '
        ' jetzt werden die Smart-Infos an das Chart angehängt ...

        Call addSmartPPTChartInfo(newPPTChart, sCInfo)


    End Sub



    ''' <summary>
    ''' zeichnet sowohl Swimlanes im BHTC Modus als auch im Normal -Modus
    ''' BHTC: Segmente customer Milestones, BHTC Milestones  
    ''' normal: Swimlane ist alles auf Hierarchie-Ebene 1 (also die Kinder der rootphase Ebene) 
    ''' es wird immer nur ein Projekt betrachtet 
    ''' es können x Swimlanes sein - es muss unterschieden werden, ob alles auf eine Seite geht oder mehrere Seiten gemacht werden 
    ''' Rahmenbedingung bei dieser Routine: es wird nur ein Project aufgerufen, ohne Varianten 
    ''' es geht also nur darum , alle Swimlanes eines Projektes zu zeichnen bzw. die ausgewählten Swimlanes eines PRojektes zu zeichnen  
    ''' </summary>
    ''' <param name="swimLanesToDo"></param>
    ''' <param name="swimLanesDone"></param>
    ''' <param name="pptFirstTime"></param>
    ''' <param name="zeilenhoehe"></param>
    ''' <param name="legendFontSize"></param>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="selectedRoles"></param>
    ''' <param name="selectedCosts"></param>
    ''' <param name="selectedBUs"></param>
    ''' <param name="selectedTyps"></param>
    ''' <param name="isMultiprojektSicht"></param>
    ''' <param name="hproj"></param>
    ''' <param name="kennzeichnung"></param>
    ''' <remarks></remarks>
    Private Sub zeichneSwimlane2SichtinPPT(ByRef swimLanesToDo As Integer, ByRef swimLanesDone As Integer, ByRef pptFirstTime As Boolean,
                                                 ByRef zeilenhoehe As Double, ByRef legendFontSize As Double,
                                                 ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection,
                                                 ByVal selectedRoles As Collection, ByVal selectedCosts As Collection,
                                                 ByVal selectedBUs As Collection, ByVal selectedTyps As Collection,
                                                 ByVal isMultiprojektSicht As Boolean, ByVal hproj As clsProjekt,
                                                 ByVal kennzeichnung As String,
                                                 ByVal minCal As Boolean)


        ' Wichtig für Kalendar 
        Dim pptStartofCalendar As Date = Nothing, pptEndOfCalendar As Date = Nothing
        Dim errorShape As PowerPoint.Shape = Nothing

        Dim curFormatSize(1) As Double

        Dim maxZeilen As Integer = 0
        Dim anzZeilen As Integer = 0
        Dim gesamtAnzZeilen As Integer = 0

        Dim msgTxt As String


        ' Ende Übernahme

        Dim format As Integer = 4
        'Dim tmpslideID As Integer

        ' an der Variablen lässt sich in der Folge erkennen, ob die Segmente BHTC Milestones gezeichnet werden müssen oder 
        ' ob ganz allgemein nach Swimlanes gesucht wird ... 
        Dim isSwimlanes2 As Boolean = (kennzeichnung = "Swimlanes2")

        Dim rds As New clsPPTShapes
        Dim considerZeitraum As Boolean = (showRangeLeft > 0 And showRangeRight >= showRangeLeft)
        Dim cphase As clsPhase

        ' mit disem Befehl werden auch die ganzen Hilfsshapes in der Klasse gesetzt , sofern bereits vorhanden ..
        ' das Ganze funktioniert also noch mit alten Vorlagen wie mit neuen ... 
        rds.pptSlide = currentSlide


        ' jetzt werden die noch fehlenden Shapes erstellt .. 
        If rds.getMissingShpNames(kennzeichnung).Count > 0 Then
            Dim msHeight As Single = 9.0
            Dim phHeight As Single = 5.6
            Call rds.createMandatoryDrawingShapes(kennzeichnung, msHeight, phHeight)
        End If



        ' jetzt muss geprüft werden, ob überhaupt alle Angaben gemacht wurden ... 
        'If completeMppDefinition.Sum = completeMppDefinition.Length Then
        Dim missingShapes As String = rds.getMissingShpNames(kennzeichnung)
        If missingShapes.Length = 0 Then
            ' es fehlt nichts ... andernfalls stehen hier die Namen mit den Shapes, die fehlen ...

            Dim considerAll As Boolean = (selectedPhases.Count + selectedMilestones.Count + selectedRoles.Count + selectedCosts.Count = 0)
            Dim selectedPhaseIDs As New Collection
            Dim selectedMilestoneIDs As New Collection
            Dim breadcrumbArray As String() = Nothing

            If Not considerAll Then
                Dim tmpPhaseIDs As New Collection
                If selectedPhases.Count > 0 Then
                    selectedPhaseIDs = hproj.getElemIdsOf(selectedPhases, False)
                End If
                If selectedRoles.Count > 0 Then
                    tmpPhaseIDs = hproj.getPhaseIdsWithRoleCost(selectedRoles, True)
                    For Each tmpPhaseID As String In tmpPhaseIDs
                        If Not selectedPhaseIDs.Contains(tmpPhaseID) Then
                            selectedPhaseIDs.Add(tmpPhaseID, tmpPhaseID)
                        End If
                    Next
                End If
                If selectedCosts.Count > 0 Then
                    tmpPhaseIDs = hproj.getPhaseIdsWithRoleCost(selectedCosts, False)
                    For Each tmpPhaseID As String In tmpPhaseIDs
                        If Not selectedPhaseIDs.Contains(tmpPhaseID) Then
                            selectedPhaseIDs.Add(tmpPhaseID, tmpPhaseID)
                        End If
                    Next
                End If

                selectedMilestoneIDs = hproj.getElemIdsOf(selectedMilestones, True)
                breadcrumbArray = hproj.getBreadCrumbArray(selectedPhaseIDs, selectedMilestoneIDs)
            End If

            ' Änderung tk 23.2.16: wenn mehrere Projekte mit swimlanes gezeichnet werden, so muss hier bestimmt werden
            ' wieviele Swimlanes zu zeichnen sind; ab dem 2. Projekt kann man sich nicht mehr auf pptFirsttime abstützen ! 
            ' wenn ein Projekt erstmalig hier reinkommt, ist swimlanestodo = 0, pptFirsttime kann true oder false sein   
            If swimLanesToDo = 0 Then
                swimLanesToDo = hproj.getSwimLanesCount(considerAll, breadcrumbArray, isSwimlanes2)
            End If


            ' wenn Kalenderlinie oder Legend-Linie über Container Grenzen gehen, werden die Koordinaten der Lines entsprechend angepasst 
            Call rds.plausibilityAdjustments()

            ' ermittelt die Koordinaten für Kalender, linker Rand Projektbeschriftung, Projekt-Fläche, Legenden-Fläche
            Call rds.bestimmeZeichenKoordinaten()

            Dim projCollection As New SortedList(Of Double, String)
            Dim minDate As Date, maxDate As Date


            ' bestimmt für den angegebenen Zeitraum die Projekte, die eine der angegeben Phasen oder Meilensteine im Zeitraum enthalten. 
            ' bestimmt darüber hinaus das minimale bzw. maximale Datum , das die Phasen der Projekte aufspannen , die den Zeitraum "berühren"  
            Call bestimmeProjekteAndMinMaxDates(selectedPhases, selectedMilestones,
                                                selectedRoles, selectedCosts,
                                                selectedBUs, selectedTyps,
                                                showRangeLeft, showRangeRight, awinSettings.mppSortiertDauer,
                                                projCollection, minDate, maxDate,
                                                isMultiprojektSicht, False, hproj)


            ' wird benötigt für die Bestimmung der Anzahl zielen und das Zeichnen der Swimlane Phase / Meilensteine
            ' wenn mppshowallIFOne = false, dann sollte zeitRaumGrenzeL = showrangeL und zeitRaumGrenzeR = showrangeR
            ' andernfalls ist der Zeitraum ggf. deutlich größer als Showrange 
            Dim zeitraumGrenzeL As Integer
            Dim zeitraumGrenzeR As Integer

            If awinSettings.mppShowAllIfOne Then

                zeitraumGrenzeL = getColumnOfDate(minDate)
                zeitraumGrenzeR = getColumnOfDate(maxDate)

            Else

                zeitraumGrenzeL = showRangeLeft
                zeitraumGrenzeR = showRangeRight

            End If



            ' tk:1.2.16 ExtendedMode macht nur Sinn, wenn mindestens 1 Phase selektiert wurde. oder aber considerAll gilt: 
            ' dabei müssen aber auch die selectedPhaseIDs berücksichtigt werden 
            awinSettings.mppExtendedMode = (awinSettings.mppExtendedMode And (selectedPhases.Count > 0 Or selectedPhaseIDs.Count > 0)) Or
                                            (awinSettings.mppExtendedMode And considerAll)



            ' muss nur bestimmt werden, wenn zum ersten Mal reinkommt 


            '
            ' bestimme das Start und Ende Datum des PPT Kalenders
            Call calcStartEndePPTKalender(minDate, maxDate,
                                          pptStartofCalendar, pptEndOfCalendar)

            ' jetzt für Swimlanes Behandlung Kalender in der Klasse setzen 

            Call rds.setCalendarDates(pptStartofCalendar, pptEndOfCalendar)

            ' die neue Art Zeilenhöhe und die Offset Werte zu bestimmen 
            ' dabei muss berücksichtigt werden dass selectedPhases.count = 0 sein kann, aber selectedPhaseIDs.count > 0 

            Call rds.bestimmeZeilenHoehe(System.Math.Max(selectedPhases.Count, selectedPhaseIDs.Count),
                                         System.Math.Max(selectedMilestones.Count, hproj.getPhase(1).countMilestones), considerAll)


            ' tk 11.10.19
            ' jetzt muss ermittelt werden, ob bei der angegebenen Zeilenhöhe alle Swimlanes gezeichnet werden können
            ' wenn nein, wird die Zeilenhöhe entsprechend reduziert , so dass alles in den Container passt 
            ' jetzt muss die Gesamt-Zahl an Zeilen ermittelt werden , die die einzelnen Swimlanes bentötigen 

            If awinSettings.mppExtendedMode Then

                For i = 1 To swimLanesToDo

                    cphase = hproj.getSwimlane(i, considerAll, breadcrumbArray, isSwimlanes2)

                    Dim segmentID As String = ""
                    If isSwimlanes2 Then
                        If hproj.isSegment(cphase.nameID) Then
                            segmentID = cphase.nameID
                        End If
                    End If
                    Dim swimLaneZeilen As Integer = hproj.calcNeededLinesSwl(cphase.nameID, selectedPhaseIDs, selectedMilestoneIDs,
                                                                                 awinSettings.mppExtendedMode,
                                                                                 considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR,
                                                                                 considerAll, segmentID)

                    anzZeilen = anzZeilen + swimLaneZeilen
                Next

            Else
                anzZeilen = swimLanesToDo
            End If

            Dim neededSpace As Double
            Dim anzSegments As Integer = 0


            If isSwimlanes2 Then
                ' jetzt müssen noch die Segment Höhen  berechnet werden 
                anzSegments = hproj.getSegmentsCount(considerAll, breadcrumbArray, isSwimlanes2)
                neededSpace = anzZeilen * rds.zeilenHoehe +
                                    anzSegments * rds.segmentHoehe
            Else
                neededSpace = anzZeilen * rds.zeilenHoehe
            End If

            Dim weitermachen As Boolean = True


            ' jetzt muss die Zeilenhöhe solange reduziert werden, bis alles reinpasst oder aber es gar nicht geht ... 
            If neededSpace > rds.availableSpace Then
                ' reduzieren der Zeilenhöhe 
                weitermachen = False

                ' zuerst die Beschriftungen raus nehmen 
                If awinSettings.mppShowMsDate Or awinSettings.mppShowMsName Or awinSettings.mppShowPhDate Or awinSettings.mppShowPhName Then
                    With awinSettings
                        .mppShowMsDate = False
                        .mppShowMsName = False
                        .mppShowPhDate = False
                        .mppShowPhName = False
                    End With
                End If

                ' dann die Zeilenhöhe auf das Minimum 1: Platz für eine Beschriftung, oben oder unten   setzen ...

                Call rds.setZeilenhöhe(absoluteMinimum:=False)

                neededSpace = anzZeilen * rds.zeilenHoehe +
                                    anzSegments * rds.segmentHoehe


                ' immer noch zu viel Platz benötigt ? 
                If neededSpace > rds.availableSpace Then
                    ' auf Minimum 2 setzen 

                    Call rds.setZeilenhöhe(absoluteMinimum:=True)

                    neededSpace = anzZeilen * rds.zeilenHoehe +
                                    anzSegments * rds.segmentHoehe

                    If neededSpace > rds.availableSpace Then
                        weitermachen = False
                    Else
                        ' alles in ORdnung 
                        weitermachen = True
                    End If
                Else
                    ' alles in Ordnung ...
                    weitermachen = True
                End If


            End If

            If weitermachen Then
                If pptFirstTime Then

                    ' jetzt erst mal den Kalender zeichnen 
                    ' zeichne den Kalender
                    'Dim calendargroup As pptNS.Shape = Nothing

                    Try

                        With rds

                            Call draw3RowsCalendar(rds, minCal)
                            'Call zeichne3RowsCalendar(rds, minCal)

                        End With



                    Catch ex As Exception

                    End Try


                    Dim smartInfoCRD As Date = Date.MinValue
                    ' jetzt wird hier die Date Info eingetragen ... 
                    Try
                        For Each kvp As KeyValuePair(Of Double, String) In projCollection
                            Dim tmpProj As clsProjekt = AlleProjekte.getProject(kvp.Value)
                            If Not IsNothing(tmpProj) Then
                                If smartInfoCRD < tmpProj.timeStamp Then
                                    smartInfoCRD = tmpProj.timeStamp
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try

                    ' 
                    ' jetzt wird das Slide gekennzeichnet als Smart Slide 
                    ' eigentlich müsste das ContainerShpae gezeichnet werden , nicht die Seite 
                    Call addSmartPPTSlideCalInfo(rds.pptSlide, rds.PPTStartOFCalendar, rds.PPTEndOFCalendar, smartInfoCRD)

                End If



                ' hier ist die Schleife, die alle swimlanes von swimlanesdone+1 bis todo zeichnet 
                ' jetzt wird das aufgerufen mit dem gesamten fertig gezeichneten Kalender, der fertig positioniert ist 

                Dim curYPosition As Double = rds.drawingAreaTop
                Dim curSwl As clsPhase
                Dim prevSwl As clsPhase = Nothing

                ' steuert im Wechsel, dass eine Zeilendifferenzierung gezeichnet wird / nicht gezeichnet wird 
                ' hat nur dann einen Effekt, wenn rds.rowDifferentiator <> Nothing 

                Dim toggleRow As Boolean = False

                Dim curSwimlaneIndex As Integer = swimLanesDone + 1
                curSwl = hproj.getSwimlane(curSwimlaneIndex, considerAll, breadcrumbArray, isSwimlanes2)
                prevSwl = hproj.getSwimlane(curSwimlaneIndex - 1, considerAll, breadcrumbArray, isSwimlanes2)


                Dim curSegmentID As String = ""

                If Not IsNothing(curSwl) Then
                    ' wird weiter unten auch noch gebraucht 
                    Dim segmentChanged As Boolean = False

                    If isSwimlanes2 Then

                        If hproj.isSegment(curSwl.nameID) Then
                            curSegmentID = curSwl.nameID
                        Else
                            curSegmentID = hproj.hierarchy.getParentIDOfID(curSwl.nameID)
                        End If


                        If Not IsNothing(prevSwl) Then
                            segmentChanged = hproj.hierarchy.getParentIDOfID(prevSwl.nameID) <>
                                            hproj.hierarchy.getParentIDOfID(curSwl.nameID)
                        End If

                        If swimLanesDone = 0 Or segmentChanged Then
                            Call zeichneSwlSegmentinAktZeile(rds, curYPosition, curSegmentID)
                            segmentChanged = False
                        End If
                    End If




                    ' jetzt werden soviele wie möglich Swimlanes gezeichnet ... 
                    Dim swimLaneZeilen As Integer = hproj.calcNeededLinesSwl(curSwl.nameID, selectedPhaseIDs, selectedMilestoneIDs,
                                                                                 awinSettings.mppExtendedMode,
                                                                                 considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR,
                                                                                 considerAll, curSegmentID)

                    Do While (curSwimlaneIndex <= swimLanesToDo) And
                        (swimLaneZeilen * rds.zeilenHoehe + curYPosition <= rds.drawingAreaBottom)


                        ' jetzt die Swimlane zeichnen
                        ' hier ist ja gewährleistet, dass alle Phasen und Meilensteine dieser Swimlane Platz finden 
                        Call zeichneSwimlaneOfProject(rds, curYPosition, toggleRow,
                                                  hproj, curSwl.nameID, considerAll,
                                                  breadcrumbArray,
                                                  considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR,
                                                  selectedPhaseIDs, selectedMilestoneIDs,
                                                  selectedRoles, selectedCosts,
                                                  swimLaneZeilen, curSegmentID)

                        ' merken, ob die letzte gezeichnete Swimlane eigentlich die Meilensteine des Segments waren ...
                        Dim lastSwimlaneWasSegment As Boolean = isSwimlanes2 And (curSwl.nameID = curSegmentID)


                        prevSwl = curSwl

                        curSwimlaneIndex = curSwimlaneIndex + 1
                        curSwl = hproj.getSwimlane(curSwimlaneIndex, considerAll, breadcrumbArray, isSwimlanes2)

                        If Not IsNothing(curSwl) Then

                            Dim segmentID As String = ""
                            If isSwimlanes2 Then
                                segmentChanged = (hproj.hierarchy.getParentIDOfID(prevSwl.nameID) <>
                                        hproj.hierarchy.getParentIDOfID(curSwl.nameID) And Not lastSwimlaneWasSegment) Or
                                        (hproj.isSegment(prevSwl.nameID) And hproj.isSegment(curSwl.nameID))

                                If hproj.isSegment(curSwl.nameID) Then
                                    segmentID = curSwl.nameID
                                End If
                            End If


                            swimLaneZeilen = hproj.calcNeededLinesSwl(curSwl.nameID, selectedPhaseIDs, selectedMilestoneIDs,
                                                                                 awinSettings.mppExtendedMode,
                                                                                 considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR,
                                                                                 considerAll, segmentID)

                            If isSwimlanes2 Then
                                If segmentChanged And
                                (swimLaneZeilen * rds.zeilenHoehe + curYPosition + rds.segmentVorlagenShape.Height <= rds.drawingAreaBottom) Then

                                    If hproj.isSegment(curSwl.nameID) Then
                                        curSegmentID = curSwl.nameID
                                    Else
                                        curSegmentID = hproj.hierarchy.getParentIDOfID(curSwl.nameID)
                                    End If

                                    Call zeichneSwlSegmentinAktZeile(rds, curYPosition, curSegmentID)
                                    segmentChanged = False
                                End If
                            End If

                        Else
                            segmentChanged = False
                        End If


                    Loop

                    If curSwimlaneIndex = swimLanesDone + 1 Then
                        ' es wurde in der Schleife keine Swimmlane gezeichnet, da sie zu groß ist für eine Seite
                        ' Abbruch provoziere
                        ' Zwischenbericht abgeben ...
                        msgTxt = "Swimlane '" & elemNameOfElemID(curSwl.nameID) & "' kann nicht gezeichnet werden; kein Platz  ...."
                        If awinSettings.englishLanguage Then
                            msgTxt = "Swimlane '" & elemNameOfElemID(curSwl.nameID) & "' could not be drawn: not enough space ...."
                        End If

                        Throw New ArgumentException(msgTxt)
                        swimLanesDone = 0
                        swimLanesToDo = 0

                    Else

                        ' jetzt die Anzahl ..Done bestimmen
                        swimLanesDone = curSwimlaneIndex - 1

                    End If

                End If

            Else
                Call MsgBox("not enough space to draw elements  ... ")
            End If


        ElseIf Not IsNothing(rds.errorVorlagenShape) Then
            ''rds.errorVorlagenShape.Copy()
            ''errorShape = pptslide.Shapes.Paste

            errorShape = createPPTShapeFromShape(rds.errorVorlagenShape, currentSlide)
            With errorShape
                .TextFrame2.TextRange.Text = missingShapes
            End With
        End If

        ' jetzt werden alle für das Zeichnen notwendigen Hilfs-Shapes unsichtbar gemacht 
        ' sie können dann beim Ändern des Reports wieder verwendet werden 
        Call rds.setShapesInvisible()

        ' jezt wird das containershape in den Hintergrund gesetzt 
        Call rds.containerShape.ZOrder(MsoZOrderCmd.msoSendToBack)


    End Sub


    ''' <summary>
    ''' zeichnet den Multiprojekt Sicht Container
    ''' </summary>
    ''' <param name="objectsToDo"></param>
    ''' <param name="objectsDone"></param>
    ''' <param name="pptFirstTime"></param>
    ''' <param name="zeilenhoehe_sav"></param>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="selectedRoles"></param>
    ''' <param name="selectedCosts"></param>
    ''' <param name="selectedBUs"></param>
    ''' <param name="isMultiprojektSicht">gibt an, ob es sich um eine Einzelprojekt/Varianten Sicht oder 
    ''' um eine Multiprojektsicht handelt </param>
    ''' <param name="isMultivariantenSicht">nur relevant, wenn multiprojektsicht = false; gibt an ob es sich um eine Multivariantensicht oder 
    ''' eine Einzelprojeksicht handelt </param>
    ''' <param name="projMitVariants">das Projekt, dessen Varianten alle dargestellt werden sollen; nur besetzt wenn isMultiprojektSicht = false</param>
    ''' <remarks></remarks>
    Private Sub drawMultiprojectViewinPPT(ByRef objectsToDo As Integer, ByRef objectsDone As Integer, ByRef pptFirstTime As Boolean,
                                             ByRef zeilenhoehe_sav As Double, ByRef legendFontSize As Double,
                                             ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection,
                                             ByVal selectedRoles As Collection, ByVal selectedCosts As Collection,
                                             ByVal selectedBUs As Collection, ByVal selectedTyps As Collection,
                                             ByVal isMultiprojektSicht As Boolean,
                                             ByVal isMultivariantenSicht As Boolean, ByVal projMitVariants As clsProjekt,
                                             ByVal kennzeichnung As String,
                                             ByVal minCal As Boolean)



        ' ur:5.10.2015: ExtendedMode macht nur Sinn, wenn mindestens 1 Phase selektiert wurde. deshalb diese Code-Zeile
        awinSettings.mppExtendedMode = awinSettings.mppExtendedMode And (selectedPhases.Count > 0)


        ' Wichtig für Kalendar 
        Dim pptStartofCalendar As Date = Nothing, pptEndOfCalendar As Date = Nothing
        Dim errorShape As PowerPoint.Shape = Nothing



        Dim format As Integer = 4
        'Dim tmpslideID As Integer



        Dim rds As New clsPPTShapes
        rds.pptSlide = currentSlide

        ' jetzt werden die Hilfs-Shapes erstellt .. 


        If rds.getMissingShpNames(kennzeichnung).Count > 0 Then
            Dim msHeight As Single = 9
            Dim phHeight As Single = 5.6
            Call rds.createMandatoryDrawingShapes(kennzeichnung, msHeight, phHeight)
        End If


        ' jetzt muss geprüft werden, ob überhaupt alle Angaben gemacht wurden ... 
        'If completeMppDefinition.Sum = completeMppDefinition.Length Then
        Dim missingShapes As String = rds.getMissingShpNames(kennzeichnung)



        If missingShapes.Length = 0 Then
            ' es fehlt nichts ... andernfalls stehen hier die Namen mit den Shapes, die fehlen ...


            ' wenn Kalenderlinie oder Legendenlinie über Container rausragt: anpassen ! 
            Call rds.plausibilityAdjustments()

            Call rds.bestimmeZeichenKoordinaten()

            Dim projCollection As New SortedList(Of Double, String)
            Dim minDate As Date, maxDate As Date

            Dim considerAll As Boolean = (selectedPhases.Count + selectedMilestones.Count = 0)

            ' bestimme die Projekte, die gezeichnet werden sollen
            ' und bestimme das kleinste / resp größte auftretende Datum 
            Call bestimmeProjekteAndMinMaxDates(selectedPhases, selectedMilestones,
                                                selectedRoles, selectedCosts,
                                                selectedBUs, selectedTyps,
                                                showRangeLeft, showRangeRight, awinSettings.mppSortiertDauer,
                                                projCollection, minDate, maxDate,
                                                isMultiprojektSicht, isMultivariantenSicht, projMitVariants)



            If objectsToDo <> projCollection.Count Then
                objectsToDo = projCollection.Count
            End If


            '
            ' bestimme das Start und Ende Datum des PPT Kalenders
            Call calcStartEndePPTKalender(minDate, maxDate,
                                          pptStartofCalendar, pptEndOfCalendar)


            ' jetzt für Swimlanes Behandlung Kalender in der Klasse setzen 


            Call rds.setCalendarDates(pptStartofCalendar, pptEndOfCalendar)


            ' bestimme die benötigte Höhe einer Zeile im Report ( nur wenn nicht schon bestimmt also zeilenhoehe <> 0
            If pptFirstTime And zeilenhoehe_sav = 0.0 Then

                Call rds.bestimmeZeilenHoehe(selectedPhases.Count, selectedMilestones.Count, considerAll)
                zeilenhoehe_sav = rds.zeilenHoehe
                ' tk alt: 26.11.16
                'With rds

                '    zeilenhoehe = bestimmeMppZeilenHoehe(.pptSlide, .phaseVorlagenShape, .milestoneVorlagenShape,
                '                                        selectedPhases.Count, selectedMilestones.Count, _
                '                                        .MsDescVorlagenShape, .MsDateVorlagenShape, _
                '                                        .PhDescVorlagenShape, .PhDateVorlagenShape,
                '                                        .projectNameVorlagenShape, _
                '                                        .durationArrowShape, .durationTextShape)
                'End With

                ' ur: 1.12.2016
            ElseIf zeilenhoehe_sav <> 0.0 And rds.zeilenHoehe = 0.0 Then

                Call rds.bestimmeZeilenHoehe(selectedPhases.Count, selectedMilestones.Count, considerAll)
                zeilenhoehe_sav = rds.zeilenHoehe
            Else
                Call MsgBox("pptfirstime = " & pptFirstTime.ToString & "; zeilenhoehe_sav = " & zeilenhoehe_sav.ToString)

            End If


            Dim hproj As New clsProjekt
            Dim hhproj As New clsProjekt
            Dim maxZeilen As Integer = 0
            Dim anzZeilen As Integer = 0
            Dim gesamtAnzZeilen As Integer = 0
            Dim projekthoehe As Double = zeilenhoehe_sav

         
            ' neu 14.10.19 
            ' über alle ausgewählte Projekte sehen und maximale Anzahl Zeilen je Projekt bestimmen
            For Each kvp As KeyValuePair(Of Double, String) In projCollection
                Try

                    hproj = AlleProjekte.getProject(kvp.Value)
                Catch ex As Exception

                End Try

                anzZeilen = hproj.calcNeededLines(selectedPhases, selectedMilestones, awinSettings.mppExtendedMode, Not awinSettings.mppShowAllIfOne)

                maxZeilen = System.Math.Max(maxZeilen, anzZeilen)
                gesamtAnzZeilen = gesamtAnzZeilen + anzZeilen

            Next
            ' Ende neu 14.10.19 

            '
            ' bestimme die relativen Abstände der Text-Shapes zu ihrem Phase/Milestone Element
            '
            ' tk 2.11.19 hier muss es mit autoSet aufgerufen werden; siehe oben wenn die manadatory shapes alle gesetzt werden , müssen die relativen Distanzen, wo der Text das Datum gesetzt werden soll angegeben werden ... 
            Call rds.calcRelDisTxtToElm()


            '
            ' bestimme das Format  

            Dim neededSpace As Double
            ' tk 14.10.19 hier soll immer alles auf eine seite gehen .. 
            neededSpace = gesamtAnzZeilen * zeilenhoehe_sav

            ' neu 
            Dim weitermachen As Boolean = True


            ' jetzt muss die Zeilenhöhe solange reduziert werden, bis alles reinpasst oder aber es gar nicht geht ... 
            If neededSpace > rds.availableSpace Then
                ' reduzieren der Zeilenhöhe 
                weitermachen = False

                ' zuerst die Beschriftungen raus nehmen 
                If awinSettings.mppShowMsDate Or awinSettings.mppShowMsName Or awinSettings.mppShowPhDate Or awinSettings.mppShowPhName Then
                    With awinSettings
                        .mppShowMsDate = False
                        .mppShowMsName = False
                        .mppShowPhDate = False
                        .mppShowPhName = False
                    End With
                End If

                ' dann die Zeilenhöhe auf das Minimum 1: Platz für eine Beschriftung, oben oder unten   setzen ...

                Call rds.setZeilenhöhe(absoluteMinimum:=False)

                neededSpace = gesamtAnzZeilen * rds.zeilenHoehe


                ' immer noch zu viel Platz benötigt ? 
                If neededSpace > rds.availableSpace Then
                    ' auf Minimum 2 setzen 
                    Call rds.setZeilenhöhe(absoluteMinimum:=True)
                    neededSpace = gesamtAnzZeilen * rds.zeilenHoehe

                    If neededSpace > rds.availableSpace Then
                        weitermachen = False
                    Else
                        ' alles in ORdnung 
                        weitermachen = True
                    End If
                Else
                    ' alles in Ordnung ...
                    weitermachen = True
                End If
                ' wenn es immer noch nicht reicht, auf Minimum 2, Höhe des Meilensteins *1,1  setzen 

            End If
            ' Ende neu 

            If weitermachen Then

                Try

                    With rds

                        Call draw3RowsCalendar(rds, minCal)

                    End With



                Catch ex As Exception

                End Try


                ' jetzt wird das aufgerufen mit dem gesamten fertig gezeichneten Kalender, der fertig positioniert ist 
                ' zeichne die Projekte 

                ' jetzt wird das Slide gekennzeichnet als Smart Slide

                ' jetzt wird hier die Date Info eingetragen ... 
                Dim smartInfoCRD As Date = Date.MinValue
                Try
                    For Each kvp As KeyValuePair(Of Double, String) In projCollection
                        Dim tmpProj As clsProjekt = AlleProjekte.getProject(kvp.Value)
                        If Not IsNothing(tmpProj) Then
                            If smartInfoCRD < tmpProj.timeStamp Then
                                smartInfoCRD = tmpProj.timeStamp
                            End If
                        End If
                    Next
                Catch ex As Exception

                End Try


                Call addSmartPPTSlideCalInfo(rds.pptSlide, rds.PPTStartOFCalendar, rds.PPTEndOFCalendar, smartInfoCRD)

                Try



                    Call drawProjectsInPPT(projCollection, objectsDone,
                                rds, selectedPhases, selectedMilestones, selectedRoles, selectedCosts, kennzeichnung)


                Catch ex As Exception

                    If Not IsNothing(rds.errorVorlagenShape) Then

                        errorShape = createPPTShapeFromShape(rds.errorVorlagenShape, rds.pptSlide)
                        With errorShape
                            .TextFrame2.TextRange.Text = ex.Message
                        End With
                    Else
                        ' erstmal sonst nichts 
                    End If


                End Try

            Else
                Call MsgBox("not enough space to draw elements  ... ")
            End If


        ElseIf Not IsNothing(rds.errorVorlagenShape) Then

            errorShape = createPPTShapeFromShape(rds.errorVorlagenShape, currentSlide)
            With errorShape
                .TextFrame2.TextRange.Text = missingShapes
            End With
        Else
            'Call MsgBox("es fehlen Shapes: " & vbLf & missingShapes)
            Call MsgBox(repMessages.getmsg(19) & vbLf & missingShapes)
        End If


        ' jetzt werden alle Shapes invisible gesetzt  ... 
        Call rds.setShapesInvisible()

        ' jezt wird das containershape in den Hintergrund gesetzt 
        Call rds.containerShape.ZOrder(MsoZOrderCmd.msoSendToBack)

    End Sub


    ''' <summary>
    ''' zeichnet die Projekte der Multiprojekt Sicht ( auch für extended Mode )
    ''' </summary>
    ''' <param name="projectCollection">der ganz zahlige Teil-1 ist die Zeile, in dei auf der ppt gezeichnet werden soll </param>
    ''' <param name="projDone"></param>
    ''' <param name="rds"></param>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="selectedRoles"></param>
    ''' <param name="selectedCosts"></param>
    ''' <remarks></remarks>
    Sub drawProjectsInPPT(ByRef projectCollection As SortedList(Of Double, String),
                                ByRef projDone As Integer,
                                ByVal rds As clsPPTShapes,
                                ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, ByVal selectedRoles As Collection, ByVal selectedCosts As Collection,
                                ByVal kennzeichnung As String)

        Dim addOn As Double = 0.0
        Dim msgTxt As String = ""

        If Not IsNothing(rds.durationArrowShape) And Not IsNothing(rds.durationTextShape) Then

            'addOn = System.Math.Max(durationArrowShape.Height, durationTextShape.Height) * 11 / 15
            addOn = System.Math.Max(rds.durationArrowShape.Height, rds.durationTextShape.Height) ' tk Änderung 

        End If

        Dim istEinzelProjektSicht As Boolean = (kennzeichnung = "Einzelprojektsicht" Or kennzeichnung = "AllePlanElemente")

        ' Bestimmen der Zeichenfläche
        Dim drawingAreaWidth As Double = rds.drawingAreaWidth
        'Dim drawingAreaHeight As Double = rds.drawingAreaBottom - rds.drawingAreaTop
        Dim drawingAreaHeight As Double = rds.availableSpace


        'Dim tagesEinheit As Double
        Dim projectsToDraw As Integer
        Dim copiedShape As PowerPoint.Shape = Nothing
        Dim fullName As String
        Dim hproj As clsProjekt

        Dim phaseShape As PowerPoint.Shape = Nothing
        Dim appear As clsAppearance
        Dim currentProjektIndex As Integer

        ' notwendig für das Positionieren des Duration Pfeils bzw. des DurationTextes
        Dim minX1 As Double
        Dim maxX2 As Double


        'Dim anzahlTage As Integer = DateDiff(DateInterval.Day, StartofPPTCalendar, endOFPPTCalendar) + 1
        Dim anzahlTage As Integer = CInt(DateDiff(DateInterval.Day, rds.PPTStartOFCalendar, rds.PPTEndOFCalendar))
        If anzahlTage <= 0 Then
            ''Throw New ArgumentException("Kalender Start bis Ende kann nicht 0 oder kleiner sein ..")
            Throw New ArgumentException("Problems with PPT StartOfCalendar, EndOf Calendar")
        End If



        ' Bestimmen der Position für den Projekt-Namen
        Dim projektNamenXPos As Double = rds.projectListLeft
        Dim projektNamenYPos As Double
        Dim projektNamenYrelPos As Double
        Dim x1 As Double
        Dim x2 As Double
        Dim projektGrafikYPos As Double
        Dim projektGrafikYrelPos As Double
        Dim phasenGrafikYPos As Double
        Dim phasenGrafikYrelPos As Double
        Dim milestoneGrafikYPos As Double
        Dim milestoneGrafikYrelPos As Double
        Dim ampelGrafikYPos As Double
        Dim ampelGrafikYrelPos As Double
        Dim rowYPos As Double
        Dim grafikrelOffset As Double

        Dim arrayOfNames() As String
        Dim phShapeNames As New Collection
        Dim msShapeNames As New Collection
        Dim drawRowDifferentiator As Boolean
        Dim toggleRowDifferentiator As Boolean
        Dim drawBUShape As Boolean
        Dim buFarbe As Long
        Dim buName As String
        Dim lastProjectName As String = ""
        Dim lastPhase As clsPhase = Nothing

        Dim lastProjectNameShape As PowerPoint.Shape = Nothing

        ' tk 6.12.2020
        Dim alreadyDrawnMilestones As New List(Of String)



        ' bestimme jetzt Y Start-Position für den Text bzw. die Grafik
        ' Änderung tk: die ProjektName, -Grafik, Milestone, Phasen Position wird jetzt relativ angegeben zum rowYPOS 
        With rds
            rowYPos = .drawingAreaTop
            projektNamenYrelPos = 0.5 * (.zeilenHoehe - .projectNameVorlagenShape.Height) + addOn
            projektGrafikYrelPos = 0.5 * (.zeilenHoehe - .projectVorlagenShape.Height) + addOn
            phasenGrafikYrelPos = 0.5 * (.zeilenHoehe - .phaseVorlagenShape.Height) + addOn
            milestoneGrafikYrelPos = 0.5 * (.zeilenHoehe - .milestoneVorlagenShape.Height) + addOn
            ampelGrafikYrelPos = 0.5 * (.zeilenHoehe - .ampelVorlagenShape.Height) + addOn
            grafikrelOffset = 0.5 * (.zeilenHoehe - .projectVorlagenShape.Height) + addOn
        End With

        ' initiales Setzen der YPositionen 
        projektNamenYPos = rowYPos + projektNamenYrelPos
        projektGrafikYPos = rowYPos + projektGrafikYrelPos
        phasenGrafikYPos = rowYPos + phasenGrafikYrelPos
        milestoneGrafikYPos = rowYPos + milestoneGrafikYrelPos
        ampelGrafikYPos = rowYPos + ampelGrafikYrelPos

        projectsToDraw = projectCollection.Count

        If Not IsNothing(rds.rowDifferentiatorShape) Then
            drawRowDifferentiator = True
        Else
            drawRowDifferentiator = False
        End If
        toggleRowDifferentiator = False

        If Not IsNothing(rds.buColorShape) Then
            drawBUShape = True
            projektNamenXPos = projektNamenXPos + rds.buColorShape.Width + 3
        Else
            drawBUShape = False
        End If

        Dim startIX As Integer = projDone + 1
        For currentProjektIndex = startIX To projectsToDraw

            ' zurücksetzen minX1, maxX2 
            minX1 = 100000.0
            maxX2 = -100000.0

            ' zurücksetzen der vergangenen Phase
            lastPhase = Nothing


            fullName = projectCollection.ElementAt(currentProjektIndex - 1).Value

            If AlleProjekte.Containskey(fullName) Then


                hproj = AlleProjekte.getProject(fullName)

                ' die müssen jetzt zurückgesetzt werden 
                alreadyDrawnMilestones.Clear()

                ' ur:23.03.2015: Test darauf, ob der Rest der Seite für dieses Projekt ausreicht'
                If awinSettings.mppExtendedMode Then
                    Dim neededSpace As Double = hproj.calcNeededLines(selectedPhases, selectedMilestones, True, Not awinSettings.mppShowAllIfOne) * rds.zeilenHoehe
                    If neededSpace - drawingAreaHeight > 1 Then

                        ' Projekt kann nicht gezeichnet werden, da nicht alle Phasen auf eine Seite passen, 
                        ' trotzdem muss das Projekt weitergezählt werden, damit das nächste zu zeichnende Projekt angegangen wird
                        projDone = projDone + 1
                        ' zuwenig Platz auf der Seite
                        ''Throw New ArgumentException("Für Projekt '" & fullName & "' ist zuwenig Platz auf einer Seite")
                        Throw New ArgumentException("not enough space for drawing " & fullName)

                    Else

                        If projektGrafikYPos - grafikrelOffset + hproj.calcNeededLines(selectedPhases, selectedMilestones, True, Not awinSettings.mppShowAllIfOne) * rds.zeilenHoehe > rds.drawingAreaBottom Then
                            Exit For
                        End If
                    End If
                End If



                Dim severalProjectsInOneLine As Boolean = False

                If Not istEinzelProjektSicht Then

                    If currentProjektIndex > 1 Then

                        If CInt(projectCollection.ElementAt(currentProjektIndex - 1).Key) = CInt(projectCollection.ElementAt(currentProjektIndex - 2).Key) And
                        Not IsNothing(lastProjectNameShape) Then
                            ' mehrere Projekte in einer Zeile 
                            severalProjectsInOneLine = True
                        Else
                            ' normal Mode ... nur 1 Projekt pro Zeile 
                        End If

                    Else
                        ' normal Mode ... nur 1 Projekt pro Zeile 
                    End If


                    copiedShape = createPPTShapeFromShape(rds.projectNameVorlagenShape, currentSlide)

                    ' wenn mehrere Projekte nacheinander in einer Zeile stehen 
                    If severalProjectsInOneLine Then

                        ' zuerst das lastProjectNAmeShape die MArgin lösche nund ganz nach oben schieben .. 
                        Dim offset As Double = projektNamenYrelPos

                        If Not IsNothing(lastProjectNameShape) Then
                            With lastProjectNameShape
                                If .TextFrame2.MarginTop > 0 Then
                                    .TextFrame2.MarginTop = 0
                                End If
                                If .TextFrame2.MarginBottom > 0 Then
                                    .TextFrame2.MarginBottom = 0
                                End If

                                .Top = CSng(rowYPos + 2)
                            End With
                        End If

                        ' jetzt das eigentliche Shape zeichnen 
                        With copiedShape

                            If currentProjektIndex > 1 And lastProjectName = hproj.name Then
                                .TextFrame2.TextRange.Text = "+ ... " & hproj.variantName
                            Else
                                .TextFrame2.TextRange.Text = "+ " & hproj.getShapeText
                            End If

                            ' die Oben und unten -Marge auf Null setzen, so dass der Text möglichst gut in die Zeile passt 
                            If .TextFrame2.MarginTop > 0 Then
                                .TextFrame2.MarginTop = 0
                            End If
                            If .TextFrame2.MarginBottom > 0 Then
                                .TextFrame2.MarginBottom = 0
                            End If

                            ' das jetzt so positionieren, dass es nach rechts versetzt und bündig unten mit dem Zeilenrand abschliesst 
                            .Left = lastProjectNameShape.Left + 8
                            If lastProjectNameShape.Top + lastProjectNameShape.Height + 2 + .Height > rowYPos + rds.zeilenHoehe Then
                                .Top = CSng(rowYPos + rds.zeilenHoehe - .Height)
                            Else
                                .Top = lastProjectNameShape.Top + lastProjectNameShape.Height + 2
                            End If



                            lastProjectName = hproj.name
                            .Name = .Name & .Id

                            If awinSettings.mppEnableSmartPPT Then

                                Call addSmartPPTMsPhInfo(copiedShape, hproj,
                                                        Nothing, hproj.getShapeText, Nothing, Nothing,
                                                        Nothing, Nothing,
                                                        hproj.startDate, hproj.endeDate,
                                                        hproj.ampelStatus, hproj.ampelErlaeuterung, Nothing,
                                                        hproj.leadPerson, hproj.getPhase(1).percentDone, hproj.getPhase(1).DocURL)

                            End If

                        End With
                    Else

                        With copiedShape
                            .Top = CSng(projektNamenYPos)
                            .Left = CSng(projektNamenXPos)
                            If currentProjektIndex > 1 And lastProjectName = hproj.name Then
                                .TextFrame2.TextRange.Text = "... " & hproj.variantName
                            Else
                                .TextFrame2.TextRange.Text = hproj.getShapeText
                            End If

                            lastProjectName = hproj.name
                            .Name = .Name & .Id

                            If awinSettings.mppEnableSmartPPT Then

                                Call addSmartPPTMsPhInfo(copiedShape, hproj,
                                                        Nothing, hproj.getShapeText, Nothing, Nothing,
                                                        Nothing, Nothing,
                                                        hproj.startDate, hproj.endeDate,
                                                        hproj.ampelStatus, hproj.ampelErlaeuterung, Nothing,
                                                        hproj.leadPerson, hproj.getPhase(1).percentDone, hproj.getPhase(1).DocURL)

                            End If

                        End With
                    End If

                    Dim projectNameShape As PowerPoint.Shape = copiedShape


                    ' zeichne jetzt ggf die Projekt-Ampel 
                    If awinSettings.mppShowAmpel And Not IsNothing(rds.ampelVorlagenShape) Then
                        Dim statusColor As Long
                        With hproj
                            If .ampelStatus = 0 Then
                                statusColor = awinSettings.AmpelNichtBewertet
                            ElseIf .ampelStatus = 1 Then
                                statusColor = awinSettings.AmpelGruen
                            ElseIf .ampelStatus = 2 Then
                                statusColor = awinSettings.AmpelGelb
                            Else
                                statusColor = awinSettings.AmpelRot
                            End If
                        End With


                        copiedShape = createPPTShapeFromShape(rds.ampelVorlagenShape, currentSlide)
                        With copiedShape
                            .Top = CSng(ampelGrafikYPos)
                            If severalProjectsInOneLine Then
                                .Left = CSng(rds.drawingAreaLeft - 3)
                            Else
                                .Left = CSng(rds.drawingAreaLeft - (.Width + 3))
                            End If
                            .Left = CSng(rds.drawingAreaLeft - (.Width + 3))
                            .Width = .Height
                            .Line.ForeColor.RGB = CInt(statusColor)
                            .Fill.ForeColor.RGB = CInt(statusColor)
                            .Name = .Name & .Id
                        End With

                        ampelGrafikYPos = ampelGrafikYPos + rds.zeilenHoehe

                    End If


                    '
                    ' zeichne jetzt das Projekt 
                    Call rds.calculatePPTx1x2(hproj.startDate, hproj.endeDate, x1, x2)


                    ' jetzt muss überprüft werden, ob projectName zu lang ist - dann wird der Name entsprechend abgekürzt ...
                    With projectNameShape
                        ' alternative Behandlung: der Projekt-Name wird umgebrochen 
                        If .Left + .Width > x1 Then
                            ' jetzt muss der Name entsprechend gekürzt werden 
                            .TextFrame2.WordWrap = MsoTriState.msoTrue
                            .Width = CSng(x1 - .Left)
                        End If

                        ' jetzt, wenn es in die nächste Zeile reingeht, so weit hochschieben, dass der Name nicht mehr in die nächste Zeile reicht 
                        If .Top + .Height > rowYPos + rds.zeilenHoehe Then
                            .Top = CSng(rowYPos + rds.zeilenHoehe - .Height)
                        End If

                    End With

                    ' hier ggf die ProjectLine zeichnen 
                    If awinSettings.mppShowProjectLine Then


                        copiedShape = createPPTShapeFromShape(rds.projectVorlagenShape, currentSlide)
                        With copiedShape
                            .Top = CSng(projektGrafikYPos)
                            .Left = CSng(x1)
                            .Width = CSng(x2 - x1)
                            .Name = .Name & .Id

                            '.Title = hproj.getShapeText
                            '.AlternativeText = hproj.startDate.ToShortDateString & " - " & hproj.endeDate.ToShortDateString

                            If awinSettings.mppEnableSmartPPT Then

                                Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(rootPhaseName)
                                Dim shortText As String = hproj.name
                                Dim originalName As String = hproj.name

                                Dim bestShortName As String = shortText
                                Dim bestLongName As String = shortText


                                Call addSmartPPTMsPhInfo(copiedShape, hproj,
                                                   fullBreadCrumb, hproj.getShapeText, shortText, shortText,
                                                   shortText, shortText,
                                                   hproj.startDate, hproj.endeDate,
                                                   hproj.ampelStatus, hproj.ampelErlaeuterung, hproj.description,
                                                   hproj.leadPerson, hproj.getPhase(1).percentDone, hproj.getPhase(1).DocURL)

                            End If


                            ' wenn Projektstart vor dem Kalender-Start liegt: kein Projektstart Symbol zeichnen
                            If DateDiff(DateInterval.Day, hproj.startDate, rds.PPTStartOFCalendar) > 0 Then
                                .Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone
                            End If

                            ' wenn Projektende nach dem Kalender-Ende liegt: kein Projektende Symbol zeichnen
                            If DateDiff(DateInterval.Day, hproj.endeDate, rds.PPTEndOFCalendar) < 0 Then
                                .Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone
                            End If
                        End With

                    End If


                End If
                '
                ' zeichne jetzt die Phasen 
                '

                Dim anzZeilenGezeichnet As Integer = 1

                For i = 0 To hproj.CountPhases - 1

                    Dim cphase As clsPhase = hproj.getPhase(i + 1)

                    Dim phaseName As String = cphase.name
                    If Not IsNothing(cphase) Then

                        ' herausfinden, ob cphase in den selektierten Phasen enthalten ist
                        Dim found As Boolean = False
                        Dim j As Integer = 1
                        Dim breadcrumb As String = ""
                        ' gibt den vollen Breadcrumb zurück 
                        Dim vglBreadCrumb As String = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                        ' falls in selPhases Categories stehen 
                        Dim vglCategoryName As String = calcHryCategoryName(cphase.appearance, False)
                        Dim selPhaseName As String = ""

                        While j <= selectedPhases.Count And Not found

                            Dim type As Integer = -1
                            Dim pvName As String = ""
                            Call splitHryFullnameTo2(CStr(selectedPhases(j)), selPhaseName, breadcrumb, type, pvName)

                            If type = -1 Or
                                (type = PTItemType.projekt And pvName = calcProjektKey(hproj.name, hproj.variantName)) Or
                                (type = PTItemType.vorlage) Then

                                If cphase.name = selPhaseName Then
                                    If vglBreadCrumb.EndsWith(breadcrumb) Then
                                        found = True
                                    Else
                                        j = j + 1
                                    End If
                                Else
                                    j = j + 1
                                End If

                            ElseIf type = PTItemType.categoryList Then

                                If selectedPhases.Contains(vglCategoryName) Then
                                    found = True
                                Else
                                    j = j + 1
                                End If

                            Else
                                j = j + 1
                            End If


                        End While

                        If found Then           ' cphase ist eine der selektierten Phasen

                            Dim projektstart As Integer = hproj.Start + hproj.StartOffset


                            Dim zeichnen As Boolean = True
                            ' erst noch prüfen , ob diese Phase tatsächlich im Zeitraum enthalten ist 
                            If awinSettings.mppShowAllIfOne Then
                                zeichnen = True
                            Else
                                If phaseWithinTimeFrame(projektstart, cphase.relStart, cphase.relEnde, showRangeLeft, showRangeRight) Then
                                    zeichnen = True
                                Else
                                    zeichnen = False
                                End If
                            End If



                            If zeichnen Then

                                Dim missingPhaseDefinition As Boolean = PhaseDefinitions.Contains(phaseName)

                                If awinSettings.mppExtendedMode Then
                                    'phasenName = cphase.name
                                    If Not IsNothing(lastPhase) Then

                                        ' Nachfragen, ob cphase und lastPhase überlappen

                                        If DateDiff(DateInterval.Day, lastPhase.getEndDate, cphase.getStartDate) < 0 Then
                                            ' Phase muss in neue Zeile eingetragen werden
                                            Dim tmpint As Integer
                                            Dim drawliste As New SortedList(Of String, SortedList)

                                            Call hproj.selMilestonesToselPhase(selectedPhases, selectedMilestones, True, tmpint, drawliste)
                                            If drawliste.ContainsKey(lastPhase.nameID) Then
                                                ' es müssen zur letzten Phase noch Meilensteine gezeichnet werden, die in einer nicht selektierten Phase liegen, die Child von der lastphase ist
                                                ' dafür: weiterschalten der Zeile
                                                phasenGrafikYPos = phasenGrafikYPos + rds.zeilenHoehe
                                                ' Y-Position für BU und Hintergrund-einfärbung erhöhen je gezeichneter Zeile
                                                '''' ur:20.04.2015:  rowYPos = rowYPos + zeilenhoehe
                                                ' Y-Position für Projektnamen erhöhen je gezeichneter Phase
                                                projektNamenYPos = projektNamenYPos + rds.zeilenHoehe
                                                ' Y-Position für Meilensteine der aktuellen Phase erhöhen je gezeichneter Phase
                                                milestoneGrafikYPos = milestoneGrafikYPos + rds.zeilenHoehe
                                                ' Y-Position der Ampel, sofern sie zu dem Projekt gezeichnet werden soll
                                                ampelGrafikYPos = ampelGrafikYPos + rds.zeilenHoehe
                                                anzZeilenGezeichnet = anzZeilenGezeichnet + 1


                                                ' ur: Meilensteine aus drawliste.value zeichnen
                                                Dim zeichnenMS As Boolean = False
                                                Dim msliste As SortedList
                                                Dim msi As Integer
                                                msliste = drawliste(lastPhase.nameID)

                                                For msi = 0 To msliste.Count - 1
                                                    Dim msID As String = CStr(msliste.GetByIndex(msi))
                                                    Dim milestone As clsMeilenstein = hproj.getMilestoneByID(msID)

                                                    ' Nachsehen, ob MS -Datum existiert und größer StartofCalender ist und im Zeitraum liegt, oder evt. trotzdem gezeichnet werden soll
                                                    If IsNothing(milestone.getDate) Then
                                                        zeichnenMS = False
                                                    Else
                                                        If DateDiff(DateInterval.Day, StartofCalendar, milestone.getDate) >= 0 Then

                                                            ' erst noch prüfen , ob dieser Meilenstein tatsächlich im Zeitraum enthalten ist 
                                                            If awinSettings.mppShowAllIfOne Then
                                                                zeichnenMS = True
                                                            Else
                                                                If milestoneWithinTimeFrame(milestone.getDate, showRangeLeft, showRangeRight) Then
                                                                    zeichnenMS = True
                                                                Else
                                                                    zeichnenMS = False
                                                                End If
                                                            End If
                                                        Else
                                                            zeichnenMS = False
                                                        End If
                                                    End If

                                                    If zeichnenMS Then

                                                        Call zeichneMeilensteininAktZeile(currentSlide, msShapeNames, minX1, maxX2,
                                                                                      milestone, hproj, milestoneGrafikYPos, rds)

                                                        Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(milestone.nameID) & milestone.nameID

                                                        If Not alreadyDrawnMilestones.Contains(fullBreadCrumb) Then
                                                            alreadyDrawnMilestones.Add(fullBreadCrumb)
                                                        End If


                                                    End If

                                                Next

                                            End If
                                            phasenGrafikYPos = phasenGrafikYPos + rds.zeilenHoehe
                                            ' Y-Position für BU und Hintergrund-einfärbung erhöhen je gezeichneter Zeile
                                            '''' ur:20.04.2015:  rowYPos = rowYPos + zeilenhoehe
                                            ' Y-Position für Projektnamen erhöhen je gezeichneter Phase
                                            projektNamenYPos = projektNamenYPos + rds.zeilenHoehe
                                            ' Y-Position für Meilensteine der aktuellen Phase erhöhen je gezeichneter Phase
                                            milestoneGrafikYPos = milestoneGrafikYPos + rds.zeilenHoehe
                                            ' Y-Position der Ampel, sofern sie zu dem Projekt gezeichnet werden soll
                                            ampelGrafikYPos = ampelGrafikYPos + rds.zeilenHoehe
                                            anzZeilenGezeichnet = anzZeilenGezeichnet + 1
                                        Else
                                            ' cphase und lastphase überlappen nicht, also auch kein weiterschalten der yPositionen

                                            ' noch zu tun:ur: 01.10.2015:hier muss man sich merken, welche Phasen nun in dieser Zeile alle gezeichnet wurden, damit die Meilensteine der Phasen in der Hierarchie
                                            ' unterhalb dieser Phasen passend dazugezeichnet werden können.
                                        End If
                                    End If

                                End If



                                ' Änderung tk 26.11 
                                If PhaseDefinitions.Contains(phaseName) Then
                                    'phaseShape = PhaseDefinitions.getShape(phaseName)
                                    appear = PhaseDefinitions.getShapeApp(phaseName)
                                Else
                                    'phaseShape = missingPhaseDefinitions.getShape(phaseName)
                                    appear = missingPhaseDefinitions.getShapeApp(phaseName)
                                End If

                                ' Ergänzung 19.4.16
                                Dim phShapeName As String = calcPPTShapeName(hproj, cphase.nameID)


                                Dim phaseStart As Date = cphase.getStartDate
                                Dim phaseEnd As Date = cphase.getEndDate
                                'Dim phShortname As String = PhaseDefinitions.getAbbrev(phaseName).Trim
                                ' erhänzt tk
                                Dim phShortname As String = ""
                                phShortname = hproj.getBestNameOfID(cphase.nameID, Not awinSettings.mppUseOriginalNames,
                                                                              awinSettings.mppUseAbbreviation)

                                Call rds.calculatePPTx1x2(phaseStart, phaseEnd, x1, x2)



                                If minX1 > x1 Then
                                    minX1 = x1
                                End If

                                If maxX2 < x2 Then
                                    maxX2 = x2
                                End If

                                ' jetzt müssen ggf der Phasen Name und das  Datum angebracht werden 
                                If awinSettings.mppShowPhName Then

                                    If phShortname.Trim.Length = 0 Then
                                        phShortname = phaseName
                                    End If


                                    copiedShape = createPPTShapeFromShape(rds.PhDescVorlagenShape, currentSlide)
                                    With copiedShape

                                        '.Name = .Name & .Id
                                        Try
                                            .Name = phShapeName & PTpptAnnotationType.text
                                        Catch ex As Exception
                                            ' Fehler abfangen ..
                                        End Try

                                        .Title = "Beschriftung"
                                        .AlternativeText = ""

                                        .TextFrame2.TextRange.Text = phShortname
                                        .TextFrame2.MarginLeft = 0.0
                                        .TextFrame2.MarginRight = 0.0
                                        '.Top = CSng(projektGrafikYPos) - .Height
                                        .Top = CSng(phasenGrafikYPos) + CSng(rds.yOffsetPhToText) - 2
                                        .Left = CSng(x1)
                                        If .Left < rds.drawingAreaLeft Then
                                            .Left = CSng(rds.drawingAreaLeft)
                                        End If
                                        .TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft

                                    End With


                                End If

                                Dim phDateText As String = ""
                                ' jetzt muss ggf das Datum angebracht werden 
                                If awinSettings.mppShowPhDate Then
                                    'Dim phDateText As String = phaseStart.ToShortDateString
                                    phDateText = phaseStart.Day.ToString & "." & phaseStart.Month.ToString & " - " &
                                                                phaseEnd.Day.ToString & "." & phaseEnd.Month.ToString
                                    Dim rightX As Double, addHeight As Double

                                    copiedShape = createPPTShapeFromShape(rds.PhDateVorlagenShape, currentSlide)
                                    With copiedShape

                                        '.Name = .Name & .Id
                                        Try
                                            .Name = phShapeName & PTpptAnnotationType.datum
                                        Catch ex As Exception

                                        End Try

                                        .Title = "Datum"
                                        .AlternativeText = ""

                                        .TextFrame2.TextRange.Text = phDateText
                                        .TextFrame2.MarginLeft = 0.0
                                        .TextFrame2.MarginRight = 0.0
                                        '.Top = CSng(projektGrafikYPos)
                                        .Top = CSng(phasenGrafikYPos) + CSng(rds.yOffsetPhToDate) + 1
                                        .Left = CSng(x1) + 1
                                        If .Left < rds.drawingAreaLeft Then
                                            .Left = CSng(rds.drawingAreaLeft + 1)
                                        End If
                                        .TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft

                                        rightX = .Left + .Width
                                        addHeight = .Height * 0.7

                                    End With



                                End If

                                ' jetzt muss ggf das Phase Delimiter Shape angebracht werden 
                                If Not IsNothing(rds.phaseDelimiterShape) And selectedPhases.Count > 1 Then

                                    ' linker Delimiter 

                                    copiedShape = createPPTShapeFromShape(rds.phaseDelimiterShape, currentSlide)
                                    With copiedShape

                                        .Height = CSng(1.3 * appear.height)
                                        .Top = CSng(phasenGrafikYPos)
                                        .Left = CSng(x1 - .Width * 0.5)
                                        .Name = .Name & .Id

                                    End With

                                    ' rechter Delimiter 

                                    copiedShape = createPPTShapeFromShape(rds.phaseDelimiterShape, currentSlide)
                                    With copiedShape

                                        .Height = CSng(1.3 * appear.height)
                                        .Top = CSng(phasenGrafikYPos)
                                        .Left = CSng(x2 + .Width * 0.5)
                                        .Name = .Name & .Id

                                    End With

                                End If



                                phaseShape = currentSlide.Shapes.AddShape(appear.shpType,
                                                                      CSng(x1),
                                                                      CSng(phasenGrafikYPos),
                                                                      CSng(x2 - x1),
                                                                      rds.phaseVorlagenShape.Height)
                                Call definePhPPTAppearance(phaseShape, appear)

                                With phaseShape

                                    Try
                                        .Name = phShapeName
                                    Catch ex As Exception

                                    End Try


                                    If missingPhaseDefinition Then
                                        .Fill.ForeColor.RGB = cphase.farbe
                                    End If

                                End With

                                If awinSettings.mppEnableSmartPPT Then

                                    Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                                    Dim shortText As String = cphase.shortName
                                    Dim originalName As String = cphase.originalName

                                    Dim bestShortName As String = hproj.getBestNameOfID(cphase.nameID, True, True)
                                    Dim bestLongName As String = hproj.getBestNameOfID(cphase.nameID, True, False)

                                    If originalName = cphase.name Then
                                        originalName = Nothing
                                    End If

                                    Call addSmartPPTMsPhInfo(phaseShape, hproj,
                                                                fullBreadCrumb, cphase.name, shortText, originalName,
                                                                bestShortName, bestLongName,
                                                                phaseStart, phaseEnd,
                                                                cphase.ampelStatus, cphase.ampelErlaeuterung, cphase.getAllDeliverables("#"),
                                                                cphase.verantwortlich, cphase.percentDone, cphase.DocURL)
                                End If

                                phShapeNames.Add(phaseShape.Name)

                                '  Phase merken, damit bei der nächsten zu zeichnenden Phase nachgesehen werden
                                '  kann, ob diese überlappt

                                If i < hproj.CountPhases - 1 Then
                                    lastPhase = hproj.getPhase(i + 1)   ' zu diesem Zeitpunkt ist ebenfalls cphase = hproj.getPhase(i+1)
                                End If

                            End If



                            ' zeichne jetzt die Meilensteine der aktuellen Phase
                            ' ur: 29.04.2015: und baue eine Collection auf, die alle selektierten Meilensteine aus den unterschiedlichen Phasen beinhaltet.
                            ' Sobald der Meilenstein/Phase gezeichnet wurde, wird er daraus gelöscht.
                            ' ur: 19.03.2015: diese Schleife muss innerhalb der für die Phase liegen

                            Dim milestoneName As String = ""
                            Dim ms As clsMeilenstein = Nothing


                            For ix As Integer = 1 To selectedMilestones.Count

                                Dim breadcrumbMS As String = ""

                                Dim type As Integer = -1
                                Dim pvName As String = ""
                                Call splitHryFullnameTo2(CStr(selectedMilestones.Item(ix)), milestoneName, breadcrumbMS, type, pvName)

                                If type = -1 Or
                                     (type = PTItemType.projekt And pvName = calcProjektKey(hproj.name, hproj.variantName)) Or
                                     (type = PTItemType.vorlage) Then

                                    ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste
                                    Dim milestoneIndices(,) As Integer = hproj.hierarchy.getMilestoneIndices(milestoneName, breadcrumbMS)

                                    Dim phaseNameID As String = ""

                                    For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                                        ms = hproj.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))

                                        If Not IsNothing(ms) Then

                                            Dim msDate As Date = ms.getDate

                                            Dim phaseNameID1 As String = hproj.hierarchy.getParentIDOfID(ms.nameID)
                                            phaseNameID = hproj.getPhase(milestoneIndices(0, mx)).nameID

                                            If phaseNameID <> phaseNameID1 Then
                                                'Call MsgBox(" Schleife über Meilensteine,  Fehler in zeichnePPTprojects,")
                                            End If

                                            If phaseNameID = cphase.nameID Then
                                                Dim zeichnenMS As Boolean

                                                Dim hilfsvar As Integer = hproj.hierarchy.getIndexOfID(cphase.nameID)

                                                If IsNothing(msDate) Then
                                                    zeichnenMS = False
                                                Else
                                                    If DateDiff(DateInterval.Day, StartofCalendar, msDate) >= 0 Then

                                                        ' erst noch prüfen , ob dieser Meilenstein tatsächlich im Zeitraum enthalten ist 
                                                        If awinSettings.mppShowAllIfOne Then
                                                            zeichnenMS = True
                                                        Else
                                                            If milestoneWithinTimeFrame(msDate, showRangeLeft, showRangeRight) Then
                                                                zeichnenMS = True
                                                            Else
                                                                zeichnenMS = False
                                                            End If
                                                        End If
                                                    Else
                                                        zeichnenMS = False
                                                    End If
                                                End If


                                                If zeichnenMS Then

                                                    Call zeichneMeilensteininAktZeile(currentSlide, msShapeNames, minX1, maxX2,
                                                                                      ms, hproj, milestoneGrafikYPos, rds)

                                                    Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(ms.nameID) & ms.nameID
                                                    If Not alreadyDrawnMilestones.Contains(fullBreadCrumb) Then
                                                        alreadyDrawnMilestones.Add(fullBreadCrumb)
                                                    End If

                                                End If


                                            Else
                                                ' selektierter Meilenstein 'milestoneName' nicht in dieser Phase enthalten
                                                ' also: nichts tun
                                            End If

                                        End If


                                    Next mx

                                End If



                            Next ix  ' nächsten selektieren Meilenstein überprüfen und ggfs. einzeichnen 

                        End If
                    End If


                Next i      ' nächste Phase bearbeiten


                ''''ur:30.09.2015: Es müssen zur letzten Phase noch Meilensteine gezeichnet werden, die in einer nicht selektierten Phase liegen,
                ''''               die Child von der lastphase ist.


                If awinSettings.mppExtendedMode Then


                    Dim tmpint As Integer
                    Dim drawliste As New SortedList(Of String, SortedList)

                    Call hproj.selMilestonesToselPhase(selectedPhases, selectedMilestones, True, tmpint, drawliste)


                    If Not IsNothing(lastPhase) Then
                        ' Abfrage, ob zur letzten gezeichneten Phase noch Meilensteine aus untergeordneten Phasen gezeichnet werden müssen

                        If drawliste.ContainsKey(lastPhase.nameID) Then

                            ' es müssen zur letzten Phase noch Meilensteine gezeichnet werden, die in einer nicht selektierten Phase liegen, die Child von der lastphase ist
                            ' dafür: weiterschalten der Zeile
                            phasenGrafikYPos = phasenGrafikYPos + rds.zeilenHoehe
                            ' Y-Position für BU und Hintergrund-einfärbung erhöhen je gezeichneter Zeile
                            '''' ur:20.04.2015:  rowYPos = rowYPos + zeilenhoehe
                            ' Y-Position für Projektnamen erhöhen je gezeichneter Phase
                            projektNamenYPos = projektNamenYPos + rds.zeilenHoehe
                            ' Y-Position für Meilensteine der aktuellen Phase erhöhen je gezeichneter Phase
                            milestoneGrafikYPos = milestoneGrafikYPos + rds.zeilenHoehe
                            ' Y-Position der Ampel, sofern sie zu dem Projekt gezeichnet werden soll
                            ampelGrafikYPos = ampelGrafikYPos + rds.zeilenHoehe
                            anzZeilenGezeichnet = anzZeilenGezeichnet + 1


                            ' ur: Meilensteine aus drawliste.value zeichnen
                            Dim zeichnenMS As Boolean = False
                            Dim msliste As SortedList
                            Dim msi As Integer
                            msliste = drawliste(lastPhase.nameID)

                            For msi = 0 To msliste.Count - 1

                                Dim msID As String = CStr(msliste.GetByIndex(msi))
                                Dim milestone As clsMeilenstein = hproj.getMilestoneByID(msID)

                                ' Nachsehen, ob MS -Datum existiert und größer StartofCalender ist und im Zeitraum liegt, oder evt. trotzdem gezeichnet werden soll
                                If IsNothing(milestone.getDate) Then
                                    zeichnenMS = False
                                Else
                                    If DateDiff(DateInterval.Day, StartofCalendar, milestone.getDate) >= 0 Then

                                        ' erst noch prüfen , ob dieser Meilenstein tatsächlich im Zeitraum enthalten ist 
                                        If awinSettings.mppShowAllIfOne Then
                                            zeichnenMS = True
                                        Else
                                            If milestoneWithinTimeFrame(milestone.getDate, showRangeLeft, showRangeRight) Then
                                                zeichnenMS = True
                                            Else
                                                zeichnenMS = False
                                            End If
                                        End If
                                    Else
                                        zeichnenMS = False
                                    End If
                                End If


                                If zeichnenMS Then

                                    Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(milestone.nameID) & milestone.nameID
                                    If Not alreadyDrawnMilestones.Contains(fullBreadCrumb) Then

                                        Call zeichneMeilensteininAktZeile(currentSlide, msShapeNames, minX1, maxX2,
                                                                      milestone, hproj, milestoneGrafikYPos, rds)

                                        alreadyDrawnMilestones.Add(fullBreadCrumb)
                                    End If

                                End If
                            Next


                        End If


                    End If

                    '''' ur: 01.10.2015: selektierte Meilensteine zeichnen, die zu keiner der selektierten Phasen gehören.

                    If drawliste.ContainsKey(rootPhaseName) Then

                        phasenGrafikYPos = phasenGrafikYPos + rds.zeilenHoehe
                        ' Y-Position für BU und Hintergrund-einfärbung erhöhen je gezeichneter Zeile
                        '''' ur:20.04.2015:  rowYPos = rowYPos + zeilenhoehe
                        ' Y-Position für Projektnamen erhöhen je gezeichneter Phase
                        projektNamenYPos = projektNamenYPos + rds.zeilenHoehe
                        ' Y-Position für Meilensteine der aktuellen Phase erhöhen je gezeichneter Phase
                        milestoneGrafikYPos = milestoneGrafikYPos + rds.zeilenHoehe
                        ' Y-Position der Ampel, sofern sie zu dem Projekt gezeichnet werden soll
                        ampelGrafikYPos = ampelGrafikYPos + rds.zeilenHoehe
                        anzZeilenGezeichnet = anzZeilenGezeichnet + 1


                        ' ur: Meilensteine aus drawliste.value zeichnen
                        Dim zeichnenMS As Boolean = False
                        Dim msliste As SortedList
                        Dim msi As Integer
                        msliste = drawliste(rootPhaseName)

                        For msi = 0 To msliste.Count - 1

                            Dim msID As String = CStr(msliste.GetByIndex(msi))
                            Dim milestone As clsMeilenstein = hproj.getMilestoneByID(msID)

                            ' Nachsehen, ob MS -Datum existiert und größer StartofCalender ist und im Zeitraum liegt, oder evt. trotzdem gezeichnet werden soll
                            If IsNothing(milestone.getDate) Then
                                zeichnenMS = False
                            Else
                                If DateDiff(DateInterval.Day, StartofCalendar, milestone.getDate) >= 0 Then

                                    ' erst noch prüfen , ob dieser Meilenstein tatsächlich im Zeitraum enthalten ist 
                                    If awinSettings.mppShowAllIfOne Then
                                        zeichnenMS = True
                                    Else
                                        If milestoneWithinTimeFrame(milestone.getDate, showRangeLeft, showRangeRight) Then
                                            zeichnenMS = True
                                        Else
                                            zeichnenMS = False
                                        End If
                                    End If
                                Else
                                    zeichnenMS = False
                                End If
                            End If

                            If zeichnenMS Then

                                Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(milestone.nameID) & milestone.nameID
                                If Not alreadyDrawnMilestones.Contains(fullBreadCrumb) Then

                                    Call zeichneMeilensteininAktZeile(currentSlide, msShapeNames, minX1, maxX2,
                                                                      milestone, hproj, milestoneGrafikYPos, rds)

                                    alreadyDrawnMilestones.Add(fullBreadCrumb)
                                End If

                            End If
                        Next

                    End If

                Else    ' Einzeilen-Modus: alle selektierten Meilensteine zeichnen, die nicht zu einer selektieren Phase gehören

                    Dim tmpint As Integer
                    Dim drawliste As New SortedList(Of String, SortedList)

                    Call hproj.selMilestonesToselPhase(selectedPhases, selectedMilestones, True, tmpint, drawliste)

                    For Each kvp As KeyValuePair(Of String, SortedList) In drawliste

                        ' ur: Meilensteine aus drawliste.value zeichnen
                        Dim zeichnenMS As Boolean = False
                        Dim msliste As SortedList
                        Dim msi As Integer
                        msliste = kvp.Value

                        For msi = 0 To msliste.Count - 1

                            Dim msID As String = CStr(msliste.GetByIndex(msi))
                            Dim milestone As clsMeilenstein = hproj.getMilestoneByID(msID)

                            ' Nachsehen, ob MS -Datum existiert und größer StartofCalender ist und im Zeitraum liegt, oder evt. trotzdem gezeichnet werden soll
                            If IsNothing(milestone.getDate) Then
                                zeichnenMS = False
                            Else
                                If DateDiff(DateInterval.Day, StartofCalendar, milestone.getDate) >= 0 Then

                                    ' erst noch prüfen , ob dieser Meilenstein tatsächlich im Zeitraum enthalten ist 
                                    If awinSettings.mppShowAllIfOne Then
                                        zeichnenMS = True
                                    Else
                                        If milestoneWithinTimeFrame(milestone.getDate, showRangeLeft, showRangeRight) Then
                                            zeichnenMS = True
                                        Else
                                            zeichnenMS = False
                                        End If
                                    End If
                                Else
                                    zeichnenMS = False
                                End If
                            End If

                            If zeichnenMS Then

                                Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(milestone.nameID) & milestone.nameID
                                If Not alreadyDrawnMilestones.Contains(fullBreadCrumb) Then

                                    Call zeichneMeilensteininAktZeile(currentSlide, msShapeNames, minX1, maxX2,
                                                                      milestone, hproj, milestoneGrafikYPos, rds)

                                    alreadyDrawnMilestones.Add(fullBreadCrumb)
                                End If

                            End If
                        Next

                    Next kvp
                End If

                ' hier könnte jetzt eigentlich auch eine Behandlung stehen, um ggf mehrere Projekt-Namen, die in einer Zeile stehen, besser auf den zur Verfügung stehenden Platz zu verteilen
                ' das kann aber immer noch später gemacht werden 
                ' hier müsste das behandelt werden 
                ' Ende dieser Behandlung 


                ' optionales zeichnen der BU Markierung 
                If drawBUShape Then
                    buName = hproj.businessUnit
                    buFarbe = awinSettings.AmpelNichtBewertet

                    If Not IsNothing(buName) Then

                        If buName.Length > 0 Then
                            Dim found As Boolean = False
                            Dim ix As Integer = 1
                            While ix <= businessUnitDefinitions.Count And Not found
                                If businessUnitDefinitions.ElementAt(ix - 1).Value.name = buName Then
                                    found = True
                                    buFarbe = businessUnitDefinitions.ElementAt(ix - 1).Value.color
                                Else
                                    ix = ix + 1
                                End If
                            End While
                        End If

                    End If



                    copiedShape = createPPTShapeFromShape(rds.buColorShape, currentSlide)
                    With copiedShape
                        .Top = CSng(rowYPos)
                        .Left = CSng(rds.projectListLeft)
                        '' '' ''Dim neededLines As Double = hproj.calcNeededLines(selectedPhases, awinSettings.mppExtendedMode, Not awinSettings.mppShowAllIfOne)
                        '' '' ''.Height = hproj.calcNeededLines(selectedPhases, awinSettings.mppExtendedMode, Not awinSettings.mppShowAllIfOne) * zeilenhoehe
                        .Height = CSng(anzZeilenGezeichnet * rds.zeilenHoehe)
                        .Fill.ForeColor.RGB = CInt(buFarbe)
                        .Name = .Name & .Id
                        ' width ist die in der Vorlage angegebene Width 
                    End With

                End If


                ' optionales zeichnen der Zeilen-Markierung
                If drawRowDifferentiator And toggleRowDifferentiator Then
                    ' zeichnen des RowDifferentiators 

                    copiedShape = createPPTShapeFromShape(rds.rowDifferentiatorShape, currentSlide)
                    With copiedShape
                        .Top = CSng(rowYPos)
                        .Left = CSng(rds.projectListLeft)
                        '''''.Height = hproj.calcNeededLines(selectedPhases, awinSettings.mppExtendedMode, Not awinSettings.mppShowAllIfOne) * zeilenhoehe
                        .Height = CSng(anzZeilenGezeichnet * rds.zeilenHoehe)
                        .Width = CSng(rds.drawingAreaRight - .Left)
                        .Name = .Name & .Id
                        .ZOrder(MsoZOrderCmd.msoSendToBack)
                    End With
                End If

                ' jetzt muss ggf die duration eingezeichnet werden 
                If Not IsNothing(rds.durationArrowShape) And Not IsNothing(rds.durationTextShape) Then

                    ' Pfeil mit Länge der Dauer zeichnen 
                    copiedShape = createPPTShapeFromShape(rds.durationArrowShape, currentSlide)
                    Dim pfeilbreite As Double = maxX2 - minX1

                    With copiedShape
                        .Top = CSng(rowYPos + 3 + 0.5 * (addOn - .Height))
                        .Left = CSng(minX1)
                        .Width = CSng(pfeilbreite)
                        .Name = .Name & .Id
                    End With

                    ' Text für die Dauer eintragen
                    Dim dauerInTagen As Long
                    Dim dauerInM As Double
                    Dim tmpDate1 As Date, tmpDate2 As Date

                    Call hproj.getMinMaxDatesAndDuration(selectedPhases, selectedMilestones, tmpDate1, tmpDate2, dauerInTagen)
                    dauerInM = 12 * dauerInTagen / 365


                    copiedShape = createPPTShapeFromShape(rds.durationTextShape, currentSlide)
                    With copiedShape
                        .TextFrame2.TextRange.Text = dauerInM.ToString("0.0") & " M"
                        .Top = CSng(rowYPos + 3 + 0.5 * (addOn - .Height))
                        .Left = CSng(minX1 + (pfeilbreite - .Width) / 2)
                        .Name = .Name & .Id
                    End With

                End If


                projDone = projDone + 1
                ' Behandlung 


                ' weiter schalten muss nur gemacht werden, wenn das nächste Projekt in der Collection nicht in der gleichen Zeile sein sollte
                ' falls das nächste Projekt in der gleichen Zeile sein sollte, so werdendas ist in der Routine bestimmeMinMaxProjekte .. festgelegt; gezeichnet wird wie auf der PRojekt-Tafel dargestellt ... 
                ' es können also auch zwei PRojekte (z.B Projekt und Nachfolger)  in einer Zeile sein ... 
                If currentProjektIndex <= projectCollection.Count - 1 Then

                    ' dadurch wird die Zeilen - bzw. Projekt - Markierung nur bei jedem zweiten Mal gezeichnet ... 
                    toggleRowDifferentiator = Not toggleRowDifferentiator

                    If Not awinSettings.mppExtendedMode Then
                        rowYPos = rowYPos + rds.zeilenHoehe
                    Else
                        rowYPos = rowYPos + anzZeilenGezeichnet * rds.zeilenHoehe
                    End If
                    lastProjectNameShape = Nothing
                    ' in PPT kann aktuell gar nicht bestimmt werden, dass es nebeneinander sein - die Preview Fuktion fehlt ja hier .. 
                    'If CInt(projectCollection.ElementAt(currentProjektIndex - 1).Key) < CInt(projectCollection.ElementAt(currentProjektIndex).Key) Then

                    '    ' dadurch wird die Zeilen - bzw. Projekt - Markierung nur bei jedem zweiten Mal gezeichnet ... 
                    '    toggleRowDifferentiator = Not toggleRowDifferentiator

                    '    If Not awinSettings.mppExtendedMode Then
                    '        rowYPos = rowYPos + rds.zeilenHoehe
                    '    Else
                    '        rowYPos = rowYPos + anzZeilenGezeichnet * rds.zeilenHoehe
                    '    End If
                    '    lastProjectNameShape = Nothing

                    'Else
                    '    ' rowYPos bleibt unverändert 
                    '    lastProjectNameShape = projectNameShape
                    'End If
                Else
                    ' dadurch wird die Zeilen - bzw. Projekt - Markierung nur bei jedem zweiten Mal gezeichnet ... 
                    toggleRowDifferentiator = Not toggleRowDifferentiator

                    If Not awinSettings.mppExtendedMode Then
                        rowYPos = rowYPos + rds.zeilenHoehe
                    Else
                        rowYPos = rowYPos + anzZeilenGezeichnet * rds.zeilenHoehe
                    End If
                    lastProjectNameShape = Nothing
                End If


                ' Ende Behandlung 

                ' jetzt alle Werte in Abhängigkeit von rowYPos wieder setzen ... 
                projektNamenYPos = rowYPos + projektNamenYrelPos
                projektGrafikYPos = rowYPos + projektGrafikYrelPos
                phasenGrafikYPos = rowYPos + phasenGrafikYrelPos
                milestoneGrafikYPos = rowYPos + milestoneGrafikYrelPos
                ampelGrafikYPos = rowYPos + ampelGrafikYrelPos


                'phasenGrafikYPos = phasenGrafikYPos + rds.zeilenHoehe
                'milestoneGrafikYPos = milestoneGrafikYPos + rds.zeilenHoehe

                If projektGrafikYPos > rds.drawingAreaBottom Then
                    Exit For
                End If



            End If


        Next            ' nächstes Projekt zeichnen


        '
        ' wenn  Texte gezeichnet wurden, müssen jetzt die Phasen in den Vordergrund geholt werden, danach auf alle Fälle auch die Meilensteine 
        Dim anzElements As Integer
        If awinSettings.mppShowMsDate Or awinSettings.mppShowMsName Or
            awinSettings.mppShowPhDate Or awinSettings.mppShowPhName Then
            ' Phasen vorholen 

            anzElements = phShapeNames.Count

            If anzElements > 0 Then

                ReDim arrayOfNames(anzElements - 1)
                For ix = 1 To anzElements
                    arrayOfNames(ix - 1) = CStr(phShapeNames.Item(ix))
                Next

                Try
                    CType(currentSlide.Shapes.Range(arrayOfNames), PowerPoint.ShapeRange).ZOrder(MsoZOrderCmd.msoBringToFront)
                Catch ex As Exception

                End Try

            End If


        End If

        ' jetzt die Meilensteine in Vordergrund holen ...
        anzElements = msShapeNames.Count

        If anzElements > 0 Then

            ReDim arrayOfNames(anzElements - 1)
            For ix = 1 To anzElements
                arrayOfNames(ix - 1) = CStr(msShapeNames.Item(ix))
            Next

            Try
                CType(currentSlide.Shapes.Range(arrayOfNames), PowerPoint.ShapeRange).ZOrder(MsoZOrderCmd.msoBringToFront)
            Catch ex As Exception

            End Try

        End If


        If currentProjektIndex < projectCollection.Count And awinSettings.mppOnePage Then
            'Throw New ArgumentException("es konnten nur " & _
            '                            currentProjektIndex.ToString & " von " & projectsToDraw.ToString & _
            '                            " Projekten gezeichnet werden ... " & vbLf & _
            '                            "bitte verwenden Sie ein anderes Vorlagen-Format")
            Throw New ArgumentException("not all projects could be drawn ... please use other setitngs")
        End If



    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="rds"></param>
    ''' <param name="hproj"></param>
    ''' <param name="milestoneID"></param>
    ''' <param name="yPosition"></param>
    Friend Function drawMilestoneAtYPos(ByRef rds As clsPPTShapes, ByVal hproj As clsProjekt,
                                    ByVal swimlaneID As String, ByVal milestoneID As String,
                                    ByVal yPosition As Double) As PowerPoint.Shape

        Dim milestoneTypShape As PowerPoint.Shape = Nothing
        Dim milestoneTypApp As New clsAppearance
        Dim newShape As PowerPoint.Shape
        Dim milestoneName As String = elemNameOfElemID(milestoneID)
        Dim cMilestone As clsMeilenstein = hproj.getMilestoneByID(milestoneID)


        If IsNothing(cMilestone) Then
            drawMilestoneAtYPos = milestoneTypShape
            Exit Function ' einfach nichts machen 
        End If


        Dim x1 As Double
        Dim x2 As Double

        Dim stdTop As Single = 50
        Dim stdLeft As Single = 100
        Dim stdHeight As Single = 10
        Dim stdWidth As Single = 30

        Dim msShapeName As String = calcPPTShapeName(hproj, milestoneID)
        ' Es muss abgefragt werden, wie lange der NAme ist, evtl muss eine Fehlermeldung kommen .,.. 
        Dim nameLength As Integer = msShapeName.Length
        Dim msBeschriftung As String = hproj.getBestNameOfID(milestoneID, Not awinSettings.mppUseOriginalNames,
                                                             awinSettings.mppUseAbbreviation, swimlaneID)

        ' eigentlich muss es das sein ..
        Try
            If cMilestone.appearance = "" Then
                milestoneTypApp = appearanceDefinitions.Item(awinSettings.defaultMilestoneClass)
            Else
                milestoneTypApp = appearanceDefinitions.Item(cMilestone.appearance)
            End If
        Catch ex As Exception
            Dim i As Integer = 0
            milestoneTypApp = appearanceDefinitions.ElementAt(i).Value

            Do While Not milestoneTypApp.isMilestone And i < appearanceDefinitions.Count - 1
                i = i + 1
                milestoneTypApp = appearanceDefinitions.ElementAt(i).Value
            Loop

        End Try

        ' Exit , wenn nichts gefunden  
        If IsNothing(milestoneTypApp) Then
            drawMilestoneAtYPos = milestoneTypShape
            Exit Function ' einfach nichts machen 
        End If

        Dim sizeFaktor As Double

        ' die rds.milestoneVorlagenShape muss im Vorfeld bestimmt werden 
        'sizeFaktor = rds.avgMSHeight / milestoneTypApp.height

        sizeFaktor = 1.0

        Dim msDate As Date = cMilestone.getDate


        Call rds.calculatePPTx1x2(msDate, msDate, x1, x2)

        If x2 <= rds.drawingAreaLeft Or x1 >= rds.drawingAreaRight Then
            ' Fertig , es wird nix gezeichnet 
            Call MsgBox("Milestone outside drawing area ...")
        Else


            Try
                ' jetzt muss ggf die Beschriftung angebracht werden 
                ' die muss vor der Phase angebracht werden, weil der nicht von der Füllung des Schriftfeldes 
                ' überdeckt werden soll 
                If awinSettings.mppShowMsName Then

                    newShape = rds.addAnnotation(MsoTextOrientation.msoTextOrientationHorizontal, msShapeName, CStr(PTpptAnnotationType.text),
                                                 msBeschriftung, "", "Beschriftung", schriftGroesse)


                    With newShape

                        .Top = CSng(yPosition + rds.yOffsetMsToText)
                        .Left = CSng(x1) - .Width / 2


                    End With


                End If

                ' jetzt muss ggf das Datum angebracht werden 
                Dim msDateText As String = ""
                If awinSettings.mppShowMsDate Then

                    msDateText = msDate.Day.ToString & "." & msDate.Month.ToString

                    newShape = rds.addAnnotation(MsoTextOrientation.msoTextOrientationHorizontal, msShapeName, CStr(PTpptAnnotationType.datum),
                                                 msDateText, "", "Datum", schriftGroesse)

                    With newShape

                        .Top = CSng(yPosition + rds.yOffsetMsToDate)
                        .Left = CSng(x1) - .Width / 2

                    End With

                End If


                Dim height As Single = CSng(sizeFaktor * milestoneTypApp.height)
                Dim width As Single = CSng(sizeFaktor * milestoneTypApp.width)
                Dim top As Single = CSng(yPosition + rds.YMilestone)
                Dim left As Single = CSng(x1) - width / 2

                milestoneTypShape = rds.pptSlide.Shapes.AddShape(milestoneTypApp.shpType, left, top, width, height)

                If awinSettings.mppKwInMilestone Then

                    Call defineMsPPTAppearance(milestoneTypShape, milestoneTypApp, 1)

                    Dim msKwText As String = ""
                    msKwText = calcKW(msDate).ToString("0#")
                    If CInt(sizeFaktor * milestoneTypShape.TextFrame2.TextRange.Font.Size) >= 3 Then
                        milestoneTypShape.TextFrame2.TextRange.Font.Size = CInt(sizeFaktor * milestoneTypApp.TextRangeFontSize)
                        milestoneTypShape.TextFrame2.TextRange.Text = msKwText
                    End If

                Else

                    Call defineMsPPTAppearance(milestoneTypShape, milestoneTypApp)

                End If


                With milestoneTypShape

                    Try
                        .Name = msShapeName
                    Catch ex As Exception

                    End Try

                    If awinSettings.mppShowAmpel Then
                        .Glow.Color.RGB = CInt(cMilestone.getBewertung(1).color)
                        If .Glow.Radius = 0 Then
                            .Glow.Radius = 2
                        End If
                    End If

                End With

                If awinSettings.mppEnableSmartPPT Then
                    'Dim longText As String = hproj.hierarchy.getBestNameOfID(milestoneID, True, False)
                    'Dim shortText As String = hproj.hierarchy.getBestNameOfID(milestoneID, True, True)
                    'Dim originalName As String = cMilestone.originalName

                    Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(milestoneID)
                    Dim shortText As String = cMilestone.shortName
                    Dim originalName As String = cMilestone.originalName

                    Dim bestShortName As String = hproj.getBestNameOfID(cMilestone.nameID, True, True)
                    Dim bestLongName As String = hproj.getBestNameOfID(cMilestone.nameID, True, False)

                    If originalName = cMilestone.name Then
                        originalName = Nothing
                    End If

                    Dim lieferumfaenge As String = cMilestone.getAllDeliverables("#")
                    Call addSmartPPTMsPhInfo(milestoneTypShape, hproj,
                                                    fullBreadCrumb, cMilestone.name, shortText, originalName,
                                                    bestShortName, bestLongName,
                                                    Nothing, msDate,
                                                    cMilestone.getBewertung(1).colorIndex, cMilestone.getBewertung(1).description,
                                                    lieferumfaenge, cMilestone.verantwortlich, cMilestone.percentDone, cMilestone.DocURL)
                End If



            Catch ex As Exception
                Call MsgBox("fehler in zeichneMeilenstein;" & vbLf & ex.Message)
            End Try



        End If

        drawMilestoneAtYPos = milestoneTypShape

    End Function




    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="rds"></param>
    ''' <param name="hproj"></param>
    ''' <param name="phaseID"></param>
    ''' <param name="yPosition"></param>
    Friend Function drawPhaseAtYPos(ByRef rds As clsPPTShapes,
                                           ByVal hproj As clsProjekt,
                                           ByVal swimlaneID As String,
                                           ByVal phaseID As String,
                                           ByVal yPosition As Double) As PowerPoint.Shape

        Dim phShapeName As String = calcPPTShapeName(hproj, phaseID)

        Dim phaseTypShape As PowerPoint.Shape = Nothing
        Dim phaseTypApp As New clsAppearance
        Dim copiedShape As PowerPoint.Shape
        Dim phaseName As String = elemNameOfElemID(phaseID)
        Dim cphase As clsPhase = hproj.getPhaseByID(phaseID)

        If IsNothing(cphase) Then
            drawPhaseAtYPos = phaseTypShape
            Exit Function ' nichts machen 
        End If



        Dim x1 As Double
        Dim x2 As Double




        Try
            If cphase.appearance = "" Then
                phaseTypApp = appearanceDefinitions.Item(awinSettings.defaultPhaseClass)
            Else
                phaseTypApp = appearanceDefinitions.Item(cphase.appearance)
            End If
        Catch ex As Exception
            Dim i As Integer = 0
            phaseTypApp = appearanceDefinitions.ElementAt(i).Value

            Do While phaseTypApp.isMilestone And i < appearanceDefinitions.Count - 1
                i = i + 1
                phaseTypApp = appearanceDefinitions.ElementAt(i).Value
            Loop
        End Try


        If IsNothing(phaseTypApp) Then
            drawPhaseAtYPos = phaseTypShape
            Exit Function ' nichts machen 
        End If


        ' jetzt wegen evtl innerer Beschriftung den Size-Faktor bestimmen 
        Dim sizeFaktor As Double = 1.0

        If awinSettings.mppUseInnerText Then

            sizeFaktor = rds.phaseVorlagenShape.Height / phaseTypApp.height

        End If




        Dim phStartDate As Date = cphase.getStartDate
        Dim phEndDate As Date = cphase.getEndDate
        Dim phDateText As String = phStartDate.Day.ToString & "." & phStartDate.Month.ToString & " - " &
                                phEndDate.Day.ToString & "." & phEndDate.Month.ToString

        Dim phDescription As String = hproj.getBestNameOfID(phaseID, Not awinSettings.mppUseOriginalNames,
                                                                awinSettings.mppUseAbbreviation, swimlaneID)


        Call rds.calculatePPTx1x2(phStartDate, phEndDate, x1, x2)

        If x2 <= rds.drawingAreaLeft Or x1 >= rds.drawingAreaRight Then
            ' Fertig 
        Else

            ' jetzt muss ggf die Beschriftung angebracht werden 
            ' die muss vor der Phase angebracht werden, weil der nicht von der Füllung des Schriftfeldes 
            ' überdeckt werden soll 
            If awinSettings.mppShowPhName And (Not awinSettings.mppUseInnerText) Then

                copiedShape = createPPTShapeFromShape(rds.PhDescVorlagenShape, rds.pptSlide)
                With copiedShape

                    .TextFrame2.TextRange.Text = phDescription
                    .Top = CSng(yPosition + rds.YPhasenText)
                    .Left = CSng(x1)
                    If .Left + .Width > rds.drawingAreaRight + 2 Then
                        .Left = CSng(rds.drawingAreaRight - .Width + 2)
                    End If

                    '.Name = .Name & .Id
                    Try
                        .Name = phShapeName & PTpptAnnotationType.text
                    Catch ex As Exception

                    End Try

                    .Title = "Beschriftung"
                    .AlternativeText = ""



                End With


            End If

            ' jetzt muss ggf das Datum angebracht werden 
            If awinSettings.mppShowPhDate And (Not awinSettings.mppUseInnerText) Then



                copiedShape = createPPTShapeFromShape(rds.PhDateVorlagenShape, rds.pptSlide)
                With copiedShape

                    .TextFrame2.TextRange.Text = phDateText
                    .Top = CSng(yPosition + rds.YPhasenDatum)
                    .Left = CSng(x1)
                    If .Left + .Width > rds.drawingAreaRight + 2 Then
                        .Left = CSng(rds.drawingAreaRight - .Width + 2)
                    End If

                    '.Name = .Name & .Id
                    Try
                        .Name = phShapeName & PTpptAnnotationType.datum
                    Catch ex As Exception

                    End Try

                    .Title = "Datum"
                    .AlternativeText = ""



                End With

            End If



            ''End With
            Dim top As Single = CSng(yPosition + rds.YPhase)
            Dim heigth As Single = rds.phaseVorlagenShape.Height
            Dim width As Single = CSng(x2 - x1)
            Dim left As Single = CSng(x1)

            phaseTypShape = rds.pptSlide.Shapes.AddShape(phaseTypApp.shpType, left, top, width, heigth)

            Call definePhPPTAppearance(phaseTypShape, phaseTypApp)

            With phaseTypShape
                Try
                    .Name = phShapeName
                Catch ex As Exception

                End Try


                ' jetzt wird die Option gezogen, wenn keine Phasen-Beschriftung stattfinden sollte ... 
                If awinSettings.mppUseInnerText Then

                    If awinSettings.mppShowPhDate Then
                        phDescription = phDescription & " " & phDateText
                    End If

                    If sizeFaktor * .TextFrame2.TextRange.Font.Size * sizeFaktor > 3.0 Then
                        .TextFrame2.TextRange.Text = phDescription
                        .TextFrame2.TextRange.Font.Size = CInt(.TextFrame2.TextRange.Font.Size * sizeFaktor)
                    End If
                End If



            End With

            If awinSettings.mppEnableSmartPPT Then
                'Dim shortText As String = hproj.hierarchy.getBestNameOfID(cphase.nameID, True, _
                '                                          True)
                'Dim longText As String = hproj.hierarchy.getBestNameOfID(cphase.nameID, True, _
                '                                       False)
                'Dim originalName As String = cphase.originalName

                Dim fullBreadCrumb As String = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                Dim shortText As String = cphase.shortName
                Dim originalName As String = cphase.originalName

                Dim bestShortName As String = hproj.getBestNameOfID(cphase.nameID, True, True)
                Dim bestLongName As String = hproj.getBestNameOfID(cphase.nameID, True, False)

                If originalName = cphase.name Then
                    originalName = Nothing
                End If

                Call addSmartPPTMsPhInfo(phaseTypShape, hproj,
                                            fullBreadCrumb, cphase.name, shortText, originalName,
                                            bestShortName, bestLongName,
                                            phStartDate, phEndDate,
                                            cphase.ampelStatus, cphase.ampelErlaeuterung, cphase.getAllDeliverables("#"),
                                            cphase.verantwortlich, cphase.percentDone, cphase.DocURL)
            End If

        End If



        drawPhaseAtYPos = phaseTypShape

    End Function


    ''' <summary>
    ''' adds / draws  
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <param name="pvName"></param>
    ''' <param name="breadCrumb"></param>
    ''' <param name="elemName"></param>
    ''' <param name="isMilestones"></param>
    ''' <param name="atleastOneAddedElement"></param>
    Friend Sub drawElemOfProject(ByVal hproj As clsProjekt, ByVal pvName As String, ByVal breadCrumb As String, ByVal elemName As String, ByVal isMilestones As Boolean,
                                  ByRef atleastOneAddedElement As Boolean)


        Dim parentNameID As String = ""
        Dim parentName As String = ""

        Dim currentMilestone As clsMeilenstein = Nothing
        Dim currentPhase As clsPhase = Nothing
        Dim allOK As Boolean = False

        If isMilestones Then
            currentMilestone = hproj.getMilestone(elemName, breadcrumb:=breadCrumb)
            allOK = Not IsNothing(currentMilestone)
        Else
            currentPhase = hproj.getPhase(elemName, breadcrumb:=breadCrumb)
            allOK = Not IsNothing(currentPhase)
        End If


        If allOK Then
            ' jetzt muss die yPos bestimmt werden , das ist die YPos des nächstgelegenen Vaters im BreadCrumb ...
            Dim found As Boolean = False
            Dim yPos As Double = 30 ' Default Wert
            Dim myNameID As String
            Dim myName As String
            If isMilestones Then
                parentNameID = hproj.hierarchy.getParentIDOfID(currentMilestone.nameID)
                myNameID = currentMilestone.nameID
                myName = currentMilestone.name
            Else
                parentNameID = hproj.hierarchy.getParentIDOfID(currentPhase.nameID)
                myNameID = currentPhase.nameID
                myName = currentPhase.name
            End If

            If parentNameID <> "" Then
                Dim parentPhase = hproj.getPhaseByID(parentNameID)
                If Not IsNothing(parentPhase) Then
                    parentName = parentPhase.name
                End If
            End If

            ' erstmal nach Geschwistern suchen ..
            ' elemName ist leer, weil jedes Geschwister diesen Breadcrumb hat 
            Dim sisterBreadCrumb As String = smartSlideLists.bestimmeFullBreadcrumb(calcProjektKey(hproj), hproj.hierarchy.getBreadCrumb(myNameID), "")

            Dim sisterOrParentShapeName As String = smartSlideLists.getShapeNameWithBreadCrumb(sisterBreadCrumb)
            Dim foundShape As PowerPoint.Shape = Nothing

            If sisterOrParentShapeName <> "" Then
                foundShape = currentSlide.Shapes.Item(sisterOrParentShapeName)
            End If

            If Not IsNothing(foundShape) Then
                found = True
                yPos = foundShape.Top
            End If

            ' dann nach Eltern suchen ...
            If Not found Then
                Dim parentBreadcrumb As String = smartSlideLists.bestimmeFullBreadcrumb(calcProjektKey(hproj), hproj.hierarchy.getBreadCrumb(parentNameID), parentName)
                If parentBreadcrumb <> "" Then
                    sisterOrParentShapeName = smartSlideLists.getShapeNameWithBreadCrumb(parentBreadcrumb)
                    If sisterOrParentShapeName <> "" Then
                        foundShape = currentSlide.Shapes.Item(sisterOrParentShapeName)
                        If Not IsNothing(foundShape) Then
                            found = True
                            yPos = foundShape.Top
                        End If
                    End If
                End If
            End If



            If isMilestones Then
                ' draw the Milestone 
                Dim newMsShape As PowerPoint.Shape = drawMilestoneAtYPos(slideCoordInfo, hproj:=hproj, swimlaneID:=parentNameID, milestoneID:=currentMilestone.nameID, yPosition:=yPos)
                atleastOneAddedElement = True
            Else
                Dim newPhaseShape As PowerPoint.Shape = drawPhaseAtYPos(slideCoordInfo, hproj:=hproj, swimlaneID:=parentNameID, phaseID:=currentPhase.nameID, yPosition:=yPos)
                atleastOneAddedElement = True
            End If

        End If




    End Sub



End Module
