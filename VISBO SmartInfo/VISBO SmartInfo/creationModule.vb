Imports ProjectBoardDefinitions
Imports System.Windows.Forms
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports System.Collections
Imports ProjectBoardBasic
Module creationModule



    Friend curSlide As PowerPoint.Slide = Nothing

    Friend curPresentation As PowerPoint.Presentation = Nothing

    ' bestimmt, ob in englisch oder auf deutsch ..
    Friend useEnglishLanguage As Boolean = True

    Friend noDBLoginPPT As Boolean = True

    ''' <summary>
    ''' erzeugt den Bericht auf Grundlage des aktuell geladenen Powerpoints  
    ''' bei Aufruf ist sichergestellt, daß in Projekthistorie die Historie des Projektes steht 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub createPPTSlidesFromProjectWithinPPT(ByRef hproj As clsProjekt,
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
        ' tk 4.10.19 das wird hier nicht mehr gebraucht , es wird ja die aktuelle Presentation als Vorlage benutzt ...
        'Dim presentationFile As String = awinPath & requirementsOrdner & "projektdossier.pptx"
        'Dim presentationFileH As String = awinPath & requirementsOrdner & "projektdossier_Hochformat.pptx"
        'Dim newFileName As String = reportOrdnerName & "Report.pptx"

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



            If Not noDBLoginPPT Then

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

            Call addSmartPPTSlideBaseInfo(curSlide, reportCreationDate, ptPRPFType.project)

            ' jetzt werden die Charts gezeichnet 
            anzShapes = curSlide.Shapes.Count
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
                pptShape = curSlide.Shapes(i)
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
                        kennzeichnung = "Soll-Ist & Prognose" Or
                        kennzeichnung = "Multivariantensicht" Or
                        kennzeichnung = "Einzelprojektsicht" Or
                        kennzeichnung = "Multiprojektsicht" Or
                        kennzeichnung = "AllePlanElemente" Or
                        kennzeichnung = "Swimlanes" Or
                        kennzeichnung = "Swimlanes2" Or
                        kennzeichnung = "MilestoneCategories" Or
                        kennzeichnung = "Legenden-Tabelle" Or
                        kennzeichnung = "Meilenstein Trendanalyse" Or
                        kennzeichnung = "Vergleich mit Beauftragung" Or
                        kennzeichnung = "Vergleich mit letztem Stand" Or
                        kennzeichnung = "Vergleich mit Vorlage" Or
                        kennzeichnung = "TableBudgetCostAPVCV" Or
                        kennzeichnung = "TableMilestoneAPVCV" Or
                        kennzeichnung = "Tabelle Projektziele" Or
                        kennzeichnung = "Tabelle Projektstatus" Or
                        kennzeichnung = "Tabelle Veränderungen" Or
                        kennzeichnung = "Tabelle Vergleich letzter Stand" Or
                        kennzeichnung = "Tabelle Vergleich Beauftragung" Or
                        kennzeichnung = "Tabelle OneGlance Beauftragung" Or
                        kennzeichnung = "Tabelle OneGlance letzter Stand" Or
                        kennzeichnung = "Ergebnis" Or
                        kennzeichnung = "Strategie/Risiko" Or
                        kennzeichnung = "Strategie/Risiko/Ausstrahlung" Or
                        kennzeichnung = "Projektphasen" Or
                        kennzeichnung = "ProjektBedarfsChart" Or
                        kennzeichnung = "Personalbedarf" Or
                        kennzeichnung = "Personalbedarf2" Or
                        kennzeichnung = "Personalkosten" Or
                        kennzeichnung = "Personalkosten2" Or
                        kennzeichnung = "Sonstige Kosten" Or
                        kennzeichnung = "Sonstige Kosten2" Or
                        kennzeichnung = "Gesamtkosten" Or
                        kennzeichnung = "Gesamtkosten2" Or
                        kennzeichnung = "Trend Strategischer Fit/Risiko" Or
                        kennzeichnung = "Trend Kennzahlen" Or
                        kennzeichnung = "Fortschritt Personalkosten" Or
                        kennzeichnung = "Fortschritt Sonstige Kosten" Or
                        kennzeichnung = "Fortschritt Rolle" Or
                        kennzeichnung = "Fortschritt Kostenart" Or
                        kennzeichnung = "Soll-Ist1 Personalkosten" Or
                        kennzeichnung = "Soll-Ist2 Personalkosten" Or
                        kennzeichnung = "Soll-Ist1C Personalkosten" Or
                        kennzeichnung = "Soll-Ist2C Personalkosten" Or
                        kennzeichnung = "Soll-Ist1 Sonstige Kosten" Or
                        kennzeichnung = "Soll-Ist2 Sonstige Kosten" Or
                        kennzeichnung = "Soll-Ist1C Sonstige Kosten" Or
                        kennzeichnung = "Soll-Ist2C Sonstige Kosten" Or
                        kennzeichnung = "Soll-Ist1 Gesamtkosten" Or
                        kennzeichnung = "Soll-Ist2 Gesamtkosten" Or
                        kennzeichnung = "Soll-Ist1C Gesamtkosten" Or
                        kennzeichnung = "Soll-Ist2C Gesamtkosten" Or
                        kennzeichnung = "Soll-Ist1 Rolle" Or
                        kennzeichnung = "Soll-Ist2 Rolle" Or
                        kennzeichnung = "Soll-Ist1C Rolle" Or
                        kennzeichnung = "Soll-Ist2C Rolle" Or
                        kennzeichnung = "Soll-Ist1 Kostenart" Or
                        kennzeichnung = "Soll-Ist2 Kostenart" Or
                        kennzeichnung = "Soll-Ist1C Kostenart" Or
                        kennzeichnung = "Soll-Ist2C Kostenart" Or
                        kennzeichnung = "Ampel-Farbe" Or
                        kennzeichnung = "Ampel-Text" Or
                        kennzeichnung = "Beschreibung" Or
                        kennzeichnung = "Business-Unit:" Or
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
                        kennzeichnung = "MilestoneCategories" Or
                        kennzeichnung = "Meilenstein Trendanalyse" Or
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

                Dim frmSelectionPhMs As New frmSelectPhasesMilestones
                If frmSelectionPhMs.ShowDialog = Windows.Forms.DialogResult.OK Then
                    selectedPhases = frmSelectionPhMs.selectedPhases
                    selectedMilestones = frmSelectionPhMs.selectedMilestones

                Else
                    selectedPhases = New Collection
                    selectedMilestones = New Collection
                End If

                phMSSelNeeded(1) = True
                If Not IsNothing(listOfFormerSelectedProjects) Then
                    selectedProjekte.Clear(False)

                    For Each tmpName As String In listOfFormerSelectedProjects
                        If ShowProjekte.contains(tmpName) Then
                            selectedProjekte.Add(ShowProjekte.getProject(tmpName), False)
                        End If
                    Next
                End If

                ' jetzt muss für den Multiprojekt Report noch showrangeLeft und Right gesetzt werden 
                showRangeLeft = ShowProjekte.getMinMonthColumn - 1
                showRangeRight = ShowProjekte.getMaxMonthColumn + 3

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
                                    Call zeichneMultiprojektSichtinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
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
                                    Call zeichneMultiprojektSichtinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
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
                                    Call zeichneMultiprojektSichtinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
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

                            Case "Multivariantensicht"


                                Try

                                    Dim minCal As Boolean = False
                                    If qualifier2.Length > 0 Then
                                        minCal = (qualifier2.Trim = "minCal")
                                    End If

                                    Dim pptFirstTime As Boolean = True
                                    Call zeichneMultiprojektSichtinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, CDbl(legendFontSize),
                                                                      selectedPhases, selectedMilestones,
                                                                      translateToRoleNames(selectedRoles), selectedCosts,
                                                                      selectedBUs, selectedTyps,
                                                                      False, True, hproj, kennzeichnung, minCal)
                                    .TextFrame2.TextRange.Text = ""
                                    '.ZOrder(MsoZOrderCmd.msoSendToBack)
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message
                                    objectsDone = objectsToDo
                                End Try


                            Case "MilestoneCategories"

                                Call MsgBox("not implemented ...")

                                'Try

                                '    Dim minCal As Boolean = False
                                '    If qualifier2.Length > 0 Then
                                '        minCal = (qualifier2.Trim = "minCal")
                                '    End If

                                '    Dim pptFirstTime As Boolean = True
                                '    Call zeichneCategorySwimlaneSichtinPPT(objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, legendFontSize,
                                '                                      selectedPhases, selectedMilestones,
                                '                                      translateToRoleNames(selectedRoles), selectedCosts,
                                '                                      selectedBUs, selectedTyps,
                                '                                      False, hproj, kennzeichnung, minCal)

                                '    .TextFrame2.TextRange.Text = ""
                                '    '.ZOrder(MsoZOrderCmd.msoSendToBack)

                                '    ' sonst wird pptLasttime benötigt, um bei mehreren PRojekten 
                                '    ' swimlaneMode wird erst nach Ende der While Schleife ausgewertet - in diesem Fall wird die tmpSav Folie gelöscht 
                                '    'swimlaneMode = True
                                'Catch ex As Exception
                                '    .TextFrame2.TextRange.Text = ex.Message
                                '    objectsDone = objectsToDo
                                'End Try


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


                            Case "Meilenstein Trendanalyse"

                                Call MsgBox("not implemented ...")


                                'Dim nameList As New SortedList(Of Date, String)
                                'Dim listOfItems As New Collection

                                ''boxName = "Meilenstein Trendanalyse"
                                'boxName = repMessages.getmsg(21)

                                'Try
                                '    ' Aufruf 
                                '    If qualifier = "" Then
                                '        ' alle Meilensteine anzeigen
                                '        nameList = hproj.getMilestones

                                '        If nameList.Count > 0 Then
                                '            For Each kvp As KeyValuePair(Of Date, String) In nameList
                                '                listOfItems.Add(kvp.Value)
                                '            Next
                                '        End If


                                '    Else
                                '        ' nur die anzeigen, die im qualifier mit # voneinander getrennt stehen  
                                '        Dim tmpStr(20) As String
                                '        Try

                                '            tmpStr = qualifier.Trim.Split(New Char() {CChar("#")}, 20)
                                '            kennzeichnung = tmpStr(0).Trim

                                '        Catch ex As Exception

                                '        End Try


                                '        ' die ListofItems muss die eindeutigen IDs beeinhalten
                                '        For i = 1 To tmpStr.Length

                                '            Dim fullmsName As String = tmpStr(i - 1).Trim
                                '            Dim msName As String = ""
                                '            Dim breadcrumb As String = ""
                                '            Dim type As Integer = -1
                                '            Dim pvName As String = ""
                                '            Call splitHryFullnameTo2(fullmsName, msName, breadcrumb, type, pvName)


                                '            Dim milestoneIndices(,) As Integer = hproj.hierarchy.getMilestoneIndices(msName, breadcrumb)
                                '            Dim msItem As String

                                '            For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                                '                If milestoneIndices(0, mx) > 0 And milestoneIndices(1, mx) > 0 Then

                                '                    Try
                                '                        msItem = hproj.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx)).nameID
                                '                        listOfItems.Add(msItem)
                                '                    Catch ex As Exception

                                '                    End Try


                                '                End If

                                '            Next

                                '        Next

                                '    End If

                                '    ' jetzt ist listofItems entsprechend gefüllt 
                                '    If listOfItems.Count > 0 Then
                                '        htop = 100
                                '        hleft = 50
                                '        hheight = 2 * ((listOfItems.Count - 1) * 20 + 110)
                                '        hwidth = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 10, 24 * boxWidth + 10)

                                '        Try
                                '            Call createMsTrendAnalysisOfProject(hproj, obj, listOfItems, htop, hleft, hheight, hwidth)

                                '            reportObj = obj

                                '            bigType = ptReportBigTypes.charts
                                '            compID = PTprdk.MilestoneTrendanalysis
                                '            Call addSmartPPTCompInfo(newShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                '                                     bigType, compID)

                                '        Catch ex As Exception
                                '            '.TextFrame2.TextRange.Text = "zum Projekt" & hproj.name & vbLf & "gibt es noch keine Trend-Analyse," & vbLf & _
                                '            '                            "da es noch nicht begonnen hat"
                                '            .TextFrame2.TextRange.Text = hproj.name & repMessages.getmsg(22)
                                '        End Try

                                '    Else
                                '        '.TextFrame2.TextRange.Text = "es gibt keine Meilensteine im Projekt" & vbLf & hproj.name
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(23) & vbLf & hproj.name
                                '    End If

                                'Catch ex As Exception
                                '    Throw New ArgumentException("Fehler in MeilensteinTrendAnalyse in CreatePPTSlidesFromProject")
                                'End Try

                            Case "Projektphasen"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(239)
                                'End If

                                'Dim scale As Integer
                                'Dim continueWork As Boolean = True
                                'Dim cproj As clsProjekt = Nothing
                                'Dim vproj As clsProjektvorlage
                                'auswahl = 0

                                'scale = hproj.dauerInDays

                                'If qualifier.Length > 0 Then
                                '    If qualifier = "Vorlage" Then
                                '        auswahl = 1
                                '        vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                                '        If IsNothing(vproj) Then
                                '            '.TextFrame2.TextRange.Text = "Projekt-Vorlage " & hproj.VorlagenName & " existiert nicht !"
                                '            .TextFrame2.TextRange.Text = repMessages.getmsg(24) & hproj.VorlagenName
                                '            continueWork = False
                                '        Else
                                '            vproj.copyTo(cproj)
                                '            cproj.startDate = hproj.startDate
                                '        End If

                                '    ElseIf qualifier = "Beauftragung" Then
                                '        cproj = bproj
                                '        auswahl = 2

                                '    Else
                                '        cproj = hproj
                                '        auswahl = 0

                                '    End If
                                'Else
                                '    cproj = hproj
                                '    auswahl = 0
                                'End If

                                'If continueWork Then
                                '    htop = 150
                                '    hleft = 150


                                '    hheight = 380
                                '    hwidth = 900
                                '    scale = cproj.dauerInDays

                                '    Dim noColorCollection As New Collection
                                '    reportObj = Nothing

                                '    Call createPhasesBalken(noColorCollection, cproj, reportObj, scale, htop, hleft, hheight, hwidth, auswahl)



                                '    bigType = ptReportBigTypes.charts
                                '    compID = PTprdk.Phasen
                                '    Call addSmartPPTCompInfo(newShape, hproj, Nothing, ptPRPFType.project, qualifier, qualifier2,
                                '                                bigType, compID)

                                'End If


                            Case "Vergleich mit Vorlage"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(238)
                                'End If


                                'Dim vproj As clsProjektvorlage
                                'Dim cproj As New clsProjekt
                                'Dim scale As Double
                                'Dim noColorCollection As New Collection
                                'Dim repObj1 As xlNS.ChartObject, repObj2 As xlNS.ChartObject
                                'Dim continueWork As Boolean = True

                                '' jetzt die Aktion durchführen ...


                                'Try

                                '    vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                                '    If IsNothing(vproj) Then
                                '        '.TextFrame2.TextRange.Text = "Projekt-Vorlage " & hproj.VorlagenName & " existiert nicht !"
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(24) & hproj.VorlagenName
                                '        continueWork = False
                                '    Else
                                '        cproj = New clsProjekt
                                '        vproj.copyTo(cproj)
                                '        cproj.startDate = hproj.startDate
                                '    End If

                                'Catch ex As Exception
                                '    'Throw New Exception("Vorlage konnte nicht bestimmt werden")
                                '    Throw New Exception(repMessages.getmsg(25))
                                'End Try

                                'If continueWork Then
                                '    htop = 150
                                '    hleft = 150


                                '    hheight = 380
                                '    hwidth = 900
                                '    scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)


                                '    appInstance.EnableEvents = False


                                '    noColorCollection = getPhasenUnterschiede(hproj, cproj)

                                '    repObj1 = Nothing
                                '    Call createPhasesBalken(noColorCollection, hproj, repObj1, scale, htop, hleft, hheight, hwidth, PThis.current)

                                '    With repObj1
                                '        htop = .Top + .Height + 3
                                '    End With


                                '    repObj2 = Nothing
                                '    Call createPhasesBalken(noColorCollection, cproj, repObj2, scale, htop, hleft, hheight, hwidth, PThis.vorlage)

                                '    ' jetzt wird das Shape in der Powerpoint entsprechend entsprechend aufgebaut 
                                '    Try
                                '        pptSize = CInt(.TextFrame2.TextRange.Font.Size)
                                '        .TextFrame2.TextRange.Text = " "
                                '    Catch ex As Exception
                                '        pptSize = 12
                                '    End Try


                                '    Dim widthFaktor As Double = 1.0
                                '    Dim heightFaktor As Double = 1.0
                                '    Dim topNext As Double


                                '    If Not repObj1 Is Nothing Then
                                '        Try
                                '            ''repObj1.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                '            ''newShapeRange = pptSlide.Shapes.Paste
                                '            newShapeRange = pictCopypptPaste(repObj1, pptSlide)

                                '            With newShapeRange(1)
                                '                .Top = CSng(top + 0.02 * height)
                                '                .Left = CSng(left + 0.02 * width)
                                '                .Width = CSng(width * 0.96)
                                '                topNext = CSng(top + 0.04 * height + .Height)
                                '                '.Height = height * 0.46
                                '            End With

                                '            repObj1.Delete()

                                '            If Not repObj2 Is Nothing Then
                                '                Try
                                '                    ''repObj2.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                '                    ''newShapeRange2 = pptSlide.Shapes.Paste
                                '                    newShapeRange2 = pictCopypptPaste(repObj2, pptSlide)

                                '                    With newShapeRange2(1)
                                '                        .Top = CSng(topNext)
                                '                        .Left = CSng(left + 0.02 * width)
                                '                        .Width = CSng(width * 0.96)
                                '                        ' Height wird nicht gesetzt - bei Bildern wird das proportional automatisch gesetzt 
                                '                    End With

                                '                    ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                '                    Try
                                '                        If newShapeRange(1).Height + newShapeRange2(1).Height > 0.96 * height Then
                                '                            widthFaktor = 0.96 * height / (newShapeRange(1).Height + newShapeRange2(1).Height)
                                '                            newShapeRange(1).Width = CSng(widthFaktor * newShapeRange(1).Width)
                                '                            newShapeRange2(1).Width = CSng(widthFaktor * newShapeRange2(1).Width)
                                '                            newShapeRange2(1).Top = CSng(newShapeRange(1).Top + newShapeRange(1).Height + 0.02 * height)
                                '                        End If
                                '                    Catch ex As Exception

                                '                    End Try



                                '                    repObj2.Delete()
                                '                Catch ex As Exception

                                '                End Try

                                '            End If
                                '        Catch ex As Exception

                                '        End Try

                                '    End If

                                'End If


                            Case "Vergleich mit Beauftragung"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(237)
                                'End If


                                'Dim cproj As clsProjekt
                                'Dim scale As Double
                                'Dim noColorCollection As New Collection
                                'Dim repObj1 As xlNS.ChartObject, repObj2 As xlNS.ChartObject



                                '' jetzt die Aktion durchführen ...


                                'If bproj Is Nothing Then
                                '    'Throw New Exception("es gibt keine Beauftragung")
                                '    Throw New Exception(repMessages.getmsg(26))
                                'End If

                                'cproj = bproj



                                'htop = 150
                                'hleft = 150

                                'hheight = 380
                                'hwidth = 900
                                'scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)


                                'noColorCollection = getPhasenUnterschiede(hproj, cproj)

                                'repObj1 = Nothing
                                'Call createPhasesBalken(noColorCollection, hproj, repObj1, scale, htop, hleft, hheight, hwidth, PThis.current)

                                'With repObj1
                                '    htop = .Top + .Height + 3
                                'End With

                                'repObj2 = Nothing
                                'Call createPhasesBalken(noColorCollection, cproj, repObj2, scale, htop, hleft, hheight, hwidth, PThis.beauftragung)

                                'Try
                                '    pptSize = CInt(.TextFrame2.TextRange.Font.Size)
                                '    .TextFrame2.TextRange.Text = " "
                                'Catch ex As Exception
                                '    pptSize = 12
                                'End Try



                                'Dim widthFaktor As Double = 1.0
                                'Dim heightFaktor As Double = 1.0
                                'Dim topNext As Double

                                'If Not repObj1 Is Nothing Then
                                '    Try
                                '        ''repObj1.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                '        ''newShapeRange = pptSlide.Shapes.Paste
                                '        newShapeRange = pictCopypptPaste(repObj1, pptSlide)

                                '        With newShapeRange(1)
                                '            .Top = CSng(top + 0.02 * height)
                                '            .Left = CSng(left + 0.02 * width)
                                '            .Width = CSng(width * 0.96)
                                '            topNext = CSng(top + 0.04 * height + .Height)
                                '            '.Height = height * 0.46
                                '        End With

                                '        repObj1.Delete()

                                '        If Not repObj2 Is Nothing Then
                                '            Try
                                '                ''repObj2.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                '                ''newShapeRang2 = pptSlide.Shapes.Paste
                                '                newShapeRange2 = pictCopypptPaste(repObj2, pptSlide)

                                '                With newShapeRange2(1)
                                '                    .Top = CSng(topNext)
                                '                    .Left = CSng(left + 0.02 * width)
                                '                    .Width = CSng(width * 0.96)
                                '                    '.Height = height * 0.46
                                '                End With

                                '                repObj2.Delete()

                                '                ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                '                If newShapeRange(1).Height + newShapeRange2(1).Height > 0.96 * height Then
                                '                    widthFaktor = 0.96 * height / (newShapeRange(1).Height + newShapeRange2(1).Height)
                                '                    newShapeRange(1).Width = CSng(widthFaktor * newShapeRange(1).Width)
                                '                    newShapeRange2(1).Width = CSng(widthFaktor * newShapeRange2(1).Width)
                                '                    newShapeRange2(1).Top = CSng(newShapeRange(1).Top + newShapeRange(1).Height + 0.02 * height)
                                '                End If
                                '            Catch ex As Exception

                                '            End Try

                                '        End If


                                '    Catch ex As Exception

                                '    End Try

                                'End If


                            Case "Vergleich mit letztem Stand"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(236)
                                'End If


                                'Dim cproj As clsProjekt
                                'Dim scale As Double
                                'Dim noColorCollection As New Collection
                                'Dim repObj1 As xlNS.ChartObject, repObj2 As xlNS.ChartObject



                                '' jetzt die Aktion durchführen ...

                                'If lastproj Is Nothing Then
                                '    Try
                                '        .TextFrame2.TextRange.Text = "Fehler: ... " & repMessages.getmsg(27)
                                '    Catch ex As Exception
                                '        pptSize = 12
                                '    End Try
                                '    'Throw New Exception("es gibt keinen letzten Stand")
                                '    Throw New Exception(repMessages.getmsg(27))
                                'End If

                                'cproj = lastproj

                                'htop = 150
                                'hleft = 150

                                'hheight = 380
                                'hwidth = 900
                                'scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)


                                'noColorCollection = getPhasenUnterschiede(hproj, cproj)

                                'repObj1 = Nothing
                                'Call createPhasesBalken(noColorCollection, hproj, repObj1, scale, htop, hleft, hheight, hwidth, PThis.current)

                                'With repObj1
                                '    htop = .Top + .Height + 3
                                'End With

                                'repObj2 = Nothing
                                'Call createPhasesBalken(noColorCollection, cproj, repObj2, scale, htop, hleft, hheight, hwidth, PThis.letzterStand)

                                'Try
                                '    pptSize = .TextFrame2.TextRange.Font.Size
                                '    .TextFrame2.TextRange.Text = " "
                                'Catch ex As Exception
                                '    pptSize = 12
                                'End Try



                                'Dim widthFaktor As Double = 1.0
                                'Dim heightFaktor As Double = 1.0
                                'Dim topNext As Double

                                'If Not repObj1 Is Nothing Then
                                '    Try
                                '        ''repObj1.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                '        ''newShapeRange = pptSlide.Shapes.Paste
                                '        newShapeRange = pictCopypptPaste(repObj1, pptSlide)

                                '        With newShapeRange(1)
                                '            .Top = CSng(top + 0.02 * height)
                                '            .Left = CSng(left + 0.02 * width)
                                '            .Width = CSng(width * 0.96)
                                '            topNext = CSng(top + 0.04 * height + .Height)
                                '            '.Height = height * 0.46
                                '        End With

                                '        repObj1.Delete()

                                '        If Not repObj2 Is Nothing Then
                                '            Try
                                '                ''repObj2.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                '                ''newShapeRange2 = pptSlide.Shapes.Paste                                             
                                '                newShapeRange2 = pictCopypptPaste(repObj2, pptSlide)

                                '                With newShapeRange2(1)
                                '                    .Top = CSng(topNext)
                                '                    .Left = CSng(left + 0.02 * width)
                                '                    .Width = CSng(width * 0.96)
                                '                    '.Height = height * 0.46
                                '                End With

                                '                repObj2.Delete()

                                '                ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                '                If newShapeRange(1).Height + newShapeRange2(1).Height > 0.96 * height Then
                                '                    widthFaktor = 0.96 * height / (newShapeRange(1).Height + newShapeRange2(1).Height)
                                '                    newShapeRange(1).Width = CSng(widthFaktor * newShapeRange(1).Width)
                                '                    newShapeRange2(1).Width = CSng(widthFaktor * newShapeRange2(1).Width)
                                '                    newShapeRange2(1).Top = CSng(newShapeRange(1).Top + newShapeRange(1).Height + 0.02 * height)
                                '                End If
                                '            Catch ex As Exception

                                '            End Try

                                '        End If


                                '    Catch ex As Exception

                                '    End Try

                                'End If


                            Case "Tabelle Projektziele"

                                Try
                                    ' wenn es im Qualifier angegebene Meilensteine gibt, dann haben die Prio vor der interaktiven Auswahl 
                                    ' 
                                    Dim sMilestones As Collection = selectedMilestones

                                    If Not IsNothing(qualifier2) Then
                                        If qualifier2.Length > 0 Then
                                            sMilestones = New Collection
                                            Dim tmpStr() As String = qualifier2.Split(New Char() {CChar(vbLf), CChar(vbCr)})
                                            For Each tmpMsName As String In tmpStr
                                                If Not sMilestones.Contains(tmpMsName) Then
                                                    sMilestones.Add(tmpMsName, tmpMsName)
                                                End If

                                            Next
                                        End If
                                    End If


                                    ' die smart Powerpoint Table Info wird in dieser MEthode gesetzt ...
                                    ' tk 24.6.18 damit man unabhängig von selectedMilestones in der PPT-Vorlage feste Meilensteine angeben kann 
                                    Call zeichneProjektTabelleZiele(pptShape, hproj, sMilestones, "", "")
                                    'Call zeichneProjektTabelleZiele(pptShape, hproj, selectedMilestones, qualifier, qualifier2)


                                Catch ex As Exception

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

                            Case "Tabelle Vergleich letzter Stand"

                                Try
                                    Call zeichneProjektTabelleVergleich(curSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, lastproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle Vergleich Beauftragung"

                                Try
                                    Call zeichneProjektTabelleVergleich(curSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, bproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle OneGlance letzter Stand"

                                Try
                                    Call zeichneProjektTabelleOneGlance(curSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, lastproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle OneGlance Beauftragung"


                                Try
                                    Call zeichneProjektTabelleOneGlance(curSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, bproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle Veränderungen"


                                Try

                                    ' Für englische Version muss Template auf Englisch sein
                                    Call zeichneProjektTerminAenderungen(pptShape, hproj, bproj, lproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle Projektstatus"


                                Try
                                    Call zeichneProjektTabelleStatus(pptShape, hproj)
                                Catch ex As Exception

                                End Try

                            Case "Soll-Ist & Prognose"
                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(235)
                                'End If

                                'If istWerteexistieren Then
                                'Else
                                '    '.TextFrame2.TextRange.Text = "Prognose"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(38)
                                'End If

                            Case "Ergebnis"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(212)
                                'End If



                                'Try

                                '    If qualifier = "letzter Stand" Then

                                '        Call createProjektErgebnisCharakteristik2(lproj, obj, PThis.letzterStand,
                                '                                                  5, 5, 280, 180, True)

                                '    ElseIf qualifier = "Beauftragung" Then
                                '        Call createProjektErgebnisCharakteristik2(bproj, obj, PThis.beauftragung,
                                '                                                  5, 5, 280, 180, True)

                                '    Else
                                '        Call createProjektErgebnisCharakteristik2(hproj, obj, PThis.current,
                                '                                                  5, 5, 280, 180, True)

                                '    End If



                                '    reportObj = obj

                                '    Dim ax As xlNS.Axis = CType(reportObj.Chart.Axes(xlNS.XlAxisType.xlCategory), Excel.Axis)
                                '    With ax
                                '        .TickLabels.Font.Size = 12
                                '    End With

                                '    notYetDone = True
                                '    bigType = ptReportBigTypes.charts
                                '    compID = PTprdk.Ergebnis

                                'Catch ex As Exception

                                'End Try



                            Case "Strategie/Risiko"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(234)
                                'End If


                                'Dim mycollection As New Collection

                                ''deleteStack.Add(.Name, .Name)
                                'Try
                                '    mycollection.Add(pname)

                                '    Call awinCreatePortfolioDiagrams(mycollection, reportObj, True, PTpfdk.FitRisiko, PTpfdk.ProjektFarbe, True, False, True, htop, hleft, hwidth, hheight, True)
                                '    notYetDone = True

                                '    bigType = ptReportBigTypes.charts
                                '    compID = PTprdk.StrategieRisiko

                                'Catch ex As Exception
                                '    Dim a As Integer = -1
                                'End Try

                            Case "Strategie/Risiko/Ausstrahlung"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(214)
                                'End If


                                'Dim mycollection As New Collection

                                ''deleteStack.Add(.Name, .Name)
                                'Try
                                '    mycollection.Add(pname)

                                '    Call awinCreatePortfolioDiagrams(mycollection, reportObj, True, PTpfdk.FitRisikoDependency, PTpfdk.ProjektFarbe, True, False, True, htop, hleft, hwidth, hheight, True)
                                '    notYetDone = True
                                '    bigType = ptReportBigTypes.charts
                                '    compID = PTprdk.Dependencies
                                'Catch ex As Exception

                                'End Try

                            Case "ProjektBedarfsChart"

                                Try
                                    Dim smartChartInfo As clsSmartPPTChartInfo = getChartParametersFromQ1(qualifier)

                                    ' Text im ShapeContainer / Platzhalter zurücksetzen 
                                    .TextFrame2.TextRange.Text = ""

                                    'If smartChartInfo.chartTyp = PTChartTypen.Pie Then
                                    '    Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                    '    bigType = ptReportBigTypes.charts
                                    '    compID = PTprdk.PersonalPie
                                    '    boxName = obj.Chart.ChartTitle.Text
                                    '    reportObj = obj
                                    '    notYetDone = True

                                    'Else

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



                            Case "Personalbedarf"

                                Call MsgBox("not implemented ...")
                                '' old
                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(159)
                                'End If


                                'Try
                                '    auswahl = 1

                                '    If qualifier.Length > 0 Then
                                '        If qualifier.Trim <> "Balken" Then
                                '            Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '            compID = PTprdk.PersonalPie
                                '        Else

                                '            qualifier2 = bestimmeRoleQ2(qualifier2, selectedRoles)



                                '            'Call createRessBalkenOfProject(hproj, bproj, obj, auswahl, htop, hleft, hheight, hwidth, True,
                                '            '                               roleName:=qualifier2,
                                '            '                               vglTyp:=PTprdk.PersonalBalken)
                                '            compID = PTprdk.PersonalBalken
                                '        End If
                                '    Else
                                '        Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '        compID = PTprdk.PersonalPie
                                '    End If

                                '    boxName = obj.Chart.ChartTitle.Text
                                '    ' immer den Text nehmen ..
                                '    ''If obj.Chart.HasTitle Then
                                '    ''    boxName = obj.Chart.ChartTitle.Text
                                '    ''Else
                                '    ''    Dim gesamtSumme As Integer = CInt(hproj.getSummeRessourcen)
                                '    ''    boxName = boxName & " (" & gesamtSumme.ToString &
                                '    ''    " " & awinSettings.kapaEinheit & ")"
                                '    ''End If


                                '    reportObj = obj
                                '    notYetDone = False
                                '    bigType = ptReportBigTypes.charts


                                'Catch ex As Exception
                                '    '.TextFrame2.TextRange.Text = "Personal-Bedarf ist Null"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(233)
                                'End Try

                            Case "Personalbedarf2"

                                Call MsgBox("not implemented ...")
                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(159)
                                'End If


                                'Try
                                '    auswahl = 1

                                '    If qualifier.Length > 0 Then
                                '        If qualifier.Trim <> "Balken" Then
                                '            Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '            compID = PTprdk.PersonalPie
                                '        Else

                                '            qualifier2 = bestimmeRoleQ2(qualifier2, selectedRoles)


                                '            'Call createRessBalkenOfProject(hproj, lproj, obj, auswahl, htop, hleft, hheight, hwidth, True,
                                '            '                               roleName:=qualifier2,
                                '            '                               vglTyp:=PTprdk.PersonalBalken2)
                                '            compID = PTprdk.PersonalBalken2
                                '        End If
                                '    Else
                                '        Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '        compID = PTprdk.PersonalPie
                                '    End If

                                '    boxName = obj.Chart.ChartTitle.Text
                                '    ' immer den Text nehmen ..
                                '    ''If obj.Chart.HasTitle Then
                                '    ''    boxName = obj.Chart.ChartTitle.Text
                                '    ''Else
                                '    ''    Dim gesamtSumme As Integer = CInt(hproj.getSummeRessourcen)
                                '    ''    boxName = boxName & " (" & gesamtSumme.ToString &
                                '    ''    " " & awinSettings.kapaEinheit & ")"
                                '    ''End If


                                '    reportObj = obj
                                '    notYetDone = False
                                '    bigType = ptReportBigTypes.charts


                                'Catch ex As Exception
                                '    '.TextFrame2.TextRange.Text = "Personal-Bedarf ist Null"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(233)
                                'End Try

                            Case "Personalkosten"
                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(164)
                                'End If


                                'Try
                                '    auswahl = 2

                                '    If qualifier.Length > 0 Then

                                '        If qualifier.Trim <> "Balken" Then
                                '            Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '            compID = PTprdk.PersonalPie
                                '        Else

                                '            qualifier2 = bestimmeRoleQ2(qualifier2, selectedRoles)


                                '            'Call createRessBalkenOfProject(hproj, bproj, obj, auswahl, htop, hleft, hheight, hwidth, True,
                                '            '                               roleName:=qualifier2,
                                '            '                               vglTyp:=PTprdk.PersonalBalken)
                                '            compID = PTprdk.PersonalBalken
                                '        End If

                                '    Else
                                '        Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '        compID = PTprdk.PersonalPie
                                '    End If

                                '    boxName = obj.Chart.ChartTitle.Text
                                '    ' tk 9.8.18
                                '    'If obj.Chart.HasTitle Then
                                '    '    boxName = obj.Chart.ChartTitle.Text
                                '    'Else
                                '    '    Dim gesamtSumme As Integer = CInt(hproj.getAllPersonalKosten.Sum)
                                '    '    boxName = boxName & " (" & gesamtSumme.ToString & " T€)"
                                '    'End If


                                '    reportObj = obj
                                '    notYetDone = False
                                '    bigType = ptReportBigTypes.charts


                                'Catch ex As Exception
                                '    '.TextFrame2.TextRange.Text = "Personal-Kosten sind Null"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(162)
                                'End Try

                            Case "Personalkosten2"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(164)
                                'End If


                                'Try
                                '    auswahl = 2

                                '    If qualifier.Length > 0 Then

                                '        If qualifier.Trim <> "Balken" Then
                                '            Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '            compID = PTprdk.PersonalPie
                                '        Else

                                '            qualifier2 = bestimmeRoleQ2(qualifier2, selectedRoles)


                                '            'Call createRessBalkenOfProject(hproj, lproj, obj, auswahl, htop, hleft, hheight, hwidth, True,
                                '            '                               roleName:=qualifier2,
                                '            '                               vglTyp:=PTprdk.PersonalBalken2)
                                '            compID = PTprdk.PersonalBalken2
                                '        End If

                                '    Else
                                '        Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '        compID = PTprdk.PersonalPie
                                '    End If

                                '    boxName = obj.Chart.ChartTitle.Text
                                '    ' tk 9.8.18
                                '    'If obj.Chart.HasTitle Then
                                '    '    boxName = obj.Chart.ChartTitle.Text
                                '    'Else
                                '    '    Dim gesamtSumme As Integer = CInt(hproj.getAllPersonalKosten.Sum)
                                '    '    boxName = boxName & " (" & gesamtSumme.ToString & " T€)"
                                '    'End If


                                '    reportObj = obj
                                '    notYetDone = False
                                '    bigType = ptReportBigTypes.charts


                                'Catch ex As Exception
                                '    '.TextFrame2.TextRange.Text = "Personal-Kosten sind Null"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(162)
                                'End Try

                            Case "Sonstige Kosten"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(165)
                                'End If

                                'Try
                                '    auswahl = 1

                                '    If qualifier.Length > 0 Then

                                '        If qualifier.Trim <> "Balken" Then
                                '            Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '            compID = PTprdk.KostenPie
                                '        Else
                                '            compID = PTprdk.KostenBalken


                                '            'Call createCostBalkenOfProject(hproj, bproj, obj, auswahl, htop, hleft, hheight, hwidth, True, compID)

                                '        End If

                                '    Else
                                '        Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '        compID = PTprdk.KostenPie
                                '    End If

                                '    If obj.Chart.HasTitle Then
                                '        boxName = obj.Chart.ChartTitle.Text
                                '    Else
                                '        Dim gesamtSumme As Integer = CInt(hproj.getGesamtAndereKosten.Sum)
                                '        boxName = boxName & " (" & gesamtSumme.ToString & " T€)"
                                '    End If

                                '    reportObj = obj
                                '    notYetDone = False

                                '    bigType = ptReportBigTypes.charts

                                'Catch ex As Exception

                                '    '.TextFrame2.TextRange.Text = "Sonstige Kosten sind Null"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(163)

                                'End Try

                            Case "Sonstige Kosten2"

                                Call MsgBox("not implemented ...")


                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(165)
                                'End If

                                'Try
                                '    auswahl = 1

                                '    If qualifier.Length > 0 Then

                                '        If qualifier.Trim <> "Balken" Then
                                '            Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '            compID = PTprdk.KostenPie
                                '        Else
                                '            compID = PTprdk.KostenBalken2


                                '            'Call createCostBalkenOfProject(hproj, lproj, obj, auswahl, htop, hleft, hheight, hwidth, True, compID)

                                '        End If

                                '    Else
                                '        Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '        compID = PTprdk.KostenPie
                                '    End If

                                '    If obj.Chart.HasTitle Then
                                '        boxName = obj.Chart.ChartTitle.Text
                                '    Else
                                '        Dim gesamtSumme As Integer = CInt(hproj.getGesamtAndereKosten.Sum)
                                '        boxName = boxName & " (" & gesamtSumme.ToString & " T€)"
                                '    End If

                                '    reportObj = obj
                                '    notYetDone = False

                                '    bigType = ptReportBigTypes.charts

                                'Catch ex As Exception

                                '    '.TextFrame2.TextRange.Text = "Sonstige Kosten sind Null"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(163)

                                'End Try

                            Case "Gesamtkosten"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(166)
                                'End If


                                ''htop = 100
                                ''hleft = 100
                                ''hwidth = boxWidth * 14
                                ''hheight = boxHeight * 10

                                'Dim formerEE As Boolean = appInstance.ScreenUpdating

                                'Try
                                '    auswahl = 2

                                '    If qualifier.Length > 0 Then

                                '        If qualifier.Trim <> "Balken" Then
                                '            Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '            compID = PTprdk.KostenPie
                                '            reportObj = obj
                                '            notYetDone = True
                                '            bigType = ptReportBigTypes.charts

                                '        Else

                                '            appInstance.ScreenUpdating = False

                                '            compID = PTprdk.KostenBalken


                                '            'Call createCostBalkenOfProjectInPPT2(hproj, bproj, pptAppfromX, pptCurrentPresentation.Name, pptSlide.Name, auswahl, pptShape, compID, qualifier, qualifier2)

                                '            appInstance.ScreenUpdating = formerEE
                                '            notYetDone = False
                                '            bigType = ptReportBigTypes.charts

                                '        End If

                                '    Else
                                '        Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '        compID = PTprdk.KostenPie
                                '        reportObj = obj
                                '        notYetDone = False
                                '        bigType = ptReportBigTypes.charts
                                '    End If

                                '    ' das Platzhalter-Objekt : den Text auf leer setzen 
                                '    Try
                                '        .TextFrame2.TextRange.Text = ""
                                '        .AlternativeText = ""
                                '    Catch ex As Exception

                                '    End Try


                                'Catch ex As Exception
                                '    '.TextFrame2.TextRange.Text = "Gesamtkosten sind Null"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(168)
                                '    If appInstance.ScreenUpdating = False Then
                                '        appInstance.ScreenUpdating = formerEE
                                '    End If
                                'End Try

                            Case "Gesamtkosten2"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(166)
                                'End If


                                ''htop = 100
                                ''hleft = 100
                                ''hwidth = boxWidth * 14
                                ''hheight = boxHeight * 10

                                'Try
                                '    auswahl = 2

                                '    If qualifier.Length > 0 Then

                                '        If qualifier.Trim <> "Balken" Then
                                '            Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '            bigType = ptReportBigTypes.charts
                                '            compID = PTprdk.KostenPie
                                '            notYetDone = True
                                '        Else
                                '            bigType = ptReportBigTypes.charts
                                '            compID = PTprdk.KostenBalken2


                                '            notYetDone = False
                                '            'Call createCostBalkenOfProject(hproj, lproj, obj, auswahl, htop, hleft, hheight, hwidth, True, compID)
                                '        End If

                                '    Else
                                '        bigType = ptReportBigTypes.charts
                                '        compID = PTprdk.KostenPie
                                '        Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth, True)
                                '        notYetDone = True
                                '    End If

                                '    If obj.Chart.HasTitle Then
                                '        boxName = obj.Chart.ChartTitle.Text
                                '    Else
                                '        Dim gesamtSumme As Integer = CInt(hproj.getGesamtKostenBedarf.Sum)
                                '        boxName = boxName & " (" & gesamtSumme.ToString & " T€)"
                                '    End If

                                '    reportObj = obj

                                '    bigType = ptReportBigTypes.charts

                                'Catch ex As Exception
                                '    '.TextFrame2.TextRange.Text = "Gesamtkosten sind Null"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(168)
                                'End Try


                            Case "Trend Strategischer Fit/Risiko"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(232)
                                'End If

                                'Dim nrSnapshots As Integer = projekthistorie.Count

                                'If nrSnapshots > 0 Then

                                '    Call createTrendSfit(obj, htop, hleft, hheight, hwidth)

                                '    reportObj = obj
                                '    notYetDone = True

                                '    bigType = -1
                                '    compID = -1

                                'Else
                                '    '.TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(171)
                                'End If



                            Case "Trend Kennzahlen"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(231)
                                'End If

                                'Dim nrSnapshots As Integer = projekthistorie.Count

                                'If nrSnapshots > 0 Then
                                '    'htop = 100
                                '    'hleft = 100
                                '    'hwidth = 300
                                '    'hheight = 400

                                '    Call createTrendKPI(obj, htop, hleft, hheight, hwidth)

                                '    reportObj = obj
                                '    notYetDone = True

                                '    bigType = -1
                                '    compID = -1

                                'Else
                                '    '.TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(171)
                                'End If

                            Case "Fortschritt Personalkosten"


                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(205)
                                'End If

                                'Dim nrSnapshots As Integer = projekthistorie.Count
                                'Dim PListe As New Collection
                                'compareToID = 1
                                'auswahl = 1

                                'If nrSnapshots > 0 Then

                                '    If istLaufendesProjekt(hproj) Then

                                '        PListe.Add(hproj.name, hproj.name)
                                '        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                '        If Not obj Is Nothing Then
                                '            reportObj = obj
                                '            notYetDone = True

                                '            With reportObj
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                '            End With
                                '        Else
                                '            '.TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                '            .TextFrame2.TextRange.Text = boxName & repMessages.getmsg(139)
                                '        End If


                                '    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                '        '.TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(179)
                                '    Else
                                '        '.TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(180)
                                '    End If

                                'Else
                                '    '.TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(171)
                                'End If

                            Case "Fortschritt Sonstige Kosten"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(206)
                                'End If
                                'Dim nrSnapshots As Integer = projekthistorie.Count
                                'Dim PListe As New Collection
                                'compareToID = 1
                                'auswahl = 2

                                'If nrSnapshots > 0 Then

                                '    If istLaufendesProjekt(hproj) Then

                                '        PListe.Add(hproj.name, hproj.name)
                                '        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                '        If Not obj Is Nothing Then
                                '            reportObj = obj
                                '            notYetDone = True

                                '            With reportObj
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                '            End With
                                '        Else
                                '            '.TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                '            .TextFrame2.TextRange.Text = boxName & repMessages.getmsg(139)
                                '        End If

                                '    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                '        '.TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(179)
                                '    Else
                                '        '.TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(180)
                                '    End If

                                'Else
                                '    '.TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(171)
                                'End If

                            Case "Fortschritt Gesamtkosten"

                                Call MsgBox("not implemented ...")

                                'If boxName = kennzeichnung Then
                                '    boxName = repMessages.getmsg(207)
                                'End If
                                'Dim nrSnapshots As Integer = projekthistorie.Count
                                'Dim PListe As New Collection
                                'compareToID = 1
                                'auswahl = 3

                                'If nrSnapshots > 0 Then

                                '    If istLaufendesProjekt(hproj) Then

                                '        PListe.Add(hproj.name, hproj.name)
                                '        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                '        If Not obj Is Nothing Then
                                '            reportObj = obj
                                '            notYetDone = True

                                '            With reportObj
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                '            End With
                                '        Else
                                '            '.TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                '            .TextFrame2.TextRange.Text = boxName & repMessages.getmsg(139)
                                '        End If

                                '    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                '        '.TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(179)
                                '    Else
                                '        '.TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(180)
                                '    End If

                                'Else
                                '    '.TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(171)
                                'End If

                            Case "Fortschritt Rolle"

                                Call MsgBox("not implemented ...")

                                'Dim nrSnapshots As Integer = projekthistorie.Count
                                'Dim PListe As New Collection
                                'compareToID = 1
                                'auswahl = 4

                                'If nrSnapshots > 0 Then

                                '    If istLaufendesProjekt(hproj) Then

                                '        PListe.Add(hproj.name, hproj.name)
                                '        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                '        'boxName = "Fortschritt " & qualifier
                                '        boxName = repMessages.getmsg(176) & qualifier

                                '        If Not obj Is Nothing Then
                                '            reportObj = obj
                                '            notYetDone = True

                                '            With reportObj
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                '            End With
                                '        Else
                                '            '.TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                '            .TextFrame2.TextRange.Text = boxName & repMessages.getmsg(139)
                                '        End If

                                '    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                '        '.TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(179)
                                '    Else
                                '        '.TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(180)
                                '    End If

                                'Else
                                '    '.TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(171)
                                'End If

                            Case "Fortschritt Kostenart"

                                Call MsgBox("not implemented ...")

                                'Dim nrSnapshots As Integer = projekthistorie.Count
                                'Dim PListe As New Collection
                                'compareToID = 1
                                'auswahl = 5

                                'If nrSnapshots > 0 Then

                                '    If istLaufendesProjekt(hproj) Then

                                '        PListe.Add(hproj.name, hproj.name)
                                '        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                '        'boxName = "Fortschritt " & qualifier
                                '        boxName = repMessages.getmsg(176) & qualifier

                                '        If Not obj Is Nothing Then
                                '            reportObj = obj
                                '            notYetDone = True

                                '            With reportObj
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                '                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                '            End With
                                '        Else
                                '            '.TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                '            .TextFrame2.TextRange.Text = boxName & repMessages.getmsg(139)
                                '        End If

                                '    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                '        '.TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(179)
                                '    Else
                                '        '.TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                '        .TextFrame2.TextRange.Text = repMessages.getmsg(180)
                                '    End If

                                'Else
                                '    '.TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                '    .TextFrame2.TextRange.Text = repMessages.getmsg(171)
                                'End If



                            Case "Ampel-Farbe"

                                If boxName = kennzeichnung Then
                                    If useEnglishLanguage Then
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
                                    If useEnglishLanguage Then
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

                            Case "Business-Unit:"

                                If boxName = kennzeichnung Then
                                    If useEnglishLanguage Then
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
                                    If useEnglishLanguage Then
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
                                    If useEnglishLanguage Then
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
                                    If useEnglishLanguage Then
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
                                    If useEnglishLanguage Then
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
        Dim newPPTChart As PowerPoint.Shape = curSlide.Shapes.AddChart(Left:=left, Top:=top, Width:=width, Height:=height)
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
        Dim errorShape As PowerPoint.ShapeRange = Nothing

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
        rds.pptSlide = curSlide


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

                            Call zeichne3RowsCalendar(rds, minCal)

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
            errorShape = pptCopypptPaste(rds.errorVorlagenShape, curSlide)

            With errorShape.Item(1)
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
    Private Sub zeichneMultiprojektSichtinPPT(ByRef objectsToDo As Integer, ByRef objectsDone As Integer, ByRef pptFirstTime As Boolean,
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
        Dim errorShape As PowerPoint.ShapeRange = Nothing



        Dim format As Integer = 4
        'Dim tmpslideID As Integer



        Dim rds As New clsPPTShapes
        rds.pptSlide = curSlide

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

            ' tk 14.10 das wird doch immer benötigt ... 
            'If awinSettings.mppExtendedMode Then

            '    ' über alle ausgewählte Projekte sehen und maximale Anzahl Zeilen je Projekt bestimmen
            '    For Each kvp As KeyValuePair(Of Double, String) In projCollection
            '        Try

            '            hproj = AlleProjekte.getProject(kvp.Value)
            '        Catch ex As Exception

            '        End Try

            '        anzZeilen = hproj.calcNeededLines(selectedPhases, selectedMilestones, awinSettings.mppExtendedMode, Not awinSettings.mppShowAllIfOne)

            '        maxZeilen = System.Math.Max(maxZeilen, anzZeilen)
            '        gesamtAnzZeilen = gesamtAnzZeilen + anzZeilen

            '    Next


            'Else
            '    projekthoehe = zeilenhoehe_sav
            'End If

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

                        ' das demnächst abändern auf 
                        Call zeichne3RowsCalendar(rds, minCal)

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



                    Call zeichnePPTprojectsInPPT(projCollection, objectsDone,
                                rds, selectedPhases, selectedMilestones, selectedRoles, selectedCosts, kennzeichnung)


                Catch ex As Exception

                    If Not IsNothing(rds.errorVorlagenShape) Then
                        ''rds.errorVorlagenShape.Copy()
                        ''errorShape = pptslide.Shapes.Paste
                        errorShape = pptCopypptPaste(rds.errorVorlagenShape, rds.pptSlide)

                        With errorShape.Item(1)
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
            ''rds.errorVorlagenShape.Copy()
            ''errorShape = pptslide.Shapes.Paste
            errorShape = pptCopypptPaste(rds.errorVorlagenShape, curSlide)

            With errorShape.Item(1)
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
    Sub zeichnePPTprojectsInPPT(ByRef projectCollection As SortedList(Of Double, String),
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
        Dim copiedShape As PowerPoint.ShapeRange = Nothing
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

                Dim msToDraw As New Collection      ' hier sind alle selektierten Meilensteine mit zugehörigen Phasen enthalten

                hproj = AlleProjekte.getProject(fullName)


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

                    copiedShape = pptCopypptPaste(rds.projectNameVorlagenShape, curSlide)

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
                        With copiedShape(1)

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

                                Call addSmartPPTMsPhInfo(copiedShape(1), hproj,
                                                        Nothing, hproj.getShapeText, Nothing, Nothing,
                                                        Nothing, Nothing,
                                                        hproj.startDate, hproj.endeDate,
                                                        hproj.ampelStatus, hproj.ampelErlaeuterung, Nothing,
                                                        hproj.leadPerson, hproj.getPhase(1).percentDone, hproj.getPhase(1).DocURL)

                            End If

                        End With
                    Else

                        With copiedShape(1)
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

                                Call addSmartPPTMsPhInfo(copiedShape(1), hproj,
                                                        Nothing, hproj.getShapeText, Nothing, Nothing,
                                                        Nothing, Nothing,
                                                        hproj.startDate, hproj.endeDate,
                                                        hproj.ampelStatus, hproj.ampelErlaeuterung, Nothing,
                                                        hproj.leadPerson, hproj.getPhase(1).percentDone, hproj.getPhase(1).DocURL)

                            End If

                        End With
                    End If

                    Dim projectNameShape As PowerPoint.Shape = copiedShape(1)


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

                        copiedShape = pptCopypptPaste(rds.ampelVorlagenShape, curSlide)

                        With copiedShape(1)
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

                        ''projectVorlagenForm.Copy()
                        ''copiedShape = pptslide.Shapes.Paste()
                        copiedShape = pptCopypptPaste(rds.projectVorlagenShape, curSlide)

                        With copiedShape(1)
                            .Top = CSng(projektGrafikYPos)
                            .Left = CSng(x1)
                            .Width = CSng(x2 - x1)
                            .Name = .Name & .Id

                            '.Title = hproj.getShapeText
                            '.AlternativeText = hproj.startDate.ToShortDateString & " - " & hproj.endeDate.ToShortDateString

                            If awinSettings.mppEnableSmartPPT Then

                                Call addSmartPPTMsPhInfo(copiedShape(1), hproj,
                                                   Nothing, hproj.getShapeText, Nothing, Nothing,
                                                   Nothing, Nothing,
                                                   hproj.startDate, hproj.endeDate,
                                                   hproj.ampelStatus, hproj.ampelErlaeuterung, Nothing,
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
                                (type = PTItemType.projekt And pvName = hproj.name) Or
                                (type = PTItemType.vorlage And pvName = hproj.VorlagenName) Then

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

                                                        Call zeichneMeilensteininAktZeile(curSlide, msShapeNames, minX1, maxX2,
                                                                                      milestone, hproj, milestoneGrafikYPos, rds)


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

                                    ''PhDescVorlagenShape.Copy()
                                    ''copiedShape = pptslide.Shapes.Paste()
                                    copiedShape = pptCopypptPaste(rds.PhDescVorlagenShape, curSlide)

                                    With copiedShape(1)

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

                                    ''PhDateVorlagenShape.Copy()
                                    ''copiedShape = pptslide.Shapes.Paste()
                                    copiedShape = pptCopypptPaste(rds.PhDateVorlagenShape, curSlide)

                                    With copiedShape(1)

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
                                    ''phasedelimiterShape.Copy()
                                    ''copiedShape = pptslide.Shapes.Paste()
                                    copiedShape = pptCopypptPaste(rds.phaseDelimiterShape, curSlide)

                                    With copiedShape(1)

                                        .Height = CSng(1.3 * appear.height)
                                        .Top = CSng(phasenGrafikYPos)
                                        .Left = CSng(x1 - .Width * 0.5)
                                        .Name = .Name & .Id

                                    End With

                                    ' rechter Delimiter 
                                    ''phasedelimiterShape.Copy()
                                    ''copiedShape = pptslide.Shapes.Paste()
                                    copiedShape = pptCopypptPaste(rds.phaseDelimiterShape, curSlide)

                                    With copiedShape(1)

                                        .Height = CSng(1.3 * appear.height)
                                        .Top = CSng(phasenGrafikYPos)
                                        .Left = CSng(x2 + .Width * 0.5)
                                        .Name = .Name & .Id

                                    End With

                                End If


                                ''copiedShape = xlnsCopypptPaste(phaseShape, pptslide)

                                ''With copiedShape(1)
                                ''    .Top = CSng(phasenGrafikYPos)
                                ''    .Left = CSng(x1)
                                ''    .Width = CSng(x2 - x1)
                                ''    .Height = rds.phaseVorlagenShape.Height
                                ''    '.Name = .Name & .Id
                                ''    Try
                                ''        .Name = phShapeName
                                ''    Catch ex As Exception

                                ''    End Try

                                ''    '.Title = phaseName
                                ''    '.AlternativeText = phDateText

                                ''    If missingPhaseDefinition Then
                                ''        .Fill.ForeColor.RGB = cphase.farbe
                                ''    End If

                                ''End With
                                phaseShape = curSlide.Shapes.AddShape(appear.shpType,
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

                                    '.Title = phaseName
                                    '.AlternativeText = phDateText

                                    If missingPhaseDefinition Then
                                        .Fill.ForeColor.RGB = cphase.farbe
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
                                    (type = PTItemType.projekt And pvName = hproj.name) Or
                                    (type = PTItemType.vorlage And pvName = hproj.VorlagenName) Then

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

                                                    Call zeichneMeilensteininAktZeile(curSlide, msShapeNames, minX1, maxX2,
                                                                                      ms, hproj, milestoneGrafikYPos, rds)




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
                                    Call zeichneMeilensteininAktZeile(curSlide, msShapeNames, minX1, maxX2,
                                                                      milestone, hproj, milestoneGrafikYPos, rds)
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
                                Call zeichneMeilensteininAktZeile(curSlide, msShapeNames, minX1, maxX2,
                                                                  milestone, hproj, milestoneGrafikYPos, rds)
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
                                Call zeichneMeilensteininAktZeile(curSlide, msShapeNames, minX1, maxX2,
                                                                  milestone, hproj, milestoneGrafikYPos, rds)
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


                    ''buColorShape.Copy()
                    ''copiedShape = pptslide.Shapes.Paste()
                    copiedShape = pptCopypptPaste(rds.buColorShape, curSlide)

                    With copiedShape(1)
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
                    ''rowDifferentiatorShape.Copy()
                    ''copiedShape = pptslide.Shapes.Paste()
                    copiedShape = pptCopypptPaste(rds.rowDifferentiatorShape, curSlide)

                    With copiedShape(1)
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
                    ''durationArrowShape.Copy()
                    ''copiedShape = pptslide.Shapes.Paste()
                    copiedShape = pptCopypptPaste(rds.durationArrowShape, curSlide)

                    Dim pfeilbreite As Double = maxX2 - minX1

                    With copiedShape(1)
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

                    ''durationTextShape.Copy()
                    ''copiedShape = pptslide.Shapes.Paste()
                    copiedShape = pptCopypptPaste(rds.durationTextShape, curSlide)

                    With copiedShape(1)
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
                    CType(curSlide.Shapes.Range(arrayOfNames), PowerPoint.ShapeRange).ZOrder(MsoZOrderCmd.msoBringToFront)
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
                CType(curSlide.Shapes.Range(arrayOfNames), PowerPoint.ShapeRange).ZOrder(MsoZOrderCmd.msoBringToFront)
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



End Module
