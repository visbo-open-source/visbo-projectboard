Imports ProjectBoardDefinitions
Imports System.Windows.Forms
Imports ClassLibrary1
Imports ProjectBoardBasic
Module creationModule

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
            'For j = 1 To anzSlidesToAdd



            ' jetzt muss die Slide als SmartPPTSlide gekennzeichnet werden 
            'Call addSmartPPTSlideBaseInfo(pptSlide, hproj.timeStamp, ptPRPFType.project)
            Call addSmartPPTSlideBaseInfo(currentSlide, reportCreationDate, ptPRPFType.project)

            ' jetzt werden die Charts gezeichnet 
            anzShapes = currentSlide.Shapes.Count
            Dim newShapeRange As PowerPoint.ShapeRange = Nothing
            Dim newShapeRange2 As PowerPoint.ShapeRange = Nothing
            Dim newShape As PowerPoint.Shape = Nothing


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
                        kennzeichnung = "Soll-Ist & Prognose" Or
                        kennzeichnung = "Multivariantensicht" Or
                        kennzeichnung = "Einzelprojektsicht" Or
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
            Next


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


                            Case "Legenden-Tabelle"

                                Call MsgBox("not implemented ...")
                                'Try
                                '    ' Einzelprojektsicht im Extended Mode
                                '    If selectedPhases.Count = 0 _
                                '        And selectedMilestones.Count = 0 _
                                '        And selectedRoles.Count = 0 _
                                '        And selectedCosts.Count = 0 _
                                '        And selectedBUs.Count = 0 _
                                '        Then
                                '        Dim i As Integer = 0
                                '        Dim tmpphases As New Collection
                                '        Dim tmpMilestones As New Collection

                                '        ' alle Phasennamen des Projektes hproj in die Collection tmpphases bringen
                                '        For Each cphase In hproj.AllPhases

                                '            Dim tmpstr = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                                '            If tmpstr <> "" Then
                                '                tmpstr = tmpstr & "#" & cphase.name
                                '                If Not tmpphases.Contains(tmpstr) Then
                                '                    tmpphases.Add(tmpstr, tmpstr)
                                '                End If
                                '            End If


                                '        Next

                                '        ' alle Meilensteine-Namen des Projektes hproj in die collection tmpMilestones bringen
                                '        Dim mSList As SortedList(Of Date, String)

                                '        mSList = hproj.getMilestones        ' holt alle Meilensteine in Form ihrer nameID sortiert nach Datum

                                '        If mSList.Count > 0 Then
                                '            For Each kvp As KeyValuePair(Of Date, String) In mSList

                                '                Dim tmpstr = hproj.hierarchy.getBreadCrumb(kvp.Value) & "#" & hproj.getMilestoneByID(kvp.Value).name
                                '                If Not tmpMilestones.Contains(tmpstr) Then
                                '                    tmpMilestones.Add(tmpstr, tmpstr)
                                '                End If


                                '            Next
                                '        End If

                                '        Call prepZeichneLegendenTabelle(pptSlide, pptShape, legendFontSize, tmpphases, tmpMilestones)
                                '    Else

                                '        Call prepZeichneLegendenTabelle(pptSlide, pptShape, legendFontSize, selectedPhases, selectedMilestones)
                                '    End If

                                'Catch ex As Exception

                                'End Try

                            Case "AllePlanElemente"

                                Call MsgBox("not implemented ...")

                                'Try

                                '    Dim i As Integer = 0
                                '    Dim tmpphases As New Collection
                                '    Dim tmpMilestones As New Collection
                                '    Dim minCal As Boolean = False
                                '    If qualifier2.Length > 0 Then
                                '        minCal = (qualifier2.Trim = "minCal")
                                '    End If

                                '    ' alle Phasennamen des Projektes hproj in die Collection tmpphases bringen
                                '    For Each cphase In hproj.AllPhases

                                '        Dim tmpstr As String = hproj.hierarchy.getBreadCrumb(cphase.nameID)
                                '        If tmpstr <> "" Then
                                '            tmpstr = tmpstr & "#" & cphase.name
                                '            If Not tmpphases.Contains(tmpstr) Then
                                '                tmpphases.Add(tmpstr, tmpstr)
                                '            End If

                                '        End If


                                '    Next



                                '    ' alle Meilensteine-Namen des Projektes hproj in die collection tmpMilestones bringen
                                '    Dim mSList As SortedList(Of Date, String)

                                '    mSList = hproj.getMilestones        ' holt alle Meilensteine in Form ihrer nameID sortiert nach Datum

                                '    If mSList.Count > 0 Then
                                '        For Each kvp As KeyValuePair(Of Date, String) In mSList

                                '            Dim tmpstr = hproj.hierarchy.getBreadCrumb(kvp.Value) & "#" & hproj.getMilestoneByID(kvp.Value).name
                                '            If Not tmpMilestones.Contains(tmpstr) Then
                                '                tmpMilestones.Add(tmpstr, tmpstr)
                                '            End If

                                '        Next
                                '    End If


                                '    ' die Slide mit Tag kennzeichnen ... 

                                '    Call zeichneMultiprojektSicht(pptAppfromX, pptCurrentPresentation, pptSlide,
                                '                                  objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, legendFontSize,
                                '                                  tmpphases, tmpMilestones,
                                '                                  translateToRoleNames(selectedRoles), selectedCosts,
                                '                                  selectedBUs, selectedTyps,
                                '                                  worker, e, False, False, hproj, kennzeichnung, minCal)
                                '    .TextFrame2.TextRange.Text = ""
                                '    '.ZOrder(MsoZOrderCmd.msoSendToBack)
                                'Catch ex As Exception
                                '    .TextFrame2.TextRange.Text = ex.Message
                                '    objectsDone = objectsToDo

                                'End Try

                            Case "Einzelprojektsicht"

                                Call MsgBox("not implemented ...")

                                'Try
                                '    Dim minCal As Boolean = False
                                '    If qualifier2.Length > 0 Then
                                '        minCal = (qualifier2.Trim = "minCal")
                                '    End If

                                '    Call zeichneMultiprojektSicht(pptAppfromX, pptCurrentPresentation, pptSlide,
                                '                                      objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, legendFontSize,
                                '                                      selectedPhases, selectedMilestones,
                                '                                      translateToRoleNames(selectedRoles), selectedCosts,
                                '                                      selectedBUs, selectedTyps,
                                '                                      worker, e, False, False, hproj, kennzeichnung, minCal)
                                '    .TextFrame2.TextRange.Text = ""
                                '    '.ZOrder(MsoZOrderCmd.msoSendToBack)
                                'Catch ex As Exception
                                '    .TextFrame2.TextRange.Text = ex.Message
                                '    objectsDone = objectsToDo
                                'End Try

                            Case "Multivariantensicht"

                                Call MsgBox("not implemented ...")

                                'Try

                                '    Dim minCal As Boolean = False
                                '    If qualifier2.Length > 0 Then
                                '        minCal = (qualifier2.Trim = "minCal")
                                '    End If

                                '    Call zeichneMultiprojektSicht(pptAppfromX, pptCurrentPresentation, pptSlide,
                                '                                      objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, legendFontSize,
                                '                                      selectedPhases, selectedMilestones,
                                '                                      translateToRoleNames(selectedRoles), selectedCosts,
                                '                                      selectedBUs, selectedTyps,
                                '                                      worker, e, False, True, hproj, kennzeichnung, minCal)
                                '    .TextFrame2.TextRange.Text = ""
                                '    '.ZOrder(MsoZOrderCmd.msoSendToBack)
                                'Catch ex As Exception
                                '    .TextFrame2.TextRange.Text = ex.Message
                                '    objectsDone = objectsToDo
                                'End Try


                            Case "MilestoneCategories"

                                Call MsgBox("not implemented ...")

                                'Try

                                '    Dim minCal As Boolean = False
                                '    If qualifier2.Length > 0 Then
                                '        minCal = (qualifier2.Trim = "minCal")
                                '    End If

                                '    Call zeichneCategorySwimlaneSicht(pptAppfromX, pptCurrentPresentation, pptSlide,
                                '                                      objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, legendFontSize,
                                '                                      selectedPhases, selectedMilestones,
                                '                                      translateToRoleNames(selectedRoles), selectedCosts,
                                '                                      selectedBUs, selectedTyps,
                                '                                      worker, e, False, hproj, kennzeichnung, minCal)

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

                                Call MsgBox("not implemented ...")

                                'Try

                                '    Dim minCal As Boolean = False
                                '    If qualifier2.Length > 0 Then
                                '        minCal = (qualifier2.Trim = "minCal")
                                '    End If

                                '    Call zeichneSwimlane2Sicht(pptAppfromX, pptCurrentPresentation, pptSlide,
                                '                                      objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, legendFontSize,
                                '                                      selectedPhases, selectedMilestones,
                                '                                      translateToRoleNames(selectedRoles), selectedCosts,
                                '                                      selectedBUs, selectedTyps,
                                '                                      worker, e, False, hproj, kennzeichnung, minCal)

                                '    .TextFrame2.TextRange.Text = ""
                                '    '.ZOrder(MsoZOrderCmd.msoSendToBack)

                                '    ' sonst wird pptLasttime benötigt, um bei mehreren PRojekten 
                                '    ' swimlaneMode wird erst nach Ende der While Schleife ausgewertet - in diesem Fall wird die tmpSav Folie gelöscht 
                                '    'swimlaneMode = True
                                'Catch ex As Exception
                                '    .TextFrame2.TextRange.Text = ex.Message & ": iDkey = " & iDkey
                                '    objectsDone = objectsToDo
                                'End Try


                            Case "Swimlanes2"
                                Dim formerSetting As Boolean = awinSettings.mppExtendedMode
                                Call MsgBox("not implemented ...")

                                'Try

                                '    Dim minCal As Boolean = False
                                '    If qualifier2.Length > 0 Then
                                '        minCal = (qualifier2.Trim = "minCal")
                                '    End If


                                '    Call zeichneSwimlane2Sicht(pptAppfromX, pptCurrentPresentation, pptSlide,
                                '                                      objectsToDo, objectsDone, pptFirstTime, zeilenhoehe_sav, legendFontSize,
                                '                                      selectedPhases, selectedMilestones,
                                '                                      translateToRoleNames(selectedRoles), selectedCosts,
                                '                                      selectedBUs, selectedTyps,
                                '                                      worker, e, False, hproj, kennzeichnung, minCal)
                                '    awinSettings.mppExtendedMode = formerSetting
                                '    .TextFrame2.TextRange.Text = ""
                                '    '.ZOrder(MsoZOrderCmd.msoSendToBack)

                                '    ' sonst wird pptLasttime benötigt, um bei mehreren Projekten 
                                '    ' swimlaneMode wird erst nach Ende der While Schleife ausgewertet - in diesem Fall wird die tmpSav Folie gelöscht 
                                '    'swimlaneMode = True
                                'Catch ex As Exception
                                '    awinSettings.mppExtendedMode = formerSetting
                                '    .TextFrame2.TextRange.Text = ex.Message & ": iDkey = " & iDkey
                                '    objectsDone = objectsToDo
                                'End Try


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
                                    Call zeichneProjektTabelleVergleich(currentSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, lastproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle Vergleich Beauftragung"

                                Try
                                    Call zeichneProjektTabelleVergleich(currentSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, bproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle OneGlance letzter Stand"

                                Try
                                    Call zeichneProjektTabelleOneGlance(currentSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, lastproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle OneGlance Beauftragung"


                                Try
                                    Call zeichneProjektTabelleOneGlance(currentSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, bproj)
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

                            Case "Business-Unit:"

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




End Module
