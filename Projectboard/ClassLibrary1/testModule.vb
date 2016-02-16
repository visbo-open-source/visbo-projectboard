Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports Microsoft.Office.Core
Imports pptNS = Microsoft.Office.Interop.PowerPoint
Imports xlNS = Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports Microsoft.Office.Interop

Public Module testModule


    ''' <summary>
    ''' erzeugt den Report aller selektieren Projekte auf Grundlage des Templates templatedossier.pptx
    ''' bei Aufruf ist sichergestellt, daß in Projekthistorie die Historie der selektierten Projekte steht 
    ''' </summary>
    ''' <param name="pptTemplate"></param>
    ''' <remarks></remarks>
    ''' 
    Public Sub createPPTReportFromProjects(ByVal pptTemplate As String, _
                                           ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                           ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                           ByVal selectedBUs As Collection, ByVal selectedTyps As Collection, _
                                           ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)

        Dim awinSelection As xlNS.ShapeRange

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim singleShp As xlNS.Shape
        Dim hproj As clsProjekt
        Dim vglName As String = " "
        Dim pName As String, variantName As String
        Dim vorlagenDateiName As String = pptTemplate
        Dim zeilenhoehe As Double = 0.0     ' zeilenhöhe muss für alle Projekte gleich sein, daher mit übergeben
        Dim legendFontSize As Single = 0.0  ' FontSize der Legenden der Schriftgröße des Projektnamens angepasst
        Dim tatsErstellt As Integer = 0

        Dim todoListe As New Collection

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, xlNS.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If Not IsNothing(awinSelection) Then

            ' hier wird bestimmt, welches der ausgewählten Projekte dasjenige ist, das im Extended Mode den meisten Platz beim Report benötigt.
            ' Der Report dieses Projektes soll dann zuerst erstellt werden, denn somit wird das Format der PowerPointPräsentation danach ausgewählt.

            Dim maxProj As clsProjekt = Nothing
            Dim maxZeilen As Integer = 1

            For Each singleShp In awinSelection

                With singleShp
                    If isProjectType(CInt(.AlternativeText)) Then
                        Try
                            hproj = ShowProjekte.getProject(singleShp.Name, True)
                            todoListe.Add(hproj.name)
                        Catch ex As Exception
                            Call MsgBox(singleShp.Name & " nicht gefunden ...")
                            Exit Sub
                        End Try
                        If hproj.calcNeededLines() > maxZeilen Then
                            maxProj = hproj
                            maxZeilen = hproj.calcNeededLines()

                        End If

                    End If

                End With
            Next

            ' Erstelle Report für das größte Projekt "maxProj"

            If Not projekthistorie Is Nothing Then
                If projekthistorie.Count > 0 Then
                    vglName = projekthistorie.First.getShapeText
                End If
            End If

            With maxProj
                pName = .name
                variantName = .variantName
            End With

            If vglName <> maxProj.getShapeText Then
                If request.pingMongoDb() Then
                    Try
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                        storedEarliest:=Date.MinValue, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, maxProj)
                    Catch ex As Exception
                        projekthistorie.clear()
                    End Try
                Else
                    Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
                End If


            Else
                ' der aktuelle Stand hproj muss hinzugefügt werden 
                Dim lastElem As Integer = projekthistorie.Count - 1
                projekthistorie.RemoveAt(lastElem)
                projekthistorie.Add(Date.Now, maxProj)
            End If

            e.Result = " Report für Projekt '" & maxProj.getShapeText & "' wird erstellt !"
            worker.ReportProgress(0, e)


            Call createPPTSlidesFromProject(maxProj, vorlagenDateiName, _
                                            selectedPhases, selectedMilestones, _
                                            selectedRoles, selectedCosts, _
                                            selectedBUs, selectedTyps, True, _
                                            (awinSelection.Count = tatsErstellt + 1), zeilenhoehe, _
                                            legendFontSize, _
                                            worker, e)
            tatsErstellt = tatsErstellt + 1


            For Each singleItem As String In todoListe

                Try
                    hproj = ShowProjekte.getProject(singleItem)
                Catch ex As Exception

                    Call MsgBox(singleItem & " nicht gefunden ...")
                    Exit Sub
                End Try

                If hproj.name <> maxProj.name Then

                    If Not projekthistorie Is Nothing Then
                        If projekthistorie.Count > 0 Then
                            vglName = projekthistorie.First.getShapeText
                        End If
                    End If

                    With hproj
                        pName = .name
                        variantName = .variantName
                    End With

                    If vglName <> hproj.getShapeText Then
                        If request.pingMongoDb() Then
                            Try
                                projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                                storedEarliest:=Date.MinValue, storedLatest:=Date.Now)
                                projekthistorie.Add(Date.Now, hproj)
                            Catch ex As Exception
                                projekthistorie.clear()
                            End Try
                        Else
                            Call MsgBox("Datenbank-Verbindung ist unterbrochen!")
                        End If


                    Else
                        ' der aktuelle Stand hproj muss hinzugefügt werden 
                        Dim lastElem As Integer = projekthistorie.Count - 1
                        projekthistorie.RemoveAt(lastElem)
                        projekthistorie.Add(Date.Now, hproj)
                    End If

                    e.Result = " Report für Projekt '" & hproj.getShapeText & "' wird erstellt !"
                    worker.ReportProgress(0, e)

                    If tatsErstellt = 0 Then

                        Call createPPTSlidesFromProject(hproj, vorlagenDateiName, _
                                                        selectedPhases, selectedMilestones, _
                                                        selectedRoles, selectedCosts, _
                                                        selectedBUs, selectedTyps, True, _
                                                        (todoListe.Count = tatsErstellt + 1), zeilenhoehe, _
                                                        legendFontSize, _
                                                        worker, e)

                    Else

                        Call createPPTSlidesFromProject(hproj, vorlagenDateiName, _
                                                        selectedPhases, selectedMilestones, _
                                                        selectedRoles, selectedCosts, _
                                                        selectedBUs, selectedTyps, False, _
                                                        (todoListe.Count = tatsErstellt + 1), zeilenhoehe, _
                                                        legendFontSize, _
                                                        worker, e)

                    End If

                    tatsErstellt = tatsErstellt + 1


                Else
                    ' maxProj wurde als erstes gezeichnet, damit das Format bei Multiprojektsicht das Richtige ist

                End If  ' if hproj = maxproj



            Next

        End If

        If tatsErstellt = 1 Then
            e.Result = " Report für " & tatsErstellt & " Projekt erstellt !"
        Else
            e.Result = " Report für " & tatsErstellt & " Projekte erstellt !"
        End If

        worker.ReportProgress(0, e)
        'frmSelectPPTTempl.statusNotification.Text = " Report mit " & tatsErstellt & " Seite erstellt !"


    End Sub


    ''' <summary>
    ''' erzeugt den Bericht Report auf Grundlage des Templates templatedossier.pptx
    ''' bei Aufruf ist sichergestellt, daß in Projekthistorie die Historie des Projektes steht 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub createPPTSlidesFromProject(ByRef hproj As clsProjekt, pptTemplateName As String, _
                                          ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                          ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                          ByVal selectedBUs As Collection, ByVal selectedTyps As Collection, _
                                          ByRef pptFirstTime As Boolean, ByVal pptLastTime As Boolean, ByRef zeilenhoehe As Double, _
                                          ByRef legendFontSize As Single, _
                                          ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)
        Dim pptApp As pptNS.Application = Nothing
        Dim pptCurrentPresentation As pptNS.Presentation = Nothing
        Dim pptTemplatePresentation As pptNS.Presentation = Nothing
        Dim pptSlide As pptNS.Slide = Nothing
        Dim shapeRange As pptNS.ShapeRange = Nothing
        Dim presentationFile As String = awinPath & requirementsOrdner & "projektdossier.pptx"
        Dim presentationFileH As String = awinPath & requirementsOrdner & "projektdossier_Hochformat.pptx"
        Dim newFileName As String = reportOrdnerName & "Report.pptx"
        Dim pptShape As pptNS.Shape
        Dim pname As String = hproj.name
        Dim fullName As String = hproj.getShapeText
        Dim top As Double, left As Double, width As Double, height As Double
        Dim htop As Double, hleft As Double, hwidth As Double, hheight As Double
        Dim pptSize As Single = 18
        ' ur: 16.04.2015: wird nun übergeben: Dim zeilenhoehe As Double = 0.0

        Dim auswahl As Integer
        Dim compareToID As Integer
        Dim qualifier As String = " ", qualifier2 As String = " "
        Dim notYetDone As Boolean = False
        Dim ze As String = " (" & awinSettings.kapaEinheit & ")"
        Dim ke As String = " (T€)"
        Dim heute As Date = Date.Now
        Dim istWerteexistieren As Boolean
        Dim boxName As String
        Dim listofShapes As New Collection

        Dim lproj As clsProjekt
        Dim bproj As clsProjekt
        Dim lastproj As clsProjekt
        Dim lastElem As Integer
        ' das sind Formen , die zur in der Tabelle Vergleich Anzeige der Tendenz verwendet werden 
        Dim gleichShape As pptNS.Shape = Nothing
        Dim steigendShape As pptNS.Shape = Nothing
        Dim fallendShape As pptNS.Shape = Nothing
        Dim ampelShape As pptNS.Shape = Nothing
        Dim sternShape As pptNS.Shape = Nothing


        ' Änderung tk 1.2.16
        ' wird benötigt, um in Ergänzung zu pptLasttime im Falle von nur einem Projekt / vielen Swimlanes die bereits erstellte Folie zu löschen 
        Dim swimlaneMode As Boolean = False

        Try
            lastElem = projekthistorie.Count - 1
            lastproj = projekthistorie.ElementAt(lastElem - 1)
        Catch ex As Exception
            lastElem = -1
            lastproj = Nothing
        End Try


        ' die Projekt Historie ist bereits gesetzt ... siehe Aufruf
        Try

            bproj = projekthistorie.beauftragung

        Catch ex As Exception
            ' es gibt keine Beauftragung
            bproj = Nothing
        End Try
        '
        '
        Try

            projekthistorie.currentIndex = projekthistorie.Count - 1
            lproj = projekthistorie.letzteFreigabe
            If lproj.timeStamp = bproj.timeStamp Then
                ' es gibt ausser der Beauftragung keinen weiteren Freigabestand
                lproj = Nothing
            End If

        Catch ex As Exception
            ' es gibt keinen letzten, freigegebenen Stand
            lproj = Nothing
        End Try
        '
        '

        If DateDiff(DateInterval.Month, hproj.startDate, heute) > 0 Then
            istWerteexistieren = True
        Else
            istWerteexistieren = False
        End If

        Try
            ' prüft, ob bereits Powerpoint geöffnet ist 
            pptApp = CType(GetObject(, "PowerPoint.Application"), pptNS.Application)
        Catch ex As Exception
            Try
                pptApp = CType(CreateObject("PowerPoint.Application"), pptNS.Application)
            Catch ex1 As Exception
                Call MsgBox("Powerpoint konnte nicht gestartet werden ..." & ex1.Message)
                Exit Sub
            End Try

        End Try


        ' entweder wird das template geöffnet ...
        ' oder aber es wird in die aktive Presentation geschrieben 

        ' jetzt wird das template geöffnet , um festzustellen , welches Format Quer oder Hoch die Vorlage hat 
        ' und dann wird die entsprechende Titelblatt Präsentation geöffnet 

        Try

            If pptApp.Presentations.Count = 0 Then

                pptTemplatePresentation = pptApp.Presentations.Open(pptTemplateName)

                If pptTemplatePresentation.PageSetup.SlideOrientation = MsoOrientation.msoOrientationHorizontal Then
                    pptCurrentPresentation = pptApp.Presentations.Open(presentationFile)
                Else
                    pptCurrentPresentation = pptApp.Presentations.Open(presentationFileH)
                End If

            Else
                pptCurrentPresentation = pptApp.ActivePresentation
                pptTemplatePresentation = pptApp.Presentations.Open(pptTemplateName)

                If pptFirstTime Then

                    If pptTemplatePresentation.PageSetup.SlideOrientation = pptCurrentPresentation.PageSetup.SlideOrientation And _
                        pptTemplatePresentation.PageSetup.SlideSize = pptCurrentPresentation.PageSetup.SlideSize Then
                        ' also in Ordnung, es kann weiter in die Current Presentation geschrieben werden ... 
                    Else
                        ' jetzt muss geprüft werden, ob die aktuelle Präsentation genauso heisst wie die zu öffnende ..
                        ' wenn ja, wird beendet - der User bekommt die Aufforderung die aktuelle Präsentation erst zu speichern  
                        Try
                            ' jetzt wird die entsprechende Template Präsentation geöffnet 
                            If pptTemplatePresentation.PageSetup.SlideOrientation = MsoOrientation.msoOrientationHorizontal Then
                                pptCurrentPresentation = pptApp.Presentations.Open(presentationFile)
                            Else
                                pptCurrentPresentation = pptApp.Presentations.Open(presentationFileH)
                            End If

                        Catch ex As Exception
                            ' in diesem Fall existiert schon eine geöffnete BoardDossier, allerdings mit anderem Format ...

                            pptTemplatePresentation.Saved = True
                            pptTemplatePresentation.Close()

                            e.Result = "Abbruch ... bitte speichern und schliessen Sie die offenen Präsentationen ... "
                            If worker.WorkerReportsProgress Then
                                worker.ReportProgress(0, e)
                            End If

                            Exit Sub

                        End Try

                    End If

                End If
            End If


        Catch ex As Exception
            e.Result = "Abbruch ... bitte speichern und schliessen Sie die offenen Präsentationen ... "
            If worker.WorkerReportsProgress Then
                worker.ReportProgress(0, e)
            End If

            Exit Sub
        End Try

        Dim anzSlidesToAdd As Integer
        Dim anzahlCurrentSlides As Integer
        Dim currentInsert As Integer = 1

        ' jetzt wird das CurrentPresentation File unter einem Dummy Namen gespeichert ..


        Try

            ' löschen, wenn der Name bereits existiert ...
            If My.Computer.FileSystem.FileExists(newFileName) And _
                pptCurrentPresentation.Name <> "Report.pptx" Then

                Try
                    My.Computer.FileSystem.DeleteFile(newFileName)
                Catch ex1 As Exception

                End Try

            End If
            ' speichern unter .. , damit Projektdossier nicht überschrieben werden kann 
            pptCurrentPresentation.SaveAs(newFileName)

            anzahlCurrentSlides = pptCurrentPresentation.Slides.Count
            anzSlidesToAdd = pptTemplatePresentation.Slides.Count
            pptTemplatePresentation.Saved = True
            pptTemplatePresentation.Close()

        Catch ex As Exception
            Throw New Exception("Probleme mit Powerpoint Template")
        End Try

        Dim reportObj As xlNS.ChartObject
        Dim obj As xlNS.ChartObject
        Dim kennzeichnung As String = ""
        Dim anzShapes As Integer

        Dim folieIX As Integer = 1
        Dim objectsToDo As Integer = 0
        Dim objectsDone As Integer = 0


        While folieIX <= anzSlidesToAdd
            'For j = 1 To anzSlidesToAdd

            If worker.WorkerSupportsCancellation Then

                If worker.CancellationPending Then
                    e.Cancel = True
                    e.Result = "Berichterstellung nach " & folieIX - 1 & " Seiten abgebrochen ..."
                    Exit While
                End If

            End If

            ' jetzt wird eine Seite aus der Vorlage ergänzt 

            ' ur:31.03.2015
            Dim tmpIX As Integer
            Dim tmpslideID As Integer

            'If Not pptFirstTime And kennzeichnung = "Multivariantensicht" Multiprojektsicht Then
            'If pptFirstTime Or _
            '    Not (kennzeichnung = "Multivariantensicht" _
            '    Or kennzeichnung = "Multiprojektsicht" _
            '    Or kennzeichnung = "AllePlanElemente" _
            '    Or kennzeichnung = "Swimlanes1" _
            '    Or kennzeichnung = "Swimlanes2") Then
            If pptFirstTime Then

                anzahlCurrentSlides = pptCurrentPresentation.Slides.Count
                tmpIX = pptCurrentPresentation.Slides.InsertFromFile(FileName:=pptTemplateName, Index:=anzahlCurrentSlides, _
                                                                              SlideStart:=folieIX, SlideEnd:=folieIX)
            Else

                pptCurrentPresentation.Slides("tmpSav").Copy()
                tmpslideID = pptCurrentPresentation.Slides("tmpSav").SlideID
                pptCurrentPresentation.Slides.Paste(pptCurrentPresentation.Slides.Count + 1)
                pptSlide = pptCurrentPresentation.Slides(pptCurrentPresentation.Slides.Count)


            End If

            '' ''Dim tmpIX As Integer
            ' '' ''tmpIX = pptCurrentPresentation.Slides.InsertFromFile(FileName:=pptTemplateName, Index:=anzahlCurrentSlides + folieIX - 1, _
            ' '' ''                                                              SlideStart:=folieIX, SlideEnd:=folieIX)

            '' ''tmpIX = pptCurrentPresentation.Slides.InsertFromFile(FileName:=pptTemplateName, Index:=anzahlCurrentSlides, _
            '' ''                                                              SlideStart:=folieIX, SlideEnd:=folieIX)



            'frmSelectPPTTempl.statusNotification.Text = "Liste der Seiten aufgebaut ...."
            e.Result = "Bericht Seite " & folieIX & " wird aufgebaut ...."

            If worker.WorkerReportsProgress Then
                worker.ReportProgress(0, e)
            End If

            anzahlCurrentSlides = pptCurrentPresentation.Slides.Count
            'pptSlide = pptCurrentPresentation.Slides(anzahlCurrentSlides + folieIX)
            pptSlide = pptCurrentPresentation.Slides(anzahlCurrentSlides)

            ' jetzt werden die Charts gezeichnet 
            anzShapes = pptSlide.Shapes.Count
            Dim newShapeRange As pptNS.ShapeRange
            Dim newShapeRange2 As pptNS.ShapeRange
            Dim newShape As pptNS.Shape


            ' jetzt wird die listofShapes aufgebaut - das sind alle Shapes, die ersetzt werden müssen ...
            For i = 1 To anzShapes
                pptShape = pptSlide.Shapes(i)
                qualifier = ""
                With pptShape

                    Dim tmpStr(3) As String
                    Try

                        If .Title <> "" Then
                            kennzeichnung = .Title
                        Else
                            tmpStr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {CChar("("), CChar(")")}, 3)
                            kennzeichnung = tmpStr(0).Trim
                        End If


                    Catch ex As Exception
                        kennzeichnung = "nicht identifizierbar"
                    End Try

                    If kennzeichnung = "Projekt-Name" Or _
                        kennzeichnung = "Soll-Ist & Prognose" Or _
                        kennzeichnung = "Multivariantensicht" Or _
                        kennzeichnung = "AllePlanElemente" Or _
                        kennzeichnung = "Swimlanes" Or _
                        kennzeichnung = "Swimlanes2" Or _
                        kennzeichnung = "Legenden-Tabelle" Or _
                        kennzeichnung = "Projekt-Grafik" Or _
                        kennzeichnung = "Meilenstein Trendanalyse" Or _
                        kennzeichnung = "Vergleich mit Beauftragung" Or _
                        kennzeichnung = "Vergleich mit letztem Stand" Or _
                        kennzeichnung = "Vergleich mit Vorlage" Or _
                        kennzeichnung = "Tabelle Projektziele" Or _
                        kennzeichnung = "Tabelle Projektstatus" Or _
                        kennzeichnung = "Tabelle Veränderungen" Or _
                        kennzeichnung = "Tabelle Vergleich letzter Stand" Or _
                        kennzeichnung = "Tabelle Vergleich Beauftragung" Or _
                        kennzeichnung = "Tabelle OneGlance Beauftragung" Or _
                        kennzeichnung = "Tabelle OneGlance letzter Stand" Or _
                        kennzeichnung = "Ergebnis" Or _
                        kennzeichnung = "Strategie/Risiko" Or _
                        kennzeichnung = "Strategie/Risiko/Ausstrahlung" Or _
                        kennzeichnung = "Projektphasen" Or _
                        kennzeichnung = "Personalbedarf" Or _
                        kennzeichnung = "Personalkosten" Or _
                        kennzeichnung = "Sonstige Kosten" Or _
                        kennzeichnung = "Gesamtkosten" Or _
                        kennzeichnung = "Trend Strategischer Fit/Risiko" Or _
                        kennzeichnung = "Trend Kennzahlen" Or _
                        kennzeichnung = "Fortschritt Personalkosten" Or _
                        kennzeichnung = "Fortschritt Sonstige Kosten" Or _
                        kennzeichnung = "Fortschritt Rolle" Or _
                        kennzeichnung = "Fortschritt Kostenart" Or _
                        kennzeichnung = "Soll-Ist1 Personalkosten" Or _
                        kennzeichnung = "Soll-Ist2 Personalkosten" Or _
                        kennzeichnung = "Soll-Ist1C Personalkosten" Or _
                        kennzeichnung = "Soll-Ist2C Personalkosten" Or _
                        kennzeichnung = "Soll-Ist1 Sonstige Kosten" Or _
                        kennzeichnung = "Soll-Ist2 Sonstige Kosten" Or _
                        kennzeichnung = "Soll-Ist1C Sonstige Kosten" Or _
                        kennzeichnung = "Soll-Ist2C Sonstige Kosten" Or _
                        kennzeichnung = "Soll-Ist1 Gesamtkosten" Or _
                        kennzeichnung = "Soll-Ist2 Gesamtkosten" Or _
                        kennzeichnung = "Soll-Ist1C Gesamtkosten" Or _
                        kennzeichnung = "Soll-Ist2C Gesamtkosten" Or _
                        kennzeichnung = "Soll-Ist1 Rolle" Or _
                        kennzeichnung = "Soll-Ist2 Rolle" Or _
                        kennzeichnung = "Soll-Ist1C Rolle" Or _
                        kennzeichnung = "Soll-Ist2C Rolle" Or _
                        kennzeichnung = "Soll-Ist1 Kostenart" Or _
                        kennzeichnung = "Soll-Ist2 Kostenart" Or _
                        kennzeichnung = "Soll-Ist1C Kostenart" Or _
                        kennzeichnung = "Soll-Ist2C Kostenart" Or _
                        kennzeichnung = "Ampel-Farbe" Or _
                        kennzeichnung = "Ampel-Text" Or _
                        kennzeichnung = "Beschreibung" Or _
                        kennzeichnung = "Business-Unit:" Or _
                        kennzeichnung = "Stand:" Or _
                        kennzeichnung = "Laufzeit:" Or _
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



            For Each tmpShape As pptNS.Shape In listofShapes


                Try

                    obj = Nothing
                    pptShape = tmpShape
                    qualifier = ""
                    qualifier2 = ""
                    kennzeichnung = ""

                    With pptShape
                        .Name = "Shape" & .Id.ToString

                        If .Title <> "" Then
                            kennzeichnung = .Title
                            qualifier = .AlternativeText
                            boxName = kennzeichnung
                        Else
                            ' Start neu
                            Dim tmpStr(10) As String
                            Try

                                tmpStr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {CChar("("), CChar(")")}, 10)
                                kennzeichnung = tmpStr(0).Trim

                            Catch ex As Exception
                                kennzeichnung = "nicht identifizierbar"
                                tmpStr(0) = " "
                            End Try

                            Try
                                If tmpStr.Length < 2 Then
                                    qualifier = ""
                                    qualifier2 = ""
                                ElseIf tmpStr.Length = 2 Then
                                    qualifier = tmpStr(1).Trim
                                ElseIf tmpStr.Length >= 3 Then
                                    qualifier = tmpStr(1).Trim
                                    qualifier2 = tmpStr(2).Trim
                                End If

                            Catch ex As Exception
                                qualifier = ""
                                qualifier2 = ""
                            End Try
                            ' Ende neu 
                        End If



                        top = .Top
                        left = .Left
                        height = .Height
                        width = .Width

                        Try
                            boxName = .TextFrame2.TextRange.Text
                        Catch ex As Exception
                            boxName = " "
                        End Try


                        notYetDone = False
                        reportObj = Nothing

                        htop = 100
                        hleft = 100
                        hwidth = 300
                        hheight = 400

                        Select Case kennzeichnung

                            Case "Projekt-Name"

                                If qualifier.Length > 0 Then
                                    .TextFrame2.TextRange.Text = fullName & ": " & qualifier
                                Else
                                    .TextFrame2.TextRange.Text = fullName
                                End If

                            Case "Projekt-Grafik"

                                Try

                                    Call zeichneProjektGrafik(pptSlide, pptShape, hproj, selectedMilestones)

                                Catch ex As Exception

                                End Try


                            Case "Legenden-Tabelle"

                                Try
                                    ' Einzelprojektsicht im Extended Mode
                                    If selectedPhases.Count = 0 _
                                        And selectedMilestones.Count = 0 _
                                        And selectedRoles.Count = 0 _
                                        And selectedCosts.Count = 0 _
                                        And selectedBUs.Count = 0 _
                                        Then
                                        Dim i As Integer = 0
                                        Dim tmpphases As New Collection
                                        Dim tmpMilestones As New Collection

                                        ' alle Phasennamen des Projektes hproj in die Collection tmpphases bringen
                                        For Each cphase In hproj.AllPhases

                                            Dim tmpstr = hproj.hierarchy.getBreadCrumb(cphase.nameID)
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

                                        Call prepZeichneLegendenTabelle(pptSlide, pptShape, legendFontSize, tmpphases, tmpMilestones)
                                    Else

                                        Call prepZeichneLegendenTabelle(pptSlide, pptShape, legendFontSize, selectedPhases, selectedMilestones)
                                    End If

                                Catch ex As Exception

                                End Try

                            Case "AllePlanElemente"


                                Try
                                    Dim i As Integer = 0
                                    Dim tmpphases As New Collection
                                    Dim tmpMilestones As New Collection

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

                                    Call zeichneMultiprojektSicht(pptApp, pptCurrentPresentation, pptSlide, _
                                                                  objectsToDo, objectsDone, pptFirstTime, zeilenhoehe, legendFontSize, _
                                                                  tmpphases, tmpMilestones, _
                                                                  selectedRoles, selectedCosts, _
                                                                  selectedBUs, selectedTyps, _
                                                                  worker, e, False, hproj, kennzeichnung)
                                    .TextFrame2.TextRange.Text = ""
                                    .ZOrder(MsoZOrderCmd.msoSendToBack)
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message
                                    objectsDone = objectsToDo

                                End Try


                            Case "Multivariantensicht"


                                Try

                                    Call zeichneMultiprojektSicht(pptApp, pptCurrentPresentation, pptSlide, _
                                                                      objectsToDo, objectsDone, pptFirstTime, zeilenhoehe, legendFontSize, _
                                                                      selectedPhases, selectedMilestones, _
                                                                      selectedRoles, selectedCosts, _
                                                                      selectedBUs, selectedTyps, _
                                                                      worker, e, False, hproj, kennzeichnung)
                                    .TextFrame2.TextRange.Text = ""
                                    .ZOrder(MsoZOrderCmd.msoSendToBack)
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message
                                    objectsDone = objectsToDo
                                End Try


                            Case "Swimlanes"

                                Try


                                    Call zeichneSwimlane2Sicht(pptApp, pptCurrentPresentation, pptSlide, _
                                                                      objectsToDo, objectsDone, pptFirstTime, zeilenhoehe, legendFontSize, _
                                                                      selectedPhases, selectedMilestones, _
                                                                      selectedRoles, selectedCosts, _
                                                                      selectedBUs, selectedTyps, _
                                                                      worker, e, False, hproj, kennzeichnung)

                                    .TextFrame2.TextRange.Text = ""
                                    .ZOrder(MsoZOrderCmd.msoSendToBack)

                                    ' sonst wird pptLasttime benötigt, um bei mehreren PRojekten 
                                    ' swimlaneMode wird erst nach Ende der While Schleife ausgewertet - in diesem Fall wird die tmpSav Folie gelöscht 
                                    swimlaneMode = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = ex.Message
                                    objectsDone = objectsToDo
                                End Try


                            Case "Swimlanes2"
                                Dim formerSetting As Boolean = awinSettings.mppExtendedMode
                                Try

                                    awinSettings.mppExtendedMode = True

                                    
                                    Call zeichneSwimlane2Sicht(pptApp, pptCurrentPresentation, pptSlide, _
                                                                      objectsToDo, objectsDone, pptFirstTime, zeilenhoehe, legendFontSize, _
                                                                      selectedPhases, selectedMilestones, _
                                                                      selectedRoles, selectedCosts, _
                                                                      selectedBUs, selectedTyps, _
                                                                      worker, e, False, hproj, kennzeichnung)
                                    awinSettings.mppExtendedMode = formerSetting
                                    .TextFrame2.TextRange.Text = ""
                                    .ZOrder(MsoZOrderCmd.msoSendToBack)

                                    ' sonst wird pptLasttime benötigt, um bei mehreren PRojekten 
                                    ' swimlaneMode wird erst nach Ende der While Schleife ausgewertet - in diesem Fall wird die tmpSav Folie gelöscht 
                                    swimlaneMode = True
                                Catch ex As Exception
                                    awinSettings.mppExtendedMode = formerSetting
                                    .TextFrame2.TextRange.Text = ex.Message
                                    objectsDone = objectsToDo
                                End Try


                            Case "Meilenstein Trendanalyse"

                                Dim nameList As New SortedList(Of Date, String)
                                Dim listOfItems As New Collection

                                boxName = "Meilenstein Trendanalyse"

                                Try
                                    ' Aufruf 
                                    If qualifier = "" Then
                                        ' alle Meilensteine anzeigen
                                        nameList = hproj.getMilestones

                                        If nameList.Count > 0 Then
                                            For Each kvp As KeyValuePair(Of Date, String) In nameList
                                                listOfItems.Add(kvp.Value)
                                            Next
                                        End If


                                    Else
                                        ' nur die anzeigen, die im qualifier mit # voneinander getrennt stehen  
                                        Dim tmpStr(20) As String
                                        Try

                                            tmpStr = qualifier.Trim.Split(New Char() {CChar("#")}, 20)
                                            kennzeichnung = tmpStr(0).Trim

                                        Catch ex As Exception

                                        End Try


                                        ' die ListofItems muss die eindeutigen IDs beeinhalten
                                        For i = 1 To tmpStr.Length

                                            Dim fullmsName As String = tmpStr(i - 1).Trim
                                            Dim msName As String = ""
                                            Dim breadcrumb As String = ""
                                            Call splitHryFullnameTo2(fullmsName, msName, breadcrumb)
                                            Dim milestoneIndices(,) As Integer = hproj.hierarchy.getMilestoneIndices(msName, breadcrumb)
                                            Dim msItem As String

                                            For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                                                If milestoneIndices(0, mx) > 0 And milestoneIndices(1, mx) > 0 Then

                                                    Try
                                                        msItem = hproj.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx)).nameID
                                                        listOfItems.Add(msItem)
                                                    Catch ex As Exception

                                                    End Try


                                                End If

                                            Next

                                        Next

                                    End If

                                    ' jetzt ist listofItems entsprechend gefüllt 
                                    If listOfItems.Count > 0 Then
                                        htop = 100
                                        hleft = 50
                                        hheight = 2 * ((listOfItems.Count - 1) * 20 + 110)
                                        hwidth = System.Math.Max(hproj.anzahlRasterElemente * boxWidth + 10, 24 * boxWidth + 10)

                                        Try
                                            Call createMsTrendAnalysisOfProject(hproj, obj, listOfItems, htop, hleft, hheight, hwidth)

                                            reportObj = obj
                                            notYetDone = True
                                        Catch ex As Exception
                                            .TextFrame2.TextRange.Text = "zum Projekt" & hproj.name & vbLf & "gibt es noch keine Trend-Analyse," & vbLf & _
                                                                        "da es noch nicht begonnen hat"
                                        End Try

                                    Else
                                        .TextFrame2.TextRange.Text = "es gibt keine Meilensteine im Projekt" & vbLf & hproj.name
                                    End If

                                Catch ex As Exception

                                End Try

                            Case "Projektphasen"

                                Dim scale As Integer
                                Dim continueWork As Boolean = True
                                Dim cproj As clsProjekt = Nothing
                                Dim vproj As clsProjektvorlage
                                auswahl = 0

                                scale = hproj.dauerInDays

                                If qualifier.Length > 0 Then
                                    If qualifier = "Vorlage" Then
                                        auswahl = 1
                                        vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                                        If IsNothing(vproj) Then
                                            .TextFrame2.TextRange.Text = "Projekt-Vorlage " & hproj.VorlagenName & " existiert nicht !"
                                            continueWork = False
                                        Else
                                            vproj.copyTo(cproj)
                                            cproj.startDate = hproj.startDate
                                        End If

                                    ElseIf qualifier = "Beauftragung" Then
                                        cproj = bproj
                                        auswahl = 2

                                    Else
                                        cproj = hproj
                                        auswahl = 0

                                    End If
                                Else
                                    cproj = hproj
                                    auswahl = 0
                                End If

                                If continueWork Then
                                    htop = 150
                                    hleft = 150


                                    hheight = 380
                                    hwidth = 900
                                    scale = cproj.dauerInDays

                                    Dim noColorCollection As New Collection
                                    reportObj = Nothing
                                    Call createPhasesBalken(noColorCollection, cproj, reportObj, scale, htop, hleft, hheight, hwidth, auswahl)


                                    notYetDone = True
                                End If


                            Case "Vergleich mit Vorlage"

                                Dim vproj As clsProjektvorlage
                                Dim cproj As New clsProjekt
                                Dim scale As Double
                                Dim noColorCollection As New Collection
                                Dim repObj1 As xlNS.ChartObject, repObj2 As xlNS.ChartObject
                                Dim continueWork As Boolean = True

                                ' jetzt die Aktion durchführen ...


                                Try

                                    vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                                    If IsNothing(vproj) Then
                                        .TextFrame2.TextRange.Text = "Projekt-Vorlage " & hproj.VorlagenName & " existiert nicht !"
                                        continueWork = False
                                    Else
                                        cproj = New clsProjekt
                                        vproj.copyTo(cproj)
                                        cproj.startDate = hproj.startDate
                                    End If

                                Catch ex As Exception
                                    Throw New Exception("Vorlage konnte nicht bestimmt werden")
                                End Try

                                If continueWork Then
                                    htop = 150
                                    hleft = 150


                                    hheight = 380
                                    hwidth = 900
                                    scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)


                                    appInstance.EnableEvents = False


                                    noColorCollection = getPhasenUnterschiede(hproj, cproj)

                                    repObj1 = Nothing
                                    Call createPhasesBalken(noColorCollection, hproj, repObj1, scale, htop, hleft, hheight, hwidth, PThis.current)

                                    With repObj1
                                        htop = .Top + .Height + 3
                                    End With


                                    repObj2 = Nothing
                                    Call createPhasesBalken(noColorCollection, cproj, repObj2, scale, htop, hleft, hheight, hwidth, PThis.vorlage)

                                    ' jetzt wird das Shape in der Powerpoint entsprechend entsprechend aufgebaut 
                                    Try
                                        pptSize = CInt(.TextFrame2.TextRange.Font.Size)
                                        .TextFrame2.TextRange.Text = " "
                                    Catch ex As Exception
                                        pptSize = 12
                                    End Try


                                    Dim widthFaktor As Double = 1.0
                                    Dim heightFaktor As Double = 1.0
                                    Dim topNext As Double


                                    If Not repObj1 Is Nothing Then
                                        Try
                                            repObj1.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                            newShapeRange = pptSlide.Shapes.Paste

                                            With newShapeRange(1)
                                                .Top = CSng(top + 0.02 * height)
                                                .Left = CSng(left + 0.02 * width)
                                                .Width = CSng(width * 0.96)
                                                topNext = CSng(top + 0.04 * height + .Height)
                                                '.Height = height * 0.46
                                            End With

                                            repObj1.Delete()

                                            If Not repObj2 Is Nothing Then
                                                Try
                                                    repObj2.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                                    newShapeRange2 = pptSlide.Shapes.Paste

                                                    With newShapeRange2(1)
                                                        .Top = CSng(topNext)
                                                        .Left = CSng(left + 0.02 * width)
                                                        .Width = CSng(width * 0.96)
                                                        ' Height wird nicht gesetzt - bei Bildern wird das proportional automatisch gesetzt 
                                                    End With

                                                    ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                                    Try
                                                        If newShapeRange(1).Height + newShapeRange2(1).Height > 0.96 * height Then
                                                            widthFaktor = 0.96 * height / (newShapeRange(1).Height + newShapeRange2(1).Height)
                                                            newShapeRange(1).Width = CSng(widthFaktor * newShapeRange(1).Width)
                                                            newShapeRange2(1).Width = CSng(widthFaktor * newShapeRange2(1).Width)
                                                            newShapeRange2(1).Top = CSng(newShapeRange(1).Top + newShapeRange(1).Height + 0.02 * height)
                                                        End If
                                                    Catch ex As Exception

                                                    End Try



                                                    repObj2.Delete()
                                                Catch ex As Exception

                                                End Try

                                            End If
                                        Catch ex As Exception

                                        End Try

                                    End If

                                End If


                            Case "Vergleich mit Beauftragung"


                                Dim cproj As clsProjekt
                                Dim scale As Double
                                Dim noColorCollection As New Collection
                                Dim repObj1 As xlNS.ChartObject, repObj2 As xlNS.ChartObject



                                ' jetzt die Aktion durchführen ...


                                If bproj Is Nothing Then
                                    Throw New Exception("es gibt keine Beauftragung")
                                End If

                                cproj = bproj



                                htop = 150
                                hleft = 150

                                hheight = 380
                                hwidth = 900
                                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)


                                noColorCollection = getPhasenUnterschiede(hproj, cproj)

                                repObj1 = Nothing
                                Call createPhasesBalken(noColorCollection, hproj, repObj1, scale, htop, hleft, hheight, hwidth, PThis.current)

                                With repObj1
                                    htop = .Top + .Height + 3
                                End With

                                repObj2 = Nothing
                                Call createPhasesBalken(noColorCollection, cproj, repObj2, scale, htop, hleft, hheight, hwidth, PThis.beauftragung)

                                Try
                                    pptSize = CInt(.TextFrame2.TextRange.Font.Size)
                                    .TextFrame2.TextRange.Text = " "
                                Catch ex As Exception
                                    pptSize = 12
                                End Try



                                Dim widthFaktor As Double = 1.0
                                Dim heightFaktor As Double = 1.0
                                Dim topNext As Double

                                If Not repObj1 Is Nothing Then
                                    Try
                                        repObj1.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                        newShapeRange = pptSlide.Shapes.Paste

                                        With newShapeRange(1)
                                            .Top = CSng(top + 0.02 * height)
                                            .Left = CSng(left + 0.02 * width)
                                            .Width = CSng(width * 0.96)
                                            topNext = CSng(top + 0.04 * height + .Height)
                                            '.Height = height * 0.46
                                        End With

                                        repObj1.Delete()

                                        If Not repObj2 Is Nothing Then
                                            Try
                                                repObj2.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                                newShapeRange2 = pptSlide.Shapes.Paste

                                                With newShapeRange2(1)
                                                    .Top = CSng(topNext)
                                                    .Left = CSng(left + 0.02 * width)
                                                    .Width = CSng(width * 0.96)
                                                    '.Height = height * 0.46
                                                End With

                                                repObj2.Delete()

                                                ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                                If newShapeRange(1).Height + newShapeRange2(1).Height > 0.96 * height Then
                                                    widthFaktor = 0.96 * height / (newShapeRange(1).Height + newShapeRange2(1).Height)
                                                    newShapeRange(1).Width = CSng(widthFaktor * newShapeRange(1).Width)
                                                    newShapeRange2(1).Width = CSng(widthFaktor * newShapeRange2(1).Width)
                                                    newShapeRange2(1).Top = CSng(newShapeRange(1).Top + newShapeRange(1).Height + 0.02 * height)
                                                End If
                                            Catch ex As Exception

                                            End Try

                                        End If


                                    Catch ex As Exception

                                    End Try

                                End If


                            Case "Vergleich mit letztem Stand"


                                Dim cproj As clsProjekt
                                Dim scale As Double
                                Dim noColorCollection As New Collection
                                Dim repObj1 As xlNS.ChartObject, repObj2 As xlNS.ChartObject



                                ' jetzt die Aktion durchführen ...

                                If lastproj Is Nothing Then
                                    Throw New Exception("es gibt keinen letzten Strand")
                                End If

                                cproj = lastproj

                                htop = 150
                                hleft = 150

                                hheight = 380
                                hwidth = 900
                                scale = System.Math.Max(hproj.dauerInDays, cproj.dauerInDays)


                                noColorCollection = getPhasenUnterschiede(hproj, cproj)

                                repObj1 = Nothing
                                Call createPhasesBalken(noColorCollection, hproj, repObj1, scale, htop, hleft, hheight, hwidth, PThis.current)

                                With repObj1
                                    htop = .Top + .Height + 3
                                End With

                                repObj2 = Nothing
                                Call createPhasesBalken(noColorCollection, cproj, repObj2, scale, htop, hleft, hheight, hwidth, PThis.letzterStand)

                                Try
                                    pptSize = .TextFrame2.TextRange.Font.Size
                                    .TextFrame2.TextRange.Text = " "
                                Catch ex As Exception
                                    pptSize = 12
                                End Try



                                Dim widthFaktor As Double = 1.0
                                Dim heightFaktor As Double = 1.0
                                Dim topNext As Double

                                If Not repObj1 Is Nothing Then
                                    Try
                                        repObj1.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                        newShapeRange = pptSlide.Shapes.Paste

                                        With newShapeRange(1)
                                            .Top = CSng(top + 0.02 * height)
                                            .Left = CSng(left + 0.02 * width)
                                            .Width = CSng(width * 0.96)
                                            topNext = CSng(top + 0.04 * height + .Height)
                                            '.Height = height * 0.46
                                        End With

                                        repObj1.Delete()

                                        If Not repObj2 Is Nothing Then
                                            Try
                                                repObj2.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                                newShapeRange2 = pptSlide.Shapes.Paste

                                                With newShapeRange2(1)
                                                    .Top = CSng(topNext)
                                                    .Left = CSng(left + 0.02 * width)
                                                    .Width = CSng(width * 0.96)
                                                    '.Height = height * 0.46
                                                End With

                                                repObj2.Delete()

                                                ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                                If newShapeRange(1).Height + newShapeRange2(1).Height > 0.96 * height Then
                                                    widthFaktor = 0.96 * height / (newShapeRange(1).Height + newShapeRange2(1).Height)
                                                    newShapeRange(1).Width = CSng(widthFaktor * newShapeRange(1).Width)
                                                    newShapeRange2(1).Width = CSng(widthFaktor * newShapeRange2(1).Width)
                                                    newShapeRange2(1).Top = CSng(newShapeRange(1).Top + newShapeRange(1).Height + 0.02 * height)
                                                End If
                                            Catch ex As Exception

                                            End Try

                                        End If


                                    Catch ex As Exception

                                    End Try

                                End If


                            Case "Tabelle Projektziele"

                                Try

                                    Call zeichneProjektTabelleZiele(pptShape, hproj, selectedMilestones)

                                Catch ex As Exception

                                End Try

                            Case "Tabelle Vergleich letzter Stand"

                                Try
                                    Call zeichneProjektTabelleVergleich(pptSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, lastproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle Vergleich Beauftragung"

                                Try
                                    Call zeichneProjektTabelleVergleich(pptSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, bproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle OneGlance letzter Stand"

                                Try
                                    Call zeichneProjektTabelleOneGlance(pptSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, lastproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle OneGlance Beauftragung"

                                Try
                                    Call zeichneProjektTabelleOneGlance(pptSlide, pptShape, gleichShape, steigendShape, fallendShape, ampelShape, sternShape, hproj, bproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle Veränderungen"

                                Try
                                    Call zeichneProjektTerminAenderungen(pptShape, hproj, bproj, lproj)
                                Catch ex As Exception

                                End Try

                            Case "Tabelle Projektstatus"

                                Try
                                    Call zeichneProjektTabelleStatus(pptShape, hproj)
                                Catch ex As Exception

                                End Try

                            Case "Soll-Ist & Prognose"

                                If istWerteexistieren Then
                                Else
                                    .TextFrame2.TextRange.Text = "Prognose"
                                End If

                            Case "Ergebnis"


                                Try
                                    If qualifier = "letzter Stand" Then
                                        Call createProjektErgebnisCharakteristik2(lproj, obj, PThis.letzterStand)

                                    ElseIf qualifier = "Beauftragung" Then
                                        Call createProjektErgebnisCharakteristik2(bproj, obj, PThis.beauftragung)

                                    Else
                                        Call createProjektErgebnisCharakteristik2(hproj, obj, PThis.current)

                                    End If



                                    reportObj = obj

                                    Dim ax As xlNS.Axis = CType(reportObj.Chart.Axes(xlNS.XlAxisType.xlCategory), Excel.Axis)
                                    With ax
                                        .TickLabels.Font.Size = 12
                                    End With

                                    notYetDone = True
                                Catch ex As Exception

                                End Try



                            Case "Strategie/Risiko"

                                Dim mycollection As New Collection

                                'deleteStack.Add(.Name, .Name)
                                Try
                                    mycollection.Add(pname)

                                    Call awinCreatePortfolioDiagrams(mycollection, reportObj, True, PTpfdk.FitRisiko, PTpfdk.ProjektFarbe, True, False, True, htop, hleft, hwidth, hheight)
                                    notYetDone = True
                                Catch ex As Exception
                                    Dim a As Integer = -1
                                End Try

                            Case "Strategie/Risiko/Ausstrahlung"

                                Dim mycollection As New Collection

                                'deleteStack.Add(.Name, .Name)
                                Try
                                    mycollection.Add(pname)

                                    Call awinCreatePortfolioDiagrams(mycollection, reportObj, True, PTpfdk.FitRisikoDependency, PTpfdk.ProjektFarbe, True, False, True, htop, hleft, hwidth, hheight)
                                    notYetDone = True
                                Catch ex As Exception

                                End Try


                            Case "Personalbedarf"

                                Try
                                    auswahl = 1

                                    If qualifier.Length > 0 Then
                                        If qualifier.Trim <> "Balken" Then
                                            Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                        Else
                                            Call createRessBalkenOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                        End If
                                    Else
                                        Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                    End If

                                    reportObj = obj
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Personal-Bedarf ist Null"
                                End Try


                            Case "Personalkosten"

                                Try
                                    auswahl = 2

                                    If qualifier.Length > 0 Then

                                        If qualifier.Trim <> "Balken" Then
                                            Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                        Else
                                            Call createRessBalkenOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                        End If

                                    Else
                                        Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                    End If


                                    reportObj = obj
                                    notYetDone = True

                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Personal-Kosten sind Null"
                                End Try

                            Case "Sonstige Kosten"



                                Try
                                    auswahl = 1

                                    If qualifier.Length > 0 Then

                                        If qualifier.Trim <> "Balken" Then
                                            Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                        Else
                                            Call createCostBalkenOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                        End If

                                    Else
                                        Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)
                                    End If



                                    reportObj = obj
                                    notYetDone = True
                                Catch ex As Exception

                                    .TextFrame2.TextRange.Text = "Sonstige Kosten sind Null"

                                End Try


                            Case "Gesamtkosten"

                                'htop = 100
                                'hleft = 100
                                'hwidth = boxWidth * 14
                                'hheight = boxHeight * 10

                                Try
                                    auswahl = 2
                                    Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)

                                    reportObj = obj
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Gesamtkosten sind Null"
                                End Try


                            Case "Trend Strategischer Fit/Risiko"

                                Dim nrSnapshots As Integer = projekthistorie.Count

                                If nrSnapshots > 0 Then

                                    Call createTrendSfit(obj, htop, hleft, hheight, hwidth)

                                    reportObj = obj
                                    notYetDone = True

                                Else
                                    .TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                End If



                            Case "Trend Kennzahlen"

                                Dim nrSnapshots As Integer = projekthistorie.Count

                                If nrSnapshots > 0 Then
                                    'htop = 100
                                    'hleft = 100
                                    'hwidth = 300
                                    'hheight = 400

                                    Call createTrendKPI(obj, htop, hleft, hheight, hwidth)

                                    reportObj = obj
                                    notYetDone = True

                                Else
                                    .TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                End If

                            Case "Fortschritt Personalkosten"

                                Dim nrSnapshots As Integer = projekthistorie.Count
                                Dim PListe As New Collection
                                compareToID = 1
                                auswahl = 1

                                If nrSnapshots > 0 Then

                                    If istLaufendesProjekt(hproj) Then

                                        PListe.Add(hproj.name, hproj.name)
                                        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                        If Not obj Is Nothing Then
                                            reportObj = obj
                                            notYetDone = True

                                            With reportObj
                                                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                            End With
                                        Else
                                            .TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                        End If


                                    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                        .TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                    Else
                                        .TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                    End If

                                Else
                                    .TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                End If

                            Case "Fortschritt Sonstige Kosten"
                                Dim nrSnapshots As Integer = projekthistorie.Count
                                Dim PListe As New Collection
                                compareToID = 1
                                auswahl = 2

                                If nrSnapshots > 0 Then

                                    If istLaufendesProjekt(hproj) Then

                                        PListe.Add(hproj.name, hproj.name)
                                        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                        If Not obj Is Nothing Then
                                            reportObj = obj
                                            notYetDone = True

                                            With reportObj
                                                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                            End With
                                        Else
                                            .TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                        End If

                                    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                        .TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                    Else
                                        .TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                    End If

                                Else
                                    .TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                End If

                            Case "Fortschritt Gesamtkosten"
                                Dim nrSnapshots As Integer = projekthistorie.Count
                                Dim PListe As New Collection
                                compareToID = 1
                                auswahl = 3

                                If nrSnapshots > 0 Then

                                    If istLaufendesProjekt(hproj) Then

                                        PListe.Add(hproj.name, hproj.name)
                                        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                        If Not obj Is Nothing Then
                                            reportObj = obj
                                            notYetDone = True

                                            With reportObj
                                                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                            End With
                                        Else
                                            .TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                        End If

                                    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                        .TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                    Else
                                        .TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                    End If

                                Else
                                    .TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                End If

                            Case "Fortschritt Rolle"

                                Dim nrSnapshots As Integer = projekthistorie.Count
                                Dim PListe As New Collection
                                compareToID = 1
                                auswahl = 4

                                If nrSnapshots > 0 Then

                                    If istLaufendesProjekt(hproj) Then

                                        PListe.Add(hproj.name, hproj.name)
                                        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                        boxName = "Fortschritt " & qualifier

                                        If Not obj Is Nothing Then
                                            reportObj = obj
                                            notYetDone = True

                                            With reportObj
                                                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                            End With
                                        Else
                                            .TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                        End If

                                    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                        .TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                    Else
                                        .TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                    End If

                                Else
                                    .TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                End If

                            Case "Fortschritt Kostenart"

                                Dim nrSnapshots As Integer = projekthistorie.Count
                                Dim PListe As New Collection
                                compareToID = 1
                                auswahl = 5

                                If nrSnapshots > 0 Then

                                    If istLaufendesProjekt(hproj) Then

                                        PListe.Add(hproj.name, hproj.name)
                                        Call awinCreateStatusDiagram1(PListe, obj, compareToID, auswahl, qualifier, False, False, htop, hleft, hwidth, hheight)

                                        boxName = "Fortschritt " & qualifier

                                        If Not obj Is Nothing Then
                                            reportObj = obj
                                            notYetDone = True

                                            With reportObj
                                                .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                                                .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                                            End With
                                        Else
                                            .TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                        End If

                                    ElseIf hproj.Start > getColumnOfDate(Date.Now) Then
                                        .TextFrame2.TextRange.Text = "Projekt hat noch nicht begonnen ... "
                                    Else
                                        .TextFrame2.TextRange.Text = "Projekt ist bereits beendet"
                                    End If

                                Else
                                    .TextFrame2.TextRange.Text = "es existiert noch keine Projekt-Historie"
                                End If


                            Case "Soll-Ist1 Personalkosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                    Dim vglBaseline As Boolean = True

                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 1, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Personalkosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Personalkosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist2 Personalkosten"


                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                    Dim vglBaseline As Boolean = False


                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 1, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Personalkosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Personalkosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist1C Personalkosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                    Dim vglBaseline As Boolean = True

                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 1, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Personalkosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Personalkosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist2C Personalkosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                    Dim vglBaseline As Boolean = False

                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 1, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Personalkosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Personalkosten nicht möglich ..."
                                End Try



                            Case "Soll-Ist1 Sonstige Kosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                    Dim vglBaseline As Boolean = True

                                    reportObj = Nothing
                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 2, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Sonstige Kosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Sonstige Kosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist2 Sonstige Kosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                    Dim vglBaseline As Boolean = False

                                    reportObj = Nothing
                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 2, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Sonstige Kosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Sonstige Kosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist1C Sonstige Kosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                    Dim vglBaseline As Boolean = True


                                    reportObj = Nothing
                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 2, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Sonstige Kosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Sonstige Kosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist2C Sonstige Kosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der last freigabe verglichen
                                    Dim vglBaseline As Boolean = False


                                    reportObj = Nothing
                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 2, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Sonstige Kosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Sonstige Kosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist1 Gesamtkosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                    Dim vglBaseline As Boolean = True

                                    reportObj = Nothing
                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 3, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Gesamtkosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Gesamtkosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist2 Gesamtkosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                    Dim vglBaseline As Boolean = False

                                    reportObj = Nothing
                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 3, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Gesamtkosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Gesamtkosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist1C Gesamtkosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                    Dim vglBaseline As Boolean = True

                                    reportObj = Nothing
                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 3, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Gesamtkosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Gesamtkosten nicht möglich ..."
                                End Try


                            Case "Soll-Ist2C Gesamtkosten"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der last freigabe verglichen
                                    Dim vglBaseline As Boolean = False

                                    reportObj = Nothing
                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 3, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Gesamtkosten" & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Gesamtkosten nicht möglich ..."
                                End Try



                            Case "Soll-Ist1 Rolle"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Beauftragung verglichen
                                    Dim vglBaseline As Boolean = True

                                    reportObj = Nothing
                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 4, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Rolle " & qualifier & ze
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Rolle " & qualifier & " nicht möglich ..."
                                End Try


                            Case "Soll-Ist2 Rolle"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                    Dim vglBaseline As Boolean = False

                                    reportObj = Nothing
                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 4, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Rolle " & qualifier & ze
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Rolle " & qualifier & " nicht möglich ..."
                                End Try


                            Case "Soll-Ist1C Rolle"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Beauftragung verglichen
                                    Dim vglBaseline As Boolean = True


                                    reportObj = Nothing
                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 4, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Rolle " & qualifier & ze
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Rolle " & qualifier & " nicht möglich ..."
                                End Try


                            Case "Soll-Ist2C Rolle"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der last freigabe verglichen
                                    Dim vglBaseline As Boolean = False


                                    reportObj = Nothing
                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 4, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Rolle " & qualifier & ze
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Rolle " & qualifier & " nicht möglich ..."
                                End Try


                            Case "Soll-Ist1 Kostenart"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Beauftragung verglichen
                                    Dim vglBaseline As Boolean = True

                                    reportObj = Nothing
                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 5, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Kostenart " & qualifier & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Kostenart " & qualifier & " nicht möglich ..."
                                End Try


                            Case "Soll-Ist2 Kostenart"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                    Dim vglBaseline As Boolean = False

                                    reportObj = Nothing
                                    Call createSollIstOfProject(hproj, reportObj, Date.Now, 5, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Kostenart " & qualifier & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Kostenart " & qualifier & " nicht möglich ..."
                                End Try


                            Case "Soll-Ist1C Kostenart"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der Beauftragung verglichen
                                    Dim vglBaseline As Boolean = True

                                    reportObj = Nothing
                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 5, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Kostenart " & qualifier & ke
                                    notYetDone = True

                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Kostenart " & qualifier & " nicht möglich ..."
                                End Try


                            Case "Soll-Ist2C Kostenart"

                                Try
                                    ' bei bereits beauftragten Projekten: es wird Current mit der last freigabe verglichen
                                    Dim vglBaseline As Boolean = False

                                    reportObj = Nothing
                                    Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 5, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                    boxName = "Kostenart " & qualifier & ke
                                    notYetDone = True
                                Catch ex As Exception
                                    .TextFrame2.TextRange.Text = "Soll-Ist Kostenart " & qualifier & " nicht möglich ..."
                                End Try


                            Case "Ampel-Farbe"

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

                            Case "Ampel-Text"
                                .TextFrame2.TextRange.Text = hproj.ampelErlaeuterung

                            Case "Business-Unit:"
                                .TextFrame2.TextRange.Text = boxName & " " & hproj.businessUnit

                            Case "Beschreibung"
                                .TextFrame2.TextRange.Text = hproj.description

                            Case "Stand:"
                                .TextFrame2.TextRange.Text = boxName & " " & hproj.timeStamp.ToShortDateString

                            Case "Laufzeit:"
                                .TextFrame2.TextRange.Text = boxName & " " & textZeitraum(hproj.Start, hproj.Start + hproj.anzahlRasterElemente - 1)

                            Case "Verantwortlich:"
                                .TextFrame2.TextRange.Text = boxName & " " & hproj.leadPerson
                            Case Else
                        End Select

                        If notYetDone Then

                            Try
                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "
                            Catch ex As Exception
                                pptSize = 12
                            End Try

                            If Not reportObj Is Nothing Then
                                Try
                                    With reportObj
                                        .Chart.ChartTitle.Text = boxName
                                        .Chart.ChartTitle.Font.Size = pptSize
                                    End With

                                    reportObj.Copy()
                                    newShapeRange = pptSlide.Shapes.Paste
                                    newShape = newShapeRange.Item(1)

                                    With newShape
                                        .Top = CSng(top + 0.02 * height)
                                        .Left = CSng(left + 0.02 * width)
                                        .Width = CSng(width * 0.96)
                                        .Height = CSng(height * 0.96)
                                    End With

                                    reportObj.Delete()
                                Catch ex As Exception

                                End Try
                            Else
                                Try
                                    .TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                                Catch ex As Exception

                                End Try

                            End If



                        End If


                    End With

                Catch ex As Exception

                    'tmpShape.TextFrame2.TextRange.Text = tmpShape.Title & ": Fehler ..."

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
                pptFirstTime = True  ' damit die Folie für die Legende geholt wird
                objectsToDo = 0
                objectsDone = 0
            End If

        End While

        If pptLastTime Or swimlaneMode Then
            Try
                If Not IsNothing(pptCurrentPresentation.Slides("tmpSav")) Then
                    pptCurrentPresentation.Slides("tmpSav").Delete()   ' Vorlage in passender Größe wird nun nicht mehr benötigt
                End If
            Catch ex As Exception

            End Try

        End If

    End Sub
    '
    '
    ' 
    ''' <summary>
    ''' erzeugt mit den übergebenen Daten den Report aus der übergebenen Report vorlage
    ''' </summary>
    ''' <param name="pptTemplateName">Name der PPT Vorlage</param>
    ''' <param name="selectedPhases">Liste mit den Phasen Namen</param>
    ''' <param name="selectedMilestones">Liste mit den Milestone Namen</param>
    ''' <param name="selectedRoles">Liste der Rollen</param>
    ''' <param name="selectedCosts">Liste der Kosten</param>
    ''' <param name="worker">background Worker</param>
    ''' <param name="e">hier werden die Progress Meldungen zurückgegeben</param>
    ''' <remarks></remarks>
    Public Sub createPPTSlidesFromConstellation(ByVal pptTemplateName As String, _
                                                    ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                                    ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                                    ByVal selectedBUs As Collection, ByVal selectedTyps As Collection, _
                                                    ByRef pptFirstTime As Boolean, _
                                                    ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)
        'ByVal showNames As Boolean, ByVal showProjectLine As Boolean,
        'ByVal showAmpeln As Boolean, ByVal showDates As Boolean, ByVal strict As Boolean, _

        Dim pptApp As pptNS.Application = Nothing
        Dim pptCurrentPresentation As pptNS.Presentation = Nothing
        Dim pptTemplatePresentation As pptNS.Presentation = Nothing
        Dim pptSlide As pptNS.Slide = Nothing
        Dim shapeRange As pptNS.ShapeRange = Nothing
        Dim presentationFile As String = awinPath & requirementsOrdner & "boarddossier.pptx"
        Dim presentationFileH As String = awinPath & requirementsOrdner & "boarddossier_Hochformat.pptx"
        Dim newFileName As String = reportOrdnerName & "MP Report.pptx"

        Dim pptShape As pptNS.Shape
        Dim portfolioName As String = currentConstellation
        Dim top As Double, left As Double, width As Double, height As Double
        Dim htop As Double, hleft As Double, hwidth As Double, hheight As Double
        Dim pptSize As Single = 18
        Dim zeilenhoehe As Double = 0.0
        Dim legendFontSize As Single = 0.0

        Dim von As Integer, bis As Integer
        Dim myCollection As New Collection
        Dim notYetDone As Boolean = False
        Dim listofShapes As New Collection


        Try
            ' prüft, ob bereits Powerpoint geöffnet ist 
            pptApp = CType(GetObject(, "PowerPoint.Application"), pptNS.Application)
        Catch ex As Exception
            Try
                pptApp = CType(CreateObject("PowerPoint.Application"), pptNS.Application)
            Catch ex1 As Exception
                Call MsgBox("Powerpoint konnte nicht gestartet werden ..." & ex1.Message)
                Exit Sub
            End Try

        End Try


        'frmSelectPPTTempl.statusNotification.Text = "PowerPoint nun geöffnet ...."
        e.Result = "PowerPoint ist nun geöffnet ...."
        If worker.WorkerReportsProgress Then
            worker.ReportProgress(0, e)
        End If

        ' jetzt wird das template geöffnet , um festzustellen , welches Format Quer oder Hoch die Vorlage hat 
        ' und dann wird die entsprechende Titelblatt Präsentation geöffnet 
        Try

            If pptApp.Presentations.Count = 0 Then

                pptTemplatePresentation = pptApp.Presentations.Open(pptTemplateName)
                If pptTemplatePresentation.PageSetup.SlideOrientation = MsoOrientation.msoOrientationHorizontal Then
                    pptCurrentPresentation = pptApp.Presentations.Open(presentationFile)
                Else
                    pptCurrentPresentation = pptApp.Presentations.Open(presentationFileH)
                End If

            Else
                pptCurrentPresentation = pptApp.ActivePresentation
                pptTemplatePresentation = pptApp.Presentations.Open(pptTemplateName)

                If pptTemplatePresentation.PageSetup.SlideOrientation = pptCurrentPresentation.PageSetup.SlideOrientation And _
                    pptTemplatePresentation.PageSetup.SlideSize = pptCurrentPresentation.PageSetup.SlideSize Then
                    ' also in Ordnung, es kann weiter in die Current Presentation geschrieben werden ... 
                Else
                    ' jetzt muss geprüft werden, ob die aktuelle Präsentation genauso heisst wie die zu öffnende ..
                    ' wenn ja, wird beendet - der User bekommt die Aufforderung die aktuelle Präsentation erst zu speichern  
                    Try
                        ' jetzt wird die entsprechende Template Präsentation geöffnet 
                        If pptTemplatePresentation.PageSetup.SlideOrientation = MsoOrientation.msoOrientationHorizontal Then
                            pptCurrentPresentation = pptApp.Presentations.Open(presentationFile)
                        Else
                            pptCurrentPresentation = pptApp.Presentations.Open(presentationFileH)
                        End If

                    Catch ex As Exception
                        ' in diesem Fall existiert schon eine geöffnete BoardDossier, allerdings mit anderem Format ...

                        pptTemplatePresentation.Saved = True
                        pptTemplatePresentation.Close()

                        e.Result = "Abbruch ... bitte speichern und schliessen Sie die offenen Präsentationen ... "
                        If worker.WorkerReportsProgress Then
                            worker.ReportProgress(0, e)
                        End If

                        Exit Sub

                    End Try

                End If
            End If


        Catch ex As Exception
            e.Result = "Abbruch ... bitte speichern und schliessen Sie die offenen Präsentationen ... "
            If worker.WorkerReportsProgress Then
                worker.ReportProgress(0, e)
            End If

            Exit Sub
        End Try


        ' jetzt ist sichergestellt, daß die Vorlage geöffnet ist: pptTemplatePresentation
        ' und das Titelblatt: 

        Dim anzSlidesToAdd As Integer
        Dim anzahlCurrentSlides As Integer
        Dim currentInsert As Integer = 1

        Try

            ' löschen, wenn der Name bereits existiert ...
            If My.Computer.FileSystem.FileExists(newFileName) And _
                pptCurrentPresentation.Name <> "MP Report.pptx" Then

                Try
                    My.Computer.FileSystem.DeleteFile(newFileName)
                Catch ex1 As Exception

                End Try

            End If
            ' speichern unter .. , damit Projektdossier nicht überschrieben werden kann 
            pptCurrentPresentation.SaveAs(newFileName)

            anzahlCurrentSlides = pptCurrentPresentation.Slides.Count
            anzSlidesToAdd = pptTemplatePresentation.Slides.Count
            pptTemplatePresentation.Saved = True
            pptTemplatePresentation.Close()

        Catch ex As Exception
            Throw New Exception("Probleme mit Powerpoint Template")
        End Try



        Dim reportObj As xlNS.ChartObject = Nothing
        Dim obj As xlNS.ChartObject = Nothing
        Dim kennzeichnung As String = ""
        Dim qualifier As String = ""
        Dim anzShapes As Integer
        Dim tatsErstellt As Integer = 0
        Dim folieIX As Integer = 1

        ' bei den objectsToDo kann es sich um Swimlanes oder Projekte handeln 
        Dim objectsToDo As Integer = 0
        Dim objectsDone As Integer = 0

        While folieIX <= anzSlidesToAdd

            tatsErstellt = tatsErstellt + 1
            If worker.WorkerSupportsCancellation Then

                If worker.CancellationPending Then
                    e.Cancel = True
                    e.Result = "Berichterstellung nach " & tatsErstellt & " Seiten abgebrochen ..."

                    Exit While
                End If

            End If

            ' jetzt wird eine Seite aus der Vorlage ergänzt 
            Dim tmpIX As Integer
            Dim tmpslideID As Integer

            ' ur:31.03.2015????

            If Not pptFirstTime Then
                '  pptSlide.Delete()
                pptCurrentPresentation.Slides("tmpSav").Copy()
                tmpslideID = pptCurrentPresentation.Slides("tmpSav").SlideID
                pptCurrentPresentation.Slides.Paste(pptCurrentPresentation.Slides.Count + 1)
                pptSlide = pptCurrentPresentation.Slides(pptCurrentPresentation.Slides.Count)


            Else
                anzahlCurrentSlides = pptCurrentPresentation.Slides.Count
                tmpIX = pptCurrentPresentation.Slides.InsertFromFile(FileName:=pptTemplateName, Index:=anzahlCurrentSlides, _
                                                                              SlideStart:=folieIX, SlideEnd:=folieIX)
            End If


            'frmSelectPPTTempl.statusNotification.Text = "Liste der Seiten aufgebaut ...."
            e.Result = "Bericht Seite " & tatsErstellt & " wird aufgebaut ...."

            If worker.WorkerReportsProgress Then
                worker.ReportProgress(0, e)
            End If


            anzahlCurrentSlides = pptCurrentPresentation.Slides.Count
            pptSlide = pptCurrentPresentation.Slides(anzahlCurrentSlides)


            ' jetzt werden die Charts gezeichnet 
            anzShapes = pptSlide.Shapes.Count


            ' jetzt wird die listofShapes aufgebaut - das sind alle Shapes, die ersetzt werden müssen ...
            ' bzw. alle Shapes, die "gemerkt" werden müssen
            For i = 1 To anzShapes
                pptShape = pptSlide.Shapes(i)
                qualifier = ""
                With pptShape



                    Dim tmpStr(3) As String
                    Try

                        If .Title <> "" Then
                            kennzeichnung = .Title
                        Else
                            tmpStr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {CChar("("), CChar(")")}, 3)
                            kennzeichnung = tmpStr(0).Trim
                        End If


                    Catch ex As Exception
                        kennzeichnung = "nicht identifizierbar"
                    End Try

                    If kennzeichnung = "Portfolio-Name" Or _
                        kennzeichnung = "Szenario-Projekt-Tabelle" Or _
                        kennzeichnung = "Legenden-Tabelle" Or _
                        kennzeichnung = "Multiprojektsicht" Or _
                        kennzeichnung = "Projekt-Tafel" Or _
                        kennzeichnung = "Projekt-Tafel Phasen" Or _
                        kennzeichnung = "Tabelle Zielerreichung" Or _
                        kennzeichnung = "Tabelle Projektstatus" Or _
                        kennzeichnung = "Tabelle Projektabhängigkeiten" Or _
                        kennzeichnung = "Übersicht Besser/Schlechter" Or _
                        kennzeichnung = "Tabelle Besser/Schlechter" Or _
                        kennzeichnung = "Tabelle ProjekteMitMsImMonat" Or _
                        kennzeichnung = "Tabelle ProjekteMitPhImMonat" Or _
                        kennzeichnung = "Tabelle ProjekteMitRolleImMonat" Or _
                        kennzeichnung = "Tabelle ProjekteMitKostenartImMonat" Or _
                        kennzeichnung = "Fortschritt Personalkosten" Or _
                        kennzeichnung = "Fortschritt Sonstige Kosten" Or _
                        kennzeichnung = "Fortschritt Gesamtkosten" Or _
                        kennzeichnung = "Fortschritt Rolle" Or _
                        kennzeichnung = "Fortschritt Kostenart" Or _
                        kennzeichnung = "Übersicht Budget" Or _
                        kennzeichnung = "Ergebnis Verbesserungspotential" Or _
                        kennzeichnung = "Ergebnis" Or _
                        kennzeichnung = "Strategie/Risiko/Marge" Or _
                        kennzeichnung = "Strategie/Risiko/Volumen" Or _
                        kennzeichnung = "Zeit/Risiko/Volumen" Or _
                        kennzeichnung = "Strategie/Risiko/Ausstrahlung" Or _
                        kennzeichnung = "Übersicht Auslastung" Or _
                        kennzeichnung = "Details Unterauslastung" Or _
                        kennzeichnung = "Details Überauslastung" Or _
                        kennzeichnung = "Bisherige Zielerreichung" Or _
                        kennzeichnung = "Prognose Zielerreichung" Or _
                        kennzeichnung = "Phase" Or _
                        kennzeichnung = "Rolle" Or _
                        kennzeichnung = "Kostenart" Or _
                        kennzeichnung = "Meilenstein" Or _
                        kennzeichnung = "Stand:" Or _
                        kennzeichnung = "Zeitraum:" Then

                        listofShapes.Add(pptShape)

                    End If



                End With
            Next



            Dim newShapeRange As pptNS.ShapeRange
            Dim newShape As pptNS.Shape

            Dim boxName As String


            For Each tmpShape As pptNS.Shape In listofShapes

                Dim tmpanz As Integer = listofShapes.Count
                pptShape = tmpShape
                qualifier = ""
                kennzeichnung = ""
                With pptShape
                    .Name = "Shape" & .Id.ToString
                    Dim tmpStr(3) As String
                    Try

                        If .Title <> "" Then
                            kennzeichnung = .Title
                            qualifier = .AlternativeText
                            boxName = kennzeichnung
                        Else
                            tmpStr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {CChar("("), CChar(")")}, 3)
                            kennzeichnung = tmpStr(0).Trim
                            boxName = .TextFrame2.TextRange.Text
                            If tmpStr.Length > 1 Then
                                Try
                                    qualifier = tmpStr(1)
                                Catch ex2 As Exception
                                    qualifier = ""
                                End Try
                            End If
                        End If

                    Catch ex As Exception
                        kennzeichnung = "nicht identifizierbar"
                        boxName = " "
                    End Try

                    ' Fortschrittsmeldung im Formular SelectPPTTempl

                    'frmSelectPPTTempl.statusNotification.Text = "Liste der Seiten aufgebaut ...."
                    e.Result = "Chart '" & kennzeichnung & "' wird aufgebaut ...."
                    If worker.WorkerReportsProgress Then
                        worker.ReportProgress(0, e)
                    End If



                    reportObj = Nothing
                    top = .Top
                    left = .Left
                    height = .Height
                    width = .Width

                    Dim nameList As New Collection

                    Select Case kennzeichnung

                        Case "Legenden-Tabelle"

                            Try
                                Call prepZeichneLegendenTabelle(pptSlide, pptShape, legendFontSize, selectedPhases, selectedMilestones)
                            Catch ex As Exception

                            End Try


                        Case "Multiprojektsicht"

                            Try
                                Dim tmpProjekt As New clsProjekt
                                Call zeichneMultiprojektSicht(pptApp, pptCurrentPresentation, pptSlide, _
                                                              objectsToDo, objectsDone, pptFirstTime, zeilenhoehe, legendFontSize, _
                                                              selectedPhases, selectedMilestones, _
                                                              selectedRoles, selectedCosts, _
                                                              selectedBUs, selectedTyps, _
                                                              worker, e, True, tmpProjekt, kennzeichnung)
                                .TextFrame2.TextRange.Text = ""
                                .ZOrder(MsoZOrderCmd.msoSendToBack)
                            Catch ex As Exception
                                .TextFrame2.TextRange.Text = ex.Message
                                objectsDone = objectsToDo
                            End Try

                        Case "Szenario-Projekt-Tabelle"

                            Try
                                Call zeichneSzenarioTabelle(pptShape, pptSlide)
                            Catch ex As Exception
                                ' in einer Tabelle führt der folgende Befehl zu einem Fehler 
                                '.TextFrame2.TextRange.Text = ex.Message
                            End Try


                        Case "Tabelle ProjekteMitMsImMonat"


                            myCollection.Clear()
                            myCollection = buildNameCollection(PTpfdk.Meilenstein, qualifier, selectedMilestones)

                            Try
                                Call zeichneTabelleProjekteMitElemImMonat(pptShape, pptSlide, myCollection, DiagrammTypen(5))
                            Catch ex As Exception
                                ' in einer Tabelle führt der folgende Befehl zu einem Fehler 
                                '.TextFrame2.TextRange.Text = ex.Message
                            End Try

                        Case "Tabelle ProjekteMitPhImMonat"


                            myCollection.Clear()
                            myCollection = buildNameCollection(PTpfdk.Phasen, qualifier, selectedPhases)

                            Try
                                Call zeichneTabelleProjekteMitElemImMonat(pptShape, pptSlide, myCollection, DiagrammTypen(0))
                            Catch ex As Exception
                                ' in einer Tabelle führt der folgende Befehl zu einem Fehler 
                                '.TextFrame2.TextRange.Text = ex.Message
                            End Try


                        Case "Tabelle ProjekteMitRolleImMonat"

                            myCollection.Clear()
                            myCollection = buildNameCollection(PTpfdk.Rollen, qualifier, selectedRoles)

                            Try
                                Call zeichneTabelleProjekteMitElemImMonat(pptShape, pptSlide, myCollection, DiagrammTypen(1))
                            Catch ex As Exception
                                ' in einer Tabelle führt der folgende Befehl zu einem Fehler 
                                '.TextFrame2.TextRange.Text = ex.Message
                            End Try

                        Case "Tabelle ProjekteMitKostenartImMonat"

                            myCollection.Clear()
                            myCollection = buildNameCollection(PTpfdk.Kosten, qualifier, selectedCosts)

                            Try
                                Call zeichneTabelleProjekteMitElemImMonat(pptShape, pptSlide, myCollection, DiagrammTypen(2))
                            Catch ex As Exception
                                ' in einer Tabelle führt der folgende Befehl zu einem Fehler 
                                '.TextFrame2.TextRange.Text = ex.Message
                            End Try

                        Case "Portfolio-Name"
                            .TextFrame2.TextRange.Text = portfolioName


                        Case "Projekt-Tafel"


                            Dim farbtyp As Integer
                            Dim rng As xlNS.Range
                            Dim colorrng As xlNS.Range
                            Dim selectionType As Integer = -1 ' keine Einschränkung
                            Dim formerSetting As Boolean = awinSettings.mppShowAmpel
                            awinSettings.mppShowAmpel = True

                            von = showRangeLeft
                            bis = showRangeRight
                            myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

                            If myCollection.Count > 0 Then
                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "

                                Dim minColumn As Integer = 10000, maxColumn As Integer = 0, maxzeile As Integer = 0

                                ' wenn nur die Projekt-Tafel und Zeitraum im großen Bild gezeigt werden soll, also ohne Qualifier aufgerufen wird , 
                                ' dann stelle die Projekt-Tafel mit den Projekten dar, ohne sie abzuschneiden ...
                                If qualifier = "" Then
                                    Call calcPictureCoord(myCollection, minColumn, maxColumn, maxzeile, False)
                                Else
                                    Call calcPictureCoord(myCollection, minColumn, maxColumn, maxzeile, True)
                                End If


                                ' set Gridlines to white 
                                With appInstance.ActiveWindow
                                    .GridlineColor = RGB(255, 255, 255)
                                End With

                                With CType(appInstance.Worksheets(arrWsNames(3)), xlNS.Worksheet)



                                    rng = CType(.Range(.Cells(1, minColumn), .Cells(maxzeile, maxColumn)), xlNS.Range)
                                    colorrng = CType(.Range(.Cells(2, showRangeLeft), .Cells(maxzeile, showRangeRight)), xlNS.Range)

                                    If Not awinSettings.showTimeSpanInPT Then

                                        Try
                                            colorrng.Interior.Color = showtimezone_color
                                        Catch ex1 As Exception

                                        End Try
                                    End If

                                    ' hier werden die Milestones gezeichnet 
                                    If qualifier = "Milestones R" Then
                                        Call awinDeleteProjectChildShapes(0)

                                        farbtyp = 3
                                        Call awinZeichneMilestones(nameList, farbtyp, True, False)



                                    ElseIf qualifier = "Milestones GR" Then
                                        Call awinDeleteProjectChildShapes(0)

                                        farbtyp = 2
                                        Call awinZeichneMilestones(nameList, farbtyp, False, False)
                                        farbtyp = 3
                                        Call awinZeichneMilestones(nameList, farbtyp, False, False)

                                    ElseIf qualifier = "Milestones GGR" Then
                                        Call awinDeleteProjectChildShapes(0)

                                        farbtyp = 1
                                        Call awinZeichneMilestones(nameList, farbtyp, False, False)
                                        farbtyp = 2
                                        Call awinZeichneMilestones(nameList, farbtyp, False, False)
                                        farbtyp = 3
                                        Call awinZeichneMilestones(nameList, farbtyp, False, False)

                                    ElseIf qualifier = "Milestones ALL" Then
                                        Call awinDeleteProjectChildShapes(0)

                                        farbtyp = 4
                                        Call awinZeichneMilestones(nameList, farbtyp, False, False)

                                    ElseIf qualifier = "Status" Then
                                        Call awinDeleteProjectChildShapes(0)
                                        Call awinZeichneStatus(True)
                                    End If


                                    rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
                                    newShapeRange = pptSlide.Shapes.Paste
                                    newShape = newShapeRange.Item(1)

                                    If Not awinSettings.showTimeSpanInPT Then
                                        colorrng.Interior.ColorIndex = -4142
                                    End If


                                    ' lösche alle Milestones wieder 
                                    If qualifier <> "" Then
                                        Call awinDeleteProjectChildShapes(0)
                                    End If
                                End With


                                ' set back 
                                With appInstance.ActiveWindow
                                    .GridlineColor = RGB(220, 220, 220)
                                End With



                                Dim ratio As Double
                                ratio = height / width


                                With newShape

                                    ratio = height / width

                                    If ratio < .Height / .Width Then
                                        ' orientieren an width 
                                        .Width = CSng(width * 0.96)
                                        .Height = CSng(ratio * .Width)
                                        ' left anpassen
                                        .Top = CSng(top + 0.02 * height)
                                        .Left = CSng(left + 0.98 * (width - .Width) / 2)

                                    Else
                                        .Height = CSng(height * 0.96)
                                        .Width = CSng(.Height / ratio)
                                        ' top anpassen 
                                        .Left = CSng(left + 0.02 * width)
                                        .Top = CSng(top + 0.98 * (height - .Height) / 2)
                                    End If

                                End With


                            Else
                                .TextFrame2.TextRange.Text = "Keine Projekte im angegebenen Zeitraum vorhanden"
                            End If


                            awinSettings.mppShowAmpel = formerSetting

                        Case "Projekt-Tafel Phasen"

                            Dim rng As xlNS.Range
                            Dim colorrng As xlNS.Range
                            Dim selectionType As Integer = -1 ' keine Einschränkung
                            Dim ok As Boolean = True

                            von = showRangeLeft
                            bis = showRangeRight
                            If von < 0 Or bis < 0 Then
                                .TextFrame2.TextRange.Text = " bitte geben Sie einen Zeitraum an ..."
                            Else
                                myCollection = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

                                If myCollection.Count > 0 Then
                                    pptSize = .TextFrame2.TextRange.Font.Size
                                    .TextFrame2.TextRange.Text = " "

                                    Dim minColumn As Integer = 10000, maxColumn As Integer = 0, maxzeile As Integer = 0

                                    Call calcPictureCoord(myCollection, minColumn, maxColumn, maxzeile, True)


                                    ' set Gridlines to white 
                                    With appInstance.ActiveWindow
                                        .GridlineColor = RGB(255, 255, 255)
                                    End With

                                    With CType(appInstance.Worksheets(arrWsNames(3)), xlNS.Worksheet)
                                        rng = CType(.Range(.Cells(1, minColumn), .Cells(maxzeile, maxColumn)), xlNS.Range)
                                        colorrng = CType(.Range(.Cells(2, showRangeLeft), .Cells(maxzeile, showRangeRight)), xlNS.Range)

                                        If Not awinSettings.showTimeSpanInPT Then

                                            Try
                                                colorrng.Interior.Color = showtimezone_color
                                            Catch ex1 As Exception

                                            End Try
                                        End If

                                        ' hier werden die Phasen gezeichnet 
                                        Call awinDeleteProjectChildShapes(0)

                                        Dim qstr(20) As String
                                        Dim phNameCollection As New Collection
                                        Dim phName As String = " "
                                        qstr = qualifier.Trim.Split(New Char() {CChar("#")}, 18)

                                        ' Aufbau der Collection 
                                        For i = 0 To qstr.Length - 1

                                            Try
                                                phName = qstr(i).Trim
                                                If PhaseDefinitions.Contains(phName) Then
                                                    phNameCollection.Add(phName, phName)
                                                End If
                                            Catch ex As Exception
                                                Call MsgBox("Fehler: Phasen Name " & phName & " konnte nicht erkannt werden ...")
                                            End Try

                                        Next

                                        Call awinDeSelect()

                                        If phNameCollection.Count > 0 Then

                                            Call awinZeichnePhasen(phNameCollection, False, True)
                                            rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)

                                            If Not awinSettings.showTimeSpanInPT Then

                                                Try
                                                    colorrng.Interior.ColorIndex = -4142
                                                Catch ex1 As Exception

                                                End Try
                                            End If


                                            ' lösche alle Phase Shapes wieder wieder 
                                            If qualifier <> "" Then
                                                Call awinDeleteProjectChildShapes(0)
                                            End If

                                        Else
                                            ok = False
                                        End If

                                    End With

                                    ' set back 
                                    With appInstance.ActiveWindow
                                        .GridlineColor = RGB(220, 220, 220)
                                    End With

                                    If ok Then
                                        newShapeRange = pptSlide.Shapes.Paste

                                        Dim ratio As Double
                                        ratio = height / width

                                        With newShapeRange.Item(1)

                                            If ratio < .Height / .Width Then
                                                ' orientieren an width 
                                                .Width = CSng(width * 0.96)
                                                .Height = CSng(ratio * .Width)
                                                ' left anpassen
                                                .Top = CSng(top + 0.02 * height)
                                                .Left = CSng(left + 0.98 * (width - .Width) / 2)

                                            Else
                                                .Height = CSng(height * 0.96)
                                                .Width = CSng(.Height / ratio)
                                                ' top anpassen 
                                                .Left = CSng(left + 0.02 * width)
                                                .Top = CSng(top + 0.98 * (height - .Height) / 2)
                                            End If

                                        End With
                                    Else
                                        .TextFrame2.TextRange.Text = "es konnten keine Phasen erkannt werden ... "
                                    End If


                                Else
                                    .TextFrame2.TextRange.Text = "Keine Projekte im angegebenen Zeitraum vorhanden"
                                End If
                            End If



                        Case "Tabelle Zielerreichung"

                            Dim farbtyp As Integer
                            Try

                                If qualifier = "rot" Then
                                    farbtyp = 3
                                ElseIf qualifier = "gelb" Then
                                    farbtyp = 2
                                ElseIf qualifier = "gruen" Then
                                    farbtyp = 1
                                ElseIf qualifier = "gelb/rot" Then
                                    farbtyp = 12
                                Else
                                    farbtyp = 3
                                End If
                                Call zeichneTabelleZielErreichung(pptShape, farbtyp)

                            Catch ex As Exception

                            End Try


                        Case "Tabelle Projektstatus"

                            Try
                                Call zeichneTabelleStatus(pptShape)
                            Catch ex As Exception

                            End Try

                        Case "Tabelle Projektabhängigkeiten"

                            Try
                                Call zeichneTabelleProjektabhaengigkeiten(pptShape)
                            Catch ex As Exception

                            End Try


                        Case "Fortschritt Personalkosten"

                            Dim compareToID As Integer = 1
                            Dim auswahl As Integer = 1

                            Call zeichneFortschrittDiagramm(boxName, compareToID, auswahl, qualifier, pptShape, reportObj, notYetDone)


                        Case "Fortschritt Sonstige Kosten"

                            Dim compareToID As Integer = 1 ' Vergleich mit Beauftragung
                            Dim auswahl As Integer = 2 ' Sonstige Kosten 

                            Call zeichneFortschrittDiagramm(boxName, compareToID, auswahl, qualifier, pptShape, reportObj, notYetDone)


                        Case "Fortschritt Gesamtkosten"

                            Dim compareToID As Integer = 1 ' Vergleich mit Beauftragung
                            Dim auswahl As Integer = 3 ' Gesamt Kosten 

                            Call zeichneFortschrittDiagramm(boxName, compareToID, auswahl, qualifier, pptShape, reportObj, notYetDone)


                        Case "Fortschritt Rolle"

                            Dim compareToID As Integer = 1 ' Vergleich mit Beauftragung
                            Dim auswahl As Integer = 4 ' Rolle ; in qualifier steht welche Rolle  

                            Call zeichneFortschrittDiagramm(boxName, compareToID, auswahl, qualifier, pptShape, reportObj, notYetDone)


                        Case "Fortschritt Kostenart"

                            Dim compareToID As Integer = 1 ' Vergleich mit Beauftragung
                            Dim auswahl As Integer = 5 ' Kostenart ; in qualifier steht welche Kostenart  

                            Call zeichneFortschrittDiagramm(boxName, compareToID, auswahl, qualifier, pptShape, reportObj, notYetDone)


                        Case "Ergebnis Verbesserungspotential"


                            boxName = boxName & " (T€)"
                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "

                            htop = 100
                            hleft = 100
                            hwidth = 450
                            hheight = awinSettings.ChartHoehe1
                            obj = Nothing
                            Call awinCreateVerbesserungsPotentialDiagramm(obj, top, left, width, height, False)

                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With

                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try

                        Case "Übersicht Budget"

                            boxName = boxName & " (T€)"
                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "

                            htop = 100
                            hleft = 100
                            hwidth = 450
                            hheight = awinSettings.ChartHoehe1
                            obj = Nothing
                            Call awinCreateBudgetErgebnisDiagramm(obj, htop, hleft, hwidth, hheight, False, True)

                            reportObj = obj

                            With reportObj
                                '.Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With

                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try

                        Case "Ergebnis"

                            boxName = boxName & " (T€)"
                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "

                            htop = 100
                            hleft = 100
                            hwidth = 450
                            hheight = awinSettings.ChartHoehe1
                            obj = Nothing
                            Call awinCreateErgebnisDiagramm(obj, htop, hleft, hwidth, hheight, False, True)

                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With

                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try


                        Case "Strategie/Risiko/Marge"

                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "


                            Dim selectionType As Integer = -1 ' keine Einschränkung
                            von = showRangeLeft
                            bis = showRangeRight
                            myCollection = ShowProjekte.withinTimeFrame(selectionType, von, bis)

                            htop = 50
                            hleft = (showRangeRight - 1) * boxWidth
                            hwidth = 0.4 * maxScreenWidth
                            hheight = 0.6 * maxScreenHeight
                            obj = Nothing

                            If qualifier = "Ampel" Then
                                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisiko, PTpfdk.AmpelFarbe, False, True, True, htop, hleft, hwidth, hheight)
                            Else
                                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisiko, PTpfdk.ProjektFarbe, False, True, True, htop, hleft, hwidth, hheight)
                            End If



                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With

                            'Call awinDeleteChart(reportObj)

                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try

                        Case "Strategie/Risiko/Ausstrahlung"

                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "


                            Dim selectionType As Integer = -1 ' keine Einschränkung
                            von = showRangeLeft
                            bis = showRangeRight
                            myCollection = ShowProjekte.withinTimeFrame(selectionType, von, bis)

                            htop = 50
                            hleft = (showRangeRight - 1) * boxWidth
                            hwidth = 0.4 * maxScreenWidth
                            hheight = 0.6 * maxScreenHeight
                            obj = Nothing

                            If qualifier = "Ampel" Then
                                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoDependency, PTpfdk.AmpelFarbe, False, True, True, htop, hleft, hwidth, hheight)
                            Else
                                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoDependency, PTpfdk.ProjektFarbe, False, True, True, htop, hleft, hwidth, hheight)
                            End If



                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With

                            'Call awinDeleteChart(reportObj)

                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try

                        Case "Strategie/Risiko/Volumen"

                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "


                            Dim selectionType As Integer = -1 ' keine Einschränkung
                            von = showRangeLeft
                            bis = showRangeRight
                            myCollection = ShowProjekte.withinTimeFrame(selectionType, von, bis)

                            htop = 50
                            hleft = (showRangeRight - 1) * boxWidth
                            hwidth = 0.4 * maxScreenWidth
                            hheight = 0.6 * maxScreenHeight
                            obj = Nothing

                            Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoVol, PTpfdk.ProjektFarbe, False, True, True, htop, hleft, hwidth, hheight)

                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With


                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try


                        Case "Zeit/Risiko/Volumen"

                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "


                            Dim selectionType As Integer = -1 ' keine Einschränkung
                            von = showRangeLeft
                            bis = showRangeRight
                            myCollection = ShowProjekte.withinTimeFrame(selectionType, von, bis)

                            htop = 50
                            hleft = (showRangeRight - 1) * boxWidth
                            hwidth = 0.4 * maxScreenWidth
                            hheight = 0.6 * maxScreenHeight
                            obj = Nothing

                            Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.ComplexRisiko, PTpfdk.ProjektFarbe, False, True, True, htop, hleft, hwidth, hheight)
                            'Call awinCreateZeitRiskVolumeDiagramm(myCollection, obj, False, False, True, True, htop, hleft, hwidth, hheight)

                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With

                            'Call awinDeleteChart(reportObj)

                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try


                        Case "Übersicht Besser/Schlechter"

                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "


                            Dim showAbsoluteDiff As Boolean = True
                            Dim isTimeTimeVgl As Boolean = False
                            Dim vglTyp As Integer = 1
                            Dim charttype As Integer = PTpfdk.betterWorseB
                            Dim bubbleValueTyp As Integer = PTbubble.strategicFit
                            Dim showLabels As Boolean = True

                            Dim qstr(20) As String
                            qstr = qualifier.Trim.Split(New Char() {CChar("#")}, 18)

                            ' Bestimmen der Parameter  
                            For i = 0 To qstr.Length - 1

                                Select Case i
                                    Case 0

                                        If qstr(i).Length > 0 Then
                                            showAbsoluteDiff = CBool(qstr(i))
                                        End If

                                    Case 1
                                        If qstr(i).Length > 0 Then
                                            isTimeTimeVgl = CBool(qstr(i))
                                        End If

                                    Case 2
                                        If qstr(i).Length > 0 Then
                                            vglTyp = CInt(qstr(i))
                                        End If

                                    Case 3
                                        If qstr(i).Length > 0 Then

                                            If CBool(qstr(i)) Then
                                                charttype = PTpfdk.betterWorseB
                                            Else
                                                charttype = PTpfdk.betterWorseL
                                            End If

                                        End If

                                    Case 4
                                        If qstr(i).Length > 0 Then
                                            bubbleValueTyp = CInt(qstr(i))
                                        End If

                                    Case 5
                                        If qstr(i).Length > 0 Then
                                            showLabels = CBool(qstr(i))
                                        End If


                                End Select

                            Next


                            Dim selectionType As Integer = PTpsel.lfundab ' nur laufende und abgeschlossene Projekte 
                            von = showRangeLeft
                            bis = showRangeRight
                            myCollection = ShowProjekte.withinTimeFrame(selectionType, von, bis)

                            htop = 50
                            hleft = (showRangeRight - 1) * boxWidth
                            hwidth = 0.4 * maxScreenWidth
                            hheight = 0.6 * maxScreenHeight
                            obj = Nothing


                            Try
                                Call awinCreateBetterWorsePortfolio(ProjektListe:=myCollection, repChart:=obj, showAbsoluteDiff:=showAbsoluteDiff, isTimeTimeVgl:=isTimeTimeVgl, vglTyp:=vglTyp, _
                                                charttype:=charttype, bubbleColor:=0, bubbleValueTyp:=bubbleValueTyp, showLabels:=showLabels, chartBorderVisible:=True, _
                                                top:=htop, left:=hleft, width:=hwidth, height:=hheight)


                                reportObj = obj

                                With reportObj
                                    '.Chart.ChartTitle.Text = boxName
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShapeRange = pptSlide.Shapes.Paste

                                With newShapeRange.Item(1)
                                    .Top = CSng(top + 0.02 * height)
                                    .Left = CSng(left + 0.02 * width)
                                    .Width = CSng(width * 0.96)
                                    .Height = CSng(height * 0.96)
                                End With


                                Try
                                    reportObj.Delete()
                                    'DiagramList.Remove(DiagramList.Count)
                                Catch ex As Exception

                                End Try

                            Catch ex As Exception
                                .TextFrame2.TextRange.Text = ex.Message
                            End Try




                        Case "Übersicht Auslastung"

                            boxName = boxName & " (PT)"

                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "

                            htop = 100
                            hleft = 100
                            hheight = awinSettings.ChartHoehe2
                            hwidth = 340
                            obj = Nothing
                            Call awinCreateAuslastungsDiagramm(obj, htop, hleft, hwidth, hheight, True)

                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With


                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try


                        Case "Details Unterauslastung"

                            boxName = boxName & " (PT)"
                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "

                            htop = 100
                            hleft = 100
                            hheight = awinSettings.ChartHoehe2
                            hwidth = 340
                            obj = Nothing
                            Call createAuslastungsDetailPie(obj, 2, htop, hleft, hheight, hwidth, True)

                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With

                            'Call awinDeleteChart(reportObj)

                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try


                        Case "Details Überauslastung"

                            boxName = boxName & " (PT)"
                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "

                            htop = 100
                            hleft = 100
                            hheight = awinSettings.ChartHoehe2
                            hwidth = 340
                            obj = Nothing
                            Call createAuslastungsDetailPie(obj, 1, htop, hleft, hheight, hwidth, True)

                            reportObj = obj

                            With reportObj
                                .Chart.ChartTitle.Text = boxName
                                .Chart.ChartTitle.Font.Size = pptSize
                            End With

                            reportObj.Copy()
                            newShapeRange = pptSlide.Shapes.Paste

                            With newShapeRange.Item(1)
                                .Top = CSng(top + 0.02 * height)
                                .Left = CSng(left + 0.02 * width)
                                .Width = CSng(width * 0.96)
                                .Height = CSng(height * 0.96)
                            End With

                            'Call awinDeleteChart(reportObj)

                            Try
                                reportObj.Delete()
                                'DiagramList.Remove(DiagramList.Count)
                            Catch ex As Exception

                            End Try



                        Case "Bisherige Zielerreichung"

                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "

                            htop = 100
                            hleft = 100
                            hheight = awinSettings.ChartHoehe2
                            hwidth = 340
                            obj = Nothing


                            Try

                                Call awinCreateZielErreichungsDiagramm(obj, -1, htop, hleft, hheight, hwidth, False, True)

                                reportObj = obj

                                With reportObj
                                    '.Chart.ChartTitle.Text = boxName
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShapeRange = pptSlide.Shapes.Paste

                                With newShapeRange.Item(1)
                                    .Top = CSng(top + 0.02 * height)
                                    .Left = CSng(left + 0.02 * width)
                                    .Width = CSng(width * 0.96)
                                    .Height = CSng(height * 0.96)
                                End With

                                'Call awinDeleteChart(reportObj)

                                Try
                                    reportObj.Delete()
                                    'DiagramList.Remove(DiagramList.Count)
                                Catch ex As Exception

                                End Try

                            Catch ex As Exception

                                .TextFrame2.TextRange.Text = ex.Message

                            End Try



                        Case "Prognose Zielerreichung"

                            pptSize = .TextFrame2.TextRange.Font.Size
                            .TextFrame2.TextRange.Text = " "

                            htop = 100
                            hleft = 100
                            hheight = awinSettings.ChartHoehe2
                            hwidth = 340
                            obj = Nothing


                            Try

                                Call awinCreateZielErreichungsDiagramm(obj, 1, htop, hleft, hheight, hwidth, False, True)

                                reportObj = obj

                                With reportObj
                                    '.Chart.ChartTitle.Text = boxName
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShapeRange = pptSlide.Shapes.Paste

                                With newShapeRange.Item(1)
                                    .Top = CSng(top + 0.02 * height)
                                    .Left = CSng(left + 0.02 * width)
                                    .Width = CSng(width * 0.96)
                                    .Height = CSng(height * 0.96)
                                End With



                                Try

                                    reportObj.Delete()

                                Catch ex As Exception

                                End Try
                            Catch ex As Exception

                                .TextFrame2.TextRange.Text = ex.Message

                            End Try


                        Case "Phase"


                            myCollection.Clear()

                            myCollection = buildNameCollection(PTpfdk.Phasen, qualifier, selectedPhases)


                            If myCollection.Count > 0 Then

                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "

                                htop = 100
                                hleft = 100
                                hheight = miniHeight  ' height of all charts
                                hwidth = miniWidth   ' width of all charts
                                obj = Nothing
                                Call awinCreateprcCollectionDiagram(myCollection, obj, htop, hleft, hwidth, hheight, False, DiagrammTypen(0), True)

                                reportObj = obj

                                With reportObj
                                    If myCollection.Count > 1 Then
                                        .Chart.ChartTitle.Text = "Phasen Übersicht"
                                    ElseIf myCollection.Count = 1 Then
                                        .Chart.ChartTitle.Text = "Phase " & CStr(myCollection.Item(1)).Replace("#", "-")
                                    Else
                                        .Chart.ChartTitle.Text = boxName
                                    End If

                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShapeRange = pptSlide.Shapes.Paste

                                With newShapeRange.Item(1)
                                    .Top = CSng(top + 0.02 * height)
                                    .Left = CSng(left + 0.02 * width)
                                    .Width = CSng(width * 0.96)
                                    .Height = CSng(height * 0.96)
                                End With

                                'Call awinDeleteChart(reportObj)
                                ' der Titel wird geändert im Report, deswegen wird das Diagramm  nicht gefunden in awinDeleteChart 

                                Try
                                    reportObj.Delete()
                                    'DiagramList.Remove(DiagramList.Count)
                                Catch ex As Exception

                                End Try

                            Else
                                .TextFrame2.TextRange.Text = "nicht definiert: " & qualifier
                            End If



                        Case "Meilenstein"

                            myCollection.Clear()
                            myCollection = buildNameCollection(PTpfdk.Meilenstein, qualifier, selectedMilestones)


                            If myCollection.Count > 0 Then

                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "

                                htop = 100
                                hleft = 100
                                hheight = miniHeight  ' height of all charts
                                hwidth = miniWidth   ' width of all charts
                                obj = Nothing
                                Call awinCreateprcCollectionDiagram(myCollection, obj, htop, hleft, hwidth, hheight, False, DiagrammTypen(5), True)

                                reportObj = obj

                                With reportObj
                                    If myCollection.Count > 1 Then
                                        .Chart.ChartTitle.Text = "Meilenstein Übersicht"
                                    ElseIf myCollection.Count = 1 Then
                                        .Chart.ChartTitle.Text = "Meilenstein " & CStr(myCollection.Item(1)).Replace("#", "-")
                                    Else
                                        .Chart.ChartTitle.Text = boxName
                                    End If

                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShapeRange = pptSlide.Shapes.Paste

                                With newShapeRange.Item(1)
                                    .Top = CSng(top + 0.02 * height)
                                    .Left = CSng(left + 0.02 * width)
                                    .Width = CSng(width * 0.96)
                                    .Height = CSng(height * 0.96)
                                End With


                                Try
                                    reportObj.Delete()
                                Catch ex As Exception

                                End Try

                            Else
                                .TextFrame2.TextRange.Text = "nicht definiert: " & qualifier
                            End If



                        Case "Rolle"


                            myCollection.Clear()
                            myCollection = buildNameCollection(PTpfdk.Rollen, qualifier, selectedRoles)


                            If myCollection.Count > 0 Then

                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "

                                htop = 100
                                hleft = 100
                                hheight = miniHeight  ' height of all charts
                                hwidth = miniWidth   ' width of all charts
                                obj = Nothing
                                Call awinCreateprcCollectionDiagram(myCollection, obj, htop, hleft, hwidth, hheight, False, DiagrammTypen(1), True)

                                reportObj = obj
                                ' jetzt wird die Größe der Überschrift neu bestimmt ...
                                With reportObj
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShapeRange = pptSlide.Shapes.Paste

                                With newShapeRange.Item(1)
                                    .Top = CSng(top + 0.02 * height)
                                    .Left = CSng(left + 0.02 * width)
                                    .Width = CSng(width * 0.96)
                                    .Height = CSng(height * 0.96)
                                End With

                                'Call awinDeleteChart(reportObj)
                                ' der Titel wird geändert im Report, deswegen wird das Diagramm  nicht gefunden in awinDeleteChart 

                                Try
                                    reportObj.Delete()
                                Catch ex As Exception

                                End Try

                            Else
                                .TextFrame2.TextRange.Text = "nicht definiert: " & qualifier
                            End If



                        Case "Kostenart"


                            myCollection.Clear()
                            myCollection = buildNameCollection(PTpfdk.Kosten, qualifier, selectedCosts)


                            If myCollection.Count > 0 Then

                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "

                                htop = 100
                                hleft = 100
                                hheight = miniHeight  ' height of all charts
                                hwidth = miniWidth   ' width of all charts
                                obj = Nothing
                                Call awinCreateprcCollectionDiagram(myCollection, obj, htop, hleft, hwidth, hheight, False, DiagrammTypen(2), True)

                                reportObj = obj

                                With reportObj
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShapeRange = pptSlide.Shapes.Paste

                                With newShapeRange.Item(1)
                                    .Top = CSng(top + 0.02 * height)
                                    .Left = CSng(left + 0.02 * width)
                                    .Width = CSng(width * 0.96)
                                    .Height = CSng(height * 0.96)
                                End With

                                'Call awinDeleteChart(reportObj)
                                ' der Titel wird geändert im Report, deswegen wird das Diagramm  nicht gefunden in awinDeleteChart 

                                Try
                                    reportObj.Delete()
                                    'DiagramList.Remove(DiagramList.Count)
                                Catch ex As Exception

                                End Try

                            Else
                                .TextFrame2.TextRange.Text = "nicht definiert: " & qualifier
                            End If


                        Case "Stand:"
                            .TextFrame2.TextRange.Text = Date.Now.ToString("d")


                        Case "Zeitraum:"
                            .TextFrame2.TextRange.Text = textZeitraum(showRangeLeft, showRangeRight)

                        Case Else

                    End Select


                    If notYetDone Then

                        pptSize = .TextFrame2.TextRange.Font.Size
                        .TextFrame2.TextRange.Text = " "

                        If Not reportObj Is Nothing Then
                            Try
                                With reportObj
                                    .Chart.ChartTitle.Text = boxName
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShapeRange = pptSlide.Shapes.Paste

                                With newShapeRange.Item(1)
                                    .Top = CSng(top + 0.02 * height)
                                    .Left = CSng(left + 0.02 * width)
                                    .Width = CSng(width * 0.96)
                                    .Height = CSng(height * 0.96)
                                End With

                                reportObj.Delete()
                            Catch ex As Exception

                            End Try
                        Else
                            .TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                        End If

                        notYetDone = False

                    End If


                End With



            Next

            listofShapes.Clear()
            If objectsDone >= objectsToDo Or awinSettings.mppOnePage Then
                folieIX = folieIX + 1
                pptFirstTime = True         ' für die LegendenVorlage
                objectsToDo = 0
                objectsDone = 0

            End If
            ' Next 

        End While
        ' hier muss die While Schleife beendet werden 

        ' pptTemplate muss noch geschlossen werden

        If tatsErstellt = 1 Then
            e.Result = " Report mit " & tatsErstellt & " Seite erstellt !"
        Else
            e.Result = " Report mit " & tatsErstellt & " Seiten erstellt !"
        End If

        If worker.WorkerReportsProgress Then
            worker.ReportProgress(0, e)
        End If

        ' Vorlage in passender Größe wird nun nicht mehr benötigt
        Try
            pptCurrentPresentation.Slides("tmpSav").Delete()
        Catch ex As Exception

        End Try

    End Sub



    Public Sub StoreAllProjectsinDB()

        Dim jetzt As Date = Now
        Dim zeitStempel As Date
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        enableOnUpdate = False

        ' die aktuelle Konstellation wird unter dem Namen <Last> gespeichert ..
        Call storeSessionConstellation(ShowProjekte, "Last")

        If request.pingMongoDb() Then

            Try
                ' jetzt werden die gezeigten Projekte in die Datenbank geschrieben 

                For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte.liste

                    Try
                        ' hier wird der Wert für kvp.Value.timeStamp = heute gesetzt 

                        If demoModusHistory Then
                            kvp.Value.timeStamp = historicDate
                        Else
                            kvp.Value.timeStamp = jetzt
                        End If

                        If request.storeProjectToDB(kvp.Value) Then
                        Else
                            Call MsgBox("Fehler in Schreiben Projekt " & kvp.Key)
                        End If
                    Catch ex As Exception

                        ' Call MsgBox("Fehler beim Speichern der Projekte in die Datenbank. Datenbank nicht aktiviert?")
                        Throw New ArgumentException("Fehler beim Speichern der Projekte in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
                        'Exit Sub
                    End Try

                Next

                historicDate = historicDate.AddMonths(1)

                ' jetzt werden alle definierten Constellations weggeschrieben

                For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

                    Try
                        If request.storeConstellationToDB(kvp.Value) Then
                        Else
                            Call MsgBox("Fehler in Schreiben Constellation " & kvp.Key)
                        End If
                    Catch ex As Exception
                        Throw New ArgumentException("Fehler beim Speichern der Portfolios in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
                        'Call MsgBox("Fehler beim Speichern der ProjekteConstellationen in die Datenbank. Datenbank nicht aktiviert?")
                        'Exit Sub
                    End Try

                Next


                ' jetzt werden alle Abhängigkeiten weggeschreiben 

                For Each kvp As KeyValuePair(Of String, clsDependenciesOfP) In allDependencies.getSortedList

                    Try
                        If request.storeDependencyofPToDB(kvp.Value) Then
                        Else
                            Call MsgBox("Fehler in Schreiben Dependency " & kvp.Key)
                        End If
                    Catch ex As Exception
                        Throw New ArgumentException("Fehler beim Speichern der Abhängigkeiten in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
                        'Call MsgBox("Fehler beim Speichern der Abhängigkeiten in die Datenbank. Datenbank nicht aktiviert?")
                        'Exit Sub
                    End Try


                Next

                zeitStempel = AlleProjekte.First.timeStamp

                Call MsgBox("ok, gespeichert!" & vbLf & zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString)

                ' Änderung 18.6 - wenn gespeichert wird, soll die Projekthistorie zurückgesetzt werden 
                Try
                    If projekthistorie.Count > 0 Then
                        projekthistorie.clear()
                    End If
                Catch ex As Exception

                End Try

            Catch ex As Exception
                Throw New ArgumentException("Fehler beim Speichern der Projekte in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
                'Call MsgBox(" Fehler beim Speichern in die Datenbank")
            End Try
        Else

            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen")

        End If

        enableOnUpdate = True

    End Sub

    Public Function StoreSelectedProjectsinDB() As Integer

        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt
        Dim hilfshproj As clsProjekt
        Dim jetzt As Date = Now
        Dim zeitStempel As Date
        Dim anzSelectedProj As Integer = 0
        Dim anzStoredProj As Integer = 0
        Dim variantCollection As Collection

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        Dim awinSelection As Excel.ShapeRange

        enableOnUpdate = False

        Try
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, Excel.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        If request.pingMongoDb() Then

            If Not awinSelection Is Nothing Then

                anzSelectedProj = awinSelection.Count

                For i = 1 To awinSelection.Count

                    ' jetzt die Aktion durchführen ...
                    singleShp1 = awinSelection.Item(i)

                    Try
                        hilfshproj = ShowProjekte.getProject(singleShp1.Name, True)

                    Catch ex As Exception
                        Throw New ArgumentException("Projekt nicht gefunden ...")
                        enableOnUpdate = True
                    End Try

                    ' alle geladenen Variante in variantCollection holen
                    variantCollection = AlleProjekte.getVariantNames(hilfshproj.name, True)

                    For vi = 1 To variantCollection.Count

                        Dim hVname As String
                        Dim tmpStr(5) As String
                        Dim trennzeichen1 As String = "("
                        Dim trennzeichen2 As String = ")"

                        ' VariantenNamen von den () befreien
                        tmpStr = variantCollection(vi).Split(New Char() {CChar(trennzeichen1)}, 4)
                        tmpStr = tmpStr(1).Split(New Char() {CChar(trennzeichen2)}, 4)
                        hVname = tmpStr(0)

                        ' gesamte ProjektInfo der Variante aus Liste AlleProjekte lesen
                        hproj = AlleProjekte.getProject(calcProjektKey(hilfshproj.name, hVname))

                        Try
                            ' hier wird der Wert für kvp.Value.timeStamp = heute gesetzt 

                            If demoModusHistory Then
                                hproj.timeStamp = historicDate
                            Else
                                hproj.timeStamp = jetzt
                            End If

                            If request.storeProjectToDB(hproj) Then

                                anzStoredProj = anzStoredProj + 1
                                'Call MsgBox("ok, Projekt '" & hproj.name & "' gespeichert!" & vbLf & hproj.timeStamp.ToShortDateString)
                            Else
                                Call MsgBox("Fehler in Schreiben Projekt " & hproj.name)
                            End If
                        Catch ex As Exception

                            ' Call MsgBox("Fehler beim Speichern der Projekte in die Datenbank. Datenbank nicht aktiviert?")
                            Throw New ArgumentException("Fehler beim Speichern der Projekte in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
                            'Exit Sub
                        End Try

                    Next vi

                Next i

            Else
                'Call MsgBox("Es wurde kein Projekt selektiert")
                ' die Anzahl selektierter und auch gespeicherter Projekte ist damit = 0
                anzStoredProj = anzSelectedProj
                Return anzSelectedProj
            End If


        Else

            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen")

        End If


        enableOnUpdate = True

        If AlleProjekte.Count > 0 Then
            zeitStempel = AlleProjekte.First.timeStamp
        End If


        Call MsgBox("ok, " & anzStoredProj & " Projekte und Varianten gespeichert!" & vbLf & zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString)
        Return anzStoredProj

    End Function


    ' ''' <summary>
    ' ''' alte Version: wird mit der Änderung vom 19.10 zum Löschen von Projekten, Varianten, Snapshots nicht mehr benötigt 
    ' ''' </summary>
    ' ''' <param name="selectedToDelete"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function RemoveSelectedProjectsfromDB(ByRef selectedToDelete As clsProjektDBInfos) As Integer


    '    Dim hproj As New clsProjekt
    '    Dim jetzt As Date = Date.Now
    '    'Dim zeitStempel As Date
    '    Dim anzSelectedProj As Integer = 0
    '    Dim anzDeletedProj As Integer = 0
    '    Dim anzDeletedTS As Integer = 0
    '    Dim anzElements As Integer
    '    Dim found As Boolean = False
    '    Dim iSel As Integer = 0
    '    Dim key As String

    '    Dim selCollection As SortedList(Of Date, String)
    '    enableOnUpdate = False
    '    Dim tmpstr(4) As String

    '    Dim request As New Request(awinSettings.databaseName)
    '    Dim requestTrash As New Request(awinSettings.databaseName & "Trash")

    '    If request.pingMongoDb() Then

    '        If selectedToDelete.Count > 0 Then

    '            anzSelectedProj = selectedToDelete.Count


    '            For Each kvpSelToDel As KeyValuePair(Of String, SortedList(Of Date, String)) In selectedToDelete.Liste

    '                selCollection = selectedToDelete.getTimeStamps(kvpSelToDel.Key)
    '                anzElements = selCollection.Count

    '                'If AlleProjekte.ContainsKey(kvpSelToDel.Key) Then
    '                '    ' Projekt ist bereits im Hauptspeicher geladen
    '                '    hproj = AlleProjekte(kvpSelToDel.Key)
    '                'End If

    '                If Not projekthistorie Is Nothing Then
    '                    projekthistorie.clear() ' alte Historie löschen
    '                Else
    '                    projekthistorie = New clsProjektHistorie
    '                End If

    '                'tmpstr = title.Trim.Split(New Char() {"#"}, 4)
    '                tmpstr = kvpSelToDel.Key.Trim.Split(New Char() {CChar("#")}, 4)   ' Projektnamen aus key separieren

    '                projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=tmpstr(0), variantName:="", storedEarliest:=Date.MinValue, storedLatest:=Date.Now)

    '                anzDeletedTS = 0    ' Anzahl gelöschter TimeStamps dieses Projekts

    '                For i = 1 To anzElements  ' Schleife über die zu löschenden TimeStamps dieses Projekts

    '                    'Dim ms As Long = selCollection.ElementAt(i - 1).Key.Millisecond

    '                    found = False
    '                    iSel = 0

    '                    While Not found
    '                        hproj = projekthistorie.ElementAt(iSel)
    '                        If hproj.timeStamp = selCollection.ElementAt(i - 1).Key Then
    '                            found = True
    '                        End If
    '                        iSel = iSel + 1
    '                    End While

    '                    If requestTrash.storeProjectToDB(hproj) Then

    '                        If request.deleteProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
    '                                                                     storedEarliest:=selCollection.ElementAt(i - 1).Key, storedLatest:=selCollection.ElementAt(i - 1).Key) Then
    '                            anzDeletedTS = anzDeletedTS + 1

    '                        Else
    '                            Call MsgBox("Fehler beim Löschen von " & hproj.name)
    '                        End If

    '                    Else
    '                        Call MsgBox("Fehler beim Speichern von " & hproj.name & " im Papierkorb")
    '                    End If

    '                Next i      'nächsten TimeStamp holen


    '                Call MsgBox("ok, " & anzDeletedTS & " TimeStamps zu Projekt " & hproj.name & " gelöscht")

    '                key = calcProjektKey(hproj)
    '                If Not request.projectNameAlreadyExists(hproj.name, hproj.variantName) Then
    '                    If AlleProjekte.Containskey(key) Then
    '                        AlleProjekte.Remove(key)
    '                        Try
    '                            ShowProjekte.Remove(hproj.name)
    '                        Catch ex As Exception
    '                        End Try
    '                    End If
    '                End If

    '                anzDeletedProj = anzDeletedProj + 1

    '            Next

    '            'projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
    '            '                                                 storedEarliest:=StartofCalendar, storedLatest:=Date.Now)

    '            'For Each kvpHist As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste

    '            '    If kvpHist.Value.timeStamp = kvpSelToDel.Value.timeStamp Then
    '            '        If requestTrash.storeProjectToDB(kvpHist.Value) Then

    '            '            If request.deleteProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
    '            '                                                 storedEarliest:=kvpHist.Value.timeStamp, storedLatest:=kvpHist.Value.timeStamp) Then
    '            '                anzDeleted = anzDeleted + 1
    '            '                'Call MsgBox("ok, Projekt '" & hproj.name & "' gespeichert!" & vbLf & hproj.timeStamp.ToShortDateString)

    '            '            Else
    '            '                Call MsgBox("Fehler beim Löschen von Projekt " & kvpSelToDel.Value.name & vbLf & kvpHist.Value.timeStamp.ToShortDateString)
    '            '            End If

    '            '        Else

    '            '            Call MsgBox("Fehler in Löschen von Projekt " & hproj.name)
    '            '        End If
    '            '    Else
    '            '        ' Es ist nicht der richtige TimeStamp von hproj.name

    '            '    End If


    '            'Next kvpHist

    '            '    anzDeletedProj = anzDeletedProj + 1
    '            '    'Call MsgBox("ok, Projekt '" & hproj.name & "' gelöscht!" & vbLf & hproj.timeStamp.ToShortDateString)
    '            'End If

    '            '    Catch ex As Exception

    '            '    ' Call MsgBox("Fehler beim Speichern der Projekte in die Datenbank. Datenbank nicht aktiviert?")
    '            '    Throw New ArgumentException("Fehler beim Löschen der Projekte in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
    '            '    'Exit Sub
    '            'End Try


    '        Else
    '            'Call MsgBox("Es wurde kein Projekt selektiert")
    '            ' die Anzahl selektierter und auch gespeicherter Projekte ist damit = 0
    '            anzDeletedProj = anzSelectedProj
    '            Return anzDeletedProj
    '        End If

    '    Else

    '        Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen")

    '    End If


    '    enableOnUpdate = True

    '    Return anzDeletedProj

    'End Function

    ' Prozedur zum Erzeugen einer Status Übersicht 

    ' 
    ' 
    ''' <summary>
    ''' erzeugt für jedes aktuell laufende Projekt das Status Diagramm , woraus man sofort erkennt wie das Projekt im Vergleich zur Beauftragung bzw. dem letzten freigebenen Stand steht  
    ''' Vorbedingung: in Projektliste stehen nur Projekte, die aktuell bereits seit über einen Monat laufen 
    ''' </summary>
    ''' <param name="ProjektListe">
    ''' enthält die Liste aller Projektnamen, die aktuell laufen 
    ''' </param>
    ''' <param name="repChart">
    ''' Verweis auf das Chart Objekt, das zur Einbettung in die ppt bentigt wird 
    ''' </param>
    ''' <param name="compareToID">
    ''' =0 : Vergleich mit erstem Projekt-Stand überhaupt
    ''' =1 : Vergleich mit Beauftragung 
    ''' =2 : Vergleich mit letzter Freigabe
    ''' =3 : Vergleich mit letztem Planungs-Stand
    ''' </param>
    ''' <param name="auswahl">
    ''' einer der Werte aus 1=Personalkosten, 2=Sonstige Kosten, 3=Gesamtkosten, 4=Rolle, 5=Kostenart 
    ''' </param>
    ''' <param name="qualifier"></param>
    ''' wenn kennzeichnung = rolle oder Kostenart ist, so gibt qualifier an , welche Rolle/Kostenart betrachtet werden soll  
    ''' <param name="showLabels">
    ''' gibt an, ob der Label für den Datenpunkt (=ProjektName) gezeigt werden soll 
    ''' </param>
    ''' <param name="chartBorderVisible"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Sub awinCreateStatusDiagram1(ByRef ProjektListe As Collection, ByRef repChart As Excel.ChartObject, ByVal compareToID As Integer, _
                                         ByVal auswahl As Integer, ByVal qualifier As String, _
                                         ByVal showLabels As Boolean, ByVal chartBorderVisible As Boolean, _
                                         ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double)

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim pname As String
        Dim hproj As New clsProjekt, vglProj As New clsProjekt
        Dim anzSnapshots As Integer, anzBubbles As Integer
        Dim currentValues() As Double, formerValues() As Double, tempArray() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String
        Dim pfDiagram As clsDiagramm
        Dim pfChart As clsEventsPfCharts
        'Dim ptype As String
        Dim chtTitle As String
        Dim chtobjName As String = windowNames(3)
        Dim smallfontsize As Double, titlefontsize As Double
        Dim kennung As String
        Dim singleProject As Boolean
        Dim vglName As String = " "
        Dim variantName As String
        Dim index As Integer
        Dim heuteColumn As Integer = getColumnOfDate(Date.Now)
        Dim werteH() As Double, werteV() As Double ' nimmt die Werte für hproj bzw vglProj auf 
        Dim tmpxValues() As Double
        Dim tmpyValues() As Double
        Dim kennzeichnung As String


        ' Checken, ob überhaupt was in der Projektliste drin ist ...
        ' wenn nein, Exit 
        If ProjektListe.Count = 0 Then
            Exit Sub
        End If


        If ProjektListe.Count > 1 Then
            singleProject = False
        Else
            singleProject = True
        End If


        If width > 450 Then
            titlefontsize = 20
            smallfontsize = 10
        ElseIf width > 250 Then
            titlefontsize = 14
            smallfontsize = 8
        Else
            titlefontsize = 12
            smallfontsize = 8
        End If



        diagramTitle = "Fortschritt"
        kennung = "Fortschritt"


        ' hier werden die Werte aufgenommen ...
        Try
            ReDim currentValues(ProjektListe.Count - 1)
            ReDim formerValues(ProjektListe.Count - 1)
            ReDim nameValues(ProjektListe.Count - 1)
            ReDim colorValues(ProjektListe.Count - 1)
            ReDim positionValues(ProjektListe.Count - 1)
        Catch ex As Exception

            Call MsgBox("Fehler in CreateStatusDiagram1 " & ex.Message)
            Exit Sub

        End Try


        anzBubbles = 0


        For i = 1 To ProjektListe.Count

            pname = CStr(ProjektListe.Item(i))
            Try
                hproj = ShowProjekte.getProject(pname)
                variantName = hproj.variantName

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.getShapeText
                    End If
                Else
                    projekthistorie = New clsProjektHistorie
                End If


                If vglName <> hproj.getShapeText Then
                    If request.pingMongoDb() Then
                        ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                        projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pname, variantName:=variantName, _
                                                                            storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                        projekthistorie.Add(Date.Now, hproj)
                    Else
                        Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
                    End If

                Else
                    ' es muss nichts gemacht werden - es ist bereits die richtige Historie 
                End If

                ' jetzt sind in der Projekt-Historie die richtigen Snapshots 
                ' jetzt muss das Vergleichs-Projekt gesetzt werden 


                Try
                    anzSnapshots = projekthistorie.Count
                Catch ex1 As Exception
                    anzSnapshots = 0
                End Try



                If anzSnapshots > 0 Then

                    Select Case compareToID

                        Case 0
                            ' mit erstem Planungs-Stand vergleichen
                            vglProj = projekthistorie.ElementAt(0)

                        Case 1
                            ' mit Beauftragung vergleichen 

                            vglProj = projekthistorie.beauftragung

                        Case 2
                            ' mit letzter Freigabe vergleichen
                            index = getIndexPrevFreigabe(projekthistorie.liste, anzSnapshots - 1)
                            vglProj = projekthistorie.ElementAt(index)

                        Case 3
                            ' mit letztem Stand vergleichen , das ist das vorletzte Element , da hproj auf der letzten Position ist
                            vglProj = projekthistorie.ElementAt(anzSnapshots - 2)

                        Case Else
                            ' mit Beauftragung vergleichen

                    End Select

                    ReDim werteH(hproj.anzahlRasterElemente - 1)
                    ReDim werteV(vglProj.anzahlRasterElemente - 1)
                    Dim hsum As Double = 0.0
                    Dim vsum As Double = 0.0

                    Select Case auswahl
                        Case 1
                            werteH = hproj.getAllPersonalKosten
                            werteV = vglProj.getAllPersonalKosten
                            diagramTitle = diagramTitle & " Personalkosten"
                            kennzeichnung = "Personalkosten"
                        Case 2
                            werteH = hproj.getGesamtAndereKosten
                            werteV = vglProj.getGesamtAndereKosten
                            diagramTitle = diagramTitle & " Sonstige Kosten"
                            kennzeichnung = "Sonstige Kosten"
                        Case 3
                            werteH = hproj.getGesamtKostenBedarf
                            werteV = vglProj.getGesamtKostenBedarf
                            diagramTitle = diagramTitle & " Gesamtkosten"
                            kennzeichnung = "Gesamtkosten"
                        Case 4
                            If RoleDefinitions.Contains(qualifier) Then
                                werteH = hproj.getRessourcenBedarf(qualifier)
                                werteV = vglProj.getRessourcenBedarf(qualifier)
                                diagramTitle = diagramTitle & " " & qualifier
                                kennzeichnung = "Rolle"
                            End If
                        Case 5
                            If CostDefinitions.Contains(qualifier) Then
                                werteH = hproj.getKostenBedarf(qualifier)
                                werteV = vglProj.getKostenBedarf(qualifier)
                                diagramTitle = diagramTitle & " " & qualifier
                                kennzeichnung = "Kostenart"
                            End If
                        Case Else
                            ' wie Gesamtkosten
                            werteH = hproj.getGesamtKostenBedarf
                            werteV = vglProj.getGesamtKostenBedarf
                            diagramTitle = diagramTitle & " Gesamtkosten"
                            kennzeichnung = "Gesamtkosten"
                    End Select


                    ' jetzt muss abgefangen werden, daß in dem Vergleichs-Projekt gar keine Werte dafür da sind 
                    ' in dem aktuellen Projekt dagegen schon ; oder umgekehrt 
                    ' es muss natürlich auch abgefangen werden, daß der Wert bei beiden nicht existiert 


                    If werteH.Sum <= 0 And werteV.Sum <= 0 Then
                        ' beide existieren nicht 
                    ElseIf werteH.Sum <= 0 Then
                    ElseIf werteV.Sum <= 0 Then
                    Else

                        For h = hproj.Start To heuteColumn - 1
                            hsum = hsum + werteH(h - hproj.Start)
                        Next
                        currentValues(anzBubbles) = hsum / werteH.Sum

                        For v = vglProj.Start To heuteColumn - 1
                            vsum = vsum + werteV(v - vglProj.Start)
                        Next
                        formerValues(anzBubbles) = vsum / werteV.Sum

                        nameValues(anzBubbles) = hproj.name
                        colorValues(anzBubbles) = hproj.farbe


                        anzBubbles = anzBubbles + 1

                    End If


                End If




            Catch ex As Exception
                Call MsgBox("Projekt " & pname & " existiert nicht !")
            End Try


        Next

        If anzBubbles > 0 Then

            ' bestimmen der besten Position für die Werte ...
            Dim labelPosition(4) As String
            labelPosition(0) = "oben"
            labelPosition(1) = "rechts"
            labelPosition(2) = "unten"
            labelPosition(3) = "links"
            labelPosition(4) = "mittig"


            ' Das folgende wird gemacht, damit die Routine pfchartistfrei genutzt werden kann 
            ReDim tmpxValues(anzBubbles - 1)
            ReDim tmpyValues(anzBubbles - 1)

            For i = 0 To anzBubbles - 1
                tmpxValues(i) = formerValues(i) * 10
                tmpyValues(i) = currentValues(i) * 10
            Next


            For i = 0 To anzBubbles - 1

                positionValues(i) = pfchartIstFrei(i, formerValues, currentValues)

            Next



            With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)

                anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
                '
                ' um welches Diagramm handelt es sich ...
                '
                i = 1
                found = False
                While i <= anzDiagrams And Not found

                    Try
                        chtTitle = CType(.ChartObjects(i), Excel.ChartObject).Chart.ChartTitle.Text
                    Catch ex As Exception
                        chtTitle = " "
                    End Try

                    If chtTitle Like ("*" & diagramTitle & "*") Then
                        found = True
                        repChart = CType(.ChartObjects(i), Excel.ChartObject)
                        Exit Sub
                    Else
                        i = i + 1
                    End If
                End While


                ReDim tempArray(anzBubbles - 1)


                With appInstance.Charts.Add

                    CType(.SeriesCollection, Excel.SeriesCollection).NewSeries()
                    CType(.SeriesCollection, Excel.SeriesCollection).Item(1).Name = diagramTitle
                    CType(.SeriesCollection, Excel.SeriesCollection).Item(1).ChartType = xlNS.XlChartType.xlXYScatter

                    For i = 1 To anzBubbles
                        tempArray(i - 1) = formerValues(i - 1)
                    Next i
                    CType(.SeriesCollection, Excel.SeriesCollection).Item(1).XValues = tempArray ' strategic

                    For i = 1 To anzBubbles
                        tempArray(i - 1) = currentValues(i - 1)
                    Next i
                    CType(.SeriesCollection, Excel.SeriesCollection).Item(1).Values = tempArray




                    'Dim series1 As xlNS.Series = _
                    '        CType(.SeriesCollection(1),  _
                    '                xlNS.Series)
                    'Dim point1 As xlNS.Point = _
                    '            CType(series1.Points(1), xlNS.Point)

                    'Dim testName As String
                    For i = 1 To anzBubbles

                        With CType(CType(.SeriesCollection, Excel.SeriesCollection).Item(1).Points(i), Excel.Point)

                            If showLabels Then
                                Try
                                    .HasDataLabel = True
                                    With .DataLabel
                                        .Text = nameValues(i - 1)
                                        If singleProject Then
                                            .Font.Size = awinSettings.CPfontsizeItems + 4
                                        Else
                                            .Font.Size = awinSettings.CPfontsizeItems
                                        End If

                                        Select Case positionValues(i - 1)
                                            Case labelPosition(0)
                                                .Position = xlNS.XlDataLabelPosition.xlLabelPositionAbove
                                            Case labelPosition(1)
                                                .Position = xlNS.XlDataLabelPosition.xlLabelPositionRight
                                            Case labelPosition(2)
                                                .Position = xlNS.XlDataLabelPosition.xlLabelPositionBelow
                                            Case labelPosition(3)
                                                .Position = xlNS.XlDataLabelPosition.xlLabelPositionLeft
                                            Case Else
                                                .Position = xlNS.XlDataLabelPosition.xlLabelPositionCenter
                                        End Select
                                    End With
                                Catch ex As Exception

                                End Try
                            Else
                                .HasDataLabel = False
                            End If

                            .Interior.Color = colorValues(i - 1)
                        End With
                    Next i


                    Try
                        With .PlotArea.Format.Fill
                            .UserPicture(awinPath & "backgroundStatusChart.jpg")
                            .Visible = Microsoft.Office.Core.MsoTriState.msoCTrue
                            .TextureAlignment = Microsoft.Office.Core.MsoTextureAlignment.msoTextureBottomLeft
                            .TextureTile = Microsoft.Office.Core.MsoTriState.msoFalse
                        End With
                    Catch ex As Exception

                    End Try


                    .HasAxis(xlNS.XlAxisType.xlCategory) = True
                    .HasAxis(xlNS.XlAxisType.xlValue) = True

                    With CType(.Axes(xlNS.XlAxisType.xlCategory), Excel.Axis)
                        .HasTitle = True
                        .HasMajorGridlines = False
                        Try
                            .MinimumScale = 0.0
                            .MaximumScale = 1.0
                            .MajorUnitIsAuto = False
                            .MajorUnit = 0.2
                        Catch ex As Exception

                        End Try

                        With .AxisTitle
                            .Characters.Text = "geplant"
                            .Characters.Font.Size = titlefontsize
                            .Characters.Font.Bold = False
                        End With
                        With .TickLabels.Font
                            .FontStyle = "Normal"
                            .Bold = True
                            .Size = awinSettings.fontsizeItems

                        End With

                    End With


                    With CType(.Axes(xlNS.XlAxisType.xlValue), Excel.Axis)
                        .HasTitle = True
                        .HasMajorGridlines = False

                        Try
                            .MinimumScale = 0.0
                            .MaximumScale = 1.0
                            .MajorUnitIsAuto = False
                            .MajorUnit = 0.2
                        Catch ex As Exception

                        End Try

                        With .AxisTitle
                            .Characters.Text = "tatsächlich"
                            .Characters.Font.Size = titlefontsize
                            .Characters.Font.Bold = False
                        End With

                        With .TickLabels.Font
                            .FontStyle = "Normal"
                            .Bold = True
                            .Size = awinSettings.fontsizeItems
                        End With
                    End With
                    .HasLegend = False
                    .HasTitle = True
                    .ChartTitle.Text = diagramTitle
                    .ChartTitle.Characters.Font.Size = awinSettings.fontsizeTitle
                    .Location(Where:=xlNS.XlChartLocation.xlLocationAsObject, _
                          Name:=CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Name)
                End With


                'appInstance.ShowChartTipNames = False
                'appInstance.ShowChartTipValues = False

                With CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                    .Top = top
                    .Left = left
                    .Width = width
                    .Height = height
                    .Name = chtobjName
                End With



                With CType(appInstance.ActiveSheet, Excel.Worksheet)
                    Try
                        Dim obj As Object = chtobjName
                        CType(.Shapes(chtobjName), Excel.Shape).Line.Visible = CType(chartBorderVisible, Microsoft.Office.Core.MsoTriState)

                    Catch ex As Exception

                    End Try
                End With

                pfDiagram = New clsDiagramm

                'pfChart = New clsAwinEvent
                pfChart = New clsEventsPfCharts
                pfChart.PfChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart

                'pfDiagram.setpfDiagramEvent = pfChart
                pfDiagram.setDiagramEvent = pfChart

                With pfDiagram
                    .DiagrammTitel = diagramTitle
                    .diagrammTyp = DiagrammTypen(3) ' Portfolio
                    .gsCollection = ProjektListe
                    .isCockpitChart = False
                End With

                DiagramList.Add(pfDiagram)
                'pfDiagram = Nothing

                repChart = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

            End With
        Else
            'Call MsgBox("es waren keine Projekte darzustellen ... (in awinCreateStatusDiagram1)")
        End If




    End Sub

    ''' <summary>
    ''' gibt für ein gegebenes Projekt die errechnete Farbe und den errechneten Status zurück
    ''' dabei wird das aktuelle Projekt in Relation zur Beauftragung/letzten Freigabe gesetzt
    ''' wenn es noch keine Projekt-Historie gibt, so wird grün und "0" zurückgegeben   
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' das Projekt in Question
    ''' <param name="compareTo">
    ''' =0 : Vergleich mit erstem Projekt-Stand überhaupt
    ''' =1 : Vergleich mit Beauftragung 
    ''' =2 : Vergleich mit letzter Freigabe
    ''' =3 : Vergleich mit letztem Planungs-Stand
    ''' </param>
    ''' <param name="auswahl">
    ''' einer der Werte aus 1=Personalkosten, 2=Sonstige Kosten, 3=Gesamtkosten, 4=Rolle, 5=Kostenart 
    ''' </param>
    ''' <param name="statusValue">
    ''' Rückgabe Paarmeter - ein Wert zwischen 0 und sehr groß; je größer über 1, desto besser im Fortschritt 
    ''' je kleiner unter 1, desto schlechter im Fortschritt
    ''' </param>
    ''' <param name="statusColor">
    ''' Rückgabe Parameter: entweder grün, gelb oder rot
    ''' </param>
    ''' <remarks></remarks>
    Public Sub getStatusColorProject(ByRef hproj As clsProjekt, ByVal compareTo As Integer, ByVal auswahl As Integer, ByVal qualifier As String, _
                                  ByRef statusValue As Double, ByRef statusColor As Long)

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim currentValues() As Double
        Dim formerValues() As Double
        Dim vglProj As clsProjekt
        Dim vglName As String, pname As String, variantName As String
        Dim anzSnapshots As Integer, index As Integer
        Dim heuteColumn As Integer = getColumnOfDate(Date.Now)
        Dim cValue As Double, fValue As Double

        With hproj
            pname = .name
            variantName = .variantName
        End With
        vglName = " "

        Try
            ReDim currentValues(hproj.anzahlRasterElemente - 1)

        Catch ex As Exception

            statusValue = 1.0
            statusColor = awinSettings.AmpelGruen
            Exit Sub

        End Try


        If Not projekthistorie Is Nothing Then
            If projekthistorie.Count > 0 Then
                vglName = projekthistorie.First.getShapeText
            End If
        Else
            projekthistorie = New clsProjektHistorie
        End If


        If vglName <> hproj.getShapeText Then
            If request.pingMongoDb() Then
                ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde
                projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pname, variantName:=variantName, _
                                                                   storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                If projekthistorie.Count > 0 Then
                    projekthistorie.Add(Date.Now, hproj)
                End If

            Else
                Call MsgBox(" Datenbank-Verbindung ist unterbrochen!" & vbLf & " Projekthistorie kann nicht geladen werden")
            End If

        Else
            ' es muss nichts gemacht werden - es ist bereits die richtige Historie 
        End If


        ' jetzt sind in der Projekt-Historie die richtigen Snapshots 
        ' jetzt muss das Vergleichs-Projekt gesetzt werden 




        Try
            anzSnapshots = projekthistorie.Count
        Catch ex1 As Exception
            anzSnapshots = 0
        End Try



        If anzSnapshots > 0 Then

            Select Case compareTo

                Case 0
                    ' mit erstem Planungs-Stand vergleichen
                    vglProj = projekthistorie.ElementAt(0)

                Case 1
                    ' mit Beauftragung vergleichen 
                    Try
                        vglProj = projekthistorie.beauftragung
                    Catch ex As Exception
                        vglProj = projekthistorie.ElementAt(0)
                    End Try


                Case 2
                    ' mit letzter Freigabe vergleichen
                    index = getIndexPrevFreigabe(projekthistorie.liste, anzSnapshots - 1)
                    vglProj = projekthistorie.ElementAt(index)

                Case 3
                    ' mit letztem Stand vergleichen , das ist das vorletzte Element , da hproj auf der letzten Position ist
                    vglProj = projekthistorie.ElementAt(anzSnapshots - 2)

                Case Else
                    ' mit erstem Element vergleichen 
                    vglProj = projekthistorie.ElementAt(0)
            End Select


            ReDim formerValues(vglProj.anzahlRasterElemente - 1)
            Dim hsum As Double = 0.0
            Dim vsum As Double = 0.0

            Select Case auswahl
                Case 1
                    currentValues = hproj.getAllPersonalKosten
                    formerValues = vglProj.getAllPersonalKosten

                Case 2
                    currentValues = hproj.getGesamtAndereKosten
                    formerValues = vglProj.getGesamtAndereKosten

                Case 3
                    currentValues = hproj.getGesamtKostenBedarf
                    formerValues = vglProj.getGesamtKostenBedarf

                Case 4
                    If RoleDefinitions.Contains(qualifier) Then
                        currentValues = hproj.getRessourcenBedarf(qualifier)
                        formerValues = vglProj.getRessourcenBedarf(qualifier)
                    End If
                Case 5
                    If CostDefinitions.Contains(qualifier) Then
                        currentValues = hproj.getKostenBedarf(qualifier)
                        formerValues = vglProj.getKostenBedarf(qualifier)
                    End If
                Case Else
                    ' wie Gesamtkosten
                    currentValues = hproj.getGesamtKostenBedarf
                    formerValues = vglProj.getGesamtKostenBedarf

            End Select


            ' jetzt muss abgefangen werden, daß in dem Vergleichs-Projekt gar keine Werte dafür da sind 
            ' in dem aktuellen Projekt dagegen schon ; oder umgekehrt 
            ' es muss natürlich auch abgefangen werden, daß der Wert bei beiden nicht existiert 


            If currentValues.Sum <= 0 And formerValues.Sum <= 0 Then
                statusValue = 0.0
                statusColor = awinSettings.AmpelNichtBewertet
                ' beide existieren nicht 

            ElseIf currentValues.Sum <= 0 Then
                statusValue = 0.0
                statusColor = awinSettings.AmpelRot

            ElseIf formerValues.Sum <= 0 Then
                statusValue = 2.0
                statusColor = awinSettings.AmpelGruen

            Else
                Dim korrFaktor As Double = formerValues.Sum / currentValues.Sum

                For h = hproj.Start To heuteColumn - 1
                    hsum = hsum + currentValues(h - hproj.Start)
                Next
                cValue = hsum / currentValues.Sum

                For v = vglProj.Start To heuteColumn - 1
                    vsum = vsum + formerValues(v - vglProj.Start)
                Next
                fValue = vsum / formerValues.Sum

                If fValue > 0 Then
                    statusValue = korrFaktor * cValue / fValue
                Else
                    statusValue = 2
                End If

                If statusValue >= 1.0 Then
                    statusColor = awinSettings.AmpelGruen
                ElseIf statusValue >= 0.9 Then
                    statusColor = awinSettings.AmpelGelb
                Else
                    statusColor = awinSettings.AmpelRot
                End If

            End If

        Else
            statusValue = 1.0
            statusColor = awinSettings.AmpelGruen
        End If

    End Sub

    Sub zeichneFortschrittDiagramm(ByVal boxName As String, ByVal compareToID As Integer, ByVal auswahl As Integer, ByVal qualifier As String, _
                                   ByRef pptShape As pptNS.Shape, ByRef reportObj As xlNS.ChartObject, ByRef notYetDone As Boolean)

        Dim PListe As New Collection
        Dim htop As Double, hleft As Double, hheight As Double, hwidth As Double


        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            If istLaufendesProjekt(kvp.Value) Then

                PListe.Add(kvp.Key, kvp.Key)

            End If

        Next
        If PListe.Count > 0 Then

            reportObj = Nothing
            qualifier = ""
            htop = 100
            hleft = 100
            hwidth = 450
            hheight = awinSettings.ChartHoehe1

            Call awinCreateStatusDiagram1(PListe, reportObj, compareToID, auswahl, qualifier, True, False, htop, hleft, hwidth, hheight)

            If Not reportObj Is Nothing Then

                notYetDone = True

                With reportObj
                    .Chart.HasAxis(xlNS.XlAxisType.xlCategory) = False
                    .Chart.HasAxis(xlNS.XlAxisType.xlValue) = False
                End With

                Try
                    ' muss hier gemacht werden - im notYetDone Block muss das nicht für alle Diagramme gemacht werden 
                    DiagramList.Remove(DiagramList.Count)
                Catch ex As Exception

                End Try
            Else

                If pptShape.TextFrame2.HasText Then
                    pptShape.TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
                End If

            End If

        Else
            If pptShape.TextFrame2.HasText Then
                pptShape.TextFrame2.TextRange.Text = "es gibt keine laufenden Projekte im betrachteten Zeitraum ... "
            End If

        End If


    End Sub



    Sub awinZeichneStatus(ByVal numberIt As Boolean)
        Dim heute As Date = Date.Now
        Dim index As Integer = 0

        Dim todoListe As New SortedList(Of Integer, clsProjekt)

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste


            If istLaufendesProjekt(kvp.Value) Then

                todoListe.Add(kvp.Value.tfZeile, kvp.Value)


            End If

        Next

        ' jetzt wird die todoListe abgearbeitet 

        For Each kvp As KeyValuePair(Of Integer, clsProjekt) In todoListe

            index = index + 1
            If numberIt Then
                Call zeichneStatusSymbolInPlantafel(kvp.Value, index)
            Else
                Call zeichneStatusSymbolInPlantafel(kvp.Value, 0)
            End If

        Next


    End Sub

    ''' <summary>
    ''' zeichnet eine Legenden Tabelle mit Darstellung von Shape sowie Short- wie Long-Name
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="pptslide"></param>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="legendPhaseVorlage"></param>
    ''' <param name="legendMilestoneVorlage"></param>
    ''' <remarks></remarks>
    Sub zeichneLegendenTabelle(ByRef pptShape As pptNS.Shape, ByVal pptslide As pptNS.Slide, _
                                   ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                   ByVal legendPhaseVorlage As pptNS.Shape, ByVal legendMilestoneVorlage As pptNS.Shape, _
                                   ByVal legendBuColorShape As pptNS.Shape)

        Dim tabelle As pptNS.Table
        Dim anzZeilen As Integer
        Dim anzMaxZeilen As Integer
        Dim anzSpalten As Integer
        Dim zeilenHoehe As Double
        Dim zeilenHoeheTitel As Double
        Dim toDraw As Integer
        Dim anzTabellenElements As Integer
        Dim anzDrawn As Integer = 0
        Dim modRest As Integer
        Dim copiedShape As pptNS.ShapeRange
        Dim uniquePhases As New Collection
        Dim uniqueMilestones As New Collection
        Dim fullName As String = ""
        Dim elemName As String = ""
        Dim breadcrumb As String = ""

        ' die eindeutigen Klassen-Namen bestimmen
        ' nur für die muss eine Legende geschrieben werden 
        If selectedPhases.Count > 0 Then
            For i As Integer = 1 To selectedPhases.Count
                fullName = CStr(selectedPhases.Item(i))
                Call splitHryFullnameTo2(fullName, elemName, breadcrumb)

                If uniquePhases.Contains(elemName) Then
                    ' nichts tun
                Else
                    uniquePhases.Add(elemName, elemName)
                End If
            Next
        End If

        '
        If selectedMilestones.Count > 0 Then
            For i As Integer = 1 To selectedMilestones.Count
                fullName = CStr(selectedMilestones.Item(i))
                Call splitHryFullnameTo2(fullName, elemName, breadcrumb)

                If uniqueMilestones.Contains(elemName) Then
                    ' nichts tun
                Else
                    uniqueMilestones.Add(elemName, elemName)
                End If
            Next
        End If


        Try
            tabelle = pptShape.Table
            ' alle Cells der Tabelle mit Schriftgröße von legendPhaseVorlage besetzen
            For i = 2 To tabelle.Rows.Count
                For j = 1 To tabelle.Columns.Count
                    CType(tabelle.Cell(i, j), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size = legendPhaseVorlage.TextFrame2.TextRange.Font.Size
                Next j
                tabelle.Rows(i).Height = 0.5
            Next i
            ' Zellenhöhe auf eine Minimum setzen, so dass Font.Size hineinpasst
        Catch ex As Exception
            Throw New Exception("Shape für Legenden-Liste hat keine Tabelle")
        End Try

        anzSpalten = tabelle.Columns.Count
        anzZeilen = tabelle.Rows.Count
        anzTabellenElements = System.Math.DivRem(anzSpalten, 3, modRest)

        If modRest <> 0 Then
            Throw New Exception("Tabelle hat keine durch 3 teilbare Anzahl Spalten" & vbLf & "Symbol, Short- und Long-Name")
        End If

        If anzZeilen < 2 Then
            Throw New Exception("Tabelle muss mindestens 2 Zeilen haben .... ")
        End If

        zeilenHoehe = tabelle.Rows(tabelle.Rows.Count).Height
        zeilenHoeheTitel = tabelle.Rows(1).Height
        anzMaxZeilen = (pptslide.CustomLayout.Height - (pptShape.Top + zeilenHoeheTitel)) / zeilenHoehe - 1
        toDraw = uniquePhases.Count + uniqueMilestones.Count




        Dim curZeile As Integer = 2
        Dim curSpalte As Integer = 1

        Dim phaseShape As xlNS.Shape
        Dim milestoneShape As xlNS.Shape
        Dim phaseName As String = ""

        Dim milestoneName As String = ""
        Dim breadcrumbMS As String = ""
        Dim shortName As String
        Dim factor As Double
        Dim tmpBU As clsBusinessUnit


        breadcrumb = ""
        ' jetzt werden ggf zunächst die BU Symbole gezeichnet 
        If Not IsNothing(legendBuColorShape) Then

            For i = 1 To businessUnitDefinitions.Count
                tmpBU = businessUnitDefinitions.ElementAt(i - 1).Value
                ' jetzt das Shape eintragen 
                legendBuColorShape.Copy()
                copiedShape = pptslide.Shapes.Paste()
                With copiedShape(1)
                    .Height = zeilenHoehe * 0.8
                    .Top = tabelle.Cell(curZeile, curSpalte).Shape.Top + (tabelle.Cell(curZeile, curSpalte).Shape.Height - .Height) * 0.5
                    .Left = tabelle.Cell(curZeile, curSpalte).Shape.Left + (tabelle.Cell(curZeile, curSpalte).Shape.Width - .Width) * 0.5
                    .Fill.ForeColor.RGB = tmpBU.color
                End With
                ' jetzt den Business Unit Name eintragen 
                CType(tabelle.Cell(curZeile, curSpalte + 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Produktlinie " & tmpBU.name

                curSpalte = curSpalte + 3
                If curSpalte > anzSpalten Then
                    curSpalte = 1
                    curZeile = curZeile + 1
                    If curZeile > anzZeilen Then
                        tabelle.Rows.Add()
                        anzZeilen = anzZeilen + 1
                    End If

                End If
            Next

            ' jetzt die undefinierte Produktlinie noch zeichnen ...

            ' jetzt das Shape eintragen 
            legendBuColorShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)
                .Height = zeilenHoehe * 0.8
                .Top = tabelle.Cell(curZeile, curSpalte).Shape.Top + (tabelle.Cell(curZeile, curSpalte).Shape.Height - .Height) * 0.5
                .Left = tabelle.Cell(curZeile, curSpalte).Shape.Left + (tabelle.Cell(curZeile, curSpalte).Shape.Width - .Width) * 0.5
                .Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
            End With
            ' jetzt den Business Unit Name eintragen 
            CType(tabelle.Cell(curZeile, curSpalte + 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Produktlinie ist undefiniert"

            curSpalte = curSpalte + 3
            If curSpalte > anzSpalten Then
                curSpalte = 1
                curZeile = curZeile + 1
                If curZeile > anzZeilen Then
                    tabelle.Rows.Add()
                    anzZeilen = anzZeilen + 1
                End If

            End If

        End If

        ' Überprüfung, ob die restlichen Zeilen für die Legende ausreichen

        If anzMaxZeilen - (curZeile - 1) < toDraw / anzTabellenElements Then
            Throw New Exception("Anzahl Zeilen in der Tabelle sind nicht ausreichend." & vbLf & "Tabelle muss anders definiert werden .... ")
        End If


        For j = 1 To uniquePhases.Count

            phaseName = CStr(uniquePhases(j))
            Dim isMissingDefinition As Boolean

            If PhaseDefinitions.Contains(phaseName) Then
                phaseShape = PhaseDefinitions.getShape(phaseName)
                shortName = PhaseDefinitions.getAbbrev(phaseName)
                isMissingDefinition = False
            Else
                phaseShape = missingPhaseDefinitions.getShape(phaseName)
                shortName = missingPhaseDefinitions.getAbbrev(phaseName)
                isMissingDefinition = True
            End If


            ' Phasen-Shape 
            phaseShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .Height = legendPhaseVorlage.Height
                .Width = legendPhaseVorlage.Width
                .Top = tabelle.Cell(curZeile, curSpalte).Shape.Top + (tabelle.Cell(curZeile, curSpalte).Shape.Height - .Height) * 0.5
                .Left = tabelle.Cell(curZeile, curSpalte).Shape.Left + (tabelle.Cell(curZeile, curSpalte).Shape.Width - .Width) * 0.5

                If .Top > pptslide.CustomLayout.Height Then
                    Throw New Exception("Die LegendenTabelle wird zu groß für eine Seite." & vbLf & "Tabelle muss anders definiert werden .... ")
                End If
            End With

            ' jetzt den Abkürzungstext eintragen 
            CType(tabelle.Cell(curZeile, curSpalte + 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size = legendPhaseVorlage.TextFrame2.TextRange.Font.Size
            CType(tabelle.Cell(curZeile, curSpalte + 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = shortName

            ' jetzt den Long Name eintragen 
            CType(tabelle.Cell(curZeile, curSpalte + 2), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size = legendPhaseVorlage.TextFrame2.TextRange.Font.Size
            CType(tabelle.Cell(curZeile, curSpalte + 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = phaseName

            curSpalte = curSpalte + 3
            If curSpalte > anzSpalten Then
                curSpalte = 1
                curZeile = curZeile + 1
                If curZeile > anzZeilen Then
                    tabelle.Rows.Add()
                    anzZeilen = anzZeilen + 1
                End If
            End If

        Next j ' nächste selektierte Phase bearbeiten


        ' jetzt die Meilensteine eintragen 
        For j = 1 To uniqueMilestones.Count

            milestoneName = CStr(uniqueMilestones(j))

            ' Änderung tk 26.11.15
            If MilestoneDefinitions.Contains(milestoneName) Then
                milestoneShape = MilestoneDefinitions.getShape(milestoneName)
            Else
                milestoneShape = missingMilestoneDefinitions.getShape(milestoneName)
            End If


            factor = milestoneShape.Width / milestoneShape.Height
            shortName = MilestoneDefinitions.getAbbrev(milestoneName)
            ' Phasen-Shape 
            milestoneShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .Height = legendMilestoneVorlage.Height
                .Width = factor * .Height
                .Top = tabelle.Cell(curZeile, curSpalte).Shape.Top + (tabelle.Cell(curZeile, curSpalte).Shape.Height - .Height) * 0.5
                .Left = tabelle.Cell(curZeile, curSpalte).Shape.Left + (tabelle.Cell(curZeile, curSpalte).Shape.Width - .Width) * 0.5

                If .Top > pptslide.CustomLayout.Height Then
                    Throw New Exception("Die LegendenTabelle wird zu groß für eine Seite." & vbLf & "Tabelle muss anders definiert werden .... ")
                End If
            End With

            ' jetzt den Abkürzungstext eintragen 
            CType(tabelle.Cell(curZeile, curSpalte + 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size = legendMilestoneVorlage.TextFrame2.TextRange.Font.Size
            CType(tabelle.Cell(curZeile, curSpalte + 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = shortName

            ' jetzt den Long Name eintragen 
            CType(tabelle.Cell(curZeile, curSpalte + 2), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size = legendMilestoneVorlage.TextFrame2.TextRange.Font.Size
            CType(tabelle.Cell(curZeile, curSpalte + 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = milestoneName

            curSpalte = curSpalte + 3
            If curSpalte > anzSpalten Then
                curSpalte = 1
                curZeile = curZeile + 1
                If curZeile > anzZeilen Then
                    tabelle.Rows.Add()
                    anzZeilen = anzZeilen + 1
                End If
            End If

        Next


    End Sub


    ''' <summary>
    ''' zeichnet die Projekte, die eines der angegebenen Phasen, Meilensteine im Zeitraum enthält 
    ''' </summary>
    ''' <param name="pptshape"></param>
    ''' <param name="pptslide"></param>
    ''' <param name="myCollection"></param>
    ''' <param name="prcTyp">gibt an, ob es sich um Phasen, Meilensteine, Rollen oder Kosten handelt</param>
    ''' <remarks></remarks>
    Sub zeichneTabelleProjekteMitElemImMonat(ByRef pptshape As pptNS.Shape, ByVal pptslide As pptNS.Slide, ByVal myCollection As Collection, ByVal prcTyp As String)

        Dim tabelle As pptNS.Table
        Dim tabheight As Double = pptshape.Height, tabwidth As Double = pptshape.Width
        Dim anzZeilen As Integer
        Dim zeilenHoehe As Double
        Dim zeilenHoeheBottom As Double

        Dim anzDrawn As Integer = 0
        Dim neededSpalten As Integer = showRangeRight - showRangeLeft + 1
        Dim neededZeilen As Integer = 0

        Dim ergebnisListe(,) As String
        Dim nrOfZeilen(neededSpalten - 1) As Integer


        If showRangeRight = 0 Or showRangeLeft = 0 Or showRangeRight - showRangeLeft = 0 Then
            Throw New Exception("kein Zeitraum in Tabelle Anzeigen der Elemente angegeben ")
        End If

        If myCollection.Count = 0 Then
            Throw New Exception("keine Elemente angegeben ... ")
        End If


        If prcTyp = DiagrammTypen(0) Or prcTyp = DiagrammTypen(5) Or _
            prcTyp = DiagrammTypen(1) Or prcTyp = DiagrammTypen(2) Then

            ergebnisListe = ShowProjekte.getProjectsWithElemNameInMonth(myCollection, prcTyp)

        Else
            ReDim ergebnisListe(0, 0)
        End If


        Dim anzProjekte As Integer = ShowProjekte.Count
        Dim tmpValue As Integer = 0

        For i As Integer = 1 To neededSpalten

            Dim found As Boolean = False
            Dim ix As Integer = 1

            While ix <= anzProjekte And Not found
                If IsNothing(ergebnisListe(i - 1, ix - 1)) Then
                    found = True
                Else
                    If ergebnisListe(i - 1, ix - 1) = "" Then
                        found = True
                    Else
                        ix = ix + 1
                    End If
                End If

            End While
            nrOfZeilen(i - 1) = ix - 1
            ' jetzt gibt ix die Anzahl Zeilen in dem Monat wieder 
            If neededZeilen < ix - 1 Then
                neededZeilen = ix
            End If
        Next

        neededZeilen = neededZeilen + 2
        ' jetzt sind in neededzeilen die Anzahl Zeilen inkl der Bottom-Line und Header-Line für Angabe, welche Meilensteine für den Zeitraum 

        Dim curZeile As Integer = 2
        Dim curSpalte As Integer = 1

        Try
            tabelle = pptshape.Table
        Catch ex As Exception
            Throw New Exception("Shape für hat keine Tabelle")
        End Try

        Dim anzSpalten As Integer

        anzSpalten = tabelle.Columns.Count
        anzZeilen = tabelle.Rows.Count

        ' jetzt muss die Tabelle ggf angepasst werden 
        If anzSpalten <> neededSpalten Then

            If anzSpalten < neededSpalten Then
                Do While anzSpalten < neededSpalten
                    tabelle.Columns.Add()
                    anzSpalten = anzSpalten + 1
                Loop
            ElseIf anzSpalten > neededSpalten Then
                Do While anzSpalten > neededSpalten
                    tabelle.Columns.Item(1).Delete()
                    anzSpalten = anzSpalten - 1
                Loop
            End If

            pptshape.Width = tabwidth

        End If

        If anzZeilen <> neededZeilen Then
            If anzZeilen < neededZeilen Then
                Do While anzZeilen < neededZeilen
                    tabelle.Rows.Add(anzZeilen - 1)
                    anzZeilen = anzZeilen + 1
                Loop
            ElseIf anzZeilen > neededZeilen Then
                Do While anzZeilen > neededZeilen
                    tabelle.Rows.Item(2).Delete()
                    anzZeilen = anzZeilen - 1
                Loop
            End If
        End If

        If pptshape.Height > tabheight Then
            Do While pptshape.Height > 1.05 * tabheight

                Try
                    pptshape.Height = tabheight
                Catch ex As Exception

                End Try

                ' wenn das nicht funktioniert hat ... 
                If pptshape.Height > 1.05 * tabheight Then
                    For ize As Integer = 2 To anzZeilen

                        For isp As Integer = 1 To neededSpalten
                            With tabelle
                                Dim oldRowHeight As Double = .Rows(ize).Height
                                .Rows(ize).Height = 0.87 * oldRowHeight
                                If .Rows(ize).Height > 0.9 * oldRowHeight Then
                                    ' die Schrift muss verkleinert werden 
                                    CType(.Cell(ize, isp), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size = _
                                            CType(.Cell(ize, isp), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size - 1

                                    CType(.Cell(ize, isp), pptNS.Cell).Shape.TextFrame2.MarginTop = 0.05
                                    CType(.Cell(ize, isp), pptNS.Cell).Shape.TextFrame2.MarginBottom = 0.05

                                End If

                                .Rows(ize).Height = 0.87 * oldRowHeight
                            End With
                        Next

                    Next
                    Try
                        pptshape.Height = tabheight
                    Catch ex As Exception

                    End Try
                End If

            Loop
        End If


        zeilenHoehe = tabelle.Rows(1).Height
        zeilenHoeheBottom = tabelle.Rows(tabelle.Rows.Count).Height

        'jetzt ist die korrekte Anzahl Zeilen und Spalten gegeben 

        Dim oldBottomHeight As Double = zeilenHoeheBottom

        ' jetzt wird die Bottom Zeile geschrieben 
        Dim startDate = StartofCalendar.AddMonths(-1)
        For m As Integer = showRangeLeft To showRangeRight
            With tabelle
                CType(.Cell(neededZeilen, m - showRangeLeft + 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = _
                            startDate.AddMonths(m).ToString("MMM yy")
            End With
        Next m

        zeilenHoeheBottom = tabelle.Rows(tabelle.Rows.Count).Height
        If zeilenHoeheBottom > oldBottomHeight Then

            Do While zeilenHoeheBottom > oldBottomHeight * 1.03
                With tabelle
                    For m As Integer = showRangeLeft To showRangeRight
                        CType(.Cell(neededZeilen, m - showRangeLeft + 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size = _
                                                CType(.Cell(neededZeilen, m - showRangeLeft + 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Size - 1

                        CType(.Cell(neededZeilen, m - showRangeLeft + 1), pptNS.Cell).Shape.TextFrame2.MarginTop = 0.05
                        CType(.Cell(neededZeilen, m - showRangeLeft + 1), pptNS.Cell).Shape.TextFrame2.MarginBottom = 0.05
                    Next
                End With

                zeilenHoeheBottom = tabelle.Rows(tabelle.Rows.Count).Height

                If zeilenHoeheBottom > oldBottomHeight Then
                    tabelle.Rows(tabelle.Rows.Count).Height = oldBottomHeight
                End If

                zeilenHoeheBottom = tabelle.Rows(tabelle.Rows.Count).Height
            Loop


        End If

        ' jetzt wird die Header-Line geschrieben 
        Dim headerzeile As String = ""
        If myCollection.Count = 1 Then
            If prcTyp = DiagrammTypen(5) Then
                headerzeile = "alle Projekte mit Meilenstein "
            ElseIf prcTyp = DiagrammTypen(0) Then
                headerzeile = "alle Projekte mit Phase "
            ElseIf prcTyp = DiagrammTypen(1) Then
                headerzeile = "alle Projekte mit Rolle "
            ElseIf prcTyp = DiagrammTypen(2) Then
                headerzeile = "alle Projekte mit Kostenart "
            End If
        Else
            If prcTyp = DiagrammTypen(5) Then
                headerzeile = "alle Projekte mit Meilensteinen "
            ElseIf prcTyp = DiagrammTypen(0) Then
                headerzeile = "alle Projekte mit Phasen "
            ElseIf prcTyp = DiagrammTypen(1) Then
                headerzeile = "alle Projekte mit Rollen "
            ElseIf prcTyp = DiagrammTypen(2) Then
                headerzeile = "alle Projekte mit Kostenarten "
            End If

        End If
        For m As Integer = 1 To myCollection.Count
            If m = 1 Then
                If prcTyp = DiagrammTypen(0) Or prcTyp = DiagrammTypen(5) Then
                    headerzeile = headerzeile & CStr(myCollection.Item(m)).Replace("#", "-")
                Else
                    headerzeile = headerzeile & CStr(myCollection.Item(m))
                End If

            Else
                If prcTyp = DiagrammTypen(0) Or prcTyp = DiagrammTypen(5) Then
                    headerzeile = headerzeile & ", " & CStr(myCollection.Item(m)).Replace("#", "-")
                Else
                    headerzeile = headerzeile & ", " & CStr(myCollection.Item(m))
                End If

            End If
        Next

        If prcTyp = DiagrammTypen(1) Then
            ' Einheit [PT] ergänzen 
            headerzeile = headerzeile & " [PT]"
        ElseIf prcTyp = DiagrammTypen(2) Then
            headerzeile = headerzeile & " [T€]"
        End If

        ' Headerzeile schreiben 
        With tabelle
            CType(.Cell(1, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = headerzeile
        End With


        ' jetzt werden die eigentlichen Inhalte geschrieben 
        With tabelle
            For isp As Integer = 1 To neededSpalten

                For ize As Integer = 1 To nrOfZeilen(isp - 1)

                    CType(.Cell(neededZeilen - ize, isp), pptNS.Cell).Shape.TextFrame2.TextRange.Text = _
                                    ergebnisListe(isp - 1, ize - 1)

                Next

            Next
        End With


    End Sub



    ''' <summary>
    ''' zeichnet die Tabelle mit den Namen der Projekt [inkl Varianten Name] im betrachteten Portfolio 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="pptslide"></param>
    ''' <remarks></remarks>
    Sub zeichneSzenarioTabelle(ByRef pptShape As pptNS.Shape, ByVal pptslide As pptNS.Slide)
        Dim tabelle As pptNS.Table
        Dim anzZeilen As Integer
        Dim anzMaxZeilen As Integer
        Dim anzSpalten As Integer
        Dim zeilenHoehe As Double
        Dim zeilenHoeheTitel As Double
        Dim pName As String
        Dim toDraw As Integer
        Dim anzTabellenElements As Integer
        Dim anzDrawn As Integer = 0


        Dim curZeile As Integer = 2
        Dim curSpalte As Integer = 1

        Try
            tabelle = pptShape.Table
        Catch ex As Exception
            Throw New Exception("Shape für Szenario-Liste hat keine Tabelle")
        End Try

        anzSpalten = tabelle.Columns.Count
        anzZeilen = tabelle.Rows.Count


        zeilenHoehe = tabelle.Rows(tabelle.Rows.Count).Height
        zeilenHoeheTitel = tabelle.Rows(1).Height
        anzMaxZeilen = (pptslide.CustomLayout.Height - (pptShape.Top + zeilenHoeheTitel)) / zeilenHoehe - 1
        anzTabellenElements = anzMaxZeilen * anzSpalten

        toDraw = ShowProjekte.Count

        With tabelle

            If currentConstellation.Trim.Length > 0 Then
                CType(.Cell(1, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = _
                CType(.Cell(1, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text & " " & currentConstellation
            Else
                CType(.Cell(1, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = _
                    CType(.Cell(1, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text & " <nicht benannt>"
            End If


            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
                pName = kvp.Value.getShapeText


                CType(.Cell(curZeile, curSpalte), pptNS.Cell).Shape.TextFrame2.TextRange.Text = pName

                curSpalte = curSpalte + 1
                anzDrawn = anzDrawn + 1

                If curSpalte > anzSpalten Then
                    curSpalte = 1
                    curZeile = curZeile + 1
                    If curZeile > anzZeilen Then
                        .Rows.Add()
                        anzZeilen = anzZeilen + 1
                    End If
                End If


            Next

        End With



    End Sub

    Sub zeichneTabelleZielErreichung(ByRef pptShape As pptNS.Shape, ByVal farbtyp As Integer)

        Dim heute As Date = Date.Now
        Dim index As Integer = 0
        Dim tabelle As pptNS.Table
        Dim farbTypenListe As New Collection
        Dim timeFrameProjekte As New Collection
        Dim hproj As clsProjekt

        Try
            tabelle = pptShape.Table
        Catch ex As Exception
            Throw New Exception("Shape hat keine Tabelle")
        End Try


        If farbtyp <= 3 Then
            farbTypenListe.Add(farbtyp)
        ElseIf farbtyp = 12 Then
            Dim tmpfarbe As Integer = 2
            farbTypenListe.Add(tmpfarbe, tmpfarbe.ToString)
            tmpfarbe = 3
            farbTypenListe.Add(tmpfarbe, tmpfarbe.ToString)
        End If

        Dim todoListe As New SortedList(Of Long, clsProjekt)
        Dim key As Long

        Dim selectionType As Integer = -1
        timeFrameProjekte = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        For Each pname As String In timeFrameProjekte
            Try
                hproj = ShowProjekte.getProject(pname)
                key = 10000 * hproj.tfZeile + hproj.Start
                todoListe.Add(key, hproj)
            Catch ex As Exception

            End Try

        Next


        Dim msNumber As Integer = 1

        ' jetzt wird die todoListe abgearbeitet 
        Dim tabellenzeile As Integer = 2
        For Each kvp As KeyValuePair(Of Long, clsProjekt) In todoListe

            hproj = kvp.Value
            For p = 1 To hproj.CountPhases

                Dim cphase As clsPhase = hproj.getPhase(p)
                Dim phaseStart As Date = hproj.startDate.AddMonths(cphase.relStart - 1)
                Dim resultColumn As Integer

                For r = 1 To cphase.countMilestones
                    Dim cResult As clsMeilenstein
                    Dim cBewertung As clsBewertung

                    cResult = cphase.getMilestone(r)

                    cBewertung = cResult.getBewertung(1)

                    resultColumn = getColumnOfDate(cResult.getDate)

                    If farbTypenListe.Contains(cBewertung.colorIndex.ToString) Then
                        ' dann muss ein Eintrag in der Tabelle gemacht werden 

                        If (resultColumn < showRangeLeft Or resultColumn > showRangeRight) Then
                            ' nichts machen 
                        Else
                            ' hier die Tabellen-Einträge machen 

                            With tabelle

                                CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = msNumber.ToString
                                CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(cBewertung.color)
                                CType(.Cell(tabellenzeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = hproj.name
                                CType(.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cResult.name
                                CType(.Cell(tabellenzeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cResult.getDate.ToShortDateString
                                CType(.Cell(tabellenzeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cBewertung.description


                            End With

                            msNumber = msNumber + 1
                            tabelle.Rows.Add()
                            tabellenzeile = tabellenzeile + 1

                        End If


                    End If



                Next

            Next




        Next

        Try
            tabelle.Rows(msNumber + 1).Delete()
        Catch ex As Exception

        End Try



    End Sub

    ''' <summary>
    ''' zeichnet die Projektgrafik mit den Meilensteinen 
    ''' </summary>
    ''' <param name="pptslide"></param>
    ''' <param name="pptShape"></param>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Sub zeichneProjektGrafik(ByRef pptslide As pptNS.Slide, ByRef pptShape As pptNS.Shape, ByVal hproj As clsProjekt, Optional ByVal selectedMilestones As Collection = Nothing)

        Dim rng As xlNS.Range
        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim pptSize As Single
        Dim newshapeRange As pptNS.ShapeRange
        Dim newShape As pptNS.Shape
        Dim originalTop As Single
        Dim originalLeft As Single


        pptSize = pptShape.TextFrame2.TextRange.Font.Size
        pptShape.TextFrame2.TextRange.Text = " "

        Dim minColumn As Integer, maxColumn As Integer
        minColumn = hproj.Start - 3
        If minColumn < 1 Then
            minColumn = 1
        End If

        maxColumn = hproj.Start + hproj.anzahlRasterElemente + 3

        ' set Gridlines to white 
        With appInstance.ActiveWindow
            .GridlineColor = RGB(255, 255, 255)
        End With

        Dim oldposition As Integer = hproj.tfZeile
        Dim projektShape As xlNS.Shape
        Dim allShapes As xlNS.Shapes
        Dim ptop As Double, pleft As Double, pwidth As Double, pheight As Double
        Dim number As Integer = 0
        Dim nameList As New Collection


        ' Änderung tk: damit nur die gewählten Milestones gezeichnet werden 
        If Not IsNothing(selectedMilestones) Then
            nameList = selectedMilestones
        End If


        Call awinDeleteProjectChildShapes(0)

        With CType(appInstance.Worksheets(arrWsNames(3)), xlNS.Worksheet)

            allShapes = .Shapes
            projektShape = allShapes.Item(hproj.name)

            ' Projekt-Shape wird jetzt in neue Zeile geschoben 
            Dim newzeile As Integer = ShowProjekte.maxZeile + 4
            hproj.tfZeile = newzeile

            hproj.CalculateShapeCoord(ptop, pleft, pwidth, pheight)
            With projektShape
                originalTop = .Top
                originalLeft = .Left
                .Top = CSng(ptop)
                .Left = CSng(pleft)
                '.Height = CSng(pheight)
                '.Width = CSng(pwidth)
            End With

            ' Änderung tk 22.11.15 das Status Symbol ist hier eigentlich nicht gut aufgehoben ... 
            'Call zeichneStatusSymbolInPlantafel(hproj, 0)
            ' das ist der aufruf, alle Meilensteine zu zeichnen, sie zu nummerieren;
            ' ausserdem wird die Kennung mitgegeben, dass dies für einen Report notwendig ist 
            Call zeichneMilestonesInProjekt(hproj, nameList, 4, 0, 0, True, number, True)


            rng = .Range(.Cells(newzeile, minColumn), .Cells(newzeile + 1, maxColumn))
            'rng.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
            rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture)
            newshapeRange = pptslide.Shapes.Paste
            newShape = newshapeRange.Item(1)

            Call awinDeleteProjectChildShapes(0)

            ' Shape wieder an die alte Position bringen 
            hproj.tfZeile = oldposition
            projektShape = allShapes.Item(hproj.name)
            With projektShape
                .Top = originalTop
                .Left = originalLeft
                '.Height = CSng(pheight)
                '.Width = CSng(pwidth)
            End With

        End With

        ' set back 
        With appInstance.ActiveWindow
            .GridlineColor = RGB(220, 220, 220)
        End With



        Dim ratio As Double
        ratio = pptShape.Height / pptShape.Width

        With newShape

            If ratio < .Height / .Width Then
                ' orientieren an width 
                .Width = CSng(pptShape.Width * 0.96)
                .Height = CSng(ratio * .Width)
                ' left anpassen
                .Top = CSng(pptShape.Top + 0.02 * pptShape.Height)
                .Left = CSng(pptShape.Left + 0.98 * (pptShape.Width - .Width) / 2)

            Else
                .Height = CSng(pptShape.Height * 0.96)
                .Width = CSng(.Height / ratio)
                ' top anpassen 
                .Left = CSng(pptShape.Left + 0.02 * pptShape.Width)
                .Top = CSng(pptShape.Top + 0.98 * (pptShape.Height - .Height) / 2)
            End If

        End With



    End Sub


    ''' <summary>
    ''' zeichnet in die Shape Tabelle die Projekt Termin-Änderungen 
    ''' übergeben wird cproj asl current Projekt
    ''' bproj als das Projekt, das den Stand zur Beauftragung repräsentierte
    ''' lproj als das Projekt, das den letzten Stand repräsentiert
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="cproj"></param>
    ''' <param name="bproj"></param>
    ''' <param name="lproj"></param>
    ''' <remarks></remarks>
    Sub zeichneProjektTerminAenderungen(ByRef pptShape As pptNS.Shape, ByVal cproj As clsProjekt, ByVal bproj As clsProjekt, ByVal lproj As clsProjekt)

        Dim heute As Date = Date.Now
        Dim index As Integer = 0
        Dim tabelle As pptNS.Table

        Try
            tabelle = pptShape.Table
        Catch ex As Exception
            Throw New Exception("Shape hat keine Tabelle")
        End Try

        pptShape.Title = ""

        Dim msNumber As Integer = 1
        Dim tabellenzeile As Integer = 1

        ' jetzt wird die Headerzeile geschrieben 
        Dim tmpstr(3) As String
        Dim title As String

        ' beauftragt am ... 
        Try
            title = CType(tabelle.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text
            tmpstr = title.Trim.Split(New Char() {CChar("#")}, 4)
            If Not IsNothing(bproj) Then
                title = tmpstr(0) & bproj.timeStamp.ToShortDateString & tmpstr(1)
            Else
                title = tmpstr(0) & "--" & tmpstr(1)
            End If

            CType(tabelle.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = title

        Catch ex As Exception

        End Try


        ' letzte Freigabe am ... 
        Try
            title = CType(tabelle.Cell(tabellenzeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text
            tmpstr = title.Trim.Split(New Char() {CChar("#")}, 4)
            If Not IsNothing(lproj) Then
                title = tmpstr(0) & lproj.timeStamp.ToShortDateString & tmpstr(1)
            Else
                title = tmpstr(0) & "--" & tmpstr(1)
            End If

            CType(tabelle.Cell(tabellenzeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = title

        Catch ex As Exception

        End Try

        ' aktueller stand vom  ... 
        Try

            title = CType(tabelle.Cell(tabellenzeile, 7), pptNS.Cell).Shape.TextFrame2.TextRange.Text
            tmpstr = title.Trim.Split(New Char() {CChar("#")}, 4)
            If Not IsNothing(cproj) Then
                title = tmpstr(0) & cproj.timeStamp.ToShortDateString & tmpstr(1)
            Else
                title = tmpstr(0) & "--" & tmpstr(1)
            End If

            CType(tabelle.Cell(tabellenzeile, 7), pptNS.Cell).Shape.TextFrame2.TextRange.Text = title

        Catch ex As Exception

        End Try


        ' jetzt wird die todoListe abgearbeitet 
        tabellenzeile = 2

        Try
            For p = 1 To cproj.CountPhases

                Dim cphase As clsPhase = cproj.getPhase(p)
                Dim phaseStart As Date = cproj.startDate.AddMonths(cphase.relStart - 1)

                Dim bphase As clsPhase
                Dim lphase As clsPhase
                Dim bdiff As Long, ldiff As Long
                Dim bphaseStart As Date
                Dim lphaseStart As Date


                Try
                    bphase = bproj.getPhaseByID(cphase.nameID)
                    bphaseStart = bproj.startDate.AddMonths(bphase.relStart - 1)
                Catch ex As Exception
                    bphase = Nothing
                End Try

                Try
                    lphase = lproj.getPhaseByID(cphase.nameID)
                    lphaseStart = lproj.startDate.AddMonths(lphase.relStart - 1)
                Catch ex As Exception
                    lphase = Nothing
                End Try



                For r = 1 To cphase.countMilestones
                    Dim cResult As clsMeilenstein = Nothing
                    Dim cBewertung As clsBewertung = Nothing

                    Dim bResult As clsMeilenstein = Nothing
                    Dim bbewertung As clsBewertung = Nothing


                    Dim lResult As clsMeilenstein = Nothing
                    Dim lbewertung As clsBewertung = Nothing
                    Dim bDate As Date, lDate As Date
                    Dim currentDate As Date

                    cResult = cphase.getMilestone(r)
                    currentDate = cResult.getDate

                    If IsNothing(bphase) Then
                    Else

                    End If

                    bResult = bphase.getMilestone(cResult.nameID)
                    If IsNothing(bResult) Then
                        bdiff = -9999
                    Else
                        bDate = bResult.getDate
                        bdiff = DateDiff(DateInterval.Day, bDate, currentDate)
                    End If


                    lResult = lphase.getMilestone(cResult.nameID)
                    If IsNothing(lResult) Then
                        ldiff = -9999
                    Else
                        lDate = lResult.getDate
                        ldiff = DateDiff(DateInterval.Day, lDate, currentDate)
                    End If


                    cBewertung = cResult.getBewertung(1)
                    'Try
                    '    cBewertung = cResult.getBewertung(1)
                    'Catch ex As Exception
                    '    cBewertung = New clsBewertung
                    'End Try

                    bbewertung = bResult.getBewertung(1)
                    'Try
                    '    bbewertung = bResult.getBewertung(1)
                    'Catch ex As Exception
                    '    bbewertung = New clsBewertung
                    'End Try

                    lbewertung = lResult.getBewertung(1)
                    'Try
                    '    lbewertung = lResult.getBewertung(1)
                    'Catch ex As Exception
                    '    lbewertung = New clsBewertung
                    'End Try

                    If bdiff <> 0 Or ldiff <> 0 Then
                        With tabelle

                            CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = msNumber.ToString
                            CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                            CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(cBewertung.color)

                            Try
                                CType(.Cell(tabellenzeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cResult.name
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "---"
                            End Try

                            ' Datum und Farbe für Beauftragung schreiben  
                            Try

                                CType(.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = bDate.ToShortDateString
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "---"
                            End Try

                            Try
                                CType(.Cell(tabellenzeile, 4), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(bbewertung.color)
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 4), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                            End Try


                            ' Datum und Farbe für letzter Stand schreiben  
                            Try

                                CType(.Cell(tabellenzeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = lDate.ToShortDateString
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "---"
                            End Try

                            Try
                                CType(.Cell(tabellenzeile, 6), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(lbewertung.color)
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 6), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                            End Try

                            ' Datum und Farbe für aktuellen Stand schreiben  
                            Try
                                CType(.Cell(tabellenzeile, 7), pptNS.Cell).Shape.TextFrame2.TextRange.Text = currentDate.ToShortDateString
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 7), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "---"
                            End Try

                            Try
                                CType(.Cell(tabellenzeile, 8), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(cBewertung.color)
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 8), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                            End Try



                        End With

                        msNumber = msNumber + 1
                        tabelle.Rows.Add()
                        tabellenzeile = tabellenzeile + 1
                    End If



                Next

            Next

            Try
                tabelle.Rows(msNumber + 1).Delete()
            Catch ex1 As Exception

            End Try

        Catch ex As Exception
            Throw New Exception("Tabelle Projektänderungen hat evtl unzulässige Anzahl Zeilen / Spalten ...")
        End Try




    End Sub

    ''' <summary>
    ''' füllt die Vergleichs-Tabelle aus und setzt die entsprechenden Trend-Markierungen gleich, fallend, steigend ein  
    ''' </summary>
    ''' <param name="pptShape">Adresse auf die Tabelle, die ausgefükllt werden soll </param>
    ''' <param name="gleichShape">Shape für gleich</param>
    ''' <param name="steigendShape">shape für steigend</param>
    ''' <param name="fallendShape">shape für fallend</param>
    ''' <param name="hproj">aktuelles Projekt</param>
    ''' <param name="vglproj">letzter Stand</param>
    ''' <remarks></remarks>
    Sub zeichneProjektTabelleVergleich(ByRef pptslide As pptNS.Slide, ByRef pptShape As pptNS.Shape, ByVal gleichShape As pptNS.Shape, ByVal steigendShape As pptNS.Shape, ByVal fallendShape As pptNS.Shape, _
                                           ByVal ampelShape As pptNS.Shape, ByVal sternShape As pptNS.Shape, ByVal hproj As clsProjekt, ByVal vglproj As clsProjekt)
        Dim anzZeilen As Integer
        Dim tabelle As pptNS.Table
        Dim zeile As Integer
        Dim tmpStr As String
        Dim tableLeft As Double = pptShape.Left
        Dim tableTop As Double = pptShape.Top
        Dim kennung As String
        Dim aktBudget As Double, vglBudget As Double
        Dim aktPersCost As Double, vglPersCost As Double
        Dim aktSonstCost As Double, vglSonstCost As Double
        Dim aktRiskCost As Double, vglRiskCost As Double
        Dim aktErgebnis As Double, vglErgebnis As Double
        Dim farbePositiv As Long
        Dim farbeNeutral As Long
        Dim farbeNegativ As Long
        Dim farbeStern As Long
        Dim unterschiede As New Collection
        Dim TimeCostColor(2) As Double
        Dim TimeTimeColor(2) As Double


        Try
            farbePositiv = steigendShape.Fill.ForeColor.RGB
            farbeNeutral = gleichShape.Fill.ForeColor.RGB
            farbeNegativ = fallendShape.Fill.ForeColor.RGB
            farbeStern = sternShape.Fill.ForeColor.RGB
        Catch ex As Exception

        End Try


        ' jetzt wird festgestellt, wo es über all Unterschiede gibt 
        ' wird für Bewertung Termine und Meilensteine benötigt 
        unterschiede = hproj.listOfDifferences(vglproj, True, 0)

        ' jetzt werden die aktuellen bzw Vergleichswerte der finanziellen KPIs bestimmt 
        Try
            hproj.calculateRoundedKPI(aktBudget, aktPersCost, aktSonstCost, aktRiskCost, aktErgebnis)

            If Not IsNothing(vglproj) Then
                vglproj.calculateRoundedKPI(vglBudget, vglPersCost, vglSonstCost, vglRiskCost, vglErgebnis)
            End If

        Catch ex As Exception



        End Try

        If CBool(pptShape.HasTable) Then
            tabelle = pptShape.Table
            anzZeilen = tabelle.Rows.Count
            If anzZeilen > 1 Then
                zeile = 1
                ' jetzt wird die Überschrift aktualisiert 
                With tabelle

                    CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Projekt" & vbLf & hproj.getShapeText

                    tmpStr = CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text
                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = tmpStr & vbLf & vglproj.timeStamp.ToShortDateString


                    ' jetzt werden die Zeilen abgearbeitet, beginnend mit 2
                    For zeile = 2 To anzZeilen

                        Try
                            kennung = CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text.Trim
                        Catch ex As Exception
                            kennung = ""
                        End Try

                        Dim aktvalue As Double
                        Dim vglValue As Double
                        Select Case kennung

                            Case "Ergebnis"

                                aktvalue = aktErgebnis
                                vglValue = vglErgebnis

                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = " nicht verfügbar"
                                Else
                                    If aktvalue = vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    ElseIf aktvalue > vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbePositiv)

                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbeNegativ)
                                    End If

                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = vglValue.ToString & " T€"
                                End If




                            Case "Budget"

                                aktvalue = aktBudget
                                vglValue = vglBudget

                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = " nicht verfügbar"
                                Else
                                    If aktvalue = vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    ElseIf aktvalue > vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbePositiv)

                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbeNegativ)
                                    End If

                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = vglValue.ToString & " T€"
                                End If




                            Case "Personalkosten"

                                aktvalue = aktPersCost
                                vglValue = vglPersCost

                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = " nicht verfügbar"
                                Else
                                    If aktvalue = vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    ElseIf aktvalue > vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ)

                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv)
                                    End If

                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = vglValue.ToString & " T€"
                                End If




                            Case "Sonstige Kosten"

                                aktvalue = aktSonstCost
                                vglValue = vglSonstCost

                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = " nicht verfügbar"
                                Else
                                    If aktvalue = vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    ElseIf aktvalue > vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ)

                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv)
                                    End If

                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = vglValue.ToString & " T€"
                                End If




                            Case "Termine Phasen"


                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "siehe folgende Charts"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "nicht verfügbar"
                                Else
                                    If unterschiede.Contains(CInt(PThcc.phasen).ToString) Then
                                        TimeTimeColor = hproj.getTimeTimeColor(vglproj, True, Date.Now)

                                        If TimeTimeColor(0) < 0 Then

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv, farbeNeutral)
                                            End If

                                        ElseIf TimeTimeColor(0) > 0 Then

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ, farbeNeutral)
                                            End If

                                        Else

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                            End If

                                        End If

                                        'Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, sternShape, farbeStern)
                                        CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "siehe folgende Charts"
                                        CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "siehe folgende Charts"
                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                        CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "identisch"
                                        CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "identisch"
                                    End If
                                End If



                            Case "Termine Meilensteine"


                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "siehe folgende Charts"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "nicht verfügbar"
                                Else
                                    If unterschiede.Contains(CInt(PThcc.resultdates).ToString) Or unterschiede.Contains(CInt(PThcc.resultampel).ToString) Then

                                        TimeTimeColor = hproj.getTimeTimeColor(vglproj, True, Date.Now)

                                        If TimeTimeColor(0) < 0 Then

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv, farbeNeutral)
                                            End If

                                        ElseIf TimeTimeColor(0) > 0 Then

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ, farbeNeutral)
                                            End If

                                        Else

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                            End If

                                        End If

                                        'Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, sternShape, farbeStern)
                                        CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "siehe folgende Charts"
                                        CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "siehe folgende Charts"
                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                        CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "identisch"
                                        CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "identisch"
                                    End If
                                End If



                            Case "Einschätzung strategischer Fit"

                                aktvalue = hproj.StrategicFit
                                vglValue = vglproj.StrategicFit

                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "nicht verfügbar"
                                Else
                                    If aktvalue = vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    ElseIf aktvalue > vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbePositiv)

                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbeNegativ)
                                    End If

                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = Format(aktvalue, "#.0")
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = Format(vglValue, "#.0")
                                End If




                            Case "Einschätzung Risiko"

                                aktvalue = hproj.Risiko
                                vglValue = vglproj.Risiko

                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "nicht verfügbar"
                                Else
                                    If aktvalue = vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    ElseIf aktvalue > vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ)

                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv)
                                    End If

                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = Format(aktvalue, "#.0")
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = Format(vglValue, "#.0")
                                End If




                            Case "Projekt-Ampel"

                                aktvalue = hproj.ampelStatus
                                vglValue = vglproj.ampelStatus
                                Dim tmpFarbe As Long

                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    If aktvalue = PTfarbe.red Then
                                        tmpFarbe = farbeNegativ
                                    ElseIf aktvalue = PTfarbe.green Then
                                        tmpFarbe = farbePositiv
                                    ElseIf aktvalue = PTfarbe.yellow Then
                                        tmpFarbe = awinSettings.AmpelGelb
                                    Else
                                        tmpFarbe = farbeNeutral
                                    End If

                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 3, ampelShape, tmpFarbe)
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "nicht verfügbar"
                                Else

                                    Dim aktFarbe As Long, vglFarbe As Long
                                    If aktvalue = PTfarbe.red Then
                                        aktFarbe = farbeNegativ
                                    ElseIf aktvalue = PTfarbe.green Then
                                        aktFarbe = farbePositiv
                                    ElseIf aktvalue = PTfarbe.yellow Then
                                        aktFarbe = awinSettings.AmpelGelb
                                    Else
                                        aktFarbe = farbeNeutral
                                    End If

                                    If vglValue = PTfarbe.red Then
                                        vglFarbe = farbeNegativ
                                    ElseIf vglValue = PTfarbe.green Then
                                        vglFarbe = farbePositiv
                                    ElseIf vglValue = PTfarbe.yellow Then
                                        vglFarbe = awinSettings.AmpelGelb
                                    Else
                                        vglFarbe = farbeNeutral
                                    End If


                                    If aktvalue = vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    ElseIf aktvalue > vglValue Then

                                        If aktvalue = 1 Then
                                            Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbePositiv)
                                        Else
                                            Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbeNegativ)
                                        End If

                                    Else

                                        If aktvalue = 0 Then
                                            Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbeNegativ)
                                        Else
                                            Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbePositiv)
                                        End If

                                    End If

                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 4, ampelShape, aktFarbe)
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 3, ampelShape, vglFarbe)

                                End If


                            Case "Projekt-Ampel Erläuterung"

                                CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = hproj.ampelErlaeuterung
                                CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = vglproj.ampelErlaeuterung

                            Case Else


                        End Select


                    Next

                End With

            End If
        End If

    End Sub

    ''' <summary>
    ''' analog zu TabelleVergleich, allerdigs reduziert auf die drei Elemente Kosten, zeit, Qualität
    ''' </summary>
    ''' <param name="pptslide"></param>
    ''' <param name="pptShape"></param>
    ''' <param name="gleichShape"></param>
    ''' <param name="steigendShape"></param>
    ''' <param name="fallendShape"></param>
    ''' <param name="ampelShape"></param>
    ''' <param name="sternShape"></param>
    ''' <param name="hproj"></param>
    ''' <param name="vglproj"></param>
    ''' <remarks></remarks>
    Sub zeichneProjektTabelleOneGlance(ByRef pptslide As pptNS.Slide, ByRef pptShape As pptNS.Shape, ByVal gleichShape As pptNS.Shape, ByVal steigendShape As pptNS.Shape, ByVal fallendShape As pptNS.Shape, _
                                               ByVal ampelShape As pptNS.Shape, ByVal sternShape As pptNS.Shape, ByVal hproj As clsProjekt, ByVal vglproj As clsProjekt)
        Dim anzZeilen As Integer
        Dim tabelle As pptNS.Table
        Dim zeile As Integer
        Dim tmpStr As String
        Dim tableLeft As Double = pptShape.Left
        Dim tableTop As Double = pptShape.Top
        Dim kennung As String
        Dim aktBudget As Double, vglBudget As Double
        Dim aktPersCost As Double, vglPersCost As Double
        Dim aktSonstCost As Double, vglSonstCost As Double
        Dim aktRiskCost As Double, vglRiskCost As Double
        Dim aktErgebnis As Double, vglErgebnis As Double
        Dim farbePositiv As Long
        Dim farbeNeutral As Long
        Dim farbeNegativ As Long
        Dim farbeStern As Long
        Dim unterschiede As New Collection
        Dim TimeCostColor(2) As Double
        Dim TimeTimeColor(2) As Double


        Try
            farbePositiv = steigendShape.Fill.ForeColor.RGB
            farbeNeutral = gleichShape.Fill.ForeColor.RGB
            farbeNegativ = fallendShape.Fill.ForeColor.RGB
            farbeStern = sternShape.Fill.ForeColor.RGB
        Catch ex As Exception

        End Try


        ' jetzt wird festgestellt, wo es über all Unterschiede gibt 
        ' wird für Bewertung Termine und Meilensteine benötigt 
        unterschiede = hproj.listOfDifferences(vglproj, True, 0)

        ' jetzt werden die aktuellen bzw Vergleichswerte der finanziellen KPIs bestimmt 
        Try
            hproj.calculateRoundedKPI(aktBudget, aktPersCost, aktSonstCost, aktRiskCost, aktErgebnis)

            If Not IsNothing(vglproj) Then
                vglproj.calculateRoundedKPI(vglBudget, vglPersCost, vglSonstCost, vglRiskCost, vglErgebnis)
            End If

        Catch ex As Exception



        End Try

        If CBool(pptShape.HasTable) Then
            tabelle = pptShape.Table
            anzZeilen = tabelle.Rows.Count
            If anzZeilen > 1 Then
                zeile = 1
                ' jetzt wird die Überschrift aktualisiert 
                With tabelle

                    CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Projekt" & vbLf & hproj.getShapeText

                    tmpStr = CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text
                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = tmpStr & vbLf & vglproj.timeStamp.ToShortDateString


                    ' jetzt werden die Zeilen abgearbeitet, beginnend mit 2
                    For zeile = 2 To anzZeilen

                        Try
                            kennung = CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text.Trim
                        Catch ex As Exception
                            kennung = ""
                        End Try

                        Dim aktvalue As Double
                        Dim vglValue As Double
                        Select Case kennung

                            Case "Gesamtkosten"

                                aktvalue = aktPersCost + aktSonstCost
                                vglValue = vglPersCost + aktSonstCost

                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = " nicht verfügbar"
                                Else
                                    If aktvalue = vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                    ElseIf aktvalue > vglValue Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ)

                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv)
                                    End If

                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = aktvalue.ToString & " T€"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = vglValue.ToString & " T€"
                                End If







                            Case "Termine"


                                If IsNothing(vglproj) Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                    CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "siehe folgende Charts"
                                    CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "nicht verfügbar"
                                Else
                                    If unterschiede.Contains(CInt(PThcc.phasen).ToString) Or unterschiede.Contains(CInt(PThcc.resultdates).ToString) Then
                                        TimeTimeColor = hproj.getTimeTimeColor(vglproj, True, Date.Now)

                                        If TimeTimeColor(0) < 0 Then

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbePositiv, farbeNeutral)
                                            End If

                                        ElseIf TimeTimeColor(0) > 0 Then

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbeNegativ, farbeNeutral)
                                            End If

                                        Else

                                            If TimeTimeColor(1) < 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral, farbePositiv)
                                            ElseIf TimeTimeColor(1) > 0 Then
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral, farbeNegativ)
                                            Else
                                                Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                            End If

                                        End If

                                        'Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, sternShape, farbeStern)
                                        CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Ende: " & hproj.endeDate.ToShortDateString
                                        CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Ende: " & vglproj.endeDate.ToShortDateString
                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)
                                        CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Ende: " & hproj.endeDate.ToShortDateString
                                        CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Ende: " & vglproj.endeDate.ToShortDateString
                                    End If
                                End If


                            Case "Erläuterung"

                                aktvalue = hproj.ampelStatus
                                vglValue = vglproj.ampelStatus


                                If aktvalue = vglValue Then
                                    Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, gleichShape, farbeNeutral)

                                ElseIf aktvalue > vglValue Then

                                    If aktvalue = 1 Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbePositiv)
                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbeNegativ)
                                    End If

                                Else

                                    If aktvalue = 0 Then
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, fallendShape, farbeNegativ)
                                    Else
                                        Call zeichneTrendSymbol(pptslide, tabelle, zeile, 2, steigendShape, farbePositiv)
                                    End If

                                End If

                                CType(.Cell(zeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = hproj.ampelErlaeuterung
                                CType(.Cell(zeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = vglproj.ampelErlaeuterung

                            Case Else


                        End Select


                    Next

                End With

            End If
        End If

    End Sub

    ''' <summary>
    ''' zeichnet das übergebene Symbol in die per zeile, spalte angegebene Tabellen-Zelle
    ''' </summary>
    ''' <param name="pptslide"></param>
    ''' <param name="tabelle"></param>
    ''' <param name="tbZeile"></param>
    ''' <param name="tbSpalte"></param>
    ''' <param name="zeichen"></param>
    ''' <param name="farbkennung"></param>
    ''' <remarks></remarks>
    Sub zeichneTrendSymbol(ByRef pptslide As pptNS.Slide, ByRef tabelle As pptNS.Table, ByVal tbZeile As Integer, ByVal tbSpalte As Integer, _
                                ByVal zeichen As pptNS.Shape, ByVal farbkennung As Long)

        Dim korrFaktor As Double = 1.0
        Dim newZeichen As pptNS.ShapeRange

        zeichen.Copy()
        newZeichen = pptslide.Shapes.Paste

        ' ist der Pfeil größer als die Zelle ? 
        If tabelle.Cell(tbZeile, tbSpalte).Shape.Width < newZeichen(1).Width Or _
             tabelle.Cell(tbZeile, tbSpalte).Shape.Height < newZeichen(1).Height Then
            ' dann am kleineren orientieren 

            Try
                korrFaktor = System.Math.Min(tabelle.Cell(tbZeile, tbSpalte).Shape.Width / newZeichen(1).Width, tabelle.Cell(tbZeile, tbSpalte).Shape.Height / newZeichen(1).Height)
            Catch ex As Exception
                ' in diesem Fall bleibt Korrfaktor auf 1.0 
            End Try


        End If

        ' Anpassen derPfeilgröße
        If korrFaktor < 1.0 Then

            korrFaktor = korrFaktor * 0.98

            With newZeichen(1)
                .Width = CSng(korrFaktor * .Width)
                .Height = CSng(korrFaktor * .Height)
            End With

        End If

        ' jetzt bestimmen der Left , Top Koordinaten des Pfeils und setzen der Farbe

        With newZeichen(1)

            .Top = tabelle.Cell(tbZeile, tbSpalte).Shape.Top + (tabelle.Cell(tbZeile, tbSpalte).Shape.Height - .Height) / 2
            .Left = tabelle.Cell(tbZeile, tbSpalte).Shape.Left + (tabelle.Cell(tbZeile, tbSpalte).Shape.Width - .Width) / 2
            .Fill.ForeColor.RGB = CInt(farbkennung)

        End With



    End Sub


    ''' <summary>
    ''' ergänzt zum Symbol noch die Linenfarbe als Hinweis wie der nächste Meilenstein aussieht 
    ''' </summary>
    ''' <param name="pptslide"></param>
    ''' <param name="tabelle"></param>
    ''' <param name="tbZeile"></param>
    ''' <param name="tbSpalte"></param>
    ''' <param name="zeichen"></param>
    ''' <param name="farbkennung"></param>
    ''' <param name="lineColor"></param>
    ''' <remarks></remarks>
    Sub zeichneTrendSymbol(ByRef pptslide As pptNS.Slide, ByRef tabelle As pptNS.Table, ByVal tbZeile As Integer, ByVal tbSpalte As Integer, _
                                    ByVal zeichen As pptNS.Shape, ByVal farbkennung As Long, ByVal lineColor As Long)

        Dim korrFaktor As Double = 1.0
        Dim newZeichen As pptNS.ShapeRange

        zeichen.Copy()
        newZeichen = pptslide.Shapes.Paste

        ' ist der Pfeil größer als die Zelle ? 
        If tabelle.Cell(tbZeile, tbSpalte).Shape.Width < newZeichen(1).Width Or _
             tabelle.Cell(tbZeile, tbSpalte).Shape.Height < newZeichen(1).Height Then
            ' dann am kleineren orientieren 

            Try
                korrFaktor = System.Math.Min(tabelle.Cell(tbZeile, tbSpalte).Shape.Width / newZeichen(1).Width, tabelle.Cell(tbZeile, tbSpalte).Shape.Height / newZeichen(1).Height)
            Catch ex As Exception
                ' in diesem Fall bleibt Korrfaktor auf 1.0 
            End Try


        End If

        ' Anpassen derPfeilgröße
        If korrFaktor < 1.0 Then

            korrFaktor = korrFaktor * 0.98

            With newZeichen(1)
                .Width = CSng(korrFaktor * .Width)
                .Height = CSng(korrFaktor * .Height)
            End With

        End If

        ' jetzt bestimmen der Left , Top Koordinaten des Pfeils und setzen der Farbe

        With newZeichen(1)

            .Top = tabelle.Cell(tbZeile, tbSpalte).Shape.Top + (tabelle.Cell(tbZeile, tbSpalte).Shape.Height - .Height) / 2
            .Left = tabelle.Cell(tbZeile, tbSpalte).Shape.Left + (tabelle.Cell(tbZeile, tbSpalte).Shape.Width - .Width) / 2
            .Fill.ForeColor.RGB = CInt(farbkennung)
            .Line.ForeColor.RGB = CInt(lineColor)
            .Line.Weight = 2

        End With



    End Sub

    ''' <summary>
    ''' zeichnet die Tabelle mit den Meilensteinen
    ''' wenn eine Collection mit den Namen übergeben wird, dann werden nur die Meilensteine mit diesen Namen betrachtet 
    ''' </summary>
    ''' <param name="pptShape"></param>
    ''' <param name="hproj"></param>
    ''' <param name="selectedItems"></param>
    ''' <remarks></remarks>
    Sub zeichneProjektTabelleZiele(ByRef pptShape As pptNS.Shape, ByVal hproj As clsProjekt, Optional ByVal selectedItems As Collection = Nothing)

        Dim heute As Date = Date.Now
        Dim anzSpalten As Integer = 0
        Dim index As Integer = 0
        Dim tabelle As pptNS.Table
        Dim todoCollection As Collection = hproj.getAllElemIDs(True)

        If IsNothing(selectedItems) Then
            todoCollection = hproj.getAllElemIDs(True)
        ElseIf selectedItems.Count = 0 Then
            todoCollection = hproj.getAllElemIDs(True)
        Else
            todoCollection = hproj.getElemIdsOf(selectedItems, True)
        End If

        Try
            tabelle = pptShape.Table
            anzSpalten = tabelle.Columns.Count

        Catch ex As Exception
            Throw New Exception("Shape hat keine Tabelle")
        End Try

        If anzSpalten < 4 Then
            Throw New Exception("Shape hat zu wenige Spalten (min 4) ")
        Else
            pptShape.Title = ""

            Dim msNumber As Integer = 1

            ' jetzt wird die todoListe abgearbeitet 
            Dim tabellenzeile As Integer = 2

            Try

                For m = 1 To todoCollection.Count
                    Dim cResult As clsMeilenstein = hproj.getMilestoneByID(CStr(todoCollection.Item(m)))
                    Dim cBewertung As clsBewertung = cResult.getBewertung(1)

                    With tabelle

                        CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = msNumber.ToString
                        CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                        CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(cBewertung.color)

                        CType(.Cell(tabellenzeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cResult.name
                        CType(.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cResult.getDate.ToShortDateString
                        CType(.Cell(tabellenzeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cBewertung.deliverables
                        If anzSpalten >= 5 Then
                            CType(.Cell(tabellenzeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cBewertung.description
                        End If


                    End With

                    msNumber = msNumber + 1
                    tabelle.Rows.Add()
                    tabellenzeile = tabellenzeile + 1

                Next


                Try
                    tabelle.Rows(msNumber + 1).Delete()
                Catch ex1 As Exception

                End Try

            Catch ex As Exception
                Throw New Exception("Tabelle Projektziele hat evtl unzulässige Anzahl Zeilen / Spalten ...")
            End Try

        End If




    End Sub

    ''' <summary>
    ''' schreibt für jedes Projekt, das Abhängigkeiten hat, diese in eine Tabelle
    ''' </summary>
    ''' <param name="pptshape"></param>
    ''' <remarks></remarks>
    Sub zeichneTabelleProjektabhaengigkeiten(ByRef pptshape As pptNS.Shape)

        Dim heute As Date = Date.Now
        Dim index As Integer = 0
        Dim tabelle As pptNS.Table

        Try
            tabelle = pptshape.Table
        Catch ex As Exception
            Throw New Exception("Shape hat keine Tabelle")
        End Try


        Dim todoListe As New SortedList(Of Integer, clsProjekt)

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            Dim anzDependencies As Integer = allDependencies.activeNumber(kvp.Value.name, PTdpndncyType.inhalt)

            If anzDependencies > 0 Then

                todoListe.Add(kvp.Value.tfZeile, kvp.Value)

            End If

        Next

        ' jetzt wird die todoListe abgearbeitet 
        Dim tabellenzeile As Integer = 2
        Dim msNumber As Integer = 1


        For Each kvp As KeyValuePair(Of Integer, clsProjekt) In todoListe

            Dim depListe As Collection = allDependencies.activeListe(kvp.Value.name, PTdpndncyType.inhalt)
            Dim ergebnisString As String = ""

            With tabelle

                If kvp.Value.ampelStatus = 0 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                ElseIf kvp.Value.ampelStatus = 1 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelGruen)
                ElseIf kvp.Value.ampelStatus = 2 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelGelb)
                ElseIf kvp.Value.ampelStatus = 3 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                Else
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                End If

                For i As Integer = 1 To depListe.Count
                    If i = 1 Then
                        ergebnisString = CStr(depListe.Item(i)).Trim
                    Else
                        ergebnisString = ergebnisString & "; " & CStr(depListe.Item(i)).Trim
                    End If
                Next

                CType(.Cell(tabellenzeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = kvp.Value.name
                CType(.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = ergebnisString


            End With
            msNumber = msNumber + 1
            tabelle.Rows.Add()
            tabellenzeile = tabellenzeile + 1

        Next

        Try
            tabelle.Rows(msNumber + 1).Delete()
        Catch ex As Exception

        End Try

    End Sub

    Sub zeichneTabelleStatus(ByRef pptshape As pptNS.Shape)
        Dim heute As Date = Date.Now
        Dim index As Integer = 0
        Dim tabelle As pptNS.Table

        Try
            tabelle = pptshape.Table
        Catch ex As Exception
            Throw New Exception("Shape hat keine Tabelle")
        End Try


        Dim todoListe As New SortedList(Of Integer, clsProjekt)

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste


            If istLaufendesProjekt(kvp.Value) Then

                todoListe.Add(kvp.Value.tfZeile, kvp.Value)

            End If

        Next

        ' jetzt wird die todoListe abgearbeitet 
        Dim tabellenzeile As Integer = 2
        Dim msNumber As Integer = 1
        For Each kvp As KeyValuePair(Of Integer, clsProjekt) In todoListe

            With tabelle
                CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = msNumber.ToString
                CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

                If kvp.Value.ampelStatus = 0 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                ElseIf kvp.Value.ampelStatus = 1 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelGruen)
                ElseIf kvp.Value.ampelStatus = 2 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelGelb)
                ElseIf kvp.Value.ampelStatus = 3 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                Else
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                End If

                CType(.Cell(tabellenzeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = kvp.Value.name
                CType(.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = kvp.Value.ampelErlaeuterung


            End With
            msNumber = msNumber + 1
            tabelle.Rows.Add()
            tabellenzeile = tabellenzeile + 1

        Next

        Try
            tabelle.Rows(msNumber + 1).Delete()
        Catch ex As Exception

        End Try

    End Sub

    Sub zeichneProjektTabelleStatus(ByRef pptshape As pptNS.Shape, ByVal hproj As clsProjekt)

        Dim heute As Date = Date.Now
        Dim index As Integer = 0
        Dim tabelle As pptNS.Table

        Try
            tabelle = pptshape.Table
        Catch ex As Exception
            Throw New Exception("Shape hat keine Tabelle")
        End Try

        pptshape.Title = ""

        ' jetzt wird die Tabelle ausgefüllt 
        Dim tabellenzeile As Integer = 2

        Try
            With tabelle

                If hproj.ampelStatus = 0 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                ElseIf hproj.ampelStatus = 1 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelGruen)
                ElseIf hproj.ampelStatus = 2 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelGelb)
                ElseIf hproj.ampelStatus = 3 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                Else
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                End If

                CType(.Cell(tabellenzeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = hproj.ampelErlaeuterung

            End With
        Catch ex As Exception
            Throw New Exception("Anzahl der Zeilen oder Spalten in der Projektstatus Tabelle passt nicht ...")
        End Try



    End Sub

    ''' <summary>
    ''' berechnet die Koordinaten des Bild, von welcher Spalte bis zu welcher Spalte und Zeile der SchnappSchuss der Projekt-Tafel aufgenommen werden soll 
    ''' </summary>
    ''' <param name="myCollection"></param>
    ''' <param name="minColumn">Ausgabe Parameter: von Spalte</param>
    ''' <param name="maxColumn">Ausgabe Parameter: bis Spalte</param>
    ''' <param name="maxZeile">Ausgabe-Parameter: bis Zeile</param>
    ''' <remarks></remarks>
    Sub calcPictureCoord(ByVal myCollection As Collection, ByRef minColumn As Integer, ByRef maxColumn As Integer, ByRef maxZeile As Integer, ByVal toBeTrimmed As Boolean)

        Dim hproj As clsProjekt
        Dim von As Integer = showRangeLeft
        Dim bis As Integer = showRangeRight

        For Each pName In myCollection
            Try
                hproj = ShowProjekte.getProject(pName.ToString)
                With hproj
                    If .Start < minColumn Then
                        minColumn = .Start
                    End If

                    If .Start + .anzahlRasterElemente - 1 > maxColumn Then
                        maxColumn = .Start + .anzahlRasterElemente - 1
                    End If

                    If .tfZeile > maxZeile Then
                        maxZeile = .tfZeile
                    End If
                End With
            Catch ex As Exception

            End Try

        Next
        maxZeile = maxZeile + 1

        If minColumn > 1 Then
            minColumn = minColumn - 1
        End If
        maxColumn = maxColumn + 1


        If von > 0 And von < minColumn Then
            minColumn = von
        End If

        If bis > maxColumn Then
            maxColumn = bis
        End If


        If toBeTrimmed Then
            If minColumn < von - 12 Then
                minColumn = von - 12
            End If

            If maxColumn > bis + 12 Then
                maxColumn = bis + 12
            End If
        End If


    End Sub


    ''' <summary>
    ''' Portfolio - Diagramme erstellen gemäß dem angegebenen charttype
    ''' </summary>
    ''' <param name="ProjektListe"></param>
    ''' <param name="repChart"></param>
    ''' <param name="showAbsoluteDiff">sollen die Unterschiede absolut oder prozentual angezeigt werden</param>
    ''' <param name="vglTyp">0: vergleiche Projekt-Ende ; 1: vergleiche mit nächstem Meilenstein </param>
    ''' <param name="charttype">betterWorseL - Vergleich mit letztem Stand
    ''' betterWorseB - Vergleich mit Beauftragungs-Stand</param>
    ''' <param name="bubbleColor"></param>
    ''' <param name="showLabels"></param>
    ''' <param name="chartBorderVisible"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Sub awinCreateBetterWorsePortfolio(ByRef ProjektListe As Collection, ByRef repChart As Excel.ChartObject, ByVal showAbsoluteDiff As Boolean, ByVal isTimeTimeVgl As Boolean, ByVal vglTyp As Integer, _
                                             ByVal charttype As Integer, ByVal bubbleColor As Integer, ByVal bubbleValueTyp As Integer, _
                                             ByVal showLabels As Boolean, ByVal chartBorderVisible As Boolean, _
                                             ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double)


        Dim anzDiagrams As Integer, i As Integer
        Dim found As Boolean
        Dim pname As String = ""
        Dim hproj As New clsProjekt, vproj As clsProjekt = Nothing
        Dim anzBubbles As Integer
        Dim yAchsenValues() As Double
        Dim xAchsenValues() As Double
        Dim bubbleValues() As Double, tempArray() As Double
        Dim nameValues() As String
        Dim colorValues() As Object
        Dim positionValues() As String
        Dim diagramTitle As String = ""
        Dim pfDiagram As clsDiagramm
        Dim pfChart As clsEventsPfCharts
        'Dim chtTitle As String
        Dim hilfsstring As String = ""
        Dim chtobjName As String = windowNames(3)
        Dim smallfontsize As Double, titlefontsize As Double
        Dim singleProject As Boolean
        Dim outOfToleranceProjekte As New SortedList(Of String, Double())
        Dim vglName As String = ""
        Dim compareToLast As Boolean = True
        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        Dim variantName As String = ""
        Dim tolerancePercent As Double = 0.02
        Dim toleranceTimeAbs As Integer = 5
        Dim toleranceCostAbs As Integer = 2
        ' die Folgenden Werte nehmen die min/max Abweichungen in Time and Cost auf; dient dazu die Skalierung korrekt darzustellen 
        Dim minTime As Double, maxTime As Double
        Dim minTC As Double, maxTC As Double

        Dim relTimeTolerance As Double = awinSettings.timeToleranzRel
        Dim absTimeTolerance As Double = awinSettings.timeToleranzAbs
        Dim relCostTolerance As Double = awinSettings.costToleranzRel
        Dim absCostTolerance As Double = awinSettings.costToleranzAbs

        Dim timeTCColor(2) As Double
        Dim xAchsenNames(1) As String
        Dim yAchsenNames(1) As String
        Dim anzkeinVproj As Integer = 0

        xAchsenNames(0) = "langsamer"
        xAchsenNames(1) = "schneller"
        yAchsenNames(0) = "teurer"
        yAchsenNames(1) = "günstiger"
        minTime = 10000
        minTC = 10000

        Dim formerSU As Boolean = appInstance.ScreenUpdating
        Dim formerEE As Boolean = appInstance.EnableEvents

        appInstance.ScreenUpdating = False

        If ProjektListe.Count > 1 Then
            singleProject = False
        Else
            singleProject = True
        End If


        If width > 450 Then
            titlefontsize = 20
            smallfontsize = 10
        ElseIf width > 250 Then
            titlefontsize = 14
            smallfontsize = 8
        Else
            titlefontsize = 12
            smallfontsize = 8
        End If

        Dim tmpanz As Integer = projekthistorie.liste.Count

        Select Case charttype
            Case PTpfdk.betterWorseL

                compareToLast = True
                If showAbsoluteDiff Then
                    diagramTitle = "Absolute " & portfolioDiagrammtitel(PTpfdk.betterWorseL)
                Else
                    diagramTitle = "Prozentuale " & portfolioDiagrammtitel(PTpfdk.betterWorseL)
                End If


            Case PTpfdk.betterWorseB

                compareToLast = False
                If showAbsoluteDiff Then
                    diagramTitle = "Absolute " & portfolioDiagrammtitel(PTpfdk.betterWorseB)
                Else
                    diagramTitle = "Prozentuale " & portfolioDiagrammtitel(PTpfdk.betterWorseB)
                End If

        End Select

        ' in der Projektliste sind jetzt laufende Projekte; jetzt wird bestimmt, welche innerhalb der 
        ' nicht-der-Rede-wert Fraktion sind 

        Dim anzOK As Integer = 0
        For i = 1 To ProjektListe.Count
            pname = ProjektListe.Item(i).ToString

            Try
                hproj = ShowProjekte.getProject(pname)
                variantName = hproj.variantName
                If request.pingMongoDb() Then

                    projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pname, variantName:=variantName, _
                                                                storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                    If compareToLast Then
                        vproj = projekthistorie.Last
                    Else
                        vproj = projekthistorie.beauftragung
                    End If
                Else
                    Call MsgBox("Datenbank-Verbindung ist unterbrochen!" & vbLf & "Projekthistorie konnte nicht geladen werden")
                End If


                If Not IsNothing(vproj) Then


                    If isTimeTimeVgl Then
                        timeTCColor = hproj.getTimeTimeColor(vproj, showAbsoluteDiff, Date.Now)
                    Else
                        timeTCColor = hproj.getTimeCostColor(vproj, vglTyp, showAbsoluteDiff, Date.Now)
                    End If


                    If showAbsoluteDiff Then

                        If isTimeTimeVgl Then
                            If timeTCColor(0) > -1 * absTimeTolerance And timeTCColor(0) < absTimeTolerance _
                            And timeTCColor(1) > -1 * absTimeTolerance And timeTCColor(1) < absTimeTolerance Then
                                ' liegt im erlaubten Toleranz-Korridor 
                                anzOK = anzOK + 1
                            Else
                                If timeTCColor(0) < minTime Then
                                    minTime = timeTCColor(0)
                                End If
                                If timeTCColor(0) > maxTime Then
                                    maxTime = timeTCColor(0)
                                End If

                                If timeTCColor(1) < minTC Then
                                    minTC = timeTCColor(1)
                                End If
                                If timeTCColor(1) > maxTC Then
                                    maxTC = timeTCColor(1)
                                End If

                                outOfToleranceProjekte.Add(vproj.name, timeTCColor)
                            End If
                        Else

                            If timeTCColor(0) > -1 * absTimeTolerance And timeTCColor(0) < absTimeTolerance _
                             And timeTCColor(1) > -1 * absCostTolerance And timeTCColor(1) < absCostTolerance Then
                                ' liegt im erlaubten Toleranz-Korridor 
                                anzOK = anzOK + 1
                            Else
                                If timeTCColor(0) < minTime Then
                                    minTime = timeTCColor(0)
                                End If
                                If timeTCColor(0) > maxTime Then
                                    maxTime = timeTCColor(0)
                                End If

                                If timeTCColor(1) < minTC Then
                                    minTC = timeTCColor(1)
                                End If
                                If timeTCColor(1) > maxTC Then
                                    maxTC = timeTCColor(1)
                                End If

                                outOfToleranceProjekte.Add(vproj.name, timeTCColor)
                            End If


                        End If

                    Else

                        If isTimeTimeVgl Then

                            If timeTCColor(0) > 1 - relTimeTolerance And timeTCColor(0) < 1 + relTimeTolerance And _
                                timeTCColor(1) > 1 - relTimeTolerance And timeTCColor(1) < 1 + relTimeTolerance Then
                                ' liegt im erlaubten Toleranz-Korridor 
                                anzOK = anzOK + 1
                            Else
                                If timeTCColor(0) < minTime Then
                                    minTime = timeTCColor(0)
                                End If
                                If timeTCColor(0) > maxTime Then
                                    maxTime = timeTCColor(0)
                                End If

                                If timeTCColor(1) < minTC Then
                                    minTC = timeTCColor(1)
                                End If
                                If timeTCColor(1) > maxTC Then
                                    maxTC = timeTCColor(1)
                                End If

                                outOfToleranceProjekte.Add(vproj.name, timeTCColor)
                            End If

                        Else
                            If timeTCColor(0) > 1 - relTimeTolerance And timeTCColor(0) < 1 + relTimeTolerance And _
                                timeTCColor(1) > 1 - relCostTolerance And timeTCColor(1) < 1 + relCostTolerance Then
                                ' liegt im erlaubten Toleranz-Korridor 
                                anzOK = anzOK + 1
                            Else
                                If timeTCColor(0) < minTime Then
                                    minTime = timeTCColor(0)
                                End If
                                If timeTCColor(0) > maxTime Then
                                    maxTime = timeTCColor(0)
                                End If

                                If timeTCColor(1) < minTC Then
                                    minTC = timeTCColor(1)
                                End If
                                If timeTCColor(1) > maxTC Then
                                    maxTC = timeTCColor(1)
                                End If

                                outOfToleranceProjekte.Add(vproj.name, timeTCColor)
                            End If
                        End If


                    End If

                Else
                    anzkeinVproj = anzkeinVproj + 1

                End If

            Catch ex As Exception
                projekthistorie.clear()
            End Try

        Next



        ' hier werden die Werte bestimmt ...
        Try
            ReDim yAchsenValues(outOfToleranceProjekte.Count - 1)
            ReDim xAchsenValues(outOfToleranceProjekte.Count - 1)
            ReDim bubbleValues(outOfToleranceProjekte.Count - 1)
            ReDim nameValues(outOfToleranceProjekte.Count - 1)
            ReDim colorValues(outOfToleranceProjekte.Count - 1)
            ReDim PfChartBubbleNames(outOfToleranceProjekte.Count - 1)
            ReDim positionValues(outOfToleranceProjekte.Count - 1)
        Catch ex As Exception

            Throw New ArgumentException("Fehler in CreateBetterWorsePortfolio " & ex.Message)

        End Try


        anzBubbles = outOfToleranceProjekte.Count

        If anzBubbles = 0 Then
            Dim logMessage As String
            Dim tmpValue1 As Double, tmpValue2 As Double


            If isTimeTimeVgl Then
                If showAbsoluteDiff Then
                    logMessage = "es gibt keine Projekte mit Abweichungen, die größer als die tolerierten Werte sind" & vbLf & _
                                    "Zeit-Toleranz Projekt-Ende: +/-" & absTimeTolerance & " Tage" & vbLf & _
                                    "Zeit-Toleranz nächster Meilenstein: +/-" & absTimeTolerance & " Tage"
                Else
                    tmpValue1 = relTimeTolerance * 100
                    logMessage = "es gibt keine Projekte mit Abweichungen, die größer als die tolerierten Werte sind" & vbLf & _
                                    "Zeit-Toleranz Projekt-Ende: +/-" & tmpValue1.ToString("##0.#") & "%" & vbLf & _
                                    "Zeit-Toleranz nächster Meilenstein: +/-" & tmpValue1.ToString("##0.#") & "%"
                End If
            Else
                If showAbsoluteDiff Then
                    logMessage = "es gibt keine Projekte mit Abweichungen, die größer als die tolerierten Werte sind" & vbLf & _
                                    "Zeit-Toleranz: +/-" & absTimeTolerance & " Tage" & vbLf & _
                                    "Kosten-Toleranz: +/-" & absCostTolerance & " T€"
                Else
                    tmpValue1 = relTimeTolerance * 100
                    tmpValue2 = relCostTolerance * 100
                    logMessage = "es gibt keine Projekte mit Abweichungen, die größer als die tolerierten Werte sind" & vbLf & _
                                    "Zeit-Toleranz: +/-" & tmpValue1.ToString("##0.#") & "%" & vbLf & _
                                    "Kosten-Toleranz: +/-" & tmpValue2.ToString("##0.#") & "%"

                End If

            End If

            appInstance.ScreenUpdating = formerSU
            appInstance.EnableEvents = formerEE
            Throw New Exception(logMessage)

        End If


        ' neuer Typ: 8.3.14 Abhängigkeiten
        Dim tmpstr(10) As String                ' nur für Zeit/Risiko Chart erforderlich

        For i = 1 To outOfToleranceProjekte.Count


            Try
                pname = outOfToleranceProjekte.ElementAt(i - 1).Key
                timeTCColor = outOfToleranceProjekte.ElementAt(i - 1).Value
                hproj = ShowProjekte.getProject(pname)

                xAchsenValues(i - 1) = timeTCColor(0)
                yAchsenValues(i - 1) = timeTCColor(1)

                If bubbleColor = PTpfdk.ProjektFarbe Then

                    ' Projekttyp wird farblich gekennzeichent
                    colorValues(anzBubbles) = hproj.farbe

                Else ' bubbleColor ist AmpelFarbe

                    ' ProjektStatus wird farblich gekennzeichnet
                    Select Case CInt(timeTCColor(2))
                        Case PTfarbe.none
                            colorValues(i - 1) = awinSettings.AmpelNichtBewertet
                        Case PTfarbe.green
                            colorValues(i - 1) = awinSettings.AmpelGruen
                        Case PTfarbe.yellow
                            colorValues(i - 1) = awinSettings.AmpelGelb
                        Case PTfarbe.red
                            colorValues(i - 1) = awinSettings.AmpelRot
                    End Select
                End If

                nameValues(i - 1) = hproj.name
                PfChartBubbleNames(i - 1) = nameValues(i - 1)

                Select Case bubbleValueTyp

                    Case PTbubble.strategicFit
                        bubbleValues(i - 1) = hproj.StrategicFit
                        'PfChartBubbleNames(i - 1) = nameValues(i - 1) & _
                        '            " (" & Format(bubbleValues(i - 1), "##0.#") & ", "

                    Case PTbubble.depencencies
                        bubbleValues(i - 1) = allDependencies.activeIndex(hproj.name, PTdpndncyType.inhalt) + 1
                        'PfChartBubbleNames(i - 1) = nameValues(i - 1) & _
                        '            " (" & Format(bubbleValues(i - 1), "##0") & ", "

                    Case PTbubble.marge
                        bubbleValues(i - 1) = hproj.ProjectMarge
                        If bubbleValues(i - 1) = 0 Then
                            bubbleValues(i - 1) = 0.005
                        End If
                        'PfChartBubbleNames(i - 1) = nameValues(i - 1) & _
                        '            " (" & Format(bubbleValues(i - 1) * 100, "##0.#") & "%, "

                End Select

                'If showAbsoluteDiff Then
                '    PfChartBubbleNames(i - 1) = PfChartBubbleNames(i - 1) & _
                '                    Format(timeTCColor(0), "##0") & ", " & Format(timeTCColor(1), "##0.#") & " T€)"
                'Else
                '    PfChartBubbleNames(i - 1) = PfChartBubbleNames(i - 1) & _
                '                    Format(timeTCColor(0) * 100, "##0.#") & "%, " & Format(timeTCColor(1) * 100, "##0.#") & "%)"
                'End If


            Catch ex As Exception

            End Try
        Next


        chtobjName = calcChartKennung("pf", charttype, ProjektListe)



        ' bestimmen der besten Position für die Werte ...
        Dim labelPosition(4) As String
        labelPosition(0) = "oben"
        labelPosition(1) = "rechts"
        labelPosition(2) = "unten"
        labelPosition(3) = "links"
        labelPosition(4) = "mittig"

        For i = 0 To anzBubbles - 1

            positionValues(i) = pfchartIstFrei(i, xAchsenValues, yAchsenValues)

        Next



        With CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet)
            anzDiagrams = CType(.ChartObjects, Excel.ChartObjects).Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False

            While i <= anzDiagrams And Not found
                If chtobjName = CType(.ChartObjects(i), Excel.ChartObject).Name Then
                    found = True
                    repChart = CType(.ChartObjects(i), Excel.ChartObject)
                    Exit Sub
                Else
                    i = i + 1
                End If

            End While


            If anzBubbles = 0 Then
                Dim logMessage As String = ""
            Else

            End If

            ReDim tempArray(anzBubbles - 1)

            With appInstance.Charts.Add

                CType(.SeriesCollection, Excel.SeriesCollection).NewSeries()
                CType(.SeriesCollection, Excel.SeriesCollection).Item(1).Name = diagramTitle


                CType(.SeriesCollection, Excel.SeriesCollection).Item(1).ChartType = xlNS.XlChartType.xlBubble3DEffect


                For i = 1 To anzBubbles
                    tempArray(i - 1) = xAchsenValues(i - 1)
                Next i
                CType(.SeriesCollection, Excel.SeriesCollection).Item(1).XValues = tempArray

                For i = 1 To anzBubbles
                    tempArray(i - 1) = yAchsenValues(i - 1)
                Next i
                CType(.SeriesCollection, Excel.SeriesCollection).Item(1).Values = tempArray

                For i = 1 To anzBubbles
                    If bubbleValues(i - 1) < 0.01 And bubbleValues(i - 1) > -0.01 Then
                        tempArray(i - 1) = 0.01
                    ElseIf bubbleValues(i - 1) < 0 Then
                        ' negative Werte werden Positiv dargestellt mit roten Beschriftung siehe unten
                        tempArray(i - 1) = System.Math.Abs(bubbleValues(i - 1))
                    Else
                        tempArray(i - 1) = bubbleValues(i - 1)
                    End If
                Next i


                CType(.SeriesCollection, Excel.SeriesCollection).Item(1).BubbleSizes = tempArray



                Dim series1 As xlNS.Series = _
                        CType(.SeriesCollection(1),  _
                                xlNS.Series)
                Dim point1 As xlNS.Point = _
                            CType(series1.Points(1), xlNS.Point)


                For i = 1 To anzBubbles

                    With CType(CType(.SeriesCollection, Excel.SeriesCollection).Item(1).Points(i), xlNS.Point)

                        If showLabels Then
                            Try
                                .HasDataLabel = True

                                With .DataLabel
                                    .Text = PfChartBubbleNames(i - 1)
                                    '.Text = nameValues(i - 1)
                                    If singleProject Then
                                        .Font.Size = awinSettings.CPfontsizeItems + 4
                                    Else
                                        .Font.Size = awinSettings.CPfontsizeItems
                                    End If

                                    Select Case positionValues(i - 1)
                                        Case labelPosition(0)
                                            .Position = xlNS.XlDataLabelPosition.xlLabelPositionAbove
                                        Case labelPosition(1)
                                            .Position = xlNS.XlDataLabelPosition.xlLabelPositionRight
                                        Case labelPosition(2)
                                            .Position = xlNS.XlDataLabelPosition.xlLabelPositionBelow
                                        Case labelPosition(3)
                                            .Position = xlNS.XlDataLabelPosition.xlLabelPositionLeft
                                        Case Else
                                            .Position = xlNS.XlDataLabelPosition.xlLabelPositionCenter
                                    End Select
                                End With
                            Catch ex As Exception

                            End Try
                        Else

                            Try
                                With .DataLabel
                                    .Text = PfChartBubbleNames(i - 1)
                                End With
                            Catch ex As Exception

                            End Try
                            .HasDataLabel = False
                        End If

                        .Interior.Color = colorValues(i - 1)

                        ' bei negativen Werten erfolgt die Beschriftung in roter Farbe  ..
                        If bubbleValues(i - 1) < 0 Then
                            .DataLabel.Font.Color = awinSettings.AmpelRot
                        End If

                    End With
                Next i


                '.ChartGroups(1).BubbleScale = sollte in Abhängigkeit der width gemacht werden 


                With CType(.ChartGroups(1), Excel.ChartGroup)

                    If singleProject Then
                        .BubbleScale = 20
                    Else
                        .BubbleScale = 20
                    End If

                    .SizeRepresents = xlNS.XlSizeRepresents.xlSizeIsArea
                    .ShowNegativeBubbles = True

                End With



                .HasAxis(xlNS.XlAxisType.xlCategory) = True
                .HasAxis(xlNS.XlAxisType.xlValue) = True

                With CType(.Axes(xlNS.XlAxisType.xlCategory), xlNS.Axis)

                    .HasMajorGridlines = False
                    .HasTitle = True

                    With .AxisTitle
                        If isTimeTimeVgl Then
                            .Characters.Text = "Zeit-Abweichung bis nächster Meilenstein"
                        Else
                            .Characters.Text = "Zeit-Abweichung bis nächster Meilenstein"
                        End If

                        .Characters.Font.Size = titlefontsize
                        .Characters.Font.Bold = False
                    End With

                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .Bold = False
                        .Size = awinSettings.fontsizeItems - 2

                    End With


                    If showAbsoluteDiff Then

                        .MajorUnit = 10.0
                        .CrossesAt = 0.0
                        .MinimumScale = minTime - 10
                        .MaximumScale = maxTime + 10


                    Else
                        .MajorUnit = 0.25
                        .CrossesAt = 1.0

                        If minTime > 0.5 Then
                            .MinimumScale = 0.5
                        Else
                            .MinimumScale = minTime - 0.1
                            If .MinimumScale <= 0 Then
                                .MinimumScale = 0.05
                            End If
                        End If

                        If maxTime < 1.5 Then
                            .MaximumScale = 1.5
                        Else
                            .MaximumScale = maxTime + 0.1
                        End If

                    End If

                    .MajorTickMark = Excel.XlTickMark.xlTickMarkCross
                    .TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNextToAxis

                End With

                With CType(.Axes(xlNS.XlAxisType.xlValue), xlNS.Axis)

                    .HasMajorGridlines = False
                    .HasTitle = True

                    With .AxisTitle

                        ' beschriftet werden muss sie mit Zeit werden, weil es intuitiv besser verstehbar ist
                        ' Achsen schneiden sich bei 1
                        '.Characters.text = "Kosten"
                        If isTimeTimeVgl Then
                            .Characters.Text = "Zeitabweichung Projektende"
                        Else
                            .Characters.Text = "Kosten-Abweichung"
                        End If

                        .Characters.Font.Size = titlefontsize
                        .Characters.Font.Bold = False
                    End With


                    With .TickLabels.Font
                        .FontStyle = "Normal"
                        .Bold = False
                        .Size = awinSettings.fontsizeItems - 2
                    End With


                    If showAbsoluteDiff Then
                        .MajorUnit = 10.0
                        .CrossesAt = 0.0
                        .MinimumScale = minTC - 10
                        .MaximumScale = maxTC + 10
                    Else
                        .MajorUnit = 0.25
                        .CrossesAt = 1.0

                        If minTC > 0.5 Then
                            .MinimumScale = 0.5
                        Else
                            .MinimumScale = minTC - 0.1
                            If .MinimumScale <= 0 Then
                                .MinimumScale = 0.05
                            End If
                        End If

                        If maxTC < 1.5 Then
                            .MaximumScale = 1.5
                        Else
                            .MaximumScale = maxTC + 0.1
                        End If


                    End If

                    .MajorTickMark = Excel.XlTickMark.xlTickMarkCross
                    .TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNextToAxis

                End With



                If anzkeinVproj > 0 Then

                    If showAbsoluteDiff Then

                        If isTimeTimeVgl Then
                            diagramTitle = diagramTitle & vbLf & _
                            anzOK.ToString & " Projekte innerhalb der Toleranz (+/-" & absTimeTolerance & " Tage)" & _
                            anzkeinVproj & " Projekte ohne letzten Stand"
                        Else
                            diagramTitle = diagramTitle & vbLf & _
                            anzOK.ToString & " Projekte innerhalb der Toleranz (+/-" & absTimeTolerance & " Tage, +/-" & absCostTolerance & " T€), " & _
                            anzkeinVproj & " Projekte ohne letzten Stand"
                        End If

                    Else
                        If isTimeTimeVgl Then
                            Dim tmpValue1 As Double = relTimeTolerance * 100
                            diagramTitle = diagramTitle & vbLf & _
                            anzOK.ToString & " Projekte innerhalb der Toleranz (+/-" & tmpValue1.ToString("##0.#") & "%)" & _
                            anzkeinVproj & " Projekte ohne letzten Stand"
                        Else
                            Dim tmpValue1 As Double = relTimeTolerance * 100
                            Dim tmpvalue2 As Double = relCostTolerance * 100
                            diagramTitle = diagramTitle & vbLf & _
                            anzOK.ToString & " Projekte innerhalb der Toleranz (+/-" & tmpValue1.ToString("##0.#") & "%, +/-" & tmpvalue2.ToString("##0.#") & "%), " & _
                            anzkeinVproj & " Projekte ohne letzten Stand"
                        End If
                    End If

                Else
                    If showAbsoluteDiff Then

                        If isTimeTimeVgl Then
                            diagramTitle = diagramTitle & vbLf & _
                            anzOK.ToString & " Projekte innerhalb der Toleranz (+/-" & absTimeTolerance & " Tage)"
                        Else
                            diagramTitle = diagramTitle & vbLf & _
                            anzOK.ToString & " Projekte innerhalb der Toleranz (+/-" & absTimeTolerance & " Tage, +/-" & absCostTolerance & " T€), "
                        End If

                    Else
                        If isTimeTimeVgl Then
                            Dim tmpValue1 As Double = relTimeTolerance * 100
                            diagramTitle = diagramTitle & vbLf & _
                            anzOK.ToString & " Projekte innerhalb der Toleranz (+/-" & tmpValue1.ToString("##0.#") & "%)"
                        Else
                            Dim tmpValue1 As Double = relTimeTolerance * 100
                            Dim tmpvalue2 As Double = relCostTolerance * 100
                            diagramTitle = diagramTitle & vbLf & _
                            anzOK.ToString & " Projekte innerhalb der Toleranz (+/-" & tmpValue1.ToString("##0.#") & "%, +/-" & tmpvalue2.ToString("##0.#") & "%), "
                        End If
                    End If
                End If




                .HasLegend = False
                .HasTitle = True
                .ChartTitle.Text = diagramTitle
                .ChartTitle.Characters.Font.Size = awinSettings.fontsizeTitle

                ' Events disablen, wegen Report erstellen
                appInstance.EnableEvents = False
                .Location(Where:=xlNS.XlChartLocation.xlLocationAsObject, _
                          Name:=CType(appInstance.Worksheets(arrWsNames(3)), Excel.Worksheet).Name)
                appInstance.EnableEvents = formerEE
                ' Events sind wieder zurückgesetzt
            End With


            'appInstance.ShowChartTipNames = False
            'appInstance.ShowChartTipValues = False

            With CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)
                .Top = top
                .Left = left
                .Width = width
                .Height = height
                .Name = chtobjName
            End With



            With appInstance.ActiveSheet
                Try
                    With CType(appInstance.ActiveSheet, Excel.Worksheet)
                        CType(.Shapes(chtobjName), Excel.Shape).Line.Visible = CType(chartBorderVisible, Microsoft.Office.Core.MsoTriState)
                    End With
                Catch ex As Exception

                End Try
            End With

            pfDiagram = New clsDiagramm

            pfChart = New clsEventsPfCharts
            pfChart.PfChartEvents = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject).Chart

            pfDiagram.setDiagramEvent = pfChart

            With pfDiagram

                .kennung = calcChartKennung("pf", charttype, ProjektListe)
                .DiagrammTitel = diagramTitle
                .diagrammTyp = DiagrammTypen(3)                     ' Portfolio
                .gsCollection = ProjektListe
                .isCockpitChart = False

            End With

            DiagramList.Add(pfDiagram)
            repChart = CType(.ChartObjects(anzDiagrams + 1), Excel.ChartObject)

        End With

        appInstance.ScreenUpdating = formerSU

    End Sub  ' Ende Prozedur awinCreatePortfolioChartDiagramm

    ''' <summary>
    ''' bestimmt für den angegebenen Zeitraum die Projekte, die eine der angegeben Phasen oder Meilensteine im Zeitraum enthalten. 
    ''' bestimmt darüber hinaus das minimale bzw. maximale Datum , das die Phasen der Projekte aufspannen , die den Zeitraum "berühren"  
    ''' </summary>
    ''' <param name="selectedPhases">die Phasen, nach denen gesúcht wird </param>
    ''' <param name="selectedMilestones">die Meilensteine, nach denen gesucht wird</param>
    ''' <param name="selectedRoles">die Rollen die gezeigt werden sollen; aktuell nicht relevant</param>
    ''' <param name="selectedCosts" >die Kostenarten, die gezeigt werden sollen; aktuell nicht relevant</param>
    ''' <param name="selectedBUs" >die Produktlininen bzw BusinessUnits, die gezeigt werden sollen</param>
    ''' <param name="selectedTyps">die Vorlagen, die gezeigt werden sollen</param>
    ''' <param name="von">linker Rand des Zeitraums</param>
    ''' <param name="bis">rechter Rand des zeitraums</param>
    ''' <param name="sortiertNachDauer" >soll nach Dauer sortiert werden: true; nach Position auf der Projekttafel: false </param>
    ''' <param name="projektListe">Ergebnis enthält alle Projekt-Namen die eine der Phasen oder einen der Meilensteine im angegebenen Zeitraum enthalten </param>
    ''' <param name="minDate">das kleinste auftretende Start-Datum einer Phase</param>
    ''' <param name="maxDate">das größte auftretende Ende-Datum einer Phase </param>
    ''' <param name="isMultiprojektSicht">gibt an, ob die Varianten angezeigt werden sollen oder die Multiprojekt-Sicht</param>
    ''' <param name="projMitVariants">im Falle Varainten-Sicht: Projekt, dessen Varianten dargestellt werden sollen</param>
    ''' <remarks></remarks>
    Public Sub bestimmeProjekteAndMinMaxDates(ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                              ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                              ByVal selectedBUs As Collection, ByVal selectedTyps As Collection, _
                                              ByVal von As Integer, ByVal bis As Integer, ByVal sortiertNachDauer As Boolean, _
                                                  ByRef projektListe As SortedList(Of Double, String), ByRef minDate As Date, ByRef maxDate As Date, _
                                                  ByVal isMultiprojektSicht As Boolean, ByVal projMitVariants As clsProjekt)

        Dim tmpMinimum As Date
        Dim tmpMaximum As Date
        Dim tmpDate As Date
        Dim currentFilter As clsFilter

        Dim hproj As clsProjekt
        Dim cphase As clsPhase
        Dim projektstart As Integer
        'Dim found As Boolean
        Dim key As Double
        Dim noTimespanDefined As Boolean
        ' selection type wird aktuell noch ignoriert .... 

        ' in der ersten Welle werden die Projektnamen aufgesammelt, die eine der Phasen oder Meilensteine enthalten 
        ' und gleichzeitig den ggf definierten filterkriterien BU und Typ entsprechen 
        currentFilter = New clsFilter("temp", selectedBUs, selectedTyps, selectedPhases, selectedMilestones, _
                                      selectedRoles, selectedCosts)

        If Not ((showRangeLeft > 0) And (showRangeRight > showRangeLeft)) Then
            noTimespanDefined = True
        Else
            noTimespanDefined = False
        End If

        If isMultiprojektSicht Then

            If noTimespanDefined Then
                tmpMinimum = StartofCalendar.AddYears(500)
                tmpMaximum = StartofCalendar.AddYears(-500)
            Else
                tmpMinimum = StartofCalendar.AddMonths(von - 1)
                tmpMaximum = StartofCalendar.AddMonths(bis).AddDays(-1)
            End If


            For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

                If currentFilter.doesNotBlock(kvp.Value) Then
                    If awinSettings.mppSortiertDauer Then
                        ' es wird aufsteigend nach der Dauer sortiert  
                        Dim tmpMinDate As Date
                        Dim tmpMaxDate As Date
                        Dim tmpDuration As Long
                        kvp.Value.getMinMaxDatesAndDuration(selectedPhases, selectedMilestones, _
                                                            tmpMinDate, tmpMaxDate, tmpDuration)

                        key = CDbl(tmpDuration)
                    Else
                        key = kvp.Value.tfZeile + kvp.Value.anzahlRasterElemente / 10000
                    End If

                    Do While projektListe.ContainsKey(key)
                        key = key + 0.000001
                    Loop
                    ' jetzt ist sicher gestellt, dass key nicht mehr vorkommen kann ... 
                    projektListe.Add(key, calcProjektKey(kvp.Value))
                End If

            Next
        Else
            ' Multivarianten Sicht 
            ' in diesem Fall soll der selektierte Zeitraum nicht betrachtet werden 
            von = 0
            bis = 0
            tmpMinimum = AlleProjekte.getMinDate(pName:=projMitVariants.name)
            tmpMaximum = AlleProjekte.getMaxDate(pName:=projMitVariants.name)

            Dim variantNames As Collection = AlleProjekte.getVariantNames(projMitVariants.name, False)
            For i As Integer = 1 To variantNames.Count
                key = i
                projektListe.Add(key, calcProjektKey(projMitVariants.name, CStr(variantNames.Item(i))))
            Next
        End If


        ' jetzt muss die zweite Welle nachkommen .. bestimmen , welches die erweiterten Min / Max Werte sind, falls fullyContained bzw. showAllIfOne 
        ' hier jetzt für alle Projekte in projektliste für jedes Element aus selectedphases und selectedmilestones das Minimum / Maximum bestimmen


        For Each kvp As KeyValuePair(Of Double, String) In projektListe

            'hproj = ShowProjekte.getProject(kvp.Value)
            ' in Projektliste sind jetzt die Keys die zusammengesetzten Schlüssel aus pname und variantName
            hproj = AlleProjekte.getProject(kvp.Value)
            projektstart = hproj.Start + hproj.StartOffset

            ' Phasen checken 
            For Each fullPhaseName As String In selectedPhases

                Dim breadcrumb As String = ""
                Dim phaseName As String = ""

                Call splitHryFullnameTo2(fullPhaseName, phaseName, breadcrumb)
                Dim phaseIndices() As Integer = hproj.hierarchy.getPhaseIndices(phaseName, breadcrumb)

                For px As Integer = 0 To phaseIndices.Length - 1
                    If phaseIndices(px) > 0 And phaseIndices(px) <= hproj.CountPhases Then
                        cphase = hproj.getPhase(phaseIndices(px))
                        If Not IsNothing(cphase) Then
                            If awinSettings.mppShowAllIfOne Or noTimespanDefined Then
                                ' das umschliesst jetzt bereits fullyContained 

                                If DateDiff(DateInterval.Day, cphase.getStartDate, tmpMinimum) > 0 Then
                                    tmpMinimum = cphase.getStartDate
                                End If

                                If DateDiff(DateInterval.Day, cphase.getEndDate, tmpMaximum) < 0 Then
                                    tmpMaximum = cphase.getEndDate
                                End If


                            Else
                                ' hier muss in Abhängigkeit von fullyContained als dem schwächeren Kriterium noch auf fullyContained geprüft werden 
                                ' andernfalls muss nichts gemacht werden 

                                If awinSettings.mppFullyContained Then
                                    If phaseWithinTimeFrame(projektstart, cphase.relStart, cphase.relEnde, von, bis) Then

                                        If DateDiff(DateInterval.Day, cphase.getStartDate, tmpMinimum) > 0 Then
                                            tmpMinimum = cphase.getStartDate
                                        End If

                                        If DateDiff(DateInterval.Day, cphase.getEndDate, tmpMaximum) < 0 Then
                                            tmpMaximum = cphase.getEndDate
                                        End If

                                    End If
                                End If
                            End If
                        End If
                    End If
                Next ' ix
            Next ' phaseName

            ' Meilensteine 
            ' das muss nur gemacht werden, wenn showAllIfOne=true 
            If awinSettings.mppShowAllIfOne Or noTimespanDefined Then

                For Each fullMsName As String In selectedMilestones

                    Dim breadcrumb As String = ""
                    Dim msName As String = ""
                    Call splitHryFullnameTo2(fullMsName, msName, breadcrumb)

                    Dim milestoneIndices(,) As Integer = hproj.hierarchy.getMilestoneIndices(msName, breadcrumb)
                    ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste

                    For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                        If milestoneIndices(0, mx) > 0 And milestoneIndices(1, mx) > 0 Then

                            Try
                                tmpDate = hproj.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx)).getDate

                                If DateDiff(DateInterval.Day, StartofCalendar, tmpDate) >= 0 Then

                                    If DateDiff(DateInterval.Day, tmpDate, tmpMinimum) > 0 Then
                                        tmpMinimum = tmpDate
                                    End If

                                    If DateDiff(DateInterval.Day, tmpDate, tmpMaximum) < 0 Then
                                        tmpMaximum = tmpDate
                                    End If

                                End If
                            Catch ex As Exception

                            End Try

                        End If

                    Next
                Next
            End If


        Next


        minDate = tmpMinimum
        maxDate = tmpMaximum

    End Sub


    ''' <summary>
    ''' bestimmt für die angegebenen Phasen und Meilensteine den Kalender-Start und das Kalender-Ende für 
    ''' für den Kalender der PPT Multiprojekt sicht  
    ''' dabei wird berücksichtigt, ob der Kalender mehr anzeigen soll als das ausgewählte Zeitfenster: ist dann der Fall, wenn fullyContained oder ShowAllIfOne gewählt wurde
    ''' </summary>
    ''' <param name="minDate">erstes, linkes Datum</param>
    ''' <param name="maxDate">zweites, rechtes Datum</param>
    ''' <param name="pptKalenderStart"></param>
    ''' <param name="pptKalenderEnde"></param>
    ''' <remarks></remarks>
    Sub calcStartEndePPTKalender(ByVal minDate As Date, ByVal maxDate As Date, _
                                     ByRef pptKalenderStart As Date, ByRef pptKalenderEnde As Date)

        Dim firstDate As Date = minDate
        Dim lastdate As Date = maxDate
        Dim linksDatum As Date
        Dim rechtsDatum As Date

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            linksDatum = StartofCalendar.AddMonths(showRangeLeft - 1)
            rechtsDatum = StartofCalendar.AddMonths(showRangeRight).AddDays(-1)
        Else
            linksDatum = minDate
            rechtsDatum = maxDate
        End If

        ' Änderung tk: es soll nicht mehr links und rechts ein zusätzlicher Bereich aufgespannt werden 
        If Not awinSettings.mppFullyContained And Not awinSettings.mppShowAllIfOne Then
            firstDate = linksDatum
            lastdate = rechtsDatum
        Else
            firstDate = minDate
            lastdate = maxDate
        End If

        If DateDiff(DateInterval.Day, linksDatum, firstDate) < 0 Then
            pptKalenderStart = firstDate.AddDays(-1 * firstDate.Day + 1)
        Else
            pptKalenderStart = linksDatum.AddDays(-1 * linksDatum.Day + 1)
        End If


        If DateDiff(DateInterval.Day, rechtsDatum, lastdate) > 0 Then
            pptKalenderEnde = lastdate.AddDays(-1 * lastdate.Day + 1).AddMonths(1).AddDays(-1)
        Else
            pptKalenderEnde = rechtsDatum.AddDays(-1 * rechtsDatum.Day + 1).AddMonths(1).AddDays(-1)
        End If



    End Sub



    ''' <summary>
    ''' zeichnet den PPT Kalender auf die übergebene Shape
    ''' </summary>
    ''' <param name="pptslide">PPT Slide, auf die gezeichnet wird</param>
    ''' <param name="calendargroup">zusammengesetztes Shape, das dem fertigen Kalender entspricht</param>
    ''' <param name="StartofPPTCalendar">Start Datum PPT Kalender</param>
    ''' <param name="endOFPPTCalendar">End Datum PPT Kalender</param>
    ''' <param name="calendarLineShape">Line, die den unteren Rand des Kalenders markiert</param>
    ''' <param name="calendarHeightShape">Line, die die Höhe des Kalenders angibt</param>
    ''' <param name="calendarStepShape">Linie, die für die Monatsabgrenzungen verwendet werden kann</param>
    ''' <param name="calendarMark">Shape, wie der ausgewählte Zeitraum markiert werden soll</param>
    ''' <param name="yearShape">Vorlage für das Jahr</param>
    ''' <param name="qmShape">Vorlage für Quartal / Monat </param>
    ''' <remarks></remarks>
    Sub zeichnePPTCalendar(ByRef pptslide As pptNS.Slide, ByRef calendargroup As pptNS.Shape, _
                           ByVal StartofPPTCalendar As Date, ByVal endOFPPTCalendar As Date, _
                               ByVal calendarLineShape As pptNS.Shape, ByVal calendarHeightShape As pptNS.Shape, _
                               ByVal calendarStepShape As pptNS.Shape, ByVal calendarMark As pptNS.Shape, _
                               ByVal yearShape As pptNS.Shape, ByVal qmShape As pptNS.Shape, _
                               ByVal yearSeparatorLine As pptNS.Shape, ByVal quartalSeparatorLine As pptNS.Shape, ByVal drawingAreaBottom As Double)

        Dim drawItem As Integer  ' 0: Monat, 1: Quartal, 2: Jahr
        'Dim anzQMs As Integer = DateDiff(DateInterval.Month, StartofPPTCalendar, endOFPPTCalendar) + 1

        Dim anzQMs As Integer
        Dim qmWidth As Double = qmShape.Width
        Dim yearWidth As Double
        Dim monthWidth As Double
        Dim textXPosition As Double, yPosition As Double
        Dim newShapes As pptNS.ShapeRange
        Dim startOfZeitraum As Integer = showRangeLeft - getColumnOfDate(StartofPPTCalendar)
        Dim zeitraumDauer As Integer = showRangeRight - showRangeLeft + 1
        Dim shapeGruppe As pptNS.ShapeRange
        Dim slideShapes As pptNS.Shapes = pptslide.Shapes
        Dim nameCollection As New Collection
        Dim arrayOFNames() As String


        Call calculateYMAeinheiten(StartofPPTCalendar, endOFPPTCalendar, calendarLineShape.Width, _
                                  yearWidth, monthWidth, anzQMs)


        If calendarLineShape.Width >= anzQMs * qmWidth Then
            qmWidth = calendarLineShape.Width / anzQMs
            drawItem = 0
        ElseIf calendarLineShape.Width >= anzQMs / 3 * qmWidth Then
            qmWidth = 3 * calendarLineShape.Width / anzQMs
            drawItem = 1
        ElseIf calendarLineShape.Width >= anzQMs / 12 * yearShape.Width Then
            qmWidth = 12 * calendarLineShape.Width / anzQMs
            drawItem = 2
        Else
            ' jetzt werden die YearShapes solange in der Schriftgröße verkleinert, bis es passt ...
            Dim fitting As Boolean = False
            Dim currentFontSize As Double = yearShape.TextFrame2.TextRange.Font.Size - 2

            While Not fitting And currentFontSize > 6

                yearShape.TextFrame2.TextRange.Font.Size = currentFontSize
                If calendarLineShape.Width >= anzQMs / 12 * yearShape.Width Then
                    fitting = True
                Else
                    currentFontSize = currentFontSize - 2
                End If

            End While

            If fitting Then
                drawItem = 2
            Else
                Throw New ArgumentException("Zeitraum ist zu groß ...")
            End If
        End If


        ' den unteren Rand des Kalenders zeichnen 
        calendarLineShape.Copy()
        newShapes = pptslide.Shapes.Paste
        With newShapes.Item(1)
            .Left = calendarLineShape.Left
            .Top = calendarLineShape.Top
            ' den Namen für PPT 2013 eindeutig machen 
            .Name = .Name & .Id
            nameCollection.Add(.Name, .Name)
        End With

        If Not IsNothing(calendarMark) Then
            calendarMark.Copy()
            newShapes = pptslide.Shapes.Paste
            With newShapes.Item(1)
                .Left = calendarLineShape.Left + startOfZeitraum * monthWidth
                .Top = calendarLineShape.Top - calendarHeightShape.Height
                .Width = zeitraumDauer * monthWidth
                .Height = calendarHeightShape.Height
                ' den Namen für PPT 2013 eindeutig machen 
                .Name = .Name & .Id
                nameCollection.Add(.Name, .Name)
            End With
        End If


        Dim lfdNr As Integer = 1
        Dim qmHeightfaktor As Double = qmShape.Height / (qmShape.Height + yearShape.Height)

        If drawItem = 0 Or drawItem = 1 Then
            ' Quartale oder Monate zeichnen 

            ' jetzt die Zwischenlinien zeichnen , wenn gewünscht 
            If Not IsNothing(calendarStepShape) Then
                textXPosition = calendarLineShape.Left
                yPosition = calendarLineShape.Top - calendarStepShape.Height

                ' Änderung tk:zeichnen der M oder Q-Linien, die auch mitten im Jahr beginnen können
                'For i As Integer = 1 To anzQMs - 1
                For i As Integer = 1 + (StartofPPTCalendar.Month - 1) To anzQMs - 1 + (StartofPPTCalendar.Month - 1)

                    If i Mod 12 <> 0 Then

                        calendarStepShape.Copy()
                        newShapes = pptslide.Shapes.Paste
                        With newShapes.Item(1)
                            .Left = textXPosition + monthWidth - 0.5 * .Width
                            If i Mod 3 = 0 Then
                                .Top = yPosition
                            Else
                                .Top = yPosition + calendarStepShape.Height * 0.5
                                .Height = calendarStepShape.Height * 0.5
                            End If
                            ' den Namen für PPT 2013 eindeutig machen 
                            .Name = .Name & .Id
                            nameCollection.Add(.Name, .Name)
                        End With
                    Else
                        ' hier ggf die Year Separator Linien zeichnen 
                    End If

                    textXPosition = textXPosition + monthWidth


                Next

            End If





            textXPosition = calendarLineShape.Left
            yPosition = calendarLineShape.Top - calendarHeightShape.Height * qmHeightfaktor _
                        + (calendarHeightShape.Height * qmHeightfaktor - qmShape.Height) * 0.5



            ' Änderung tk: Kalender soll auch im Jahr beginnen können ..



            If drawItem = 0 Then

                ' Monate zeichnen 
                lfdNr = StartofPPTCalendar.Month

                For i As Integer = 1 To anzQMs
                    qmShape.Copy()
                    newShapes = pptslide.Shapes.Paste
                    With newShapes.Item(1)
                        .TextFrame2.TextRange.Text = lfdNr.ToString
                        .Left = textXPosition + (qmWidth - .Width) * 0.5
                        .Top = yPosition
                        ' den Namen für PPT 2013 eindeutig machen 
                        .Name = .Name & .Id
                        nameCollection.Add(.Name, .Name)
                    End With

                    lfdNr = lfdNr + 1
                    If lfdNr > 12 Then
                        lfdNr = 1
                    End If
                    textXPosition = textXPosition + qmWidth
                Next

            Else
                ' Quartale zeichnen 
                lfdNr = (StartofPPTCalendar.Month - 1) \ 3 + 1
                For i As Integer = 1 To anzQMs / 3
                    qmShape.Copy()
                    newShapes = pptslide.Shapes.Paste
                    With newShapes.Item(1)
                        .TextFrame2.TextRange.Text = "Q" & lfdNr.ToString
                        .Left = textXPosition + (qmWidth - .Width) * 0.5
                        .Top = yPosition
                        ' den Namen für PPT 2013 eindeutig machen 
                        .Name = .Name & .Id
                        nameCollection.Add(.Name, .Name)
                    End With

                    lfdNr = lfdNr + 1
                    If lfdNr > 4 Then
                        lfdNr = 1
                    End If
                    textXPosition = textXPosition + qmWidth
                Next

            End If

            ' Jetzt die horizontale Line zeichnen , die die M bzw Q von den Jahren trennt
            calendarLineShape.Copy()
            newShapes = pptslide.Shapes.Paste
            With newShapes.Item(1)
                .Left = calendarLineShape.Left
                .Top = calendarLineShape.Top - qmHeightfaktor * calendarHeightShape.Height
                ' den Namen für PPT 2013 eindeutig machen 
                .Name = .Name & .Id
                nameCollection.Add(.Name, .Name)
            End With

        End If

        ' jetzt die Jahre zeichnen 
        textXPosition = calendarLineShape.Left

        If drawItem = 0 Or drawItem = 1 Then

            yPosition = calendarLineShape.Top - calendarHeightShape.Height _
                        + (calendarHeightShape.Height * (1 - qmHeightfaktor) - yearShape.Height) * 0.5
        Else

            yPosition = calendarLineShape.Top - calendarHeightShape.Height _
                        + (calendarHeightShape.Height - yearShape.Height) * 0.5
        End If

        lfdNr = StartofPPTCalendar.Year

        ' den linken Rand des Kalenders zeichnen 
        calendarHeightShape.Copy()
        newShapes = pptslide.Shapes.Paste
        With newShapes.Item(1)
            .Left = calendarLineShape.Left
            .Top = calendarLineShape.Top - calendarHeightShape.Height
            ' den Namen für PPT 2013 eindeutig machen 
            .Name = .Name & .Id
            nameCollection.Add(.Name, .Name)
        End With


        ' Änderung tk Zeichnen der 
        Dim ix As Integer = StartofPPTCalendar.Month
        Dim index As Integer = 1
        Dim partYear As Boolean = (ix > 1)
        Dim lineXPosition As Double = calendarLineShape.Left

        While index <= anzQMs

            ' den Jahrestext schreiben 
            If ix < 7 Then

                yearShape.Copy()
                newShapes = pptslide.Shapes.Paste

                ' hier wird nur ein ganzes Jahr dargestellt
                With newShapes.Item(1)
                    .TextFrame2.TextRange.Text = lfdNr.ToString
                    If partYear Then
                        .Left = textXPosition + (qmWidth * (12 - ix + 1) - .Width) * 0.5
                    Else
                        .Left = textXPosition + (yearWidth - .Width) * 0.5
                    End If

                    .Top = yPosition
                    ' den Namen für PPT 2013 eindeutig machen 
                    .Name = .Name & .Id
                    nameCollection.Add(.Name, .Name)
                End With

            End If

            lfdNr = lfdNr + 1
            If partYear Then
                textXPosition = textXPosition + qmWidth * (12 - ix + 1)
            Else
                textXPosition = textXPosition + yearWidth
            End If


            calendarHeightShape.Copy()
            newShapes = pptslide.Shapes.Paste
            With newShapes.Item(1)

                If partYear Then
                    .Left = lineXPosition + qmWidth * (12 - ix + 1)
                    lineXPosition = lineXPosition + qmWidth * (12 - ix + 1)
                Else
                    '.Left = calendarHeightShape.Left + yearWidth
                    .Left = lineXPosition + yearWidth
                    lineXPosition = lineXPosition + yearWidth
                End If

                '.Top = calendarHeightShape.
                .Top = calendarLineShape.Top - calendarHeightShape.Height
                ' den Namen für PPT 2013 eindeutig machen 
                .Name = .Name & .Id
                nameCollection.Add(.Name, .Name)

            End With


            ' jetzt index verändern 
            If partYear Then
                index = index + 12 - ix + 1
            Else
                index = index + 12
            End If

            ' ist das nächste Jahr ein volles Jahr ? 
            If anzQMs - index >= 12 Then
                partYear = False
                ix = 1
            Else
                partYear = True
                ix = 12 - (anzQMs - index)
            End If

        End While


        ' und jetzt noch die oberste Linie zeichnen  
        calendarLineShape.Copy()
        newShapes = pptslide.Shapes.Paste
        With newShapes.Item(1)
            .Left = calendarLineShape.Left
            .Top = calendarLineShape.Top - calendarHeightShape.Height
            ' den Namen für PPT 2013 eindeutig machen 
            .Name = .Name & .Id
            nameCollection.Add(.Name, .Name)
        End With

        ' jetzt ggf die vertikalen Raster in der Zeichenfläche zeichnen 
        If awinSettings.mppVertikalesRaster And _
            Not IsNothing(yearSeparatorLine) And Not IsNothing(quartalSeparatorLine) Then

            textXPosition = calendarLineShape.Left

            ' zeichnen der Q- oder M  Separators
            Dim divisor As Integer = 12


            If drawItem = 0 Then
                divisor = 1
            ElseIf drawItem = 1 Then
                divisor = 3
            End If

            For i = 0 + (StartofPPTCalendar.Month - 1) To (anzQMs + StartofPPTCalendar.Month - 1) / divisor

                If i Mod CInt(12 / divisor) = 0 Then
                    ' es handelt sich um eine JAhreslinie
                    yearSeparatorLine.Copy()
                Else
                    quartalSeparatorLine.Copy()
                End If

                newShapes = pptslide.Shapes.Paste
                With newShapes.Item(1)
                    .Left = textXPosition
                    .Top = calendarLineShape.Top
                    .Height = drawingAreaBottom - calendarLineShape.Top
                    ' den Namen für PPT 2013 eindeutig machen 
                    .Name = .Name & .Id
                    nameCollection.Add(.Name, .Name)
                End With

                textXPosition = textXPosition + qmWidth

            Next

        End If

        ' jetzt sollen alle gezeichneten Shapes gruppiert werden 
        Dim anzElements As Integer = nameCollection.Count
        If anzElements = 0 Then

            calendargroup = Nothing

        ElseIf anzElements = 1 Then

            calendargroup = pptslide.Shapes.Item(nameCollection.Item(1))

        Else

            ReDim arrayOFNames(anzElements - 1)

            For i = 1 To anzElements
                arrayOFNames(i - 1) = CStr(nameCollection.Item(i))
            Next

            shapeGruppe = pptslide.Shapes.Range(arrayOFNames)
            calendargroup = shapeGruppe.Group


        End If


    End Sub



    ''' <summary>
    ''' zeichnet einen Kalender mit drei Reihen: Jahre, Quartale oder Monate, Monate oder Kalenderwochen; 
    ''' es wird also entweder y/q/m gezeichnet oder y/m/w Monate oder 
    ''' </summary>
    ''' <param name="rds">die Powerpoint Klasse, die das Slide und alle Hilfsshapes enthält; mit deren Hilfe wird dann gezeichnet</param>
    ''' <param name="calendargroup">die Kalendergruppe, die zurückgegeben wird</param>
    ''' <remarks></remarks>
    Sub zeichne3RowsCalendar(ByRef rds As clsPPTShapes, ByRef calendargroup As pptNS.Shape)

        'Sub zeichne3RowsCalendar(ByRef pptslide As pptNS.Slide, ByRef calendargroup As pptNS.Shape, _
        '                               ByVal StartofPPTCalendar As Date, ByVal endOFPPTCalendar As Date, _
        '                                   ByVal calendarLineShape As pptNS.Shape, ByVal calendarHeightShape As pptNS.Shape, _
        '                                   ByVal calendarStepShape As pptNS.Shape, ByVal calendarMark As pptNS.Shape, _
        '                                   ByVal yearShape As pptNS.Shape, ByVal qmShape As pptNS.Shape, _
        '                                   ByVal yearSeparatorLine As pptNS.Shape, ByVal quartalSeparatorLine As pptNS.Shape, ByVal drawingAreaBottom As Double)
        Dim monthName(11) As String
        monthName(0) = "Jan"
        monthName(1) = "Feb"
        monthName(2) = "Mar"
        monthName(3) = "Apr"
        monthName(4) = "May"
        monthName(5) = "Jun"
        monthName(6) = "Jul"

        monthName(7) = "Aug"
        monthName(8) = "Sep"
        monthName(9) = "Oct"
        monthName(10) = "Nov"
        monthName(11) = "Dec"

        Dim QuartalsName() As String = {"QI", "QII", "QIII", "QIV"}


        Dim newShapes As pptNS.ShapeRange

        ' nimmt die Namen aller erzeugten Shapes auf: daraus wird später die Gruppe erzeugt 
        Dim nameCollection As New Collection

        ' wieviele Tage auf dem Kalender?
        Dim anzahlTage As Integer = DateDiff(DateInterval.Day, rds.PPTStartOFCalendar, rds.PPTEndOFCalendar)

        ' wie breit ist ein Tg auf dem Kalender? 
        Dim rasterDayWidth As Double = rds.calendarLineShape.Width / anzahlTage

        ' bestimmt ein proportionales Aussehen der Kalenderleiste 
        ' es sollte sichergestellt sein, dass die Shapes für Year und Q/M/W jeweils genügend Margin nach oben und unten haben 
        Dim KalenderHoehe As Double = (2 * rds.quarterMonthVorlagenShape.Height + rds.yearVorlagenShape.Height) * 1.05
        Dim yyHeightfaktor As Double = rds.yearVorlagenShape.Height / KalenderHoehe
        Dim qmHeightfaktor As Double = rds.quarterMonthVorlagenShape.Height / KalenderHoehe

        Dim drawKWs As Boolean
        Dim drawQuartale As Boolean
        If rds.calendarLineShape.Width >= (1 + anzahlTage / 7) * 2 * rds.quarterMonthVorlagenShape.Width Then
            drawKWs = True
            drawQuartale = False
        Else
            drawKWs = False
            drawQuartale = True
        End If

        ' ####---------------------------------------
        ' jetzt wird der Aussen-Rand gezeichnet
        ' ... die unterste horizontale Line zeichnen
        rds.calendarLineShape.Copy()
        newShapes = rds.pptSlide.Shapes.Paste
        With newShapes.Item(1)
            .Left = rds.calendarLineShape.Left
            .Top = rds.calendarLineShape.Top
            .Name = .Name & .Id
            .AlternativeText = ""
            .Title = ""
            nameCollection.Add(.Name, .Name)
        End With

        ' ... die oberste horizontale Line zeichnen
        rds.calendarLineShape.Copy()
        newShapes = rds.pptSlide.Shapes.Paste
        With newShapes.Item(1)
            .Left = rds.calendarLineShape.Left
            .Top = rds.calendarLineShape.Top - (KalenderHoehe + rds.calendarLineShape.Height / 2)
            .AlternativeText = ""
            .Title = ""
            .Name = .Name & .Id
            nameCollection.Add(.Name, .Name)
        End With

        ' ... die Trennlinie1 (Jahre) zeichnen
        rds.calendarLineShape.Copy()
        newShapes = rds.pptSlide.Shapes.Paste
        With newShapes.Item(1)
            .Left = rds.calendarLineShape.Left
            .Top = rds.calendarLineShape.Top - (KalenderHoehe + rds.calendarLineShape.Height / 2) + _
                    yyHeightfaktor * KalenderHoehe
            .AlternativeText = ""
            .Title = ""
            .Name = .Name & .Id

            ' das Format von StepShape übernehmen
            rds.calendarStepShape.PickUp()
            .Apply()

            nameCollection.Add(.Name, .Name)
        End With

        ' ... die Trennlinie2 (Q/M) zeichnen
        rds.calendarLineShape.Copy()
        newShapes = rds.pptSlide.Shapes.Paste
        With newShapes.Item(1)
            .Left = rds.calendarLineShape.Left
            .Top = rds.calendarLineShape.Top - (KalenderHoehe + rds.calendarLineShape.Height / 2) + _
                    yyHeightfaktor * KalenderHoehe + qmHeightfaktor * KalenderHoehe
            .Name = .Name & .Id
            .AlternativeText = ""
            .Title = ""

            ' das Format von StepShape übernehmen
            rds.calendarStepShape.PickUp()
            .Apply()

            nameCollection.Add(.Name, .Name)
        End With

        ' den linken und den rechten Rand zeichnen 
        rds.calendarHeightShape.Copy()
        newShapes = rds.pptSlide.Shapes.Paste
        With newShapes.Item(1)
            .Left = rds.calendarLineShape.Left
            .Top = rds.calendarLineShape.Top - (KalenderHoehe + rds.calendarLineShape.Height / 2)
            .Height = KalenderHoehe
            .Name = .Name & .Id
            .AlternativeText = ""
            .Title = ""

            nameCollection.Add(.Name, .Name)
        End With

        rds.calendarHeightShape.Copy()
        newShapes = rds.pptSlide.Shapes.Paste
        With newShapes.Item(1)
            '.Left = calendarLineShape.Left + calendarLineShape.Width - .Width / 2
            .Left = rds.calendarLineShape.Left + rds.calendarLineShape.Width
            .Top = rds.calendarLineShape.Top - (KalenderHoehe + rds.calendarLineShape.Height / 2)
            .Height = KalenderHoehe
            .Name = .Name & .Id
            .AlternativeText = ""
            .Title = ""

            nameCollection.Add(.Name, .Name)
        End With

        ' Ende Aussen-Box für KAlender schreiben 
        ' ##########################################

        Dim curDatePtr As Date = rds.PPTStartOFCalendar
        Dim curLeft As Double, curTop As Double, curRight As Double
        Dim rowHeight As Double



        Dim atleastOne As Boolean = False
        Dim beschriftung As String = ""

        ' ###########################################
        ' zeichne die Jahres-Trennlinien, schreibe die Jahreszahlen
        Dim dimension As Integer = DateDiff(DateInterval.Year, rds.PPTStartOFCalendar, rds.PPTEndOFCalendar) + 5
        Dim positionY() As Double
        ReDim positionY(dimension)
        Dim positionYPtr = 0

        rowHeight = yyHeightfaktor * KalenderHoehe

        curTop = rds.calendarLineShape.Top - (KalenderHoehe + rds.calendarLineShape.Height / 2)
        curLeft = rds.calendarLineShape.Left
        curRight = rds.calendarLineShape.Left + rds.calendarLineShape.Width

        curDatePtr = rds.PPTStartOFCalendar.AddDays(-1 * rds.PPTStartOFCalendar.DayOfYear).AddYears(1)

        Do While curDatePtr < rds.PPTEndOFCalendar
            curRight = rds.calendarLineShape.Left + DateDiff(DateInterval.Day, rds.PPTStartOFCalendar, curDatePtr) * rasterDayWidth
            positionY(positionYPtr) = curRight
            positionYPtr = positionYPtr + 1

            rds.calendarStepShape.Copy()
            newShapes = rds.pptSlide.Shapes.Paste
            With newShapes.Item(1)
                .Left = curRight
                .Top = curTop
                .Height = rowHeight
                .Name = .Name & .Id
                .AlternativeText = ""
                .Title = ""

                nameCollection.Add(.Name, .Name)
            End With

            ' jetzt die Jahreszahl schreiben 
            If curRight - curLeft >= rds.yearVorlagenShape.Width Then

                beschriftung = curDatePtr.AddMonths(-1).Year.ToString
                rds.yearVorlagenShape.Copy()
                newShapes = rds.pptSlide.Shapes.Paste
                With newShapes.Item(1)
                    .Left = curLeft + (curRight - curLeft - rds.yearVorlagenShape.Width) / 2
                    .Top = curTop + (rowHeight - rds.yearVorlagenShape.Height) / 2
                    .Name = .Name & .Id
                    .AlternativeText = ""
                    .Title = ""

                    .TextFrame2.TextRange.Text = beschriftung
                    nameCollection.Add(.Name, .Name)
                End With

            End If

            curLeft = curRight
            curDatePtr = curDatePtr.AddYears(1)

        Loop

        ' jetzt muss noch die Behandlung kommen, ob das Teil-Jahr beschriftet werden muss
        curRight = rds.calendarLineShape.Left + rds.calendarLineShape.Width
        If curRight - curLeft > 2 * rds.yearVorlagenShape.Width Then

            beschriftung = curDatePtr.AddMonths(-1).Year.ToString
            rds.yearVorlagenShape.Copy()
            newShapes = rds.pptSlide.Shapes.Paste
            With newShapes.Item(1)
                .Left = curLeft + (curRight - curLeft - rds.yearVorlagenShape.Width) / 2
                .Top = curTop + (rowHeight - rds.yearVorlagenShape.Height) / 2
                .Name = .Name & .Id
                .AlternativeText = ""
                .Title = ""

                .TextFrame2.TextRange.Text = beschriftung
                nameCollection.Add(.Name, .Name)
            End With

        End If

        ' Ende Jahres-Zeile zeichnen 
        ' ###########################################
        '

        '
        ' ###########################################
        ' zeichne die Quartals bzw. Monats-Reihe 
        rowHeight = qmHeightfaktor * KalenderHoehe

        curTop = curTop + yyHeightfaktor * KalenderHoehe
        curLeft = rds.calendarLineShape.Left
        curRight = rds.calendarLineShape.Left + rds.calendarLineShape.Width

        ' StartofPPTCalendar beginnt immer am 1. eines Monats 

        Dim position2() As Double
        dimension = DateDiff(DateInterval.Month, rds.PPTStartOFCalendar, rds.PPTEndOFCalendar) + 5
        ReDim position2(dimension)
        Dim position2Ptr = 0

        Dim monthKennzahl As Integer = rds.PPTStartOFCalendar.Month Mod 3
        Dim curQuartal As Integer

        If drawQuartale Then
            ' das erste Quartal berechnen  
            Select Case monthKennzahl
                Case 1
                    ' bringt es auf den 1. des Monats, addiert 3 Monate, geht auf den letzten Tag davor 
                    curDatePtr = rds.PPTStartOFCalendar.AddDays(-1 * rds.PPTStartOFCalendar.Day + 1).AddMonths(3).AddDays(-1)

                Case 2
                    ' bringt es auf den 1. des Monats, addiert 2 Monate, geht auf den letzten Tag davor 
                    curDatePtr = rds.PPTStartOFCalendar.AddDays(-1 * rds.PPTStartOFCalendar.Day + 1).AddMonths(2).AddDays(-1)

                Case 0
                    ' bringt es auf den 1. des Monats, addiert 1 Monat, geht auf den letzten Tag davor 
                    curDatePtr = rds.PPTStartOFCalendar.AddDays(-1 * rds.PPTStartOFCalendar.Day + 1).AddMonths(1).AddDays(-1)

            End Select

            curQuartal = curDatePtr.Month / 3

        Else
            ' den ersten Monat berechnen
            curDatePtr = rds.PPTStartOFCalendar.AddDays(-1 * rds.PPTStartOFCalendar.Day + 1).AddMonths(1).AddDays(-1)
        End If

        ' ggf müssen die Schriftgrößen angepasst werden

        ' bestimmt ob die vertical Linien in der zweiten sufen gezeichnet werden sollen
        ' nur zeichnen wenn auch genügend Platz da ist
        ' Entschediung: Überprüfung beim ersten Auftreten 

        Dim beschrifteLevel2 As Boolean = True

        Do While curDatePtr <= rds.PPTEndOFCalendar

            curRight = rds.calendarLineShape.Left + DateDiff(DateInterval.Day, rds.PPTStartOFCalendar, curDatePtr) * rasterDayWidth

            beschrifteLevel2 = beschrifteLevel2 And (curRight - curLeft >= rds.quarterMonthVorlagenShape.Width)

            ' Merken, an dieser Stelle werden ggf nachher die vertikalen Linien gezeichnet , aber nur, wenn nicht eine Dezember Linie
            ' und nur wenn nicht = endofPPTCalendar
            If curDatePtr.Month <> 12 And curDatePtr < rds.PPTEndOFCalendar Then
                position2(position2Ptr) = curRight
                position2Ptr = position2Ptr + 1
            End If

            rds.calendarStepShape.Copy()
            newShapes = rds.pptSlide.Shapes.Paste
            With newShapes.Item(1)
                .Left = curRight
                .Top = curTop
                .Height = rowHeight
                .Name = .Name & .Id
                .AlternativeText = ""
                .Title = ""

                nameCollection.Add(.Name, .Name)
            End With





            ' jetzt die Quartals- bzw. Monatszahl  schreiben 
            If beschrifteLevel2 Then

                If drawQuartale Then
                    beschriftung = QuartalsName(curQuartal - 1)
                Else
                    beschriftung = monthName(curDatePtr.Month - 1)
                End If

                rds.quarterMonthVorlagenShape.Copy()
                newShapes = rds.pptSlide.Shapes.Paste
                With newShapes.Item(1)
                    .Left = curLeft + (curRight - curLeft - rds.quarterMonthVorlagenShape.Width) / 2
                    .Top = curTop + (rowHeight - rds.quarterMonthVorlagenShape.Height) / 2
                    .Name = .Name & .Id
                    .AlternativeText = ""
                    .Title = ""

                    .TextFrame2.TextRange.Text = beschriftung
                    nameCollection.Add(.Name, .Name)
                End With

            Else
                ' hier muss gekennzeichnet werden, dass keine Beschriftung mehr stattfinden konnte. 
                ' Dann sollen auf diesem Granularitätslevel auch keine vertikalen Linien gezeichnet werden  
                beschriftung = ""
            End If


            curLeft = curRight
            If drawQuartale Then
                ' dieses scheinbare Nullsummenspiel bei den Tagen ist entscheidend, damit immer der letzte Tag des betreffenden Monats rauskommt
                ' und das kann nur sichergestellt werden, wenn man vom 1. eines Monats ausgeht und eins abzieht 
                curDatePtr = curDatePtr.AddDays(1).AddMonths(3).AddDays(-1)
                curQuartal = curQuartal + 1
                If curQuartal > 4 Then
                    curQuartal = 1
                End If
            Else
                curDatePtr = curDatePtr.AddDays(1).AddMonths(1).AddDays(-1)
            End If


        Loop

        ' jetzt muss noch die Behandlung kommen, ob das Teil-Quartal / Monat beschriftet werden muss
        ' jetzt die Quartals- bzw. Monatszahl  schreiben 
        curRight = rds.calendarLineShape.Left + rds.calendarLineShape.Width

        If curRight - curLeft > rds.quarterMonthVorlagenShape.Width Then

            If beschrifteLevel2 Then

                If drawQuartale Then
                    beschriftung = QuartalsName(curQuartal - 1)
                Else
                    beschriftung = monthName(curDatePtr.Month - 1)
                End If

                rds.quarterMonthVorlagenShape.Copy()
                newShapes = rds.pptSlide.Shapes.Paste
                With newShapes.Item(1)
                    .Left = curLeft + (curRight - curLeft - rds.quarterMonthVorlagenShape.Width) / 2
                    .Top = curTop + (rowHeight - rds.quarterMonthVorlagenShape.Height) / 2
                    .Name = .Name & .Id
                    .AlternativeText = ""
                    .Title = ""

                    .TextFrame2.TextRange.Text = beschriftung
                    nameCollection.Add(.Name, .Name)
                End With

            Else
                beschriftung = ""
            End If

        End If




        ' Ende Quartals bzw. Monats-Zeile zeichnen 
        ' ###########################################
        '

        '
        ' ###########################################
        ' zeichne die Monats- bzw. Kalenderwochen Reihe
        rowHeight = qmHeightfaktor * KalenderHoehe

        curTop = curTop + rowHeight
        curLeft = rds.calendarLineShape.Left
        curRight = rds.calendarLineShape.Left + rds.calendarLineShape.Width

        Dim position() As Double
        ' Play it safe - einfach Puffer von 5 daruf geben 
        dimension = anzahlTage / 7 + 5
        ReDim position(dimension)
        Dim positionPtr = 0

        If drawKWs Then
            ' die erste KW berechnen 
            curDatePtr = rds.PPTStartOFCalendar.AddDays(-1 * rds.PPTStartOFCalendar.DayOfWeek + 1)
            If DateDiff(DateInterval.Day, rds.PPTStartOFCalendar, curDatePtr) < 0 Then
                curDatePtr = curDatePtr.AddDays(7)
            End If
        Else
            curDatePtr = rds.PPTStartOFCalendar.AddDays(-1 * rds.PPTStartOFCalendar.Day + 1).AddMonths(1).AddDays(-1)
        End If

        Dim beschrifteLevel3 As Boolean = False

        Do While curDatePtr <= rds.PPTEndOFCalendar
            curRight = rds.calendarLineShape.Left + DateDiff(DateInterval.Day, rds.PPTStartOFCalendar, curDatePtr) * rasterDayWidth

            If curDatePtr < rds.PPTEndOFCalendar Then
                beschrifteLevel3 = beschrifteLevel3 Or (curRight - curLeft >= rds.quarterMonthVorlagenShape.Width)
            Else
                beschrifteLevel3 = (curRight - curLeft >= rds.quarterMonthVorlagenShape.Width)
            End If


            ' Merken, an dieser Stelle werden ggf nachher die vertikalen Linien gezeichnet , aber nur, wenn nicht eine Dezember Linie
            ' und nur wenn nicht = endofPPTCalendar
            If (drawKWs Or (curDatePtr.Month <> 12)) And curDatePtr < rds.PPTEndOFCalendar Then
                position(positionPtr) = curRight
                positionPtr = positionPtr + 1
            End If

            ' auch das StepShape muss nur gezeichnet werden, wenn kleiner als endofPPTCalendar
            If curDatePtr < rds.PPTEndOFCalendar Then
                rds.calendarStepShape.Copy()
                newShapes = rds.pptSlide.Shapes.Paste
                With newShapes.Item(1)
                    .Left = curRight
                    .Top = curTop
                    .Height = rowHeight
                    .Name = .Name & .Id
                    .AlternativeText = ""
                    .Title = ""

                    nameCollection.Add(.Name, .Name)
                End With
            End If


            ' jetzt die KW bzw. Monatszahl  schreiben 
            If beschrifteLevel3 Then

                If drawKWs Then
                    If curDatePtr.DayOfWeek = 1 Then
                        beschriftung = calcKW(curDatePtr.AddDays(-7)).ToString("0#")
                    Else
                        beschriftung = calcKW(curDatePtr).ToString("0#")
                    End If

                Else

                    beschriftung = monthName(curDatePtr.Month - 1)

                End If


                rds.quarterMonthVorlagenShape.Copy()
                newShapes = rds.pptSlide.Shapes.Paste
                With newShapes.Item(1)
                    .Left = curLeft + (curRight - curLeft - rds.quarterMonthVorlagenShape.Width) / 2
                    .Top = curTop + (rowHeight - rds.quarterMonthVorlagenShape.Height) / 2
                    .Name = .Name & .Id
                    .AlternativeText = ""
                    .Title = ""

                    .TextFrame2.TextRange.Text = beschriftung
                    nameCollection.Add(.Name, .Name)
                End With
            Else
                beschriftung = ""
                ' Kennzeichnen , dass diese Stufe nicht als vertikale Linie dargestellt werden soll 
            End If

            curLeft = curRight
            If drawKWs Then
                curDatePtr = curDatePtr.AddDays(7)
            Else
                curDatePtr = curDatePtr.AddDays(1).AddMonths(1).AddDays(-1)
            End If


        Loop

        ' jetzt muss noch die Behandlung kommen, ob der Rest noch beschriftet werden soll 

        If curDatePtr > rds.PPTEndOFCalendar Then
            curDatePtr = rds.PPTEndOFCalendar
        End If

        curRight = rds.calendarLineShape.Left + rds.calendarLineShape.Width
        If curRight - curLeft > rds.quarterMonthVorlagenShape.Width Then
            If beschrifteLevel3 Then

                If drawKWs Then
                    If curDatePtr.DayOfWeek = 1 Then
                        beschriftung = calcKW(curDatePtr.AddDays(-7)).ToString("0#")
                    Else
                        beschriftung = calcKW(curDatePtr).ToString("0#")
                    End If
                Else
                    beschriftung = monthName(curDatePtr.Month - 1)
                End If

                rds.quarterMonthVorlagenShape.Copy()
                newShapes = rds.pptSlide.Shapes.Paste
                With newShapes.Item(1)
                    .Left = curLeft + (curRight - curLeft - rds.quarterMonthVorlagenShape.Width) / 2
                    .Top = curTop + (rowHeight - rds.quarterMonthVorlagenShape.Height) / 2
                    .Name = .Name & .Id
                    .AlternativeText = ""
                    .Title = ""

                    .TextFrame2.TextRange.Text = beschriftung
                    nameCollection.Add(.Name, .Name)
                End With


            End If
        End If


        ' jetzt das CalendarMark zeichnen 
        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            Dim startOfZeitraum As Integer = showRangeLeft - getColumnOfDate(rds.PPTStartOFCalendar)
            Dim zeitraumDauer As Integer = showRangeRight - showRangeLeft + 1

            If Not IsNothing(rds.calendarMarkShape) Then

                rds.calendarMarkShape.Copy()
                newShapes = rds.pptSlide.Shapes.Paste

                With newShapes.Item(1)
                    .Left = rds.calendarLineShape.Left + _
                        DateDiff(DateInterval.Day, rds.PPTStartOFCalendar, StartofCalendar.AddMonths(showRangeLeft - 1)) * rasterDayWidth
                    .Top = rds.calendarLineShape.Top - KalenderHoehe
                    .Width = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(showRangeLeft - 1), StartofCalendar.AddMonths(showRangeRight).AddDays(-1)) * rasterDayWidth
                    .Height = KalenderHoehe
                    .Name = .Name & .Id
                    .AlternativeText = ""
                    .Title = ""
                    nameCollection.Add(.Name, .Name)
                End With

            End If

        End If


        ' zeichne die vertikalen Linien, wenn gewünscht ... 

        If awinSettings.mppVertikalesRaster And _
            Not IsNothing(rds.calendarYearSeparator) And Not IsNothing(rds.calendarQuartalSeparator) Then

            ' zeichne die Monats- bzw. Kalenderwochen Linien

            ' als erstes wird die linke Begrenzung gezeichnet

            rds.calendarQuartalSeparator.Copy()
            newShapes = rds.pptSlide.Shapes.Paste
            With newShapes.Item(1)
                '.Left = position(i) - .Width / 2
                .Left = rds.calendarLineShape.Left
                .Top = rds.calendarLineShape.Top
                .Height = rds.drawingAreaBottom - rds.calendarLineShape.Top
                .Name = .Name & .Id
                .AlternativeText = ""
                .Title = ""

                nameCollection.Add(.Name, .Name)
            End With


            ' dann wird die rechte Begrenzung gezeichnet
            rds.calendarQuartalSeparator.Copy()
            newShapes = rds.pptSlide.Shapes.Paste
            With newShapes.Item(1)
                '.Left = position(i) - .Width / 2
                .Left = rds.calendarLineShape.Left + rds.calendarLineShape.Width
                .Top = rds.calendarLineShape.Top
                .Height = rds.drawingAreaBottom - rds.calendarLineShape.Top
                .Name = .Name & .Id
                .AlternativeText = ""
                .Title = ""

                nameCollection.Add(.Name, .Name)
            End With


            If beschrifteLevel3 Then
                For i As Integer = 0 To positionPtr - 1

                    rds.calendarQuartalSeparator.Copy()
                    newShapes = rds.pptSlide.Shapes.Paste
                    With newShapes.Item(1)
                        '.Left = position(i) - .Width / 2
                        .Left = position(i)
                        .Top = rds.calendarLineShape.Top
                        .Height = rds.drawingAreaBottom - rds.calendarLineShape.Top
                        .Name = .Name & .Id
                        .AlternativeText = ""
                        .Title = ""

                        nameCollection.Add(.Name, .Name)
                    End With

                Next
            ElseIf beschrifteLevel2 Then

                For i As Integer = 0 To position2Ptr - 1

                    rds.calendarQuartalSeparator.Copy()
                    newShapes = rds.pptSlide.Shapes.Paste
                    With newShapes.Item(1)
                        '.Left = position(i) - .Width / 2
                        .Left = position2(i)
                        .Top = rds.calendarLineShape.Top
                        .Height = rds.drawingAreaBottom - rds.calendarLineShape.Top
                        .Name = .Name & .Id
                        .AlternativeText = ""
                        .Title = ""

                        nameCollection.Add(.Name, .Name)
                    End With

                Next

            End If


            ' zeichne die Jahres Linien - die werden auf alle fälle gezeichnet
            For i As Integer = 0 To positionYPtr - 1

                rds.calendarYearSeparator.Copy()
                newShapes = rds.pptSlide.Shapes.Paste
                With newShapes.Item(1)
                    '.Left = positionY(i) - .Width / 2
                    .Left = positionY(i)
                    .Top = rds.calendarLineShape.Top
                    .Height = rds.drawingAreaBottom - rds.calendarLineShape.Top
                    .Name = .Name & .Id
                    .AlternativeText = ""
                    .Title = ""

                    nameCollection.Add(.Name, .Name)
                End With

            Next


        End If

        ' jetzt sollen alle gezeichneten Shapes gruppiert werden 
        Dim shapeGruppe As pptNS.ShapeRange
        Dim slideShapes As pptNS.Shapes = rds.pptSlide.Shapes

        Dim arrayOFNames() As String


        Dim anzElements As Integer = nameCollection.Count
        If anzElements = 0 Then

            calendargroup = Nothing

        ElseIf anzElements = 1 Then

            calendargroup = rds.pptSlide.Shapes.Item(nameCollection.Item(1))

        Else

            ReDim arrayOFNames(anzElements - 1)

            For i = 1 To anzElements
                arrayOFNames(i - 1) = CStr(nameCollection.Item(i))
            Next

            shapeGruppe = rds.pptSlide.Shapes.Range(arrayOFNames)
            calendargroup = shapeGruppe.Group


        End If


    End Sub




    ''' <summary>
    ''' zeichnet die Swimlanes 
    ''' </summary>
    ''' <param name="rds"></param>
    ''' <param name="hproj"></param>
    ''' <param name="swimlaneNameID">die NameID der Phase, die als Swimlnae gezeichnet werden soll</param>
    ''' <param name="considerAll">sollen alle Pan-Elemente in der Swimlane gezeichnet werden </param>
    ''' <param name="breadCrumbArray">enthält ggf die Liste der BreadCrumbs aller ausgewählten Phasen bzw. Meilensteine </param>
    ''' <param name="selectedPhaseIDs">die NameIDs, die in diesem Projekt der Liste der gewählten Phasen entspricht </param>
    ''' <param name="selectedMilestoneIDs">die NameIDs, die in diesem Projekt der Liste der gewählten Meilensteine entspricht</param>
    ''' <param name="selectedRoles">für später: die ausgewählten Rollen</param>
    ''' <param name="selectedCosts">für später: die ausgewählten Kostearten</param>
    ''' <remarks></remarks>
    Sub zeichneSwimlaneOfProject(ByRef rds As clsPPTShapes, ByRef curYPosition As Double, _
                                 ByRef toggleRowDifferentiator As Boolean, _
                                 ByVal hproj As clsProjekt, swimlaneNameID As String, _
                                 ByVal considerAll As Boolean, ByVal breadCrumbArray As String(),
                                 ByVal considerZeitraum As Boolean, ByVal zeitraumGrenzeL As Integer, ByVal zeitraumGrenzeR As Integer, _
                                 ByVal selectedPhaseIDs As Collection, ByVal selectedMilestoneIDs As Collection, _
                                 ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                 ByVal kontrolleAnzZeilen As Integer)




        ' nimmt die Namen aller erzeugten Shapes auf: daraus wird später die Gruppe erzeugt 
        Dim shapeNameCollection As New Collection

        Dim swlMilestoneCollection As New Collection

        Dim extended As Boolean = awinSettings.mppExtendedMode

        ' x1, x2 sind die Anfangs- und End-Koordinaten eines Shapes auf der Zeichenfläche 
        Dim x1 As Double, x2 As Double

        ' startNr, endNr sind die Anfangs- und End-Indices der Kind-Phasen der Swimlane
        Dim startNr As Integer = 0
        Dim endNr As Integer = 0

        ' wird benutzt, um mal oben und mal unten in der Swimlane zeichnen zu können 
        Dim aktuelleYPosition As Double = curYPosition

        ' in startNr ist nachher die Phasen-Nummer der swimlane, in startNr +1 die Phasen-Nummer des ersten Kindes 
        ' in endNr ist die Phasen-Nummer des letzten Kindes 
        Call hproj.calcStartEndChildNrs(swimlaneNameID, startNr, endNr)

        'Dim fullSwlBreadCrumb As String = hproj.getBcElemName(swimlaneNameID)

        Dim copiedShape As pptNS.ShapeRange

        Dim childPhaseIDs As New Collection
        Dim childMilestoneIDs As New Collection

        If Not considerAll Then
            childPhaseIDs = hproj.schnittmengeChilds(swimlaneNameID, selectedPhaseIDs)
            childMilestoneIDs = hproj.schnittmengeChilds(swimlaneNameID, selectedMilestoneIDs)
        End If



        ' ###########################################################
        ' zeichnen des Swimlane-Namens
        '
        rds.projectNameVorlagenShape.Copy()
        copiedShape = rds.pptSlide.Shapes.Paste()

        Dim swlNameShape As pptNS.Shape = copiedShape.Item(1)

        With copiedShape.Item(1)
            .Top = CSng(curYPosition) + rds.YprojectName
            .Left = rds.projectListLeft
            .TextFrame2.TextRange.Text = elemNameOfElemID(swimlaneNameID)
            .Name = .Name & .Id
            .AlternativeText = elemNameOfElemID(swimlaneNameID)

            shapeNameCollection.Add(.Name, .Name)
        End With


        ' ###########################################################
        ' wenn diese Phase nicht existiert , dann Fehler schreiben ...  
        '
        Dim cphase As clsPhase = hproj.getPhaseByID(swimlaneNameID)

        If IsNothing(cphase) Then

            rds.projectNameVorlagenShape.Copy()
            copiedShape = rds.pptSlide.Shapes.Paste()


            With copiedShape.Item(1)
                .Top = CSng(curYPosition) + rds.YprojectName
                .Left = rds.drawingAreaLeft
                .TextFrame2.TextRange.Text = " ... existiert in diesem Projekt nicht ..."
                .Name = .Name & .Id
                .AlternativeText = "Swimlane " & elemNameOfElemID(swimlaneNameID)

                shapeNameCollection.Add(.Name, .Name)
            End With
        Else
            ' weiter mit Zeichnen der Swimlane ...

            ' ###########################################################
            ' optionales Zeichnen der Swimlane-Linie 
            '
            If awinSettings.mppShowProjectLine Then

                Call rds.calculatePPTx1x2(cphase.getStartDate, cphase.getEndDate, x1, x2)

                ' jetzt muss überprüft werden, ob projectName zu lang ist - dann wird der Name entsprechend abgekürzt ...
                With swlNameShape
                    If .Left + .Width > x1 Then
                        ' jetzt muss der Name entsprechend gekürzt werden 
                        Dim longName As String = .TextFrame2.TextRange.Text
                        Dim shortName As String = ""

                        .TextFrame2.TextRange.Text = shortName
                        Dim stringIX As Integer = 0
                        Do While .Left + .Width < x1 And stringIX <= longName.Length - 1
                            shortName = shortName & longName.Chars(stringIX)
                            stringIX = stringIX + 1
                            .TextFrame2.TextRange.Text = shortName
                        Loop

                    End If
                End With

                rds.projectVorlagenShape.Copy()
                copiedShape = rds.pptSlide.Shapes.Paste()
                With copiedShape(1)
                    .Top = CSng(curYPosition) + rds.YProjectLine
                    .Left = CSng(x1)
                    .Width = CSng(x2 - x1)
                    .Name = .Name & .Id
                    .AlternativeText = cphase.name & " von " & cphase.getStartDate.ToShortDateString & " bis " & _
                                            cphase.getEndDate.ToShortDateString
                    ' wenn Projektstart vor dem Kalender-Start liegt: kein Projektstart Symbol zeichnen
                    If DateDiff(DateInterval.Day, hproj.startDate, rds.PPTStartOFCalendar) > 0 Then
                        .Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone
                    End If

                    ' wenn Projektende nach dem Kalender-Ende liegt: kein Projektende Symbol zeichnen
                    If DateDiff(DateInterval.Day, hproj.endeDate, rds.PPTEndOFCalendar) < 0 Then
                        .Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone
                    End If


                    shapeNameCollection.Add(.Name, .Name)
                End With



            End If


            ' ###########################################################
            ' optionales zeichnen der horizontalen Zeilen - es wird immer nur die Zeile oben gezeichnet ... andernfalls hätte man 
            ' Doppelzeichnungen 
            ' bei der ersten Swimlane auf einer Seite wird die horizontale nicht gezeichnet ... 
            '
            If awinSettings.mppShowHorizontals Then

                rds.horizontalLineShape.Copy()
                copiedShape = rds.pptSlide.Shapes.Paste()

                With copiedShape.Item(1)
                    .Top = CSng(curYPosition)
                    .Left = rds.drawingAreaLeft
                    .Width = rds.drawingAreaWidth
                    .Name = .Name & .Id
                    .AlternativeText = "horizontal line" & elemNameOfElemID(swimlaneNameID)

                    shapeNameCollection.Add(.Name, .Name)
                End With

            End If

        End If




        ' ###########################################################
        ' optionales zeichnen der Zeilen-Markierung
        '
        If (Not IsNothing(rds.rowDifferentiatorShape)) And toggleRowDifferentiator Then
            ' zeichnen des RowDifferentiators 
            rds.rowDifferentiatorShape.Copy()
            copiedShape = rds.pptSlide.Shapes.Paste()
            With copiedShape.Item(1)
                .Top = CSng(curYPosition)
                .Left = rds.projectListLeft
                .Height = kontrolleAnzZeilen * rds.zeilenHoehe
                .Width = rds.drawingAreaRight - .Left
                .Name = .Name & .Id
                .AlternativeText = ""
                .Title = ""

                .ZOrder(MsoZOrderCmd.msoSendToBack)
                shapeNameCollection.Add(.Name, .Name)
            End With
        End If


        ' ###########################################################
        ' jetzt werden die Phasen und Meilensteine gezeichnet, 
        ' beginnend mit Phase <startNr+1> .. <endNr>

        ' zum Bestimmen der optimierten Zeilenanzahl 
        ' es kann in dieser Swimlane nicht mehr als endNr-startNr Zeilen geben 
        Dim dimension As Integer = endNr - startNr
        Dim lastEndDates(dimension) As Date
        For i As Integer = 0 To dimension
            lastEndDates(i) = StartofCalendar.AddDays(-1)
        Next

        Dim maxOffsetZeile As Integer = 1
        Dim curOffsetZeile As Integer = 1


        Dim zeilenoffset As Integer = 1
        Dim curPhase As clsPhase

        ' beginne mit den Meilensteinen, die direkt der Swimlane zugeordnet sind 
        curPhase = hproj.getPhase(startNr)
        If Not IsNothing(curPhase) Then

            ' für jeden Meilenstein dieser Phase untersuchen, ob er gezeigt werden soll 

            For msIX As Integer = 1 To curPhase.countMilestones
                Dim curMs As clsMeilenstein = curPhase.getMilestone(msIX)

                If Not IsNothing(curMs) Then

                    If considerAll Or childMilestoneIDs.Contains(curMs.nameID) Then
                        If Not considerZeitraum _
                                    Or _
                                    (considerZeitraum And milestoneWithinTimeFrame(curMs.getDate, _
                                                                                zeitraumGrenzeL, zeitraumGrenzeR)) Then
                            ' zeichne den Meilenstein 
                            Dim tmpCollection As New Collection
                            Call zeichneMeilensteinInSwimlane(rds, tmpCollection, hproj, _
                                                              swimlaneNameID, curMs.nameID, curYPosition)

                            ' Shape-Namen für spätere Gruppierung der gesamten Swimlane aufnehmen 
                            For Each tmpName As String In tmpCollection
                                shapeNameCollection.Add(tmpName, tmpName)
                                ' die Milestones werden nachher alle in den Vordergrund geholt ...
                                swlMilestoneCollection.Add(tmpName, tmpName)
                            Next
                        End If

                    End If

                End If

            Next

        End If


        ' hier werden jetzt alle Phasen-Kinder inkl ihrer Meilensteine untersucht, ob sie gezeichnet werden sollen ... 
        For swlIX As Integer = startNr + 1 To endNr
            curPhase = hproj.getPhase(swlIX)


            If Not IsNothing(curPhase) Then

                If considerAll Or childPhaseIDs.Contains(curPhase.nameID) Then
                    If Not considerZeitraum _
                                Or _
                                (considerZeitraum And phaseWithinTimeFrame(hproj.Start, curPhase.relStart, curPhase.relEnde, _
                                                                            zeitraumGrenzeL, zeitraumGrenzeR)) Then

                        Dim requiredZeilen As Integer = hproj.calcNeededLinesSwl(curPhase.nameID, _
                                                                                           selectedPhaseIDs, _
                                                                                           selectedMilestoneIDs, _
                                                                                           extended, _
                                                                                           considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR, _
                                                                                           considerAll)

                        ' zeichne die Phase
                        zeilenoffset = findeBesteZeile(lastEndDates, maxOffsetZeile, curPhase.getStartDate, requiredZeilen)
                        'maxOffsetZeile = System.Math.Max(zeilenoffset + requiredZeilen - 1, maxOffsetZeile)
                        ' tk: da das nicht rekursiv aufgerufen wird, sollte sich das nur auf das tatsächlich gezeichnete und deren Zeilennummer beschränken 
                        maxOffsetZeile = System.Math.Max(zeilenoffset, maxOffsetZeile)

                        If DateDiff(DateInterval.Day, lastEndDates(zeilenoffset - 1), curPhase.getEndDate) > 0 Then
                            lastEndDates(zeilenoffset - 1) = curPhase.getEndDate
                        End If

                        aktuelleYPosition = curYPosition + (zeilenoffset - 1) * rds.zeilenHoehe

                        Call zeichnePhaseinSwimlane(rds, shapeNameCollection, hproj, swimlaneNameID, _
                                                    curPhase.nameID, aktuelleYPosition)
                        'lastEndDate = curPhase.getEndDate
                    End If

                End If

                ' für jeden Meilenstein dieser Phase untersuchen, ob er gezeigt werden soll 

                For msIX As Integer = 1 To curPhase.countMilestones
                    Dim curMs As clsMeilenstein = curPhase.getMilestone(msIX)

                    If Not IsNothing(curMs) Then

                        If considerAll Or childMilestoneIDs.Contains(curMs.nameID) Then
                            If Not considerZeitraum _
                                        Or _
                                        (considerZeitraum And milestoneWithinTimeFrame(curMs.getDate, _
                                                                                    zeitraumGrenzeL, zeitraumGrenzeR)) Then


                                ' zeichne den Meilenstein 
                                ' die aktuelle Y-Position muss nicht bestimmt werden, weil das ja bereits mit der Phase geschehen ist 
                                ' es muss nur sichergestellt sein, dass aktuelleYPosition initial auf CurYPosition gesetzt wird
                                Dim tmpCollection As New Collection
                                Call zeichneMeilensteinInSwimlane(rds, tmpCollection, hproj, _
                                                                  swimlaneNameID, curMs.nameID, aktuelleYPosition)

                                ' Shape-Namen für spätere Gruppierung der gesamten Swimlane aufnehmen 
                                For Each tmpName As String In tmpCollection
                                    shapeNameCollection.Add(tmpName, tmpName)
                                    ' die Milestones werden nachher alle in den Vordergrund geholt ...
                                    swlMilestoneCollection.Add(tmpName, tmpName)
                                Next

                            End If

                        End If

                    End If

                Next

            End If

        Next

        




        ' ###########################################################
        ' Weiterschalten der CurYPosition 
        ' Umschalten des toggleRowDifferentiators: dadurch wird die Zeilen - bzw. Projekt - Markierung 
        ' nur bei jedem zweiten Mal gezeichnet ... 
        '
        toggleRowDifferentiator = Not toggleRowDifferentiator

        ' eine Zeile für die nächste Swimlane weiterschalten ...
        'curYPosition = curYPosition + rds.zeilenHoehe
        curYPosition = curYPosition + maxOffsetZeile * rds.zeilenHoehe


        ' ###########################################################
        ' alle Milestones in den Vordergrund holen 
        '
        Dim anzElements As Integer = swlMilestoneCollection.Count
        Dim arrayOFNames() As String
        Dim shapeGruppe As pptNS.ShapeRange

        If anzElements > 1 Then

            ReDim arrayOFNames(anzElements - 1)

            For i = 1 To anzElements
                arrayOFNames(i - 1) = CStr(swlMilestoneCollection.Item(i))
            Next

            shapeGruppe = rds.pptSlide.Shapes.Range(arrayOFNames)
            shapeGruppe.ZOrder(MsoZOrderCmd.msoBringToFront)

        End If

        'Dim slideShapes As pptNS.Shapes = rds.pptSlide.Shapes

        ' ###########################################################
        ' Zusammenfassen aller shapes in einer Gruppe 
        ' jetzt sollen alle gezeichneten Shapes gruppiert werden 
        '
        



        anzElements = shapeNameCollection.Count
        If anzElements > 1 Then

            ReDim arrayOFNames(anzElements - 1)

            For i = 1 To anzElements
                arrayOFNames(i - 1) = CStr(shapeNameCollection.Item(i))
            Next

            shapeGruppe = rds.pptSlide.Shapes.Range(arrayOFNames)
            shapeGruppe.Group()

        End If



    End Sub



    ''' <summary>
    ''' zeichnet die Projekte der Multiprojekt Sicht ( auch für extended Mode )
    ''' </summary>
    ''' <param name="pptslide">Powerpoint Folie</param>
    ''' <param name="projectCollection">enthält die nach der Position des Projekts auf der Projekttafel von oben nach unten, links nach rechts 
    ''' sortierte Liste an Projekten, die auf der Multiprojekt-Sicht ausgegeben werden sollen; die Namen sind die vollen Namen, pName+variantName </param>
    ''' <param name="StartofPPTCalendar">Beginn des Powerpoint Kalenders</param>
    ''' <param name="endOFPPTCalendar">Ende des Powerpoint Kalenders</param>
    ''' <param name="drawingAreaLeft"></param>
    ''' <param name="drawingAreaRight"></param>
    ''' <param name="drawingAreaTop"></param>
    ''' <param name="drawingAreaBottom"></param>
    ''' <param name="zeilenhoehe"></param>
    ''' <param name="projectListLeft"></param>
    ''' <param name="selectedPhases">welche Phasen sollen dargestellt werden; auch mehrere Phasen werden alle in eine Zeile gezeichnet</param>
    ''' <param name="selectedMilestones">welche Meilensteine sollen gezeichnet werden </param>
    ''' <param name="selectedRoles">welche Rollen sollen dargestellt werden; wenn mehrere Rollen ausgewählt sind, wird die Summe dargestellt</param>
    ''' <param name="selectedCosts">welche Kostenarten sollen dargestellt werden; wenn mehrere Kostenarten ausgewählt sind, wird die Summe dargestellt</param>
    ''' <param name="projectNameVorlagenShape"></param>
    ''' <param name="MsDescVorlagenShape">Vorlage, d.h Schriftart, Größe und relative Lage zum Meilenstein für die Element-(Meilenstein) Beschriftung </param>
    ''' <param name="MsDateVorlagenShape">Vorlage, d.h Schriftart, Größe und relative Lage zum Meilenstein für die Element-(Meilenstein) Beschriftung für das Meilenstein Datum </param>
    ''' <param name="PhDescVorlagenShape"></param>
    ''' <param name="PhDateVorlagenShape"></param>
    ''' <param name="phaseVorlagenShape">Vorlage für die Höhe der Phasen Shape; ale s Shape wird die entsprechende Darstellungsklasse verwendet </param>
    ''' <param name="milestoneVorlagenShape">Vorlage für die Höhe der Meilenstein Shape; als Shape wird die entsprechende Darstellungsklasse verwendet; dient auch zur relativen Einschätzung Meilenstein zu Phase</param>
    ''' <param name="projectVorlagenForm">Vorlage (Strichdicke, etc) für die Darstellung des Projekts; Farbe wird vom Projekt übernommen </param>
    ''' <param name="ampelVorlagenShape"></param>
    ''' <param name="yOffsetMsToText"></param>
    ''' <param name="yOffsetMsToDate"></param>
    ''' <param name="yOffsetPhToText"></param>
    ''' <param name="yOffsetPhToDate"></param>
    ''' <remarks>wenn ein Fehler auftritt wird eine Exception geworfen und im aufrufenden Programm eine entsprechende Fehlermeldung in das Shape</remarks>
    Sub zeichnePPTprojects(ByRef pptslide As pptNS.Slide, ByRef projectCollection As SortedList(Of Double, String), _
                            ByRef projDone As Integer, _
                            ByVal StartofPPTCalendar As Date, ByVal endOFPPTCalendar As Date, _
                            ByVal drawingAreaLeft As Double, ByVal drawingAreaRight As Double, ByVal drawingAreaTop As Double, ByVal drawingAreaBottom As Double, _
                            ByVal zeilenhoehe As Double, _
                            ByVal projectListLeft As Double, _
                            ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                            ByVal projectNameVorlagenShape As pptNS.Shape, _
                            ByVal MsDescVorlagenShape As pptNS.Shape, ByVal MsDateVorlagenShape As pptNS.Shape, _
                            ByVal PhDescVorlagenShape As pptNS.Shape, ByVal PhDateVorlagenShape As pptNS.Shape, _
                            ByVal phaseVorlagenShape As pptNS.Shape, ByVal milestoneVorlagenShape As pptNS.Shape, ByVal projectVorlagenForm As pptNS.Shape, _
                            ByVal ampelVorlagenShape As pptNS.Shape, ByVal rowDifferentiatorShape As pptNS.Shape, ByVal buColorShape As pptNS.Shape, _
                            ByVal phasedelimiterShape As pptNS.Shape, _
                            ByVal durationArrowShape As pptNS.Shape, ByVal durationTextShape As pptNS.Shape, _
                            ByVal yOffsetMsToText As Double, ByVal yOffsetMsToDate As Double, _
                            ByVal yOffsetPhToText As Double, ByVal yOffsetPhToDate As Double, _
                            ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)

        Dim addOn As Double = 0.0

        If Not IsNothing(durationArrowShape) And Not IsNothing(durationTextShape) Then

            'addOn = System.Math.Max(durationArrowShape.Height, durationTextShape.Height) * 11 / 15
            addOn = System.Math.Max(durationArrowShape.Height, durationTextShape.Height) ' tk Änderung 

        End If

        ' Bestimmen der Zeichenfläche
        Dim drawingAreaWidth As Double = drawingAreaRight - drawingAreaLeft
        Dim drawingAreaHeight As Double = drawingAreaBottom - drawingAreaTop


        'Dim tagesEinheit As Double
        Dim projectsToDraw As Integer
        Dim copiedShape As pptNS.ShapeRange
        Dim fullName As String
        Dim hproj As clsProjekt

        Dim phaseShape As xlNS.Shape
        Dim currentProjektIndex As Integer

        ' notwendig für das Positionieren des Duration Pfeils bzw. des DurationTextes
        Dim minX1 As Double
        Dim maxX2 As Double


        Dim anzahlTage As Integer = DateDiff(DateInterval.Day, StartofPPTCalendar, endOFPPTCalendar) + 1
        If anzahlTage <= 0 Then
            Throw New ArgumentException("Kalender Start bis Ende kann nicht 0 oder kleiner sein ..")
        End If



        ' Bestimmen der Position für den Projekt-Namen
        Dim projektNamenXPos As Double = projectListLeft
        Dim projektNamenYPos As Double
        Dim x1 As Double
        Dim x2 As Double
        Dim projektGrafikYPos As Double
        Dim phasenGrafikYPos As Double
        Dim milestoneGrafikYPos As Double
        Dim ampelGrafikYPos As Double
        Dim rowYPos As Double
        Dim grafikOffset As Double

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



        ' bestimme jetzt Y Start-Position für den Text bzw. die Grafik
        rowYPos = drawingAreaTop
        projektNamenYPos = drawingAreaTop + 0.5 * (zeilenhoehe - projectNameVorlagenShape.Height) + addOn
        projektGrafikYPos = drawingAreaTop + 0.5 * (zeilenhoehe - projectVorlagenForm.Height) + addOn
        phasenGrafikYPos = drawingAreaTop + 0.5 * (zeilenhoehe - phaseVorlagenShape.Height) + addOn
        milestoneGrafikYPos = drawingAreaTop + 0.5 * (zeilenhoehe - milestoneVorlagenShape.Height) + addOn
        ampelGrafikYPos = drawingAreaTop + 0.5 * (zeilenhoehe - ampelVorlagenShape.Height) + addOn
        grafikOffset = 0.5 * (zeilenhoehe - projectVorlagenForm.Height) + addOn

        projectsToDraw = projectCollection.Count

        If Not IsNothing(rowDifferentiatorShape) Then
            drawRowDifferentiator = True
        Else
            drawRowDifferentiator = False
        End If
        toggleRowDifferentiator = False

        If Not IsNothing(buColorShape) Then
            drawBUShape = True
            projektNamenXPos = projektNamenXPos + buColorShape.Width + 3
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
                    Dim neededSpace As Double = hproj.calcNeededLines(selectedPhases, selectedMilestones, True, Not awinSettings.mppShowAllIfOne) * zeilenhoehe
                    If neededSpace > drawingAreaBottom - drawingAreaTop Then

                        ' Projekt kann nicht gezeichnet werden, da nicht alle Phasen auf eine Seite passen, 
                        ' trotzdem muss das Projekt weitergezählt werden, damit das nächste zu zeichnende Projekt angegangen wird
                        projDone = projDone + 1
                        ' zuwenig Platz auf der Seite
                        Throw New ArgumentException("Für Projekt '" & fullName & "' ist zuwenig Platz auf einer Seite")

                    Else

                        If projektGrafikYPos - grafikOffset + hproj.calcNeededLines(selectedPhases, selectedMilestones, True, Not awinSettings.mppShowAllIfOne) * zeilenhoehe > drawingAreaBottom Then
                            Exit For
                        End If
                    End If
                End If

                If worker.WorkerSupportsCancellation Then

                    If worker.CancellationPending Then
                        e.Cancel = True
                        e.Result = "Berichterstellung abgebrochen ..."
                        Exit For
                    End If

                End If

                ' Zwischenbericht abgeben ...
                e.Result = "Projekt '" & hproj.getShapeText & "' wird gezeichnet  ...."
                If worker.WorkerReportsProgress Then
                    worker.ReportProgress(0, e)
                End If


                '
                ' zeichne den Projekt-Namen
                projectNameVorlagenShape.Copy()
                copiedShape = pptslide.Shapes.Paste()
                Dim projectNameShape As pptNS.Shape = copiedShape.Item(1)

                With copiedShape(1)
                    .Top = CSng(projektNamenYPos)
                    .Left = CSng(projektNamenXPos)
                    If currentProjektIndex > 1 And lastProjectName = hproj.name Then
                        '.TextFrame2.TextRange.Text = "... " & hproj.variantName & " " & hproj.VorlagenName
                        .TextFrame2.TextRange.Text = "... " & hproj.variantName
                    Else
                        '.TextFrame2.TextRange.Text = hproj.getShapeText & " " & hproj.VorlagenName
                        .TextFrame2.TextRange.Text = hproj.getShapeText
                    End If
                    lastProjectName = hproj.name
                    .Name = .Name & .Id
                End With

                projektNamenYPos = projektNamenYPos + zeilenhoehe


                ' zeichne jetzt ggf die Projekt-Ampel 
                If awinSettings.mppShowAmpel And Not IsNothing(ampelVorlagenShape) Then
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

                    ampelVorlagenShape.Copy()
                    copiedShape = pptslide.Shapes.Paste()

                    With copiedShape(1)
                        .Top = CSng(ampelGrafikYPos)
                        .Left = CSng(drawingAreaLeft - (.Width + 3))
                        .Width = .Height
                        .Line.ForeColor.RGB = CInt(statusColor)
                        .Fill.ForeColor.RGB = CInt(statusColor)
                        .Name = .Name & .Id
                    End With

                    ampelGrafikYPos = ampelGrafikYPos + zeilenhoehe

                End If

                '
                ' zeichne jetzt das Projekt 
                Call calculatePPTx1x2(StartofPPTCalendar, endOFPPTCalendar, hproj.startDate, hproj.endeDate, _
                                        drawingAreaLeft, drawingAreaWidth, x1, x2)


                ' jetzt muss überprüft werden, ob projectName zu lang ist - dann wird der Name entsprechend abgekürzt ...
                With projectNameShape
                    If .Left + .Width > x1 Then
                        ' jetzt muss der Name entsprechend gekürzt werden 
                        Dim longName As String = .TextFrame2.TextRange.Text
                        Dim shortName As String = ""

                        .TextFrame2.TextRange.Text = shortName
                        Dim stringIX As Integer = 0
                        Do While .Left + .Width < x1 And stringIX <= longName.Length - 1
                            shortName = shortName & longName.Chars(stringIX)
                            stringIX = stringIX + 1
                            .TextFrame2.TextRange.Text = shortName
                        Loop

                    End If
                End With




                If awinSettings.mppShowProjectLine Then

                    projectVorlagenForm.Copy()
                    copiedShape = pptslide.Shapes.Paste()
                    With copiedShape(1)
                        .Top = CSng(projektGrafikYPos)
                        .Left = CSng(x1)
                        .Width = CSng(x2 - x1)
                        .Name = .Name & .Id
                        ' wenn Projektstart vor dem Kalender-Start liegt: kein Projektstart Symbol zeichnen
                        If DateDiff(DateInterval.Day, hproj.startDate, StartofPPTCalendar) > 0 Then
                            .Line.BeginArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone
                        End If

                        ' wenn Projektende nach dem Kalender-Ende liegt: kein Projektende Symbol zeichnen
                        If DateDiff(DateInterval.Day, hproj.endeDate, endOFPPTCalendar) < 0 Then
                            .Line.EndArrowheadStyle = MsoArrowheadStyle.msoArrowheadNone
                        End If
                    End With

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
                        Dim selPhaseName As String = ""

                        While j <= selectedPhases.Count And Not found

                            Call splitHryFullnameTo2(CStr(selectedPhases(j)), selPhaseName, breadcrumb)
                            If cphase.name = selPhaseName Then
                                If vglBreadCrumb.EndsWith(breadcrumb) Then
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
                                                phasenGrafikYPos = phasenGrafikYPos + zeilenhoehe
                                                ' Y-Position für BU und Hintergrund-einfärbung erhöhen je gezeichneter Zeile
                                                '''' ur:20.04.2015:  rowYPos = rowYPos + zeilenhoehe
                                                ' Y-Position für Projektnamen erhöhen je gezeichneter Phase
                                                projektNamenYPos = projektNamenYPos + zeilenhoehe
                                                ' Y-Position für Meilensteine der aktuellen Phase erhöhen je gezeichneter Phase
                                                milestoneGrafikYPos = milestoneGrafikYPos + zeilenhoehe
                                                ' Y-Position der Ampel, sofern sie zu dem Projekt gezeichnet werden soll
                                                ampelGrafikYPos = ampelGrafikYPos + zeilenhoehe
                                                anzZeilenGezeichnet = anzZeilenGezeichnet + 1


                                                ' ur: Meilensteine aus drawliste.value zeichnen
                                                Dim zeichnenMS As Boolean = False
                                                Dim msliste As SortedList
                                                Dim msi As Integer
                                                msliste = drawliste(lastPhase.nameID)

                                                For msi = 0 To msliste.Count - 1
                                                    Dim msID As String = msliste.GetByIndex(msi)
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
                                                        Call zeichneMeilensteininAktZeile(pptslide, msShapeNames, milestone, hproj, milestoneGrafikYPos, _
                                                                                            StartofPPTCalendar, endOFPPTCalendar, _
                                                                                            drawingAreaLeft, drawingAreaRight, drawingAreaTop, drawingAreaBottom, _
                                                                                            MsDescVorlagenShape, MsDateVorlagenShape, milestoneVorlagenShape, _
                                                                                            yOffsetMsToText, yOffsetMsToDate)
                                                    End If

                                                Next

                                            End If
                                            phasenGrafikYPos = phasenGrafikYPos + zeilenhoehe
                                            ' Y-Position für BU und Hintergrund-einfärbung erhöhen je gezeichneter Zeile
                                            '''' ur:20.04.2015:  rowYPos = rowYPos + zeilenhoehe
                                            ' Y-Position für Projektnamen erhöhen je gezeichneter Phase
                                            projektNamenYPos = projektNamenYPos + zeilenhoehe
                                            ' Y-Position für Meilensteine der aktuellen Phase erhöhen je gezeichneter Phase
                                            milestoneGrafikYPos = milestoneGrafikYPos + zeilenhoehe
                                            ' Y-Position der Ampel, sofern sie zu dem Projekt gezeichnet werden soll
                                            ampelGrafikYPos = ampelGrafikYPos + zeilenhoehe
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
                                    phaseShape = PhaseDefinitions.getShape(phaseName)
                                Else
                                    phaseShape = missingPhaseDefinitions.getShape(phaseName)
                                End If


                                Dim phaseStart As Date = cphase.getStartDate
                                Dim phaseEnd As Date = cphase.getEndDate
                                'Dim phShortname As String = PhaseDefinitions.getAbbrev(phaseName).Trim
                                ' erhänzt tk
                                Dim phShortname As String = ""
                                phShortname = hproj.hierarchy.getBestNameOfID(cphase.nameID, Not awinSettings.mppUseOriginalNames, _
                                                                              awinSettings.mppUseAbbreviation)

                                Call calculatePPTx1x2(StartofPPTCalendar, endOFPPTCalendar, phaseStart, phaseEnd, _
                                                    drawingAreaLeft, drawingAreaWidth, x1, x2)



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

                                    PhDescVorlagenShape.Copy()
                                    copiedShape = pptslide.Shapes.Paste()
                                    With copiedShape(1)

                                        .Name = .Name & .Id
                                        .TextFrame2.TextRange.Text = phShortname
                                        .TextFrame2.MarginLeft = 0.0
                                        .TextFrame2.MarginRight = 0.0
                                        '.Top = CSng(projektGrafikYPos) - .Height
                                        .Top = CSng(phasenGrafikYPos) + CSng(yOffsetPhToText) - 2
                                        .Left = CSng(x1)
                                        If .Left < drawingAreaLeft Then
                                            .Left = CSng(drawingAreaLeft)
                                        End If
                                        .TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft

                                    End With


                                End If

                                ' jetzt muss ggf das Datum angebracht werden 
                                If awinSettings.mppShowPhDate Then
                                    'Dim phDateText As String = phaseStart.ToShortDateString
                                    Dim phDateText As String = phaseStart.Day.ToString & "." & phaseStart.Month.ToString
                                    Dim rightX As Double, addHeight As Double

                                    PhDateVorlagenShape.Copy()
                                    copiedShape = pptslide.Shapes.Paste()
                                    With copiedShape(1)

                                        .Name = .Name & .Id
                                        .TextFrame2.TextRange.Text = phDateText
                                        .TextFrame2.MarginLeft = 0.0
                                        .TextFrame2.MarginRight = 0.0
                                        '.Top = CSng(projektGrafikYPos)
                                        .Top = CSng(phasenGrafikYPos) + CSng(yOffsetPhToDate) + 1
                                        .Left = CSng(x1) + 1
                                        If .Left < drawingAreaLeft Then
                                            .Left = CSng(drawingAreaLeft + 1)
                                        End If
                                        .TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft

                                        rightX = .Left + .Width
                                        addHeight = .Height * 0.7

                                    End With


                                    ' Änderung tk 14.3.15 kein Voranstellen des Phasen Namens mehr ... 
                                    phDateText = phaseEnd.Day.ToString & "." & phaseEnd.Month.ToString

                                    PhDateVorlagenShape.Copy()
                                    copiedShape = pptslide.Shapes.Paste()
                                    With copiedShape(1)

                                        .Name = .Name & .Id
                                        .TextFrame2.TextRange.Text = phDateText
                                        .TextFrame2.MarginLeft = 0.0
                                        .TextFrame2.MarginRight = 0.0
                                        .Top = CSng(phasenGrafikYPos) + CSng(yOffsetPhToDate) + 1
                                        .Left = CSng(x2) - .Width - 1
                                        If .Left + .Width > drawingAreaRight Then
                                            .Left = drawingAreaRight - (.Width + 1)
                                        End If
                                        .TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignRight

                                        If rightX >= .Left Then
                                            .Top = .Top + addHeight
                                        End If

                                    End With

                                End If

                                ' jetzt muss ggf das Phase Delimiter Shape angebracht werden 
                                If Not IsNothing(phasedelimiterShape) And selectedPhases.Count > 1 Then

                                    ' linker Delimiter 
                                    phasedelimiterShape.Copy()
                                    copiedShape = pptslide.Shapes.Paste()

                                    With copiedShape(1)

                                        .Height = 1.3 * phaseShape.Height
                                        .Top = CSng(phasenGrafikYPos)
                                        .Left = CSng(x1) - .Width * 0.5
                                        .Name = .Name & .Id

                                    End With

                                    ' rechter Delimiter 
                                    phasedelimiterShape.Copy()
                                    copiedShape = pptslide.Shapes.Paste()

                                    With copiedShape(1)

                                        .Height = 1.3 * phaseShape.Height
                                        .Top = CSng(phasenGrafikYPos)
                                        .Left = CSng(x2) + .Width * 0.5
                                        .Name = .Name & .Id

                                    End With

                                End If

                                ' jetzt das Shape zeichnen 
                                phaseShape.Copy()
                                copiedShape = pptslide.Shapes.Paste()

                                With copiedShape(1)
                                    .Top = CSng(phasenGrafikYPos)
                                    .Left = CSng(x1)
                                    .Width = CSng(x2 - x1)
                                    .Height = phaseVorlagenShape.Height
                                    .Name = .Name & .Id
                                End With

                                phShapeNames.Add(copiedShape.Name)

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

                                Call splitHryFullnameTo2(CStr(selectedMilestones.Item(ix)), milestoneName, breadcrumbMS)


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

                                                Call zeichneMeilensteininAktZeile(pptslide, msShapeNames, ms, hproj, milestoneGrafikYPos, _
                                                                                                              StartofPPTCalendar, endOFPPTCalendar, _
                                                                                                              drawingAreaLeft, drawingAreaRight, drawingAreaTop, drawingAreaBottom, _
                                                                                                              MsDescVorlagenShape, MsDateVorlagenShape, milestoneVorlagenShape, _
                                                                                                              yOffsetMsToText, yOffsetMsToDate)




                                            End If


                                        Else
                                            ' selektierter Meilenstein 'milestoneName' nicht in dieser Phase enthalten
                                            ' also: nichts tun
                                        End If

                                    End If


                                Next mx

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
                            phasenGrafikYPos = phasenGrafikYPos + zeilenhoehe
                            ' Y-Position für BU und Hintergrund-einfärbung erhöhen je gezeichneter Zeile
                            '''' ur:20.04.2015:  rowYPos = rowYPos + zeilenhoehe
                            ' Y-Position für Projektnamen erhöhen je gezeichneter Phase
                            projektNamenYPos = projektNamenYPos + zeilenhoehe
                            ' Y-Position für Meilensteine der aktuellen Phase erhöhen je gezeichneter Phase
                            milestoneGrafikYPos = milestoneGrafikYPos + zeilenhoehe
                            ' Y-Position der Ampel, sofern sie zu dem Projekt gezeichnet werden soll
                            ampelGrafikYPos = ampelGrafikYPos + zeilenhoehe
                            anzZeilenGezeichnet = anzZeilenGezeichnet + 1


                            ' ur: Meilensteine aus drawliste.value zeichnen
                            Dim zeichnenMS As Boolean = False
                            Dim msliste As SortedList
                            Dim msi As Integer
                            msliste = drawliste(lastPhase.nameID)

                            For msi = 0 To msliste.Count - 1

                                Dim msID As String = msliste.GetByIndex(msi)
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
                                    Call zeichneMeilensteininAktZeile(pptslide, msShapeNames, milestone, hproj, milestoneGrafikYPos, _
                                                                        StartofPPTCalendar, endOFPPTCalendar, _
                                                                        drawingAreaLeft, drawingAreaRight, drawingAreaTop, drawingAreaBottom, _
                                                                        MsDescVorlagenShape, MsDateVorlagenShape, milestoneVorlagenShape, _
                                                                        yOffsetMsToText, yOffsetMsToDate)
                                End If
                            Next


                        End If


                    End If

                    '''' ur: 01.10.2015: selektierte Meilensteine zeichnen, die zu keiner der selektierten Phasen gehören.

                    If drawliste.ContainsKey(rootPhaseName) Then

                        phasenGrafikYPos = phasenGrafikYPos + zeilenhoehe
                        ' Y-Position für BU und Hintergrund-einfärbung erhöhen je gezeichneter Zeile
                        '''' ur:20.04.2015:  rowYPos = rowYPos + zeilenhoehe
                        ' Y-Position für Projektnamen erhöhen je gezeichneter Phase
                        projektNamenYPos = projektNamenYPos + zeilenhoehe
                        ' Y-Position für Meilensteine der aktuellen Phase erhöhen je gezeichneter Phase
                        milestoneGrafikYPos = milestoneGrafikYPos + zeilenhoehe
                        ' Y-Position der Ampel, sofern sie zu dem Projekt gezeichnet werden soll
                        ampelGrafikYPos = ampelGrafikYPos + zeilenhoehe
                        anzZeilenGezeichnet = anzZeilenGezeichnet + 1


                        ' ur: Meilensteine aus drawliste.value zeichnen
                        Dim zeichnenMS As Boolean = False
                        Dim msliste As SortedList
                        Dim msi As Integer
                        msliste = drawliste(rootPhaseName)

                        For msi = 0 To msliste.Count - 1

                            Dim msID As String = msliste.GetByIndex(msi)
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
                                Call zeichneMeilensteininAktZeile(pptslide, msShapeNames, milestone, hproj, milestoneGrafikYPos, _
                                                                    StartofPPTCalendar, endOFPPTCalendar, _
                                                                    drawingAreaLeft, drawingAreaRight, drawingAreaTop, drawingAreaBottom, _
                                                                    MsDescVorlagenShape, MsDateVorlagenShape, milestoneVorlagenShape, _
                                                                    yOffsetMsToText, yOffsetMsToDate)
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

                            Dim msID As String = msliste.GetByIndex(msi)
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
                                Call zeichneMeilensteininAktZeile(pptslide, msShapeNames, milestone, hproj, milestoneGrafikYPos, _
                                                                    StartofPPTCalendar, endOFPPTCalendar, _
                                                                    drawingAreaLeft, drawingAreaRight, drawingAreaTop, drawingAreaBottom, _
                                                                    MsDescVorlagenShape, MsDateVorlagenShape, milestoneVorlagenShape, _
                                                                    yOffsetMsToText, yOffsetMsToDate)
                            End If
                        Next

                    Next kvp
                End If




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


                    buColorShape.Copy()
                    copiedShape = pptslide.Shapes.Paste()
                    With copiedShape(1)
                        .Top = CSng(rowYPos)
                        .Left = CSng(projectListLeft)
                        '' '' ''Dim neededLines As Double = hproj.calcNeededLines(selectedPhases, awinSettings.mppExtendedMode, Not awinSettings.mppShowAllIfOne)
                        '' '' ''.Height = hproj.calcNeededLines(selectedPhases, awinSettings.mppExtendedMode, Not awinSettings.mppShowAllIfOne) * zeilenhoehe
                        .Height = anzZeilenGezeichnet * zeilenhoehe
                        .Fill.ForeColor.RGB = CInt(buFarbe)
                        .Name = .Name & .Id
                        ' width ist die in der Vorlage angegebene Width 
                    End With

                End If


                ' optionales zeichnen der Zeilen-Markierung
                If drawRowDifferentiator And toggleRowDifferentiator Then
                    ' zeichnen des RowDifferentiators 
                    rowDifferentiatorShape.Copy()
                    copiedShape = pptslide.Shapes.Paste()
                    With copiedShape(1)
                        .Top = CSng(rowYPos)
                        .Left = CSng(projectListLeft)
                        '''''.Height = hproj.calcNeededLines(selectedPhases, awinSettings.mppExtendedMode, Not awinSettings.mppShowAllIfOne) * zeilenhoehe
                        .Height = anzZeilenGezeichnet * zeilenhoehe
                        .Width = drawingAreaRight - .Left
                        .Name = .Name & .Id
                        .ZOrder(MsoZOrderCmd.msoSendToBack)
                    End With
                End If

                ' dadurch wird die Zeilen - bzw. Projekt - Markierung nur bei jedem zweiten Mal gezeichnet ... 
                toggleRowDifferentiator = Not toggleRowDifferentiator

                ' jetzt muss ggf die duration eingezeichnet werden 
                If Not IsNothing(durationArrowShape) And Not IsNothing(durationTextShape) Then

                    ' Pfeil mit Länge der Dauer zeichnen 
                    durationArrowShape.Copy()
                    copiedShape = pptslide.Shapes.Paste()

                    Dim pfeilbreite As Double = maxX2 - minX1

                    With copiedShape(1)
                        .Top = CSng(rowYPos) + 3 + 0.5 * (addOn - .Height)
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

                    durationTextShape.Copy()

                    copiedShape = pptslide.Shapes.Paste()

                    With copiedShape(1)
                        .TextFrame2.TextRange.Text = dauerInM.ToString("0.0") & " M"
                        .Top = CSng(rowYPos) + 3 + 0.5 * (addOn - .Height)
                        .Left = CSng(minX1 + (pfeilbreite - .Width) / 2)
                        .Name = .Name & .Id
                    End With

                End If


                projDone = projDone + 1
                If Not awinSettings.mppExtendedMode Then

                    projektGrafikYPos = projektGrafikYPos + zeilenhoehe
                    rowYPos = rowYPos + zeilenhoehe
                Else

                    projektGrafikYPos = projektGrafikYPos + anzZeilenGezeichnet * zeilenhoehe
                    rowYPos = rowYPos + anzZeilenGezeichnet * zeilenhoehe

                End If

                phasenGrafikYPos = phasenGrafikYPos + zeilenhoehe
                milestoneGrafikYPos = milestoneGrafikYPos + zeilenhoehe

                If projektGrafikYPos > drawingAreaBottom Then
                    Exit For
                End If



            End If


        Next            ' nächstes Projekt zeichnen


        '
        ' wenn Texte gezeichnet wurden, müssen jetzt die Phasen, dann die Meilensteine in den Vordergrund geholt werden 
        If awinSettings.mppShowMsDate Or awinSettings.mppShowMsName Or _
            awinSettings.mppShowPhDate Or awinSettings.mppShowPhName Then
            ' Phasen vorholen 
            Dim anzElements As Integer
            anzElements = phShapeNames.Count

            If anzElements > 0 Then

                ReDim arrayOfNames(anzElements - 1)
                For ix = 1 To anzElements
                    arrayOfNames(ix - 1) = CStr(phShapeNames.Item(ix))
                Next

                Try
                    CType(pptslide.Shapes.Range(arrayOfNames), pptNS.ShapeRange).ZOrder(MsoZOrderCmd.msoBringToFront)
                Catch ex As Exception

                End Try

            End If

            anzElements = msShapeNames.Count

            If anzElements > 0 Then

                ReDim arrayOfNames(anzElements - 1)
                For ix = 1 To anzElements
                    arrayOfNames(ix - 1) = CStr(msShapeNames.Item(ix))
                Next

                Try
                    CType(pptslide.Shapes.Range(arrayOfNames), pptNS.ShapeRange).ZOrder(MsoZOrderCmd.msoBringToFront)
                Catch ex As Exception

                End Try

            End If

        End If

        If currentProjektIndex < projectCollection.Count And awinSettings.mppOnePage Then
            Throw New ArgumentException("es konnten nur " & _
                                        currentProjektIndex.ToString & " von " & projectsToDraw.ToString & _
                                        " Projekten gezeichnet werden ... " & vbLf & _
                                        "bitte verwenden Sie ein anderes Vorlagen-Format")
        End If



    End Sub
    ''' <summary>
    ''' Zeichnet den Meilenstein MS im PPTslide an Position MilestoneGrafikYPOs
    ''' </summary> 
    ''' <param name="pptslide">Powerpoint Folie, in die gezeichnet werden soll</param>
    ''' <param name="msShapeNames">Name des Meilenstein -Shapes in der Powerpoint-Folie</param>
    ''' <param name="MS">zu zeichnender Meilenstein</param>
    ''' <param name="hproj">Projekt, zu dem der Meilenstein gehört</param>
    ''' <param name="milestoneGrafikYPos">Position des Meilenstein</param>
    ''' <param name="StartofPPTCalendar">Beginn des Powerpoint Kalenders</param>
    ''' <param name="endOFPPTCalendar">Ende des Powerpoint Kalenders</param>
    ''' <param name="drawingAreaLeft"></param>
    ''' <param name="drawingAreaRight"></param>
    ''' <param name="drawingAreaTop"></param>
    ''' <param name="drawingAreaBottom"></param>
    ''' <param name="MsDescVorlagenShape">Vorlage, d.h Schriftart, Größe und relative Lage zum Meilenstein für die Element-(Meilenstein) Beschriftung </param>
    ''' <param name="MsDateVorlagenShape">Vorlage, d.h Schriftart, Größe und relative Lage zum Meilenstein für die Element-(Meilenstein) Beschriftung für das Meilenstein Datum </param>
    ''' <param name="milestoneVorlagenShape">Vorlage für die Höhe der Meilenstein Shape; als Shape wird die entsprechende Darstellungsklasse verwendet; dient auch zur relativen Einschätzung Meilenstein zu Phase</param>
    ''' <param name="yOffsetMsToText"></param>
    ''' <param name="yOffsetMsToDate"></param>
    ''' <remarks>wenn ein Fehler auftritt wird eine Exception geworfen und im aufrufenden Programm eine entsprechende Fehlermeldung in das Shape</remarks>

    Private Sub zeichneMeilensteininAktZeile(ByRef pptslide As pptNS.Slide, _
                                                 ByRef msShapeNames As Collection, _
                                                 ByVal MS As clsMeilenstein, _
                                                 ByVal hproj As clsProjekt, _
                                                 ByVal milestoneGrafikYPos As Double, _
                                                 ByVal StartofPPTCalendar As Date, ByVal endOFPPTCalendar As Date, _
                                                 ByVal drawingAreaLeft As Double, ByVal drawingAreaRight As Double, _
                                                 ByVal drawingAreaTop As Double, ByVal drawingAreaBottom As Double, _
                                                 ByVal MsDescVorlagenShape As pptNS.Shape, _
                                                 ByVal MsDateVorlagenShape As pptNS.Shape, _
                                                 ByVal milestoneVorlagenShape As pptNS.Shape, _
                                                 ByVal yOffsetMsToText As Double, ByVal yOffsetMsToDate As Double)

        Dim milestoneTypShape As xlNS.Shape
        Dim copiedShape As pptNS.ShapeRange


        ' notwendig für das Positionieren des Duration Pfeils bzw. des DurationTextes
        Dim minX1 As Double
        Dim maxX2 As Double

        Dim x1 As Double
        Dim x2 As Double



        ' Änderung tk 26.11.15
        If MilestoneDefinitions.Contains(MS.name) Then
            milestoneTypShape = MilestoneDefinitions.getShape(MS.name)
        Else
            milestoneTypShape = missingMilestoneDefinitions.getShape(MS.name)
        End If


        Dim msdate As Date = MS.getDate

        Dim seitenverhaeltnis As Double
        With milestoneTypShape
            seitenverhaeltnis = .Height / .Width
        End With


        Call calculatePPTx1x2(StartofPPTCalendar, endOFPPTCalendar, msdate, msdate, _
                            drawingAreaLeft, drawingAreaRight - drawingAreaLeft, x1, x2)


        If minX1 > x1 Then
            minX1 = x1
        End If

        If maxX2 < x2 Then
            maxX2 = x2
        End If

        ' jetzt muss ggf die Beschriftung angebracht werden 
        ' die muss vor dem Meilenstein angebracht werden, weil der nicht von der Füllung des Schriftfeldes 
        ' überdeckt werden soll 
        If awinSettings.mppShowMsName Then

            Dim msBeschriftung As String
            msBeschriftung = hproj.hierarchy.getBestNameOfID(MS.nameID, Not awinSettings.mppUseOriginalNames, _
                                                             awinSettings.mppUseAbbreviation)

            MsDescVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .TextFrame2.TextRange.Text = msBeschriftung
                .Top = CSng(milestoneGrafikYPos) + CSng(yOffsetMsToText)
                '.Left = CSng(x1) - .Width / 2
                .Left = CSng(x1) - .Width / 2
                .Name = .Name & .Id

            End With


        End If

        ' jetzt muss ggf das Datum angebracht werden 
        If awinSettings.mppShowMsDate Then
            'Dim msDateText As String = msDate.ToShortDateString
            Dim msDateText As String
            msDateText = msdate.Day.ToString & "." & msdate.Month.ToString

            MsDateVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .TextFrame2.TextRange.Text = msDateText
                .Top = CSng(milestoneGrafikYPos) + CSng(yOffsetMsToDate)
                .Left = CSng(x1) - .Width / 2
                .Name = .Name & .Id

            End With

        End If


        ' Erst jetzt wird der Meilenstein gezeichnet 
        milestoneTypShape.Copy()
        copiedShape = pptslide.Shapes.Paste()



        With copiedShape.Item(1)
            .Top = CSng(milestoneGrafikYPos)
            .Height = milestoneVorlagenShape.Height
            .Width = .Height / seitenverhaeltnis
            .Left = CSng(x1) - .Width / 2
            .Name = .Name & .Id
            If awinSettings.mppShowAmpel Then
                .Glow.Color.RGB = CInt(MS.getBewertung(1).color)
                If .Glow.Radius = 0 Then
                    .Glow.Radius = 5
                End If
            End If
        End With

        msShapeNames.Add(copiedShape.Name)


    End Sub

    ''' <summary>
    ''' zeichnet das aktuelle Segment; 
    ''' optional kann ein Modus angegeben werden sowie das Projekt, um beispielsweise Darstellungsklasse der ein PRojekt gemäß Modus in di
    ''' </summary>
    ''' <param name="rds">enthält sowohl slide als auch die Hilfs-Shapes </param>
    ''' <param name="curYPosition">gibt die aktuelle Y-Position wieder , ab der gezeichnet werden kann; ist am Ende wieder auf der nächsten freien Zeile  </param>
    ''' <remarks></remarks>
    Private Sub zeichneSwlSegmentinAktZeile(ByRef rds As clsPPTShapes, ByRef curYPosition As Double, ByVal segmentPhaseID As String, _
                                     Optional ByVal modus As Integer = 0, Optional ByVal hproj As clsProjekt = Nothing)

        Dim copiedShape As pptNS.ShapeRange

        rds.segmentVorlagenShape.Copy()
        copiedShape = rds.pptSlide.Shapes.Paste()

        If modus = 0 Then
            With copiedShape.Item(1)
                .Top = CSng(curYPosition)
                .Left = CSng(rds.drawingAreaLeft)
                .Width = CSng(rds.drawingAreaWidth)
                .TextFrame2.TextRange.Text = elemNameOfElemID(segmentPhaseID)
                .Name = .Name & .Id
                .AlternativeText = "Segment " & elemNameOfElemID(segmentPhaseID)

                ' Current Y-Position aktualisieren 
                curYPosition = curYPosition + .Height
            End With
        End If


    End Sub



    
    ''' <summary>
    ''' zeichnet eine Phase in der aktuellen Swimlane 
    ''' </summary>
    ''' <param name="rds"></param>
    ''' <param name="shapeNames"></param>
    ''' <param name="hproj"></param>
    ''' <param name="phaseID"></param>
    ''' <param name="yPosition"></param>
    ''' <remarks></remarks>
    Private Sub zeichnePhaseinSwimlane(ByRef rds As clsPPTShapes, ByRef shapeNames As Collection, _
                                           ByVal hproj As clsProjekt, _
                                           ByVal swimlaneID As String, _
                                           ByVal phaseID As String, _
                                           ByVal yPosition As Double)

        Dim phaseTypShape As xlNS.Shape
        Dim copiedShape As pptNS.ShapeRange
        Dim phaseName As String = elemNameOfElemID(phaseID)
        Dim cphase As clsPhase = hproj.getPhaseByID(phaseID)

        If IsNothing(cphase) Then
            Exit Sub ' nichts machen 
        End If



        Dim x1 As Double
        Dim x2 As Double


        Dim phDescription As String = hproj.hierarchy.getBestNameOfID(phaseID, Not awinSettings.mppUseOriginalNames, _
                                                                awinSettings.mppUseAbbreviation, swimlaneID)

        If PhaseDefinitions.Contains(phaseName) Then
            phaseTypShape = PhaseDefinitions.getShape(phaseName)
        Else
            phaseTypShape = missingPhaseDefinitions.getShape(phaseName)
        End If


        ' jetzt wegen evtl innerer Beschriftung den Size-Faktor bestimmen 
        Dim sizeFaktor As Double = 1.0

        If awinSettings.mppUseInnerText Then

            phaseTypShape.Copy()
            copiedShape = rds.pptSlide.Shapes.Paste()

            With copiedShape
                If .Height > 0.0 Then
                    sizeFaktor = rds.phaseVorlagenShape.Height / .Height
                End If
                .Delete()
            End With

        End If
        



        Dim phStartDate As Date = cphase.getStartDate
        Dim phEndDate As Date = cphase.getEndDate
        Dim phDateText As String = phStartDate.Day.ToString & "." & phStartDate.Month.ToString & " - " & _
                                phEndDate.Day.ToString & "." & phEndDate.Month.ToString


        Call rds.calculatePPTx1x2(phStartDate, phEndDate, x1, x2)

        If x2 <= rds.drawingAreaLeft Or x1 >= rds.drawingAreaRight Then
            ' Fertig 
        Else

            ' jetzt muss ggf die Beschriftung angebracht werden 
            ' die muss vor der Phase angebracht werden, weil der nicht von der Füllung des Schriftfeldes 
            ' überdeckt werden soll 
            If awinSettings.mppShowPhName And (Not awinSettings.mppUseInnerText) Then

                rds.PhDescVorlagenShape.Copy()
                copiedShape = rds.pptSlide.Shapes.Paste()
                With copiedShape(1)

                    .TextFrame2.TextRange.Text = phDescription
                    .Top = CSng(yPosition + rds.YPhasenText)
                    .Left = CSng(x1)
                    If .Left + .Width > rds.drawingAreaRight + 2 Then
                        .Left = rds.drawingAreaRight - .Width + 2
                    End If
                    .Name = .Name & .Id

                    shapeNames.Add(.Name, .Name)
                End With


            End If

            ' jetzt muss ggf das Datum angebracht werden 
            If awinSettings.mppShowPhDate And (Not awinSettings.mppUseInnerText) Then

                rds.PhDateVorlagenShape.Copy()
                copiedShape = rds.pptSlide.Shapes.Paste()
                With copiedShape(1)

                    .TextFrame2.TextRange.Text = phDateText
                    .Top = CSng(yPosition + rds.YPhasenDatum)
                    .Left = CSng(x1)
                    If .Left + .Width > rds.drawingAreaRight + 2 Then
                        .Left = rds.drawingAreaRight - .Width + 2
                    End If

                    .Name = .Name & .Id

                    shapeNames.Add(.Name, .Name)
                End With

            End If


            ' Erst jetzt wird die Phase gezeichnet 
            phaseTypShape.Copy()
            copiedShape = rds.pptSlide.Shapes.Paste()

            With copiedShape.Item(1)
                .Top = CSng(yPosition + rds.YPhase)
                .Height = rds.phaseVorlagenShape.Height
                .Width = CSng(x2 - x1)
                .Left = CSng(x1)
                .Name = .Name & .Id
                .Title = phaseName
                .AlternativeText = phStartDate.ToShortDateString & " - " & phEndDate.ToShortDateString

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

                shapeNames.Add(.Name, .Name)
            End With


        End If




    End Sub

    ''' <summary>
    ''' zeichnet den angegebenen Meilenstein in der Zeile mit YPosition
    ''' es wird eine Größenanpassung gemäß Faktor im Vergleich zur Darstellungsklasse gemacht  
    ''' </summary>
    ''' <param name="rds"></param>
    ''' <param name="shapeNames">die Namen der erzeugten Shapes</param>
    ''' <param name="hproj">das Projekt selber </param>
    ''' <param name="milestoneID">die ID des Meilensteins, der gezeichnet werden soll</param>
    ''' <param name="yPosition">die yPosition auf der Zeichenfläche; die x-Position wird errechnet</param>
    ''' <remarks></remarks>
    Private Sub zeichneMeilensteinInSwimlane(ByRef rds As clsPPTShapes, ByRef shapeNames As Collection, _
                                               ByVal hproj As clsProjekt, _
                                               ByVal swimlaneID As String, _
                                               ByVal milestoneID As String, _
                                               ByVal yPosition As Double)

        Dim milestoneTypShape As xlNS.Shape
        Dim copiedShape As pptNS.ShapeRange
        Dim milestoneName As String = elemNameOfElemID(milestoneID)
        Dim cMilestone As clsMeilenstein = hproj.getMilestoneByID(milestoneID)

        If IsNothing(cMilestone) Then
            Exit Sub ' einfach nichts machen 
        End If


        Dim x1 As Double
        Dim x2 As Double


        Dim msBeschriftung As String

        If MilestoneDefinitions.Contains(milestoneName) Then
            milestoneTypShape = MilestoneDefinitions.getShape(milestoneName)
        Else
            milestoneTypShape = missingMilestoneDefinitions.getShape(milestoneName)
        End If

        Dim sizeFaktor As Double
        milestoneTypShape.Copy()
        copiedShape = rds.pptSlide.Shapes.Paste()

        With copiedShape
            If .Height <= 0.0 Then
                sizeFaktor = 1.0
            Else
                sizeFaktor = rds.milestoneVorlagenShape.Height / .Height
            End If
            .Delete()
        End With



        Dim msDate As Date = cMilestone.getDate


        Call rds.calculatePPTx1x2(msDate, msDate, x1, x2)

        If x2 <= rds.drawingAreaLeft Or x1 >= rds.drawingAreaRight Then
            ' Fertig 
        Else

            ' jetzt muss ggf die Beschriftung angebracht werden 
            ' die muss vor der Phase angebracht werden, weil der nicht von der Füllung des Schriftfeldes 
            ' überdeckt werden soll 
            If awinSettings.mppShowMsName Then


                ' im Einzeile Modus fehlt der Kontext, deswegen die etwas aufwändigere Beschriftung  
                msBeschriftung = hproj.hierarchy.getBestNameOfID(milestoneID, Not awinSettings.mppUseOriginalNames, _
                                                                 awinSettings.mppUseAbbreviation, _
                                                                 swimlaneID)

                rds.MsDescVorlagenShape.Copy()
                copiedShape = rds.pptSlide.Shapes.Paste()
                With copiedShape(1)

                    .TextFrame2.TextRange.Text = msBeschriftung
                    .Top = CSng(yPosition + rds.YMilestoneText)
                    .Left = CSng(x1) - .Width / 2
                    .Name = .Name & .Id

                    shapeNames.Add(.Name, .Name)
                End With


            End If

            ' jetzt muss ggf das Datum angebracht werden 
            Dim msDateText As String = ""
            If awinSettings.mppShowMsDate Then

                msDateText = msDate.Day.ToString & "." & msDate.Month.ToString

                rds.MsDateVorlagenShape.Copy()
                copiedShape = rds.pptSlide.Shapes.Paste()
                With copiedShape(1)

                    .TextFrame2.TextRange.Text = msDateText
                    .Top = CSng(yPosition + rds.YMilestoneDate)
                    .Left = CSng(x1) - .Width / 2
                    .Name = .Name & .Id

                    shapeNames.Add(.Name, .Name)
                End With

            End If


            ' Erst jetzt wird der Meilenstein gezeichnet 
            milestoneTypShape.Copy()
            copiedShape = rds.pptSlide.Shapes.Paste()

            With copiedShape.Item(1)
                .Height = sizeFaktor * .Height
                .Width = sizeFaktor * .Width
                .Top = CSng(yPosition + rds.YMilestone)
                .Left = CSng(x1) - .Width / 2
                .Name = .Name & .Id
                .Title = milestoneName
                .AlternativeText = msDate.ToShortDateString

                If awinSettings.mppShowAmpel Then
                    .Glow.Color.RGB = CInt(cMilestone.getBewertung(1).color)
                    If .Glow.Radius = 0 Then
                        .Glow.Radius = 5
                    End If
                End If

                Dim msKwText As String = ""
                If awinSettings.mppKwInMilestone Then

                    msKwText = calcKW(msDate).ToString("0#")
                    If CInt(sizeFaktor * .TextFrame2.TextRange.Font.Size) >= 3 Then
                        .TextFrame2.TextRange.Font.Size = CInt(sizeFaktor * .TextFrame2.TextRange.Font.Size)
                        .TextFrame2.TextRange.Text = msKwText
                    End If

                End If

                shapeNames.Add(.Name, .Name)
            End With


        End If


    End Sub


    ''' <summary>
    ''' bestimmt anhand der Shapes bzw. Einstellungen die benötigte Zeilenhöhe
    ''' bei "extended Mode" sonst ist es auch gleichzeitig die Projekt-Höhe
    ''' </summary>
    ''' <param name="phaseVorlagenShape"></param>
    ''' <param name="milestoneVorlagenShape"></param>
    ''' <param name="MsDescVorlagenShape"></param>
    ''' <param name="MsDateVorlagenShape"></param>
    ''' <param name="PhDescVorlagenShape"></param>
    ''' <param name="PhDateVorlagenShape" ></param>
    ''' <param name="projectNameVorlagenShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function bestimmeMppZeilenHoehe(ByVal pptSlide As pptNS.Slide, _
                                             ByVal phaseVorlagenShape As pptNS.Shape, ByVal milestoneVorlagenShape As pptNS.Shape, _
                                                 ByVal anzPhasen As Integer, ByVal anzMilestones As Integer, _
                                                 ByVal MsDescVorlagenShape As pptNS.Shape, ByVal MsDateVorlagenShape As pptNS.Shape, _
                                                 ByVal PhDescVorlagenShape As pptNS.Shape, ByVal PhDateVorlagenShape As pptNS.Shape, _
                                                 ByVal projectNameVorlagenShape As pptNS.Shape, _
                                                 ByVal durationArrow As pptNS.Shape, ByVal durationText As pptNS.Shape) As Double

        'Dim versatzFaktor As Double = 0.87

        Dim listOfShapes As New Collection

        Dim minTop As Double = 1.79769313486231E+308        ' Maximale Double - Wert
        Dim maxBottom As Double = 0.0


        listOfShapes.Add(projectNameVorlagenShape.Name)
        minTop = System.Math.Min(minTop, projectNameVorlagenShape.Top)
        maxBottom = System.Math.Max(maxBottom, projectNameVorlagenShape.Top + projectNameVorlagenShape.Height)

        If anzPhasen > 0 Then
            listOfShapes.Add(phaseVorlagenShape.Name)
            minTop = System.Math.Min(minTop, phaseVorlagenShape.Top)
            maxBottom = System.Math.Max(maxBottom, phaseVorlagenShape.Top + phaseVorlagenShape.Height)

            ' ur:27.03.2015: die Höhe von phaseDelimiterShape wird beim Zeichnen auf 1.3 * Phaseshapehöhe gesetzt.
            ' Also ist Zeilenhöhe nicht davon abhängig

            'If Not IsNothing(phaseDelimiterShape) Then
            '    listOfShapes.Add(phaseDelimiterShape.Name)
            'End If

            If awinSettings.mppShowPhDate Then
                listOfShapes.Add(PhDateVorlagenShape.Name)
                minTop = System.Math.Min(minTop, PhDateVorlagenShape.Top)
                maxBottom = System.Math.Max(maxBottom, PhDateVorlagenShape.Top + PhDateVorlagenShape.Height)
            End If

            If awinSettings.mppShowPhName Then
                listOfShapes.Add(PhDescVorlagenShape.Name)
                minTop = System.Math.Min(minTop, PhDescVorlagenShape.Top)
                maxBottom = System.Math.Max(maxBottom, PhDescVorlagenShape.Top + PhDescVorlagenShape.Height)
            End If
        End If

        If anzMilestones > 0 Then
            listOfShapes.Add(milestoneVorlagenShape.Name)
            minTop = System.Math.Min(minTop, milestoneVorlagenShape.Top)
            maxBottom = System.Math.Max(maxBottom, milestoneVorlagenShape.Top + milestoneVorlagenShape.Height)

            If awinSettings.mppShowMsDate Then
                listOfShapes.Add(MsDateVorlagenShape.Name)
                minTop = System.Math.Min(minTop, MsDateVorlagenShape.Top)
                maxBottom = System.Math.Max(maxBottom, MsDateVorlagenShape.Top + MsDateVorlagenShape.Height)

            End If

            If awinSettings.mppShowMsName Then
                listOfShapes.Add(MsDescVorlagenShape.Name)
                minTop = System.Math.Min(minTop, MsDescVorlagenShape.Top)
                maxBottom = System.Math.Max(maxBottom, MsDescVorlagenShape.Top + MsDescVorlagenShape.Height)

            End If
        End If

        Dim projekthoehe As Double = maxBottom - minTop
        ' Änderung tk: in der Vorlage sind die Margins top und Bottom jeweils auf 0 gesetzt
        'projekthoehe = ((maxBottom - minTop) * 13 / 15)


        ' jetzt werden noch die Höhe des Pfeiles und der Beschriftung berücksichtigt 

        Dim addOn As Double = 0.0

        If Not IsNothing(durationArrow) And Not IsNothing(durationText) Then

            addOn = System.Math.Max(durationArrow.Height, durationText.Height)

        End If

        'bestimmeMppZeilenHoehe = projekthoehe + addOn * 13 / 15
        bestimmeMppZeilenHoehe = projekthoehe + addOn * 13 / 15

    End Function


    ''' <summary>
    ''' zeichnet die Legende zu einer Multiprojekt-Sicht 
    ''' </summary>
    ''' <param name="pptslide">ppt slide, in die gezeichnet werden soll</param>
    ''' <param name="selectedPhases">Phasen, die gezeichnet werden sollen</param>
    ''' <param name="selectedMilestones">Meilensteine, die gezeichnet werden sollen</param>
    ''' <param name="selectedRoles">Rollen, zu denen die Info ausgegeben werden soll</param>
    ''' <param name="selectedCosts">Kostenarten, zu denen die Info ausgegeben werden soll </param>
    ''' <param name="legendAreaTop">Oberer Rand der Legenden Zeichenfläche </param>
    ''' <param name="legendAreaLeft">linker Rand der Legenden Zeichenfläche</param>
    ''' <param name="legendAreaRight">rechter Rand der Legenden Zeichenfläche</param>
    ''' <param name="legendAreaBottom">unterer Rand der Legenden Zeichenfläche</param>
    ''' <param name="legendTextVorlagenShape">Legenden Schriftvorlage</param>
    ''' <param name="legendPhaseVorlagenShape">Legenden Phasen Vorlage (bestimmt die Höhe und Breite)</param>
    ''' <param name="legendMilestoneVorlagenShape">Legenden Meilenstein Vorlage; betimmt Höhe und Höhen / Breitenverhältnis </param>
    ''' <param name="projectVorlagenShape">Vorlagen Shape zur Darstellung des Projekts</param>
    ''' <param name="ampelVorlagenShape">Vorlagen Shape zur Darstellung der Projekt-Ampel</param>
    ''' <remarks></remarks>
    Sub zeichnePPTlegende(ByRef pptslide As pptNS.Slide, _
                                ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                ByVal legendAreaTop As Double, ByVal legendAreaLeft As Double, legendAreaRight As Double, legendAreaBottom As Double, _
                                ByVal legendLineShape As pptNS.Shape, ByVal legendStartShape As pptNS.Shape, _
                                ByVal legendTextVorlagenShape As pptNS.Shape, ByVal legendPhaseVorlagenShape As pptNS.Shape, ByVal legendMilestoneVorlagenShape As pptNS.Shape, _
                                ByVal projectVorlagenShape As pptNS.Shape, ByVal ampelVorlagenShape As pptNS.Shape, ByVal buColorVorlagenShape As pptNS.Shape)

        Dim maxZeilen As Integer
        Dim mindestNettoHoehe As Double = System.Math.Max(legendMilestoneVorlagenShape.Height, legendPhaseVorlagenShape.Height)
        Dim zeilenHoehe As Double
        Dim xCursor As Double, yCursor As Double
        Dim copiedShape As pptNS.ShapeRange
        Dim buName As String
        Dim buColor As Long
        Dim maxDelta As Double = 0.0

        Dim tmpDbl(3) As Double
        tmpDbl(0) = legendTextVorlagenShape.Height
        tmpDbl(1) = legendMilestoneVorlagenShape.Height
        tmpDbl(2) = legendPhaseVorlagenShape.Height

        If Not IsNothing(buColorVorlagenShape) Then
            tmpDbl(3) = buColorVorlagenShape.Height
        Else
            tmpDbl(3) = 0.0
        End If

        zeilenHoehe = tmpDbl.Max

        If zeilenHoehe = mindestNettoHoehe Then
            zeilenHoehe = zeilenHoehe * 1.1
        End If

        maxZeilen = (legendAreaBottom - legendAreaTop) / zeilenHoehe

        xCursor = legendAreaLeft
        yCursor = legendAreaTop

        ' jetzt das LegendlineShape eintragen 
        legendLineShape.Copy()
        copiedShape = pptslide.Shapes.Paste()
        With copiedShape.Item(1)
            .Top = legendLineShape.Top
            .Left = legendLineShape.Left
        End With

        ' jetzt das LegendlineShape eintragen 
        legendStartShape.Copy()
        copiedShape = pptslide.Shapes.Paste()
        With copiedShape.Item(1)
            .Top = legendStartShape.Top
            .Left = legendStartShape.Left
        End With


        If Not IsNothing(buColorVorlagenShape) Then

            For i = 1 To businessUnitDefinitions.Count
                buName = businessUnitDefinitions.ElementAt(i - 1).Value.name
                buColor = businessUnitDefinitions.ElementAt(i - 1).Value.color

                ' jetzt das Shape eintragen 
                buColorVorlagenShape.Copy()
                copiedShape = pptslide.Shapes.Paste()
                With copiedShape(1)
                    .Top = yCursor
                    .Height = zeilenHoehe
                    .Left = xCursor
                    .Fill.ForeColor.RGB = buColor
                End With

                ' jetzt den Business Unit Name eintragen 
                ' Text
                legendTextVorlagenShape.Copy()
                copiedShape = pptslide.Shapes.Paste()
                With copiedShape(1)
                    .TextFrame2.TextRange.Text = buName
                    .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                    .Left = xCursor + buColorVorlagenShape.Width + 3
                    If maxDelta < buColorVorlagenShape.Width + .Width + 3 Then
                        maxDelta = buColorVorlagenShape.Width + .Width + 3
                    End If
                End With

                yCursor = yCursor + zeilenHoehe
                If yCursor + zeilenHoehe > legendAreaBottom Then
                    yCursor = legendAreaTop
                    xCursor = xCursor + maxDelta
                    maxDelta = 0.0
                End If

            Next
            buName = "ohne Name"
            buColor = awinSettings.AmpelNichtBewertet

            ' jetzt das Shape eintragen 
            buColorVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)
                .Top = yCursor
                .Height = zeilenHoehe
                .Left = xCursor
                .Fill.ForeColor.RGB = buColor
            End With

            ' jetzt den Business Unit Name eintragen 
            ' Text
            legendTextVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)
                .TextFrame2.TextRange.Text = buName
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                .Left = xCursor + buColorVorlagenShape.Width + 3
                If maxDelta < buColorVorlagenShape.Width + .Width + 3 Then
                    maxDelta = buColorVorlagenShape.Width + .Width + 3
                End If
            End With

            xCursor = xCursor + maxDelta
            maxDelta = 0.0
            yCursor = legendAreaTop

        End If

        ' jetzt ggf die Legende für das Projekt zeichnen 
        If awinSettings.mppShowProjectLine Then

            ' Grafik
            projectVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .Height = System.Math.Min(projectVorlagenShape.Height, legendPhaseVorlagenShape.Height)
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                .Left = xCursor
                .Width = legendPhaseVorlagenShape.Width

            End With

            ' Text
            legendTextVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)
                .TextFrame2.TextRange.Text = "Projekt, ggf. mit" & vbLf & "Anfangs- und Ende-Markierung"
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                .Left = xCursor + legendPhaseVorlagenShape.Width + 3
                xCursor = .Left + copiedShape(1).Width + 15
            End With



        End If

        yCursor = legendAreaTop

        ' jetzt ggf die Phasen-Legende zeichnen 
        Dim phaseShape As xlNS.Shape
        Dim phaseName As String = ""
        Dim maxBreite As Double = 0.0

        Dim breadcrumb As String = ""
        Dim selPhaseName As String = ""


        ' tk: Änderung 21.6.15
        ' jetzt muss bestimmt werden wieviele eindeutige Phasen-Klassen-Definitionen denn überhaupt da sind , mehrfache Vorkommnisse müssen 
        ' ja nicht mehrfach in der Legende immer mit der gleichen Abkürzung / Farbe gezeigt werden ..

        Dim uniqueElemClasses As New Collection
        For i = 1 To selectedPhases.Count
            Call splitHryFullnameTo2(CStr(selectedPhases(i)), phaseName, breadcrumb)

            If uniqueElemClasses.Contains(phaseName) Then
                ' nichts tun, ist schon enthalten 
            Else
                uniqueElemClasses.Add(phaseName, phaseName)
            End If
        Next

        For i = 1 To uniqueElemClasses.Count

            phaseName = CStr(uniqueElemClasses(i))
            Dim phShortname As String = PhaseDefinitions.getAbbrev(phaseName)

            ' Änderung tk 26.11.15
            If PhaseDefinitions.Contains(phaseName) Then
                phaseShape = PhaseDefinitions.getShape(phaseName)
            Else
                phaseShape = missingPhaseDefinitions.getShape(phaseName)
            End If

            ' Phasen-Shape 
            phaseShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .Height = legendPhaseVorlagenShape.Height
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                .Left = xCursor
                .Width = legendPhaseVorlagenShape.Width

            End With

            ' Phasen-Text
            legendTextVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .TextFrame2.TextRange.Text = phShortname & " (=" & phaseName & ")"
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                .Left = xCursor + legendPhaseVorlagenShape.Width + 3


                If maxBreite < legendPhaseVorlagenShape.Width + 3 + .Width Then
                    maxBreite = legendPhaseVorlagenShape.Width + 3 + .Width
                End If
            End With

            If i Mod maxZeilen = 0 And i < selectedPhases.Count Then
                xCursor = xCursor + maxBreite + 10
                If xCursor >= legendAreaRight Then
                    Throw New ArgumentException("Platz für die Legende reicht nicht aus. Evt.muss eine neue Vorlage definiert werden!")
                End If
                maxBreite = 0.0
                yCursor = legendAreaTop
            Else
                yCursor = yCursor + zeilenHoehe
            End If



        Next

        If uniqueElemClasses.Count > 0 Then
            xCursor = xCursor + maxBreite + 15
        End If
        yCursor = legendAreaTop

        ' jetzt ggf die Meilenstein-Legende zeichnen 
        Dim meilensteinShape As xlNS.Shape
        Dim msShortname As String
        maxBreite = 0.0

        Dim msName As String = ""
        Dim breadcrumbMS As String = ""

        uniqueElemClasses.Clear()
        For i = 1 To selectedMilestones.Count
            Call splitHryFullnameTo2(CStr(selectedMilestones.Item(i)), msName, breadcrumbMS)

            If uniqueElemClasses.Contains(msName) Then
                ' nichts tun, ist schon enthalten 
            Else
                uniqueElemClasses.Add(msName, msName)
            End If
        Next


        For i = 1 To uniqueElemClasses.Count

            msName = CStr(uniqueElemClasses.Item(i))

            msShortname = MilestoneDefinitions.getAbbrev(msName)

            ' Änderung tk 26.11.15
            If MilestoneDefinitions.Contains(msName) Then
                meilensteinShape = MilestoneDefinitions.getShape(msName)
            Else
                meilensteinShape = missingMilestoneDefinitions.getShape(msName)
            End If


            ' Meilenstein-Shape 
            meilensteinShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)
                .Left = xCursor
                .Height = legendMilestoneVorlagenShape.Height
                .Width = legendMilestoneVorlagenShape.Width / legendMilestoneVorlagenShape.Height * .Height
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
            End With

            ' Meilenstein-Text
            legendTextVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .TextFrame2.TextRange.Text = msShortname & " (=" & msName & ")"
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                .Left = xCursor + legendMilestoneVorlagenShape.Width + 3

                If maxBreite < legendMilestoneVorlagenShape.Width + 3 + .Width Then
                    maxBreite = legendMilestoneVorlagenShape.Width + 3 + .Width
                End If
            End With

            If i Mod maxZeilen = 0 And i < selectedMilestones.Count Then
                xCursor = xCursor + maxBreite + 10
                If xCursor >= legendAreaRight Then
                    Throw New ArgumentException("Platz für die Legende reicht nicht aus. Evt.muss eine neue Vorlage definiert werden!")
                End If
                yCursor = legendAreaTop
                maxBreite = 0.0
            Else
                yCursor = yCursor + zeilenHoehe
            End If



        Next

        If uniqueElemClasses.Count > 0 Then
            xCursor = xCursor + maxBreite + 15
            If xCursor >= legendAreaRight Then
                Throw New ArgumentException("Platz für die Legende reicht nicht aus. Evt.muss eine neue Vorlage definiert werden!")
            End If
        End If
        yCursor = legendAreaTop

        ' jetzt ggf die Ampel Legende zeichnen 


        If awinSettings.mppShowAmpel Then

            ' Ampel-Shape 

            For i = 1 To 4

                ampelVorlagenShape.Copy()
                copiedShape = pptslide.Shapes.Paste()

                With copiedShape(1)

                    .Height = legendMilestoneVorlagenShape.Height
                    .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                    .Width = legendMilestoneVorlagenShape.Height
                    .Left = xCursor + (i - 1) * (.Width + 4)

                    If i = 1 Then
                        .Fill.ForeColor.RGB = CInt(awinSettings.AmpelNichtBewertet)
                    ElseIf i = 2 Then
                        .Fill.ForeColor.RGB = CInt(awinSettings.AmpelGruen)
                    ElseIf i = 3 Then
                        .Fill.ForeColor.RGB = CInt(awinSettings.AmpelGelb)
                    Else
                        .Fill.ForeColor.RGB = CInt(awinSettings.AmpelRot)
                    End If

                End With
            Next


            ' Projekt-Ampel-Text
            legendTextVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .TextFrame2.TextRange.Text = "Projekt-Ampeln"
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                .Left = xCursor + 4 * (legendMilestoneVorlagenShape.Height + 4)

            End With

            yCursor = yCursor + zeilenHoehe


            For i = 1 To 4

                legendMilestoneVorlagenShape.Copy()
                copiedShape = pptslide.Shapes.Paste()

                With copiedShape(1)
                    .Height = legendMilestoneVorlagenShape.Height
                    .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                    .Width = legendMilestoneVorlagenShape.Height
                    .Glow.Radius = 3
                    .Left = xCursor + (i - 1) * (.Width + 4)

                    If i = 1 Then
                        .Glow.Color.RGB = CInt(awinSettings.AmpelNichtBewertet)
                    ElseIf i = 2 Then
                        .Glow.Color.RGB = CInt(awinSettings.AmpelGruen)
                    ElseIf i = 3 Then
                        .Glow.Color.RGB = CInt(awinSettings.AmpelGelb)
                    Else
                        .Glow.Color.RGB = CInt(awinSettings.AmpelRot)
                    End If

                End With
            Next


            ' Projekt-Ampel-Text
            legendTextVorlagenShape.Copy()
            copiedShape = pptslide.Shapes.Paste()
            With copiedShape(1)

                .TextFrame2.TextRange.Text = "Meilenstein-Ampeln"
                .Top = CSng(yCursor + 0.5 * (zeilenHoehe - .Height))
                .Left = xCursor + 4 * (legendMilestoneVorlagenShape.Height + 4)

            End With


        End If




    End Sub



    ''' <summary>
    ''' berechnet die x1 und x2-Koordinaten , also den Start und das Ende des Elements in x-Koordinaten
    ''' im Gegensatz zu ...OLD werden hier die Koordinaten in Abhängigkeit von dem Abstand Tagen vom linken Rand gemessen. 
    ''' bei der bisherigen ...OLD wurde gemessen, wieviel volle Monate Abstand waren plus wieviele Rest-Tage 
    ''' das wird in der neuen Art als Methode in clsPPTShapes gemacht 
    ''' </summary>
    ''' <param name="pptStartOfCalendar">linker Rand es Kalenders</param>
    ''' <param name="pptEndOfCalendar">rechter Rand des Kalenders</param>
    ''' <param name="startdate">Startdatum des Elements, wenn es vor dem linken Rand des Kalenders liegt, wird es auf den KAlenderstart gesetzt </param>
    ''' <param name="enddate">Endedatum des Elements, wenn es nach dem Ende-Datum des Kalenders liegt, wird es auf das Ende -Datum gesetzt</param>
    ''' <param name="linkerRand">linker Rand in x-Koordinaten</param>
    ''' <param name="breite">Breite zwischen Kalender-Start und Ende in x-Koordinaten</param>
    ''' <param name="x1Pos">Rückgabe Wert Start</param>
    ''' <param name="x2Pos">Rückgabe Wert Ende</param>
    ''' <remarks></remarks>
    Private Sub calculatePPTx1x2New(ByVal pptStartOfCalendar As Date, ByVal pptEndOfCalendar As Date, _
                                     ByVal startdate As Date, ByVal enddate As Date, _
                                     ByVal linkerRand As Double, ByVal breite As Double, _
                                     ByRef x1Pos As Double, ByRef x2Pos As Double)



        Dim anzahlTageImKalender As Integer = DateDiff(DateInterval.Day, pptStartOfCalendar, pptEndOfCalendar)
        Dim tagesbreite As Double = breite / anzahlTageImKalender

        Dim offset1 As Integer = DateDiff(DateInterval.Day, pptStartOfCalendar, startdate)
        If offset1 <= 0 Then
            x1Pos = linkerRand
        Else
            x1Pos = linkerRand + offset1 * tagesbreite
        End If


        Dim offset2 As Integer = DateDiff(DateInterval.Day, pptStartOfCalendar, enddate)
        If offset2 >= anzahlTageImKalender Then
            x2Pos = linkerRand + breite
        Else
            x2Pos = linkerRand + offset2 * tagesbreite
        End If

    End Sub

    ''' <summary>
    ''' Änderung tk: das wird in zeichnepptProjects verwendet 
    ''' sollte im 1. HJ 2016 ersetzt werden durch die Art und Weise, wie bei zeichneSwimlanes gearbeitet wird ...  
    ''' </summary>
    ''' <param name="pptStartOfCalendar"></param>
    ''' <param name="pptEndOfCalendar"></param>
    ''' <param name="startdate"></param>
    ''' <param name="enddate"></param>
    ''' <param name="linkerRand"></param>
    ''' <param name="breite"></param>
    ''' <param name="x1Pos"></param>
    ''' <param name="x2Pos"></param>
    ''' <remarks></remarks>
    Private Sub calculatePPTx1x2(ByVal pptStartOfCalendar As Date, ByVal pptEndOfCalendar As Date, _
                                         ByVal startdate As Date, ByVal enddate As Date, _
                                         ByVal linkerRand As Double, ByVal breite As Double, _
                                         ByRef x1Pos As Double, ByRef x2Pos As Double)

        Dim tageProMonat(12) As Integer
        tageProMonat(0) = 30 ' dummy
        tageProMonat(1) = 31
        tageProMonat(2) = 28
        tageProMonat(3) = 31
        tageProMonat(4) = 30
        tageProMonat(5) = 31
        tageProMonat(6) = 30
        tageProMonat(7) = 31
        tageProMonat(8) = 31
        tageProMonat(9) = 30
        tageProMonat(10) = 31
        tageProMonat(11) = 30
        tageProMonat(12) = 31



        Dim anzQMs As Integer

        Dim yWidth As Double, mWidth As Double
        Call calculateYMAeinheiten(pptStartOfCalendar, pptEndOfCalendar, breite, yWidth, mWidth, anzQMs)


        Dim offset1 As Integer = DateDiff(DateInterval.Month, pptStartOfCalendar, startdate)
        If offset1 < 0 Then
            x1Pos = linkerRand
        Else
            x1Pos = linkerRand + _
                    (offset1 + startdate.Day / tageProMonat(startdate.Month)) * mWidth
        End If


        Dim offset2 As Integer = DateDiff(DateInterval.Month, pptStartOfCalendar, enddate)
        If offset2 >= anzQMs Then
            x2Pos = linkerRand + breite
        Else
            x2Pos = linkerRand + _
                    (offset2 + enddate.Day / tageProMonat(enddate.Month)) * mWidth
        End If

    End Sub

    ''' <summary>
    ''' leifert eine Sammlung relativer Größen für jedes Shape mit Enlarge13 im Titel 
    ''' </summary>
    ''' <param name="pptSlide"></param>
    ''' <param name="slideHeight"></param>
    ''' <param name="slideWidth"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function saveRelSizesOfElements(ByVal pptSlide As pptNS.Slide, ByVal slideHeight As Double, ByVal slideWidth As Double) As SortedList(Of String, Double())

        Dim tmpList As New SortedList(Of String, Double())

        For Each shp As pptNS.Shape In pptSlide.Shapes
            If shp.AlternativeText.Trim = "Enlarge13" Or shp.AlternativeText.Trim = "Enlarge10und13" Then
                Dim relSizes(4) As Double
                ' top 

                relSizes(0) = shp.Top / slideHeight
                relSizes(1) = shp.Left / slideWidth
                relSizes(2) = shp.Height / slideHeight
                relSizes(3) = shp.Width / slideWidth

                If shp.Type = MsoShapeType.msoLine Then
                    If relSizes(2) = 0 Then
                        relSizes(4) = shp.Line.Weight / slideHeight
                    ElseIf relSizes(3) = 0 Then
                        relSizes(4) = shp.Line.Weight / slideWidth
                    End If

                Else
                    relSizes(4) = 0
                End If

                tmpList.Add(shp.Name, relSizes)

            End If
        Next

        saveRelSizesOfElements = tmpList

    End Function

    ''' <summary>
    ''' bringt bei Powerpoint 2013 die Shapes wieder auf ihre alte relative Größe zurück 
    ''' </summary>
    ''' <param name="sizes"></param>
    ''' <param name="slideHeight"></param>
    ''' <param name="slideWidth"></param>
    ''' <param name="pptslide"></param>
    ''' <remarks></remarks>
    Private Sub restoreRelSizesDuePPT2013(ByVal sizes As SortedList(Of String, Double()), ByVal slideHeight As Double, ByVal slideWidth As Double, ByRef pptslide As pptNS.Slide)
        Dim relSizes(4) As Double
        Dim shp As pptNS.Shape
        Dim allShapes As pptNS.Shapes = pptslide.Shapes

        For Each kvp As KeyValuePair(Of String, Double()) In sizes

            shp = allShapes.Item(kvp.Key)
            relSizes = kvp.Value

            With shp
                'Call MsgBox("jetzt " & shp.Name & ", " & shp.Title & ", " & shp.AlternativeText & vbLf & _
                '            "Top: " & relSizes(0).ToString & " * " & slideHeight.ToString & vbLf & _
                '            "Left: " & relSizes(1).ToString & " * " & slideWidth.ToString & vbLf & _
                '            "Height: " & relSizes(2).ToString & " * " & slideHeight.ToString & vbLf & _
                '            "Width: " & relSizes(3).ToString & " * " & slideWidth.ToString & vbLf & _
                '            "Weight: " & relSizes(4).ToString & " * .. whatever ")

                .Top = relSizes(0) * slideHeight
                .Left = relSizes(1) * slideWidth
                .Height = relSizes(2) * slideHeight
                .Width = relSizes(3) * slideWidth
                If .Type = MsoShapeType.msoLine Then
                    If .Height = 0 Then
                        .Line.Weight = relSizes(4) * slideHeight
                    ElseIf .Width = 0 Then
                        .Line.Weight = relSizes(4) * slideWidth
                    End If
                End If
                'Call MsgBox("jetzt " & shp.Name & ", " & shp.AlternativeText & " -> done")
            End With


        Next


    End Sub

    ''' <summary>
    ''' gibt einen Array of Single zurück, der all die Größen der übergebenen Shapes enthält (Schriftgröße bzw Liniendicke oder Weight)
    ''' </summary>
    ''' <param name="projectNameVorlagenShape"></param>
    ''' <param name="MsDescVorlagenShape"></param>
    ''' <param name="MsDateVorlagenShape"></param>
    ''' <param name="PhDescVorlagenShape"></param>
    ''' <param name="PhDateVorlagenShape"></param>
    ''' <param name="phaseVorlagenShape"></param>
    ''' <param name="milestoneVorlagenShape"></param>
    ''' <param name="projectVorlagenShape"></param>
    ''' <param name="ampelVorlagenShape"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function saveSizesOfElements(ByVal projectNameVorlagenShape As pptNS.Shape, _
                                         ByVal MsDescVorlagenShape As pptNS.Shape, ByVal MsDateVorlagenShape As pptNS.Shape, _
                                         ByVal PhDescVorlagenShape As pptNS.Shape, ByVal PhDateVorlagenShape As pptNS.Shape, _
                                         ByVal phaseVorlagenShape As pptNS.Shape, ByVal milestoneVorlagenShape As pptNS.Shape, _
                                         ByVal projectVorlagenShape As pptNS.Shape, ByVal ampelVorlagenShape As pptNS.Shape, _
                                         Optional ByVal segmentVorlagenShape As pptNS.Shape = Nothing) As Single()

        Dim sizes(9) As Single

        sizes(0) = projectNameVorlagenShape.TextFrame2.TextRange.Font.Size
        sizes(1) = MsDescVorlagenShape.TextFrame2.TextRange.Font.Size
        sizes(2) = MsDateVorlagenShape.TextFrame2.TextRange.Font.Size
        sizes(3) = PhDescVorlagenShape.TextFrame2.TextRange.Font.Size
        sizes(4) = PhDateVorlagenShape.TextFrame2.TextRange.Font.Size
        sizes(5) = phaseVorlagenShape.Height
        sizes(6) = milestoneVorlagenShape.Height
        sizes(7) = projectVorlagenShape.Line.Weight

        If IsNothing(ampelVorlagenShape) Then
            sizes(8) = 0.0
        Else
            sizes(8) = ampelVorlagenShape.Height
        End If

        If IsNothing(segmentVorlagenShape) Then
            sizes(9) = 0.0
        Else
            sizes(9) = segmentVorlagenShape.TextFrame2.TextRange.Font.Size
        End If

        saveSizesOfElements = sizes

    End Function

    ''' <summary>
    ''' vergößert die Textshapes der Texte, die ihre relative Größe bei einer Format-Änderung behalten sollen
    ''' alle Text Shapes, deren .Title  gleich Enlarge ist, werden angepasst 
    ''' </summary>
    ''' <param name="enlargeFaktor"></param>
    ''' <param name="pptSlide"></param>
    ''' <remarks></remarks>
    Private Sub enlargeTxtShapes(ByVal enlargeFaktor As Double, ByRef pptSlide As pptNS.Slide)

        Dim allShapes As pptNS.Shapes = pptSlide.Shapes
        'Dim korrektiv As Double = System.Math.Pow(0.9, enlargeFaktor)

        For Each tmpShape As pptNS.Shape In allShapes

            Try
                If (tmpShape.AlternativeText.Trim = "Enlarge" Or tmpShape.AlternativeText.Trim = "Enlarge10und13") And tmpShape.HasTextFrame Then
                    With tmpShape.TextFrame2.TextRange.Font
                        .Size = .Size * enlargeFaktor
                    End With
                End If
            Catch ex As Exception

            End Try


        Next

    End Sub



    ''' <summary>
    ''' stellt die Größen der übergebenen Shapes wieder her 
    ''' </summary>
    ''' <param name="sizes"></param>
    ''' <param name="projectNameVorlagenShape"></param>
    ''' <param name="MsDescVorlagenShape"></param>
    ''' <param name="MsDateVorlagenShape"></param>
    ''' <param name="PhDescVorlagenShape"></param>
    ''' <param name="PhDateVorlagenShape"></param>
    ''' <param name="phaseVorlagenShape"></param>
    ''' <param name="milestoneVorlagenShape"></param>
    ''' <param name="projectVorlagenShape"></param>
    ''' <param name="ampelVorlagenShape"></param>
    ''' <remarks></remarks>
    Private Sub restoreSizesOfElements(ByVal sizes() As Single, _
                                           ByRef projectNameVorlagenShape As pptNS.Shape, _
                                           ByRef MsDescVorlagenShape As pptNS.Shape, ByRef MsDateVorlagenShape As pptNS.Shape, _
                                           ByRef PhDescVorlagenShape As pptNS.Shape, ByRef PhDateVorlagenShape As pptNS.Shape, _
                                           ByRef phaseVorlagenShape As pptNS.Shape, ByRef milestoneVorlagenShape As pptNS.Shape, _
                                           ByRef projectVorlagenShape As pptNS.Shape, ByRef ampelVorlagenShape As pptNS.Shape, _
                                           Optional ByRef segmentVorlagenShape As pptNS.Shape = Nothing)


        projectNameVorlagenShape.TextFrame2.TextRange.Font.Size = sizes(0)
        MsDescVorlagenShape.TextFrame2.TextRange.Font.Size = sizes(1)
        MsDateVorlagenShape.TextFrame2.TextRange.Font.Size = sizes(2)
        PhDescVorlagenShape.TextFrame2.TextRange.Font.Size = sizes(3)
        PhDateVorlagenShape.TextFrame2.TextRange.Font.Size = sizes(4)
        phaseVorlagenShape.Height = sizes(5)
        milestoneVorlagenShape.Height = sizes(6)
        projectVorlagenShape.Line.Weight = sizes(7)

        If Not IsNothing(ampelVorlagenShape) Then
            ampelVorlagenShape.Height = sizes(8)
        End If

        If (Not IsNothing(segmentVorlagenShape)) And (sizes.Length = 10) Then
            segmentVorlagenShape.TextFrame2.TextRange.Font.Size = sizes(9)
        End If


    End Sub

    ' Änderung tk - rausgenommen , ersetzt durch MEthode in Klasse clsPPTShapes
    '
    ' ermittelt die Koordinaten für Kalender, linker Rand Projektbeschriftung, Projekt-Fläche, Legenden-Fläche
    ''Private Sub bestimmeZeichenKoordinaten(ByVal multiprojektContainerShape As pptNS.Shape, _
    ''                                           ByVal calendarLineShape As pptNS.Shape, ByVal calendarHeightShape As pptNS.Shape, _
    ''                                           ByVal legendLineShape As pptNS.Shape, _
    ''                                           ByRef containerLeft As Single, ByRef containerRight As Single, ByRef containerTop As Single, ByRef containerBottom As Single, _
    ''                                           ByRef calendarLeft As Single, ByRef calendarRight As Single, ByRef calendarTop As Single, ByRef calendarBottom As Single, _
    ''                                           ByRef drawingAreaLeft As Single, ByRef drawingAreaRight As Single, ByRef drawingAreaTop As Single, ByRef drawingAreaBottom As Single, _
    ''                                           ByRef projectListLeft As Single, _
    ''                                           ByRef legendAreaLeft As Single, ByRef legendAreaRight As Single, ByRef legendAreaTop As Single, ByRef legendAreaBottom As Single)

    ''    ' bestimme Container Area ud linker Rand der Projektliste
    ''    With multiprojektContainerShape
    ''        containerLeft = .Left
    ''        containerRight = .Left + .Width
    ''        containerTop = .Top
    ''        containerBottom = .Top + .Height
    ''        projectListLeft = .Left + 10
    ''    End With

    ''    ' bestimme KalenderArea
    ''    calendarLeft = calendarLineShape.Left
    ''    calendarRight = calendarLineShape.Left + calendarLineShape.Width
    ''    calendarTop = containerTop + 5
    ''    calendarBottom = calendarTop + calendarHeightShape.Height

    ''    ' bestimme Drawing Area
    ''    drawingAreaLeft = calendarLeft
    ''    drawingAreaRight = calendarRight
    ''    drawingAreaTop = calendarBottom + 15


    ''    If awinSettings.mppShowLegend Then
    ''        drawingAreaBottom = legendLineShape.Top - 5
    ''    Else
    ''        drawingAreaBottom = containerBottom - 10
    ''    End If



    ''    ' bestimme Legend Drawing Area 
    ''    If awinSettings.mppShowLegend Then
    ''        legendAreaTop = legendLineShape.Top + (containerBottom - legendLineShape.Top) * 0.05
    ''        legendAreaBottom = containerBottom - (containerBottom - legendLineShape.Top) * 0.1
    ''    Else
    ''        legendLineShape.Top = containerBottom - 5
    ''        legendAreaTop = containerBottom - 5
    ''        legendAreaBottom = containerBottom
    ''    End If


    ''    legendAreaLeft = drawingAreaLeft
    ''    legendAreaRight = System.Math.Min(legendLineShape.Left + legendLineShape.Width, containerRight - 5)



    ''End Sub

    ''' <summary>
    ''' berechnet die "Breite" für ein Jahr, für einen Monat, sowie die Anzahl Monate m Kalender 
    ''' </summary>
    ''' <param name="startOfPPTCalendar"></param>
    ''' <param name="endOfPPTCalendar"></param>
    ''' <param name="breite"></param>
    ''' <param name="yWidth"></param>
    ''' <param name="mWidth"></param>
    ''' <remarks></remarks>
    Private Sub calculateYMAeinheiten(ByVal startOfPPTCalendar As Date, ByVal endOfPPTCalendar As Date, _
                                          ByVal breite As Double, _
                                          ByRef yWidth As Double, ByRef mWidth As Double, ByRef anzahlM As Integer)

        Dim anzQMs As Integer = DateDiff(DateInterval.Month, startOfPPTCalendar, endOfPPTCalendar) + 1
        Dim anzahlTage As Integer = DateDiff(DateInterval.Day, startOfPPTCalendar, endOfPPTCalendar) + 1

        yWidth = 12 * breite / anzQMs
        mWidth = breite / anzQMs
        anzahlM = anzQMs

    End Sub

    ''' <summary>
    ''' gibt eine Collection zurück; der Typ der collection kann sein: Phase, Meilenstein, Rolle, Kostenart 
    ''' der Qualifier ist entweder aufgebaut mit expliziten Bezeichnern, getrennt durch #
    ''' oder durch die Variablen-Nummer, getrennt durch %
    ''' </summary>
    ''' <param name="type"></param>
    ''' <param name="qualifier"></param>
    ''' <remarks></remarks>
    Private Function buildNameCollection(ByVal type As Integer, ByVal qualifier As String, _
                                        ByVal selectedItems As Collection) As Collection

        Dim qstr(30) As String
        Dim tmpCollection As New Collection
        Dim tmpName As String = " "
        Dim explicit As Boolean = True
        Dim trennzeichen As Char = "#"

        If qualifier.Contains("#") Then
            explicit = True
            trennzeichen = "#"
        ElseIf qualifier.Contains("%") Then
            explicit = False
            trennzeichen = "%"
        ElseIf IsNumeric(qualifier) Then
            explicit = False
            trennzeichen = "%"
        Else
            explicit = True
            trennzeichen = "#"
        End If

        qstr = qualifier.Trim.Split(New Char() {CChar(trennzeichen)}, 30)

        ' Aufbau der Collection 
        For i = 0 To qstr.Length - 1

            tmpName = ""
            Try
                If Not explicit Then
                    Dim ix As Integer

                    If qstr(i).Length > 0 Then
                        If qstr(i).Trim = "Alle" Then
                            tmpCollection.Clear()
                            For ii As Integer = 1 To selectedItems.Count
                                tmpCollection.Add(selectedItems.Item(ii), selectedItems.Item(ii))
                            Next
                            Exit For
                        Else
                            Try
                                If IsNumeric(qstr(i)) Then
                                    ix = CInt(qstr(i))
                                    If ix >= 1 And ix <= selectedItems.Count Then
                                        tmpName = CStr(selectedItems.Item(ix)).Trim
                                    Else
                                        tmpName = ""
                                    End If
                                End If


                            Catch ex As Exception

                            End Try
                        End If
                    End If

                Else
                    tmpName = qstr(i).Trim
                End If

                If tmpName.Length > 0 Then
                    Select Case type

                        Case PTpfdk.Phasen
                            Dim phName As String = ""
                            Dim tmpBC As String = ""
                            Call splitHryFullnameTo2(tmpName, phName, tmpBC)
                            If PhaseDefinitions.Contains(phName) Then
                                tmpCollection.Add(tmpName, tmpName)
                            End If

                        Case PTpfdk.Meilenstein
                            Dim msName As String = ""
                            Dim tmpBC As String = ""
                            Call splitHryFullnameTo2(tmpName, msName, tmpBC)
                            If MilestoneDefinitions.Contains(msName) Then
                                tmpCollection.Add(tmpName, tmpName)
                            End If

                        Case PTpfdk.Rollen
                            If RoleDefinitions.Contains(tmpName) Then
                                tmpCollection.Add(tmpName, tmpName)
                            End If

                        Case PTpfdk.Kosten
                            If CostDefinitions.Contains(tmpName) Then
                                tmpCollection.Add(tmpName, tmpName)
                            End If

                    End Select
                End If



            Catch ex As Exception
                Call MsgBox("Fehler: Phasen Name " & tmpName & " konnte nicht erkannt werden ...")
            End Try

        Next

        buildNameCollection = tmpCollection

    End Function

    ''' <summary>
    ''' zeichnet den Multiprojekt Sicht Container
    ''' </summary>
    ''' <param name="pptApp">ist die Powerpoint Applikation</param>
    ''' <param name="pptCurrentPresentation">ist die aktuelle PPT Präsentation; das Format wird hier noch bestimmt</param>
    ''' <param name="pptslide"></param>
    ''' <param name="objectsToDo"></param>
    ''' <param name="objectsDone"></param>
    ''' <param name="pptFirstTime"></param>
    ''' <param name="zeilenhoehe"></param>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="selectedRoles"></param>
    ''' <param name="selectedCosts"></param>
    ''' <param name="selectedBUs"></param>
    ''' <param name="selectedTyps"></param>
    ''' <param name="worker"></param>
    ''' <param name="e"></param>
    ''' <param name="isMultiprojektSicht">gibt an, ob es sich um eine Einzelprojekt/Varianten Sicht oder 
    ''' um eine Multiprojektsicht handelt </param>
    ''' <param name="projMitVariants">das Projekt, dessen Varianten alle dargestellt werden sollen; nur besetzt wenn isMultiprojektSicht = false</param>
    ''' <remarks></remarks>
    Private Sub zeichneMultiprojektSicht(ByRef pptApp As pptNS.Application, ByRef pptCurrentPresentation As pptNS.Presentation, ByRef pptslide As pptNS.Slide, _
                                             ByRef objectsToDo As Integer, ByRef objectsDone As Integer, ByRef pptFirstTime As Boolean, _
                                             ByRef zeilenhoehe As Double, ByRef legendFontSize As Double, _
                                             ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                             ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                             ByVal selectedBUs As Collection, ByVal selectedTyps As Collection, _
                                             ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs, _
                                             ByVal isMultiprojektSicht As Boolean, ByVal projMitVariants As clsProjekt, _
                                             ByVal kennzeichnung As String)

        ' ur:5.10.2015: ExtendedMode macht nur Sinn, wenn mindestens 1 Phase selektiert wurde. deshalb diese Code-Zeile
        awinSettings.mppExtendedMode = awinSettings.mppExtendedMode And (selectedPhases.Count > 0)


        ' Wichtig für Kalendar 
        Dim pptStartofCalendar As Date = Nothing, pptEndOfCalendar As Date = Nothing
        Dim errorShape As pptNS.ShapeRange = Nothing


        Dim dinFormatA(4, 1) As Double
        Dim querFormat As Boolean
        Dim curFormatSize(1) As Double


        dinFormatA(0, 0) = 3120.0
        dinFormatA(0, 1) = 2206.15

        dinFormatA(1, 0) = 2206.15
        dinFormatA(1, 1) = 1560.0

        dinFormatA(2, 0) = 1560.0
        dinFormatA(2, 1) = 1103.0

        dinFormatA(3, 0) = 1103.0
        dinFormatA(3, 1) = 780.0

        dinFormatA(4, 0) = 780.0
        dinFormatA(4, 1) = 540.0

        ' Ende Übernahme

        Dim format As Integer = 4
        'Dim tmpslideID As Integer



        Dim rds As New clsPPTShapes

        ' mit disem Befehl werden auch die ganzen Hilfsshapes in der Klasse gesetzt 
        rds.pptSlide = pptslide


        ' jetzt muss geprüft werden, ob überhaupt alle Angaben gemacht wurden ... 
        'If completeMppDefinition.Sum = completeMppDefinition.Length Then
        Dim missingShapes As String = rds.getMissingShpNames(kennzeichnung)
        If missingShapes.Length = 0 Then
            ' es fehlt nichts ... andernfalls stehen hier die Namen mit den Shapes, die fehlen ...

            If pptCurrentPresentation.PageSetup.SlideOrientation = MsoOrientation.msoOrientationHorizontal Then
                querFormat = True
            Else
                querFormat = False
            End If


            curFormatSize(0) = pptCurrentPresentation.PageSetup.SlideWidth
            curFormatSize(1) = pptCurrentPresentation.PageSetup.SlideHeight

            ' jetzt werden die DinA Formate gesetzt 
            ' Voraussetzung ist allerdings, dass es sich bei der Vorlage um DIN A4 handelt 
            Dim paperSizeRatio As Double
            If pptFirstTime Then


                If pptCurrentPresentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeA4Paper Then

                    If querFormat Then
                        paperSizeRatio = curFormatSize(0) / curFormatSize(1)
                        dinFormatA(4, 0) = curFormatSize(0)
                        dinFormatA(4, 1) = curFormatSize(1)
                    Else
                        paperSizeRatio = curFormatSize(1) / curFormatSize(0)
                        dinFormatA(4, 1) = curFormatSize(0)
                        dinFormatA(4, 0) = curFormatSize(1)
                    End If

                    dinFormatA(3, 0) = dinFormatA(4, 0) * paperSizeRatio
                    dinFormatA(3, 1) = dinFormatA(4, 1) * paperSizeRatio

                ElseIf pptCurrentPresentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeA3Paper Then
                    If querFormat Then
                        paperSizeRatio = curFormatSize(0) / curFormatSize(1)
                        dinFormatA(3, 0) = curFormatSize(0)
                        dinFormatA(3, 1) = curFormatSize(1)

                    Else
                        paperSizeRatio = curFormatSize(1) / curFormatSize(0)
                        dinFormatA(3, 1) = curFormatSize(0)
                        dinFormatA(3, 0) = curFormatSize(1)
                    End If

                    dinFormatA(4, 0) = dinFormatA(3, 0) / paperSizeRatio
                    dinFormatA(4, 1) = dinFormatA(3, 1) / paperSizeRatio

                Else
                    Call MsgBox("Vorlage ist weder ein A4 noch ein A3 Format ... bitte verwenden Sie eine A4 oder A3 Vorlage")
                    'Throw New ArgumentException("Vorlage ist weder ein A4 noch ein A3 Format ... bitte verwenden Sie eine A4 oder A3 Vorlage")
                End If


                For i = 2 To 0 Step -1
                    dinFormatA(i, 0) = dinFormatA(i + 1, 0) * paperSizeRatio
                    dinFormatA(i, 1) = dinFormatA(i + 1, 1) * paperSizeRatio
                Next
            Else
                ' pptFirstTime war False, d.h. das Format wurde bereits angepasst
            End If


            ' wenn Kalenderlinie oder Legendenlinie über Container rausragt: anpassen ! 
            Call rds.plausibilityAdjustments()

            Call rds.bestimmeZeichenKoordinaten()

            Dim projCollection As New SortedList(Of Double, String)
            Dim minDate As Date, maxDate As Date

            Dim considerAll As Boolean = (selectedPhases.Count + selectedMilestones.Count = 0)

            ' bestimme die Projekte, die gezeichnet werden sollen
            ' und bestimme das kleinste / resp größte auftretende Datum 
            Call bestimmeProjekteAndMinMaxDates(selectedPhases, selectedMilestones, _
                                                selectedRoles, selectedCosts, _
                                                selectedBUs, selectedTyps, _
                                                showRangeLeft, showRangeRight, awinSettings.mppSortiertDauer, _
                                                projCollection, minDate, maxDate, _
                                                isMultiprojektSicht, projMitVariants)


            
            If objectsToDo <> projCollection.Count Then
                objectsToDo = projCollection.Count
            End If
            

            '
            ' bestimme das Start und Ende Datum des PPT Kalenders
            Call calcStartEndePPTKalender(minDate, maxDate, _
                                          pptStartofCalendar, pptEndOfCalendar)


            ' bestimme die benötigte Höhe einer Zeile im Report ( nur wenn nicht schon bestimmt also zeilenhoehe <> 0
            If pptFirstTime And zeilenhoehe = 0.0 Then
                With rds

                    zeilenhoehe = bestimmeMppZeilenHoehe(.pptSlide, .phaseVorlagenShape, .milestoneVorlagenShape,
                                                        selectedPhases.Count, selectedMilestones.Count, _
                                                        .MsDescVorlagenShape, .MsDateVorlagenShape, _
                                                        .PhDescVorlagenShape, .PhDateVorlagenShape,
                                                        .projectNameVorlagenShape, _
                                                        .durationArrowShape, .durationTextShape)
                End With

            End If


            Dim hproj As New clsProjekt
            Dim hhproj As New clsProjekt
            Dim maxZeilen As Integer = 0
            Dim anzZeilen As Integer = 0
            Dim gesamtAnzZeilen As Integer = 0
            Dim projekthoehe As Double = zeilenhoehe

            If awinSettings.mppExtendedMode Then

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


            Else
                projekthoehe = zeilenhoehe
            End If

            '
            ' bestimme die relativen Abstände der Text-Shapes zu ihrem Phase/Milestone Element
            '
            Call rds.calcRelDisTxtToElm()


            '
            ' bestimme das Format  

            Dim neededSpace As Double


            If awinSettings.mppExtendedMode Then                    ' für Berichte im extendedMode
                If awinSettings.mppOnePage Then
                    neededSpace = gesamtAnzZeilen * zeilenhoehe
                Else
                    neededSpace = maxZeilen * zeilenhoehe
                End If
            Else
                neededSpace = (projCollection.Count + 1) * zeilenhoehe ' für normale Berichte hier: projekthoehe = zeilenhoehe
            End If




            Dim availableSpace As Double
            availableSpace = rds.drawingAreaBottom - rds.drawingAreaTop

            Dim oldHeight As Double
            Dim oldwidth As Double

            oldHeight = pptCurrentPresentation.PageSetup.SlideHeight
            oldwidth = pptCurrentPresentation.PageSetup.SlideWidth


            Dim curHeight As Double = oldHeight
            Dim curWidth As Double = oldwidth



            If (availableSpace < neededSpace And awinSettings.mppOnePage) Or _
               (availableSpace < neededSpace And awinSettings.mppExtendedMode) Then

                Dim ix As Integer = format
                Dim ok As Boolean = True
                ' jetzt erst mal die Schriftgrößen und Liniendicken merken ...

                Dim sizeMemory() As Single
                Dim relativeSizeMemory As New SortedList(Of String, Double())

                With rds
                    sizeMemory = saveSizesOfElements(.projectNameVorlagenShape, _
                                                 .MsDescVorlagenShape, .MsDateVorlagenShape, _
                                                 .PhDescVorlagenShape, .PhDateVorlagenShape, _
                                                 .phaseVorlagenShape, .milestoneVorlagenShape, _
                                                 .projectVorlagenShape, .ampelVorlagenShape)
                End With


                If pptApp.Version = "14.0" Then
                    ' muss nichts machen

                Else

                    relativeSizeMemory = saveRelSizesOfElements(pptslide, oldHeight, oldwidth)

                End If


                Do While availableSpace < neededSpace And ix > 0

                    With pptCurrentPresentation

                        .PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeCustom

                        If querFormat Then
                            .PageSetup.SlideWidth = dinFormatA(ix - 1, 0)
                            .PageSetup.SlideHeight = dinFormatA(ix - 1, 1)
                        Else
                            .PageSetup.SlideWidth = dinFormatA(ix - 1, 1)
                            .PageSetup.SlideHeight = dinFormatA(ix - 1, 0)
                        End If


                    End With

                    curHeight = pptCurrentPresentation.PageSetup.SlideHeight
                    curWidth = pptCurrentPresentation.PageSetup.SlideWidth

                    ' jetzt muss bestimmt werden , ob es sich um Powerpoint 2010 oder 2013 handelt 
                    ' wenn ja, dann müssen die markierten Shapes entsprechend behandelt werden 

                    If pptApp.Version = "14.0" Then
                        ' muss nichts machen
                    Else

                        Call restoreRelSizesDuePPT2013(relativeSizeMemory, curHeight, curWidth, pptslide)
                    End If

                    ' jetzt wieder die Koordinaten neu berechnen 
                    Call rds.bestimmeZeichenKoordinaten()

                    'Call bestimmeZeichenKoordinaten(containerShape, _
                    '                                calendarLineShape, calenderHeightShape, legendLineShape, _
                    '                                containerLeft, containerRight, containerTop, containerBottom, _
                    '                                calendarLeft, calendarRight, calendarTop, calendarBottom, _
                    '                                drawingAreaLeft, drawingAreaRight, drawingAreaTop, drawingAreaBottom, _
                    '                                projectListLeft, _
                    '                                legendAreaLeft, legendAreaRight, legendAreaTop, legendAreaBottom)

                    availableSpace = rds.drawingAreaBottom - rds.drawingAreaTop

                    If availableSpace < neededSpace Then
                        ix = ix - 1
                    End If

                Loop

                ix = ix - 1
                If ix < 0 Then
                    ix = 0
                End If

                ' jetzt die Schriftgrößen und Liniendicken wieder auf den ursprünglichen Wert setzen 
                If pptApp.Version = "14.0" Then
                    With rds
                        Call restoreSizesOfElements(sizeMemory, .projectNameVorlagenShape, _
                                            .MsDescVorlagenShape, .MsDateVorlagenShape, _
                                            .PhDescVorlagenShape, .PhDateVorlagenShape, _
                                            .phaseVorlagenShape, .milestoneVorlagenShape, _
                                            .projectVorlagenShape, .ampelVorlagenShape)
                    End With


                End If


                ' jetzt alle Text Shapes, die auf der Folie ihre relative Größe behalten sollen 
                ' entsprechend um den errechneten Faktor anpassen

                Dim enlargeTxtFaktor As Double = curHeight / oldHeight
                Call enlargeTxtShapes(enlargeTxtFaktor, pptslide)

                ' ur: 30.03.2015:jetzt alle Beschriftungen der Phasen und Meilensteine wieder im richtigen Abstand positionieren 
                ' 
                With rds
                    .PhDescVorlagenShape.Top = .phaseVorlagenShape.Top + .yOffsetPhToText
                    .PhDateVorlagenShape.Top = .phaseVorlagenShape.Top + .yOffsetPhToDate

                    .MsDescVorlagenShape.Top = .milestoneVorlagenShape.Top + .yOffsetMsToText
                    .MsDateVorlagenShape.Top = .milestoneVorlagenShape.Top + .yOffsetMsToDate
                End With

            End If


            If pptFirstTime Then

                'ur: 25.03.2015: sichern der im Format veränderten Folie
                pptslide.Copy()
                pptCurrentPresentation.Slides.Paste(1).Name = "tmpSav"
                pptFirstTime = False
                legendFontSize = rds.projectNameVorlagenShape.TextFrame2.TextRange.Font.Size

            End If

            ' zeichne den Kalender
            Dim calendargroup As pptNS.Shape = Nothing

            Try

                With rds
                    
                        Call zeichnePPTCalendar(pptslide, calendargroup, _
                                            pptStartofCalendar, pptEndOfCalendar, _
                                            .calendarLineShape, .calendarHeightShape, .calendarStepShape, .calendarMarkShape, _
                                            .yearVorlagenShape, .quarterMonthVorlagenShape, .calendarYearSeparator, .calendarQuartalSeparator, _
                                            .drawingAreaBottom)

                End With



            Catch ex As Exception

            End Try


            ' jetzt wird das aufgerufen mit dem gesamten fertig gezeichneten Kalender, der fertig positioniert ist 

            ' zeichne die Projekte 

            Try

                With rds
                    
                        Call zeichnePPTprojects(pptslide, projCollection, objectsDone, _
                                        pptStartofCalendar, pptEndOfCalendar, _
                                        .drawingAreaLeft, .drawingAreaRight, .drawingAreaTop, .drawingAreaBottom, _
                                        zeilenhoehe, .projectListLeft, _
                                        selectedPhases, selectedMilestones, selectedRoles, selectedCosts, _
                                        .projectNameVorlagenShape, .MsDescVorlagenShape, .MsDateVorlagenShape, _
                                        .PhDescVorlagenShape, .PhDateVorlagenShape, _
                                        .phaseVorlagenShape, .milestoneVorlagenShape, .projectVorlagenShape, .ampelVorlagenShape,
                                        .rowDifferentiatorShape, .buColorShape, .phaseDelimiterShape, _
                                        .durationArrowShape, .durationTextShape, _
                                        .yOffsetMsToText, .yOffsetMsToDate, .yOffsetPhToText, .yOffsetPhToDate, _
                                        worker, e)

                End With





            Catch ex As Exception

                If Not IsNothing(rds.errorVorlagenShape) Then
                    rds.errorVorlagenShape.Copy()
                    errorShape = pptslide.Shapes.Paste
                    With errorShape.Item(1)
                        .TextFrame2.TextRange.Text = ex.Message
                    End With
                Else
                    ' erstmal sonst nichts 
                End If


            End Try


            ' zeichne die Legende 
            If awinSettings.mppShowLegend Then
                Try

                    With rds
                        Call zeichnePPTlegende(pptslide, _
                                        selectedPhases, selectedMilestones, selectedRoles, selectedCosts, _
                                        .legendAreaTop, .legendAreaLeft, .legendAreaRight, .legendAreaBottom, _
                                        .legendLineShape, .legendStartShape, _
                                        .legendTextVorlagenShape, .legendPhaseVorlagenShape, .legendMilestoneVorlagenShape, _
                                        .projectVorlagenShape, .ampelVorlagenShape, .legendBuColorShape)

                    End With


                Catch ex As Exception

                    If Not IsNothing(rds.errorVorlagenShape) Then
                        rds.errorVorlagenShape.Copy()
                        errorShape = pptslide.Shapes.Paste
                        With errorShape.Item(1)
                            .TextFrame2.TextRange.Text = ex.Message
                        End With
                    End If

                End Try

            End If




        ElseIf Not IsNothing(rds.errorVorlagenShape) Then
            rds.errorVorlagenShape.Copy()
            errorShape = pptslide.Shapes.Paste
            With errorShape.Item(1)
                .TextFrame2.TextRange.Text = missingShapes
            End With
        Else
            Call MsgBox("es fehlen Shapes: " & vbLf & missingShapes)
        End If

        ' jetzt werden alle Shapes gelöscht ... 
        Call rds.deleteShapes()


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
    ''' <param name="pptApp"></param>
    ''' <param name="pptCurrentPresentation"></param>
    ''' <param name="pptslide"></param>
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
    ''' <param name="worker"></param>
    ''' <param name="e"></param>
    ''' <param name="isMultiprojektSicht"></param>
    ''' <param name="hproj"></param>
    ''' <param name="kennzeichnung"></param>
    ''' <remarks></remarks>
    Private Sub zeichneSwimlane2Sicht(ByRef pptApp As pptNS.Application, ByRef pptCurrentPresentation As pptNS.Presentation, ByRef pptslide As pptNS.Slide, _
                                                 ByRef swimLanesToDo As Integer, ByRef swimLanesDone As Integer, ByRef pptFirstTime As Boolean, _
                                                 ByRef zeilenhoehe As Double, ByRef legendFontSize As Double, _
                                                 ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, _
                                                 ByVal selectedRoles As Collection, ByVal selectedCosts As Collection, _
                                                 ByVal selectedBUs As Collection, ByVal selectedTyps As Collection, _
                                                 ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs, _
                                                 ByVal isMultiprojektSicht As Boolean, ByVal hproj As clsProjekt, _
                                                 ByVal kennzeichnung As String)



        ' Wichtig für Kalendar 
        Dim pptStartofCalendar As Date = Nothing, pptEndOfCalendar As Date = Nothing
        Dim errorShape As pptNS.ShapeRange = Nothing


        Dim dinFormatA(4, 1) As Double
        Dim querFormat As Boolean
        Dim curFormatSize(1) As Double

        Dim maxZeilen As Integer = 0
        Dim anzZeilen As Integer = 0
        Dim gesamtAnzZeilen As Integer = 0



        dinFormatA(0, 0) = 3120.0
        dinFormatA(0, 1) = 2206.15

        dinFormatA(1, 0) = 2206.15
        dinFormatA(1, 1) = 1560.0

        dinFormatA(2, 0) = 1560.0
        dinFormatA(2, 1) = 1103.0

        dinFormatA(3, 0) = 1103.0
        dinFormatA(3, 1) = 780.0

        dinFormatA(4, 0) = 780.0
        dinFormatA(4, 1) = 540.0

        ' Ende Übernahme

        Dim format As Integer = 4
        'Dim tmpslideID As Integer

        ' an der Variablen lässt sich in der Folge erkennen, ob die Segmente BHTC Milestones gezeichnet werden müssen oder 
        ' ob ganz allgemein nach Swimlanes gesucht wird ... 
        Dim isBHTCSchema As Boolean = (kennzeichnung = "Swimlanes2")

        Dim rds As New clsPPTShapes
        Dim considerZeitraum As Boolean = (showRangeLeft > 0 And showRangeRight > showRangeLeft)
        Dim cphase As clsPhase

        ' mit disem Befehl werden auch die ganzen Hilfsshapes in der Klasse gesetzt 
        rds.pptSlide = pptslide


        ' jetzt muss geprüft werden, ob überhaupt alle Angaben gemacht wurden ... 
        'If completeMppDefinition.Sum = completeMppDefinition.Length Then
        Dim missingShapes As String = rds.getMissingShpNames(kennzeichnung)
        If missingShapes.Length = 0 Then
            ' es fehlt nichts ... andernfalls stehen hier die Namen mit den Shapes, die fehlen ...

            If pptCurrentPresentation.PageSetup.SlideOrientation = MsoOrientation.msoOrientationHorizontal Then
                querFormat = True
            Else
                querFormat = False
            End If


            curFormatSize(0) = pptCurrentPresentation.PageSetup.SlideWidth
            curFormatSize(1) = pptCurrentPresentation.PageSetup.SlideHeight

            ' jetzt werden die DinA Formate gesetzt 
            ' Voraussetzung ist allerdings, dass es sich bei der Vorlage um DIN A4 handelt 
            Dim paperSizeRatio As Double

            Dim considerAll As Boolean = (selectedPhases.Count + selectedMilestones.Count = 0)
            Dim selectedPhaseIDs As New Collection
            Dim selectedMilestoneIDs As New Collection
            Dim breadcrumbArray As String() = Nothing

            If Not considerAll Then
                selectedPhaseIDs = hproj.getElemIdsOf(selectedPhases, False)
                selectedMilestoneIDs = hproj.getElemIdsOf(selectedMilestones, True)
                breadcrumbArray = hproj.getBreadCrumbArray(selectedPhaseIDs, selectedMilestoneIDs)
            End If

            If pptFirstTime Then

                swimLanesToDo = hproj.getSwimLanesCount(considerAll, breadcrumbArray, isBHTCSchema)

                If pptCurrentPresentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeA4Paper Then

                    If querFormat Then
                        paperSizeRatio = curFormatSize(0) / curFormatSize(1)
                        dinFormatA(4, 0) = curFormatSize(0)
                        dinFormatA(4, 1) = curFormatSize(1)
                    Else
                        paperSizeRatio = curFormatSize(1) / curFormatSize(0)
                        dinFormatA(4, 1) = curFormatSize(0)
                        dinFormatA(4, 0) = curFormatSize(1)
                    End If

                    dinFormatA(3, 0) = dinFormatA(4, 0) * paperSizeRatio
                    dinFormatA(3, 1) = dinFormatA(4, 1) * paperSizeRatio

                ElseIf pptCurrentPresentation.PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeA3Paper Then
                    If querFormat Then
                        paperSizeRatio = curFormatSize(0) / curFormatSize(1)
                        dinFormatA(3, 0) = curFormatSize(0)
                        dinFormatA(3, 1) = curFormatSize(1)

                    Else
                        paperSizeRatio = curFormatSize(1) / curFormatSize(0)
                        dinFormatA(3, 1) = curFormatSize(0)
                        dinFormatA(3, 0) = curFormatSize(1)
                    End If

                    dinFormatA(4, 0) = dinFormatA(3, 0) / paperSizeRatio
                    dinFormatA(4, 1) = dinFormatA(3, 1) / paperSizeRatio

                Else
                    Call MsgBox("Vorlage ist weder ein A4 noch ein A3 Format ... bitte verwenden Sie eine A4 oder A3 Vorlage")
                    'Throw New ArgumentException("Vorlage ist weder ein A4 noch ein A3 Format ... bitte verwenden Sie eine A4 oder A3 Vorlage")
                End If


                For i = 2 To 0 Step -1
                    dinFormatA(i, 0) = dinFormatA(i + 1, 0) * paperSizeRatio
                    dinFormatA(i, 1) = dinFormatA(i + 1, 1) * paperSizeRatio
                Next
            Else
                ' pptFirstTime war False, d.h. das Format wurde bereits angepasst
            End If


            Call rds.plausibilityAdjustments()


            Call rds.bestimmeZeichenKoordinaten()

            Dim projCollection As New SortedList(Of Double, String)
            Dim minDate As Date, maxDate As Date

            ' bestimme die Projekte, die gezeichnet werden sollen
            ' und bestimme das kleinste / resp größte auftretende Datum 
            Call bestimmeProjekteAndMinMaxDates(selectedPhases, selectedMilestones, _
                                                selectedRoles, selectedCosts, _
                                                selectedBUs, selectedTyps, _
                                                showRangeLeft, showRangeRight, awinSettings.mppSortiertDauer, _
                                                projCollection, minDate, maxDate, _
                                                isMultiprojektSicht, hproj)


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

            ' aktuell wird davon ausgegangen , dass in projMitVariants nur eine Variante ist
            ' Swimlane wird aktuell nur für BHTC erstellt , da gibt es noch keine Varianten 



            ' tk:1.2.16 ExtendedMode macht nur Sinn, wenn mindestens 1 Phase selektiert wurde. oder aber considerAll gilt: 
            awinSettings.mppExtendedMode = (awinSettings.mppExtendedMode And (selectedPhases.Count > 0)) Or _
                                            (awinSettings.mppExtendedMode And considerAll)



            ' muss nur bestimmt werden, wenn zum ersten Mal reinkommt 


            '
            ' bestimme das Start und Ende Datum des PPT Kalenders
            Call calcStartEndePPTKalender(minDate, maxDate, _
                                          pptStartofCalendar, pptEndOfCalendar)

            ' jetzt für Swimlanes Behandlung Kalender in der Klasse setzen 

            Call rds.setCalendarDates(pptStartofCalendar, pptEndOfCalendar)

            ' die neue Art Zeilenhöhe und die Offset Werte zu bestimmen 

            Call rds.bestimmeZeilenHoehe(selectedPhases.Count, selectedMilestones.Count, considerAll)


            ' tk 1.2.16
            ' eigentlich muss er das Ganze nur machen, wenn pptFirsttime 
            If pptFirstTime Then

                If awinSettings.mppExtendedMode Then

                    ' jetzt muss die Gesamt-Zahl an Zeilen ermittelt werden , die die einzelnen Swimlanes bentötigen 

                    For i = 1 To swimLanesToDo

                        cphase = hproj.getSwimlane(i, considerAll, breadcrumbArray, isBHTCSchema)

                        Dim swimLaneZeilen As Integer = hproj.calcNeededLinesSwl(cphase.nameID, selectedPhaseIDs, selectedMilestoneIDs, _
                                                                                 awinSettings.mppExtendedMode, _
                                                                                 considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR, _
                                                                                 considerAll)

                        anzZeilen = anzZeilen + swimLaneZeilen
                    Next


                Else
                    anzZeilen = swimLanesToDo
                End If


                '
                ' bestimme das Format  

                Dim neededSpace As Double


                If isBHTCSchema Then
                    ' jetzt müssen noch die Segment Höhen  berechnet werden 

                    neededSpace = anzZeilen * rds.zeilenHoehe + _
                                    hproj.getSegmentsCount(considerAll, breadcrumbArray, isBHTCSchema) * rds.segmentHoehe
                Else

                    neededSpace = anzZeilen * rds.zeilenHoehe

                End If



                Dim oldHeight As Double
                Dim oldwidth As Double

                oldHeight = pptCurrentPresentation.PageSetup.SlideHeight
                oldwidth = pptCurrentPresentation.PageSetup.SlideWidth


                Dim curHeight As Double = oldHeight
                Dim curWidth As Double = oldwidth



                If (rds.availableSpace < neededSpace And awinSettings.mppOnePage) Then

                    ' es muss das Format angepasst werden ... 

                    Dim ix As Integer = format
                    Dim ok As Boolean = True
                    ' jetzt erst mal die Schriftgrößen und Liniendicken merken ...

                    Dim sizeMemory() As Single
                    Dim relativeSizeMemory As New SortedList(Of String, Double())

                    With rds
                        sizeMemory = saveSizesOfElements(.projectNameVorlagenShape, _
                                                     .MsDescVorlagenShape, .MsDateVorlagenShape, _
                                                     .PhDescVorlagenShape, .PhDateVorlagenShape, _
                                                     .phaseVorlagenShape, .milestoneVorlagenShape, _
                                                     .projectVorlagenShape, .ampelVorlagenShape, _
                                                     .segmentVorlagenShape)
                    End With


                    If pptApp.Version = "14.0" Then
                        ' muss nichts machen

                    Else

                        relativeSizeMemory = saveRelSizesOfElements(pptslide, oldHeight, oldwidth)

                    End If


                    Do While rds.availableSpace < neededSpace And ix > 0

                        With pptCurrentPresentation

                            .PageSetup.SlideSize = PowerPoint.PpSlideSizeType.ppSlideSizeCustom

                            If querFormat Then
                                .PageSetup.SlideWidth = dinFormatA(ix - 1, 0)
                                .PageSetup.SlideHeight = dinFormatA(ix - 1, 1)
                            Else
                                .PageSetup.SlideWidth = dinFormatA(ix - 1, 1)
                                .PageSetup.SlideHeight = dinFormatA(ix - 1, 0)
                            End If


                        End With

                        curHeight = pptCurrentPresentation.PageSetup.SlideHeight
                        curWidth = pptCurrentPresentation.PageSetup.SlideWidth

                        ' jetzt muss bestimmt werden , ob es sich um Powerpoint 2010 oder 2013 handelt 
                        ' wenn ja, dann müssen die markierten Shapes entsprechend behandelt werden 

                        If pptApp.Version = "14.0" Then
                            ' muss nichts machen
                        Else

                            Call restoreRelSizesDuePPT2013(relativeSizeMemory, curHeight, curWidth, pptslide)
                        End If

                        ' jetzt wieder die Koordinaten neu berechnen 
                        Call rds.bestimmeZeichenKoordinaten()


                        If rds.availableSpace < neededSpace Then
                            ix = ix - 1
                        End If

                    Loop

                    ix = ix - 1
                    If ix < 0 Then
                        ix = 0
                    End If

                    ' jetzt die Schriftgrößen und Liniendicken wieder auf den ursprünglichen Wert setzen 
                    If pptApp.Version = "14.0" Then
                        With rds
                            Call restoreSizesOfElements(sizeMemory, .projectNameVorlagenShape, _
                                                .MsDescVorlagenShape, .MsDateVorlagenShape, _
                                                .PhDescVorlagenShape, .PhDateVorlagenShape, _
                                                .phaseVorlagenShape, .milestoneVorlagenShape, _
                                                .projectVorlagenShape, .ampelVorlagenShape, _
                                                .segmentVorlagenShape)
                        End With


                    End If


                    ' jetzt alle Text Shapes, die auf der Folie ihre relative Größe behalten sollen 
                    ' entsprechend um den errechneten Faktor anpassen

                    Dim enlargeTxtFaktor As Double = curHeight / oldHeight
                    Call enlargeTxtShapes(enlargeTxtFaktor, pptslide)

                    ' ur: 30.03.2015:jetzt alle Beschriftungen der Phasen und Meilensteine wieder im richtigen Abstand positionieren 
                    ' tk 2.2 braucht man nicht mehr ...
                    'With rds
                    '    .PhDescVorlagenShape.Top = .phaseVorlagenShape.Top + .yOffsetPhToText
                    '    .PhDateVorlagenShape.Top = .phaseVorlagenShape.Top + .yOffsetPhToDate

                    '    .MsDescVorlagenShape.Top = .milestoneVorlagenShape.Top + .yOffsetMsToText
                    '    .MsDateVorlagenShape.Top = .milestoneVorlagenShape.Top + .yOffsetMsToDate
                    'End With

                End If

            End If



            If pptFirstTime Then

                ' jetzt erst mal den Kalender zeichnen 
                ' zeichne den Kalender
                Dim calendargroup As pptNS.Shape = Nothing

                Try

                    With rds
                        ' das demnächst abändern auf 
                        Call zeichne3RowsCalendar(rds, calendargroup)

                    End With



                Catch ex As Exception

                End Try

                ' wenn Legende gezeichnet werden soll - die Legende zeichnen 

                ' zeichne die Legende 
                If awinSettings.mppShowLegend Then
                    Try

                        With rds
                            Call zeichnePPTlegende(pptslide, _
                                            selectedPhases, selectedMilestones, selectedRoles, selectedCosts, _
                                            .legendAreaTop, .legendAreaLeft, .legendAreaRight, .legendAreaBottom, _
                                            .legendLineShape, .legendStartShape, _
                                            .legendTextVorlagenShape, .legendPhaseVorlagenShape, .legendMilestoneVorlagenShape, _
                                            .projectVorlagenShape, .ampelVorlagenShape, .legendBuColorShape)

                        End With


                    Catch ex As Exception

                        If Not IsNothing(rds.errorVorlagenShape) Then
                            rds.errorVorlagenShape.Copy()
                            errorShape = pptslide.Shapes.Paste
                            With errorShape.Item(1)
                                .Top = rds.legendLineShape.Top + 10
                                .TextFrame2.TextRange.Text = ex.Message
                            End With
                        End If

                    End Try

                End If

                'ur: 25.03.2015: sichern der im Format veränderten Folie
                rds.pptSlide.Copy()
                pptCurrentPresentation.Slides.Paste(1).Name = "tmpSav"
                pptFirstTime = False
                legendFontSize = rds.projectNameVorlagenShape.TextFrame2.TextRange.Font.Size

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
            curSwl = hproj.getSwimlane(curSwimlaneIndex, considerAll, breadcrumbArray, isBHTCSchema)
            prevSwl = hproj.getSwimlane(curSwimlaneIndex - 1, considerAll, breadcrumbArray, isBHTCSchema)


            If Not IsNothing(curSwl) Then


                Dim segmentChanged As Boolean = False
                Dim curSegmentID As String = hproj.hierarchy.getParentIDOfID(curSwl.nameID)

                If isBHTCSchema Then
                    If Not IsNothing(prevSwl) Then
                        segmentChanged = hproj.hierarchy.getParentIDOfID(prevSwl.nameID) <> _
                                            hproj.hierarchy.getParentIDOfID(curSwl.nameID)
                    End If

                    If swimLanesDone = 0 Or segmentChanged Then
                        Call zeichneSwlSegmentinAktZeile(rds, curYPosition, curSegmentID)
                        segmentChanged = False
                    End If
                End If
                
                


                ' jetzt werden soviele wie möglich Swimlanes gezeichnet ... 
                Dim swimLaneZeilen As Integer = hproj.calcNeededLinesSwl(curSwl.nameID, selectedPhaseIDs, selectedMilestoneIDs, _
                                                                                 awinSettings.mppExtendedMode, _
                                                                                 considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR, _
                                                                                 considerAll)

                Do While (curSwimlaneIndex <= swimLanesToDo) And _
                        (swimLaneZeilen * rds.zeilenHoehe + curYPosition <= rds.drawingAreaBottom)


                    ' Zwischen-Meldung ausgeben ...
                    If worker.WorkerSupportsCancellation Then

                        If worker.CancellationPending Then
                            e.Cancel = True
                            e.Result = "Berichterstellung abgebrochen ..."
                            Exit Sub
                        End If

                    End If

                    ' Zwischenbericht abgeben ...
                    e.Result = "Swimlane '" & elemNameOfElemID(curSwl.nameID) & "' wird gezeichnet  ...."
                    If worker.WorkerReportsProgress Then
                        worker.ReportProgress(0, e)
                    End If

                    ' jetzt die Swimlane zeichnen
                    ' hier ist ja gewährleistet, dass alle Phasen und Meilensteine dieser Swimlane Platz finden 
                    Call zeichneSwimlaneOfProject(rds, curYPosition, toggleRow, _
                                                  hproj, curSwl.nameID, considerAll, _
                                                  breadcrumbArray, _
                                                  considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR, _
                                                  selectedPhaseIDs, selectedMilestoneIDs, _
                                                  selectedRoles, selectedCosts, _
                                                  swimLaneZeilen)


                    prevSwl = curSwl

                    curSwimlaneIndex = curSwimlaneIndex + 1
                    curSwl = hproj.getSwimlane(curSwimlaneIndex, considerAll, breadcrumbArray, isBHTCSchema)

                    If Not IsNothing(curSwl) Then

                        If isBHTCSchema Then
                            segmentChanged = hproj.hierarchy.getParentIDOfID(prevSwl.nameID) <> _
                                        hproj.hierarchy.getParentIDOfID(curSwl.nameID)

                        End If
                        
                        swimLaneZeilen = hproj.calcNeededLinesSwl(curSwl.nameID, selectedPhaseIDs, selectedMilestoneIDs, _
                                                                                 awinSettings.mppExtendedMode, _
                                                                                 considerZeitraum, zeitraumGrenzeL, zeitraumGrenzeR, _
                                                                                 considerAll)

                        If isBHTCSchema Then
                            If segmentChanged And _
                                (swimLaneZeilen * rds.zeilenHoehe + curYPosition + rds.segmentVorlagenShape.Height <= rds.drawingAreaBottom) Then

                                curSegmentID = hproj.hierarchy.getParentIDOfID(curSwl.nameID)
                                Call zeichneSwlSegmentinAktZeile(rds, curYPosition, curSegmentID)
                                segmentChanged = False
                            End If
                        End If
                        
                    Else
                        segmentChanged = False
                    End If


                Loop

                ' jetzt die Anzahl ..Done bestimmen
                swimLanesDone = curSwimlaneIndex - 1

            End If

        ElseIf Not IsNothing(rds.errorVorlagenShape) Then
            rds.errorVorlagenShape.Copy()
            errorShape = pptslide.Shapes.Paste
            With errorShape.Item(1)
                .TextFrame2.TextRange.Text = missingShapes
            End With
        End If

        ' jetzt werden alle Shapes gelöscht ... 
        Call rds.deleteShapes()


    End Sub

    ''' <summary>
    ''' übernimmt das zeichnen der LegendenTabelle bzw. die Vorbereitugen dazu 
    ''' </summary>
    ''' <param name="pptslide"></param>
    ''' <param name="tableShape"></param>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <remarks></remarks>
    Private Sub prepZeichneLegendenTabelle(ByRef pptslide As pptNS.Slide, ByRef tableShape As pptNS.Shape, ByVal legendFontSize As Single, _
                                           ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection)

        Dim legendPhaseVorlagenShape As pptNS.Shape = Nothing
        Dim legendMilestoneVorlagenShape As pptNS.Shape = Nothing
        Dim legendBuColorShape As pptNS.Shape = Nothing
        Dim errorVorlagenShape As pptNS.Shape = Nothing
        Dim errorShape As pptNS.ShapeRange = Nothing

        ' mit completeLegendDefinition wird überprüft , ob alle Informationen/Shapes für das Erstellen der Legenden Tabelle vorhanden sind
        Dim completeLegendDefinition() As Integer
        ReDim completeLegendDefinition(1)

        Dim anzShapes As Integer = pptslide.Shapes.Count
        Dim pptShape As pptNS.Shape
        ' jetzt wird die listofShapes aufgebaut - das sind alle Shapes, die ersetzt werden müssen ...
        ' bzw. alle Shapes, die "gemerkt" werden müssen
        For i = 1 To anzShapes
            pptShape = pptslide.Shapes(i)

            With pptShape

                ' jetzt muss geprüft werden, ob es sich um ein definierendes Element für die Multiprojekt-Sichten handelt
                If .Title.Length > 0 Then
                    Select Case .Title

                        Case "LegendPhase"
                            legendPhaseVorlagenShape = pptShape
                            legendPhaseVorlagenShape.TextFrame2.TextRange.Font.Size = legendFontSize
                            completeLegendDefinition(0) = 1


                        Case "LegendMilestone"
                            legendMilestoneVorlagenShape = pptShape
                            legendMilestoneVorlagenShape.TextFrame2.TextRange.Font.Size = legendFontSize
                            completeLegendDefinition(1) = 1

                        Case "Fehlermeldung"
                            ' optional 
                            errorVorlagenShape = pptShape

                        Case "LegendBuColor"
                            ' optional
                            legendBuColorShape = pptShape
                            legendBuColorShape.TextFrame2.TextRange.Font.Size = legendFontSize
                        Case Else


                    End Select
                End If


            End With
        Next


        If completeLegendDefinition.Sum = completeLegendDefinition.Length Then
            ' alle Information vorhanden 

            Try
                Call zeichneLegendenTabelle(tableShape, pptslide, _
                                            selectedPhases, selectedMilestones, _
                                            legendPhaseVorlagenShape, legendMilestoneVorlagenShape, _
                                            legendBuColorShape)

            Catch ex As Exception
                errorVorlagenShape.Copy()
                errorShape = pptslide.Shapes.Paste
                With errorShape(1)
                    '.TextFrame2.TextRange.Text = "Fehler beim Zeichnen  Legenden Symbole Phase oder Meilenstein fehlen "
                    .TextFrame2.TextRange.Text = ex.Message
                End With
            End Try

            If Not IsNothing(legendBuColorShape) Then
                legendBuColorShape.Delete()
            End If

            legendPhaseVorlagenShape.Delete()
            legendMilestoneVorlagenShape.Delete()

        ElseIf Not IsNothing(errorVorlagenShape) Then
            errorVorlagenShape.Copy()
            errorShape = pptslide.Shapes.Paste
            With errorShape(1)
                .TextFrame2.TextRange.Text = "die Legenden Symbole Phase oder Meilenstein fehlen "
            End With


        End If

        Try
            If Not IsNothing(errorVorlagenShape) Then
                errorVorlagenShape.Delete()
            End If
        Catch ex As Exception

        End Try

    End Sub



End Module
