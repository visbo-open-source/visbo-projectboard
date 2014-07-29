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
    Public Sub createPPTReportFromProjects(ByVal pptTemplate As String, ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)

        Dim awinSelection As xlNS.ShapeRange

        Dim request As New Request(awinSettings.databaseName)
        Dim singleShp As xlNS.Shape
        Dim hproj As clsProjekt
        Dim vglName As String = " "
        Dim pName As String, variantName As String
        Dim vorlagenDateiName As String = pptTemplate
        Dim tatsErstellt As Integer = 0

        Try
            awinSelection = appInstance.ActiveWindow.Selection.ShapeRange
            awinSelection = CType(appInstance.ActiveWindow.Selection.ShapeRange, xlNS.ShapeRange)
        Catch ex As Exception
            awinSelection = Nothing
        End Try

        For Each singleShp In awinSelection
            With singleShp
                If .AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle Or _
                    (.AutoShapeType = MsoAutoShapeType.msoShapeMixed And Not .HasChart _
                     And Not .Connector = Microsoft.Office.Core.MsoTriState.msoTrue) Then

                    Try
                        hproj = ShowProjekte.getProject(singleShp.Name)
                    Catch ex As Exception

                        Call MsgBox(singleShp.Name & " nicht gefunden ...")
                        Exit Sub
                    End Try

                    If Not projekthistorie Is Nothing Then
                        If projekthistorie.Count > 0 Then
                            vglName = projekthistorie.First.name
                        End If
                    End If

                    With hproj
                        pName = .name
                        variantName = .variantName
                    End With

                    If vglName.Trim <> pName.Trim Then
                        If request.pingMongoDb() Then
                            Try
                                projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pName, variantName:=variantName, _
                                                                                storedEarliest:=Date.MinValue, storedLatest:=Date.Now)
                                projekthistorie.Add(Date.Now, hproj)
                            Catch ex As Exception
                                projekthistorie = Nothing
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

                    e.Result = " Report für Projekt '" & hproj.name & "' wird erstellt !"
                    worker.ReportProgress(0, e)
                    'frmSelectPPTTempl.statusNotification.Text = " Report für Projekt '" & hproj.name & " wird erstellt !"

                    createPPTSlidesFromProject(hproj, vorlagenDateiName)
                    tatsErstellt = tatsErstellt + 1

                End If
            End With
        Next

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
    Public Sub createPPTSlidesFromProject(ByRef hproj As clsProjekt, pptTemplate As String)
        Dim pptApp As pptNS.Application = Nothing
        Dim pptPresentation As pptNS.Presentation = Nothing
        Dim pptSlide As pptNS.Slide = Nothing
        Dim shapeRange As pptNS.ShapeRange = Nothing
        Dim presentationFile As String = awinPath & requirementsOrdner & "projektdossier.pptx"
        Dim pptShape As pptNS.Shape
        Dim pname As String = hproj.name
        Dim top As Double, left As Double, width As Double, height As Double
        Dim htop As Double, hleft As Double, hwidth As Double, hheight As Double
        Dim pptSize As Integer = 18
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
            pptApp = GetObject(, "PowerPoint.Application")
        Catch ex As Exception
            Try
                pptApp = CreateObject("PowerPoint.Application")
            Catch ex1 As Exception
                Call MsgBox("Powerpoint konnte nicht gestartet werden ..." & ex1.Message)
                Exit Sub
            End Try

        End Try


        ' entweder wird das template geöffnet ...
        ' oder aber es wird in die aktive Presentation geschrieben 

        If pptApp.Presentations.Count = 0 Then
            Try
                pptApp.Presentations.Open(presentationFile)
                pptPresentation = pptApp.ActivePresentation
            Catch ex As Exception
                pptPresentation = pptApp.Presentations.Add()
            End Try
        Else
            pptPresentation = pptApp.ActivePresentation
        End If
        Dim anzahlSlides As Integer = pptPresentation.Slides.Count
        Dim AnzAdded As Integer = pptPresentation.Slides.InsertFromFile(pptTemplate, anzahlSlides)
        Dim reportObj As xlNS.ChartObject
        Dim obj As New Object
        Dim kennzeichnung As String
        Dim anzShapes As Integer

        For j = 1 To AnzAdded
            pptSlide = pptPresentation.Slides(anzahlSlides + j)

            ' jetzt werden die Charts gezeichnet 

            anzShapes = pptSlide.Shapes.Count
            Dim newShape As pptNS.ShapeRange
            Dim newShape2 As pptNS.ShapeRange

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
                            tmpStr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {"(", ")"}, 3)
                            kennzeichnung = tmpStr(0).Trim
                        End If


                    Catch ex As Exception
                        kennzeichnung = "nicht identifizierbar"
                    End Try

                    If kennzeichnung = "Projekt-Name" Or _
                        kennzeichnung = "Soll-Ist & Prognose" Or _
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
                        kennzeichnung = "Teilprojekte" Or _
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
                        kennzeichnung = "Beschreibung" Or _
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

                                tmpStr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {"(", ")"}, 10)
                                kennzeichnung = tmpStr(0).Trim

                            Catch ex As Exception
                                kennzeichnung = "nicht identifizierbar"
                                tmpStr(0) = " "
                            End Try

                            Try
                                If tmpStr.Count < 2 Then
                                    qualifier = ""
                                    qualifier2 = ""
                                ElseIf tmpStr.Count = 2 Then
                                    qualifier = tmpStr(1).Trim
                                ElseIf tmpStr.Count >= 3 Then
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
                                    .TextFrame2.TextRange.Text = pname & ": " & qualifier
                                Else
                                    .TextFrame2.TextRange.Text = pname
                                End If

                            Case "Projekt-Grafik"

                                Try

                                    Call zeichneProjektGrafik(pptSlide, pptShape, hproj)

                                Catch ex As Exception

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

                                            tmpStr = qualifier.Trim.Split(New Char() {"#"}, 20)
                                            kennzeichnung = tmpStr(0).Trim

                                        Catch ex As Exception

                                        End Try

                                        For i = 1 To tmpStr.Count
                                            listOfItems.Add(tmpStr(i - 1).Trim)
                                        Next

                                    End If

                                    ' jetzt ist listofItems entsprechend gefüllt 
                                    If listOfItems.Count > 0 Then
                                        htop = 100
                                        hleft = 50
                                        hheight = 2 * ((listOfItems.Count - 1) * 20 + 110)
                                        hwidth = System.Math.Max(hproj.Dauer * boxWidth + 10, 24 * boxWidth + 10)

                                        Call createMsTrendAnalysisOfProject(hproj, obj, listOfItems, htop, hleft, hheight, hwidth)

                                        reportObj = obj
                                        notYetDone = True
                                    Else
                                        .TextFrame2.TextRange.Text = "es gibt keine Meilensteine im Projekt" & vbLf & hproj.name
                                    End If

                                Catch ex As Exception

                                End Try

                            Case "Teilprojekte"

                                Dim scale As Integer

                                Dim cproj As clsProjekt = Nothing
                                Dim vproj As clsProjektvorlage
                                auswahl = 0

                                scale = hproj.dauerInDays

                                If qualifier.Length > 0 Then
                                    If qualifier = "Vorlage" Then
                                        auswahl = 1
                                        vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                                        vproj.CopyTo(cproj)
                                        cproj.startDate = hproj.startDate


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

                                htop = 150
                                hleft = 150


                                hheight = 380
                                hwidth = 900
                                scale = cproj.dauerInDays

                                Dim noColorCollection As New Collection
                                reportObj = Nothing
                                Call createPhasesBalken(noColorCollection, cproj, reportObj, scale, htop, hleft, hheight, hwidth, auswahl)


                                notYetDone = True

                            Case "Vergleich mit Vorlage"

                                Dim vproj As clsProjektvorlage
                                Dim cproj As clsProjekt
                                Dim scale As Double
                                Dim noColorCollection As New Collection
                                Dim repObj1 As xlNS.ChartObject, repObj2 As xlNS.ChartObject


                                ' jetzt die Aktion durchführen ...


                                Try

                                    vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                                    cproj = New clsProjekt
                                    vproj.CopyTo(cproj)
                                    cproj.startDate = hproj.startDate

                                Catch ex As Exception
                                    Throw New Exception("Vorlage konnte nicht bestimmt werden")
                                End Try


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
                                        repObj1.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
                                        newShape = pptSlide.Shapes.Paste

                                        With newShape(1)
                                            .Top = top + 0.02 * height
                                            .Left = left + 0.02 * width
                                            .Width = width * 0.96
                                            topNext = top + 0.04 * height + .Height
                                            '.Height = height * 0.46
                                        End With

                                        repObj1.Delete()

                                        If Not repObj2 Is Nothing Then
                                            Try
                                                repObj2.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
                                                newShape2 = pptSlide.Shapes.Paste

                                                With newShape2(1)
                                                    .Top = topNext
                                                    .Left = left + 0.02 * width
                                                    .Width = width * 0.96
                                                    ' Height wird nicht gesetzt - bei Bildern wird das proportional automatisch gesetzt 
                                                End With

                                                ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                                Try
                                                    If newShape(1).Height + newShape2(1).Height > 0.96 * height Then
                                                        widthFaktor = 0.96 * height / (newShape(1).Height + newShape2(1).Height)
                                                        newShape.Width = widthFaktor * newShape.Width
                                                        newShape2.Width = widthFaktor * newShape2.Width
                                                        newShape2.Top = newShape.Top + newShape.Height + 0.02 * height
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
                                        repObj1.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
                                        newShape = pptSlide.Shapes.Paste

                                        With newShape(1)
                                            .Top = top + 0.02 * height
                                            .Left = left + 0.02 * width
                                            .Width = width * 0.96
                                            topNext = top + 0.04 * height + .Height
                                            '.Height = height * 0.46
                                        End With

                                        repObj1.Delete()

                                        If Not repObj2 Is Nothing Then
                                            Try
                                                repObj2.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
                                                newShape2 = pptSlide.Shapes.Paste

                                                With newShape2(1)
                                                    .Top = topNext
                                                    .Left = left + 0.02 * width
                                                    .Width = width * 0.96
                                                    '.Height = height * 0.46
                                                End With

                                                repObj2.Delete()

                                                ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                                If newShape(1).Height + newShape2(1).Height > 0.96 * height Then
                                                    widthFaktor = 0.96 * height / (newShape(1).Height + newShape2(1).Height)
                                                    newShape.Width = widthFaktor * newShape.Width
                                                    newShape2.Width = widthFaktor * newShape2.Width
                                                    newShape2.Top = newShape.Top + newShape.Height + 0.02 * height
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
                                        repObj1.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
                                        newShape = pptSlide.Shapes.Paste

                                        With newShape(1)
                                            .Top = top + 0.02 * height
                                            .Left = left + 0.02 * width
                                            .Width = width * 0.96
                                            topNext = top + 0.04 * height + .Height
                                            '.Height = height * 0.46
                                        End With

                                        repObj1.Delete()

                                        If Not repObj2 Is Nothing Then
                                            Try
                                                repObj2.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
                                                newShape2 = pptSlide.Shapes.Paste

                                                With newShape2(1)
                                                    .Top = topNext
                                                    .Left = left + 0.02 * width
                                                    .Width = width * 0.96
                                                    '.Height = height * 0.46
                                                End With

                                                repObj2.Delete()

                                                ' jetzt muss noch geschaut werden, ob die Shapes zu viele Höhe beanspruchen 
                                                If newShape(1).Height + newShape2(1).Height > 0.96 * height Then
                                                    widthFaktor = 0.96 * height / (newShape(1).Height + newShape2(1).Height)
                                                    newShape.Width = widthFaktor * newShape.Width
                                                    newShape2.Width = widthFaktor * newShape2.Width
                                                    newShape2.Top = newShape.Top + newShape.Height + 0.02 * height
                                                End If
                                            Catch ex As Exception

                                            End Try

                                        End If


                                    Catch ex As Exception

                                    End Try

                                End If


                            Case "Tabelle Projektziele"

                                Try
                                    Call zeichneProjektTabelleZiele(pptShape, hproj)

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



                                If qualifier = "letzter Stand" Then
                                    Call createProjektErgebnisCharakteristik2(lproj, obj, PThis.letzterStand)

                                ElseIf qualifier = "Beauftragung" Then
                                    Call createProjektErgebnisCharakteristik2(bproj, obj, PThis.beauftragung)

                                Else
                                    Call createProjektErgebnisCharakteristik2(hproj, obj, PThis.current)

                                End If



                                reportObj = obj

                                Dim ax As xlNS.Axis = reportObj.Chart.Axes(xlNS.XlAxisType.xlCategory)
                                With ax
                                    .TickLabels.Font.Size = 12
                                End With

                                notYetDone = True


                            Case "Strategie/Risiko"

                                Dim mycollection As New Collection

                                'deleteStack.Add(.Name, .Name)

                                mycollection.Add(pname)

                                'htop = topOfMagicBoard + hproj.tfZeile * boxHeight
                                'hleft = hproj.tfSpalte * boxWidth - 10
                                'hwidth = 12 * boxWidth
                                'hheight = 8 * boxHeight

                                Call awinCreatePortfolioDiagrams(mycollection, obj, True, PTpfdk.FitRisiko, 0, True, False, True, htop, hleft, hwidth, hheight)
                                reportObj = obj

                                notYetDone = True

                            Case "Personalbedarf"

                                'htop = 100
                                'hleft = 100
                                'hwidth = boxWidth * 14
                                'hheight = boxHeight * 10


                                auswahl = 1
                                Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)

                                reportObj = obj
                                notYetDone = True

                            Case "Personalkosten"

                                'htop = 100
                                'hleft = 100
                                'hwidth = boxWidth * 14
                                'hheight = boxHeight * 10

                                auswahl = 2
                                Call createRessPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)

                                reportObj = obj
                                notYetDone = True

                            Case "Sonstige Kosten"


                                'htop = 100
                                'hleft = 100
                                'hwidth = boxWidth * 14
                                'hheight = boxHeight * 10

                                Try
                                    auswahl = 1
                                    Call createCostPieOfProject(hproj, obj, auswahl, htop, hleft, hheight, hwidth)

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

                                ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                Dim vglBaseline As Boolean = True

                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 1, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Personalkosten" & ke
                                notYetDone = True

                            Case "Soll-Ist2 Personalkosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                Dim vglBaseline As Boolean = False


                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 1, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Personalkosten" & ke
                                notYetDone = True

                            Case "Soll-Ist1C Personalkosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                Dim vglBaseline As Boolean = True

                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 1, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Personalkosten" & ke
                                notYetDone = True

                            Case "Soll-Ist2C Personalkosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                Dim vglBaseline As Boolean = False

                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 1, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Personalkosten" & ke
                                notYetDone = True


                            Case "Soll-Ist1 Sonstige Kosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                Dim vglBaseline As Boolean = True

                                reportObj = Nothing
                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 2, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Sonstige Kosten" & ke
                                notYetDone = True

                            Case "Soll-Ist2 Sonstige Kosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                Dim vglBaseline As Boolean = False

                                reportObj = Nothing
                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 2, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Sonstige Kosten" & ke
                                notYetDone = True

                            Case "Soll-Ist1C Sonstige Kosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                Dim vglBaseline As Boolean = True


                                reportObj = Nothing
                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 2, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Sonstige Kosten" & ke
                                notYetDone = True

                            Case "Soll-Ist2C Sonstige Kosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der last freigabe verglichen
                                Dim vglBaseline As Boolean = False


                                reportObj = Nothing
                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 2, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Sonstige Kosten" & ke
                                notYetDone = True

                            Case "Soll-Ist1 Gesamtkosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                Dim vglBaseline As Boolean = True

                                reportObj = Nothing
                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 3, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Gesamtkosten" & ke
                                notYetDone = True

                            Case "Soll-Ist2 Gesamtkosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                Dim vglBaseline As Boolean = False

                                reportObj = Nothing
                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 3, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Gesamtkosten" & ke
                                notYetDone = True

                            Case "Soll-Ist1C Gesamtkosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Baseline verglichen
                                Dim vglBaseline As Boolean = True

                                reportObj = Nothing
                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 3, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Gesamtkosten" & ke
                                notYetDone = True

                            Case "Soll-Ist2C Gesamtkosten"

                                ' bei bereits beauftragten Projekten: es wird Current mit der last freigabe verglichen
                                Dim vglBaseline As Boolean = False

                                reportObj = Nothing
                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 3, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Gesamtkosten" & ke
                                notYetDone = True


                            Case "Soll-Ist1 Rolle"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Beauftragung verglichen
                                Dim vglBaseline As Boolean = True

                                reportObj = Nothing
                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 4, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Rolle " & qualifier & ze
                                notYetDone = True

                            Case "Soll-Ist2 Rolle"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                Dim vglBaseline As Boolean = False

                                reportObj = Nothing
                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 4, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Rolle " & qualifier & ze
                                notYetDone = True

                            Case "Soll-Ist1C Rolle"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Beauftragung verglichen
                                Dim vglBaseline As Boolean = True


                                reportObj = Nothing
                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 4, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Rolle " & qualifier & ze
                                notYetDone = True

                            Case "Soll-Ist2C Rolle"

                                ' bei bereits beauftragten Projekten: es wird Current mit der last freigabe verglichen
                                Dim vglBaseline As Boolean = False


                                reportObj = Nothing
                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 4, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Rolle " & qualifier & ze
                                notYetDone = True

                            Case "Soll-Ist1 Kostenart"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Beauftragung verglichen
                                Dim vglBaseline As Boolean = True

                                reportObj = Nothing
                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 5, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Kostenart " & qualifier & ke
                                notYetDone = True

                            Case "Soll-Ist2 Kostenart"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Last Freigabe verglichen
                                Dim vglBaseline As Boolean = False

                                reportObj = Nothing
                                Call createSollIstOfProject(hproj, reportObj, Date.Now, 5, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Kostenart " & qualifier & ke
                                notYetDone = True

                            Case "Soll-Ist1C Kostenart"

                                ' bei bereits beauftragten Projekten: es wird Current mit der Beauftragung verglichen
                                Dim vglBaseline As Boolean = True

                                reportObj = Nothing
                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 5, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Kostenart " & qualifier & ke
                                notYetDone = True

                            Case "Soll-Ist2C Kostenart"

                                ' bei bereits beauftragten Projekten: es wird Current mit der last freigabe verglichen
                                Dim vglBaseline As Boolean = False

                                reportObj = Nothing
                                Call createSollIstCurveOfProject(hproj, reportObj, Date.Now, 5, qualifier, vglBaseline, htop, hleft, hheight, hwidth)

                                boxName = "Kostenart " & qualifier & ke
                                notYetDone = True

                            Case "Ampel-Farbe"

                                Select Case hproj.ampelStatus
                                    Case 0
                                        .Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
                                    Case 1
                                        .Fill.ForeColor.RGB = awinSettings.AmpelGruen
                                    Case 2
                                        .Fill.ForeColor.RGB = awinSettings.AmpelGelb
                                    Case 3
                                        .Fill.ForeColor.RGB = awinSettings.AmpelRot
                                    Case Else
                                End Select

                            Case "Beschreibung"
                                .TextFrame2.TextRange.Text = hproj.ampelErlaeuterung

                            Case "Stand:"
                                .TextFrame2.TextRange.Text = boxName & " " & hproj.timeStamp.ToShortDateString

                            Case "Laufzeit:"
                                .TextFrame2.TextRange.Text = boxName & " " & textZeitraum(hproj.Start, hproj.Start + hproj.Dauer - 1)

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
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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

        Next





    End Sub
    '
    '
    ' 
    Public Sub createPPTSlidesFromConstellation(ByVal pptTemplate As String, ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs)
        Dim pptApp As pptNS.Application = Nothing
        Dim pptPresentation As pptNS.Presentation = Nothing
        Dim pptSlide As pptNS.Slide = Nothing
        Dim shapeRange As pptNS.ShapeRange = Nothing
        Dim presentationFile As String = awinPath & requirementsOrdner & "boarddossier.pptx"
        Dim pptShape As pptNS.Shape
        Dim portfolioName As String = currentConstellation
        Dim top As Double, left As Double, width As Double, height As Double
        Dim htop As Double, hleft As Double, hwidth As Double, hheight As Double
        Dim pptSize As Integer = 18
        'Dim hproj As clsProjekt
        'Dim pName As String
        'Dim auswahl As Integer
        Dim von As Integer, bis As Integer
        Dim myCollection As New Collection
        Dim notYetDone As Boolean = False
        Dim listofShapes As New Collection
        'Dim deleteStack As New Collection
        'Look for existing instance

        Try
            ' prüft, ob bereits Powerpoint geöffnet ist 
            pptApp = GetObject(, "PowerPoint.Application")
        Catch ex As Exception
            pptApp = CreateObject("PowerPoint.Application")
        End Try


        'frmSelectPPTTempl.statusNotification.Text = "PowerPoint nun geöffnet ...."
        e.Result = "PowerPoint ist nun geöffnet ...."
        worker.ReportProgress(0, e)

        ' entweder wird das template geöffnet ...
        ' oder aber es wird in die aktive Presentation geschrieben 

        If pptApp.Presentations.Count = 0 Then
            Try
                pptApp.Presentations.Open(presentationFile)
                pptPresentation = pptApp.ActivePresentation
            Catch ex As Exception
                pptPresentation = pptApp.Presentations.Add()
            End Try
        Else
            pptPresentation = pptApp.ActivePresentation
        End If
        ' wieviel Slides sind in der aktuellen Präsentataion
        Dim anzahlSlides As Integer
        ' wieviel Slides wurden aus der Vorlage hinzugefügt 
        Dim AnzAdded As Integer

        Try
            anzahlSlides = pptPresentation.Slides.Count
            AnzAdded = pptPresentation.Slides.InsertFromFile(pptTemplate, anzahlSlides)
        Catch ex As Exception
            Throw New Exception("Probleme mit Powerpoint Template")
        End Try


        Dim reportObj As xlNS.ChartObject = Nothing
        Dim obj As New Object
        Dim kennzeichnung As String = ""
        Dim qualifier As String = ""
        Dim anzShapes As Integer
        Dim tatsErstellt As Integer = 0

        For j = 1 To AnzAdded

            tatsErstellt = tatsErstellt + 1
            If worker.CancellationPending Then
                e.Cancel = True
                e.Result = "Berichterstellung nach " & tatsErstellt & " Seiten abgebrochen ..."
                'logMessage = "Berichterstellung nach " & tatsErstellt & " Reports abgebrochen ..."
                'Call logfileSchreiben(logMessage, " ")
                Exit For
            Else
                'frmSelectPPTTempl.statusNotification.Text = "Liste der Seiten aufgebaut ...."
                e.Result = "Bericht Seite " & tatsErstellt & " wird aufgebaut ...."
                worker.ReportProgress(0, e)
                pptSlide = pptPresentation.Slides(anzahlSlides + j)


                ' jetzt werden die Charts gezeichnet 
                anzShapes = pptSlide.Shapes.Count


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
                                tmpStr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {"(", ")"}, 3)
                                kennzeichnung = tmpStr(0).Trim
                            End If


                        Catch ex As Exception
                            kennzeichnung = "nicht identifizierbar"
                        End Try

                        If kennzeichnung = "Portfolio-Name" Or _
                            kennzeichnung = "Projekt-Tafel" Or _
                            kennzeichnung = "Projekt-Tafel Phasen" Or _
                            kennzeichnung = "Tabelle Zielerreichung" Or _
                            kennzeichnung = "Tabelle Projektstatus" Or _
                            kennzeichnung = "Übersicht Besser/Schlechter" Or _
                            kennzeichnung = "Tabelle Besser/Schlechter" Or _
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


                Dim newShape As pptNS.ShapeRange
                Dim boxName As String

                For Each tmpShape As pptNS.Shape In listofShapes

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
                                tmpStr = .TextFrame2.TextRange.Text.Trim.Split(New Char() {"(", ")"}, 3)
                                kennzeichnung = tmpStr(0).Trim
                                boxName = .TextFrame2.TextRange.Text
                                If tmpStr.Count > 1 Then
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
                        worker.ReportProgress(0, e)


                        reportObj = Nothing
                        top = .Top
                        left = .Left
                        height = .Height
                        width = .Width

                        Dim nameList As New SortedList(Of String, String)

                        Select Case kennzeichnung
                            Case "Portfolio-Name"
                                .TextFrame2.TextRange.Text = portfolioName


                            Case "Projekt-Tafel"

                                Dim farbtyp As Integer
                                Dim rng As xlNS.Range
                                Dim colorrng As xlNS.Range
                                Dim selectionType As Integer = -1 ' keine Einschränkung
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

                                    With appInstance.Worksheets(arrWsNames(3))
                                        rng = .range(.cells(1, minColumn), .cells(maxzeile, maxColumn))
                                        colorrng = .range(.cells(2, showRangeLeft), .cells(maxzeile, showRangeRight))

                                        Try
                                            colorrng.Interior.Color = showtimezone_color
                                        Catch ex1 As Exception

                                        End Try


                                        ' hier werden die Milestones gezeichnet 
                                        If qualifier = "Milestones R" Then
                                            Call awinDeleteMilestoneShapes(0)

                                            farbtyp = 3
                                            Call awinZeichneMilestones(nameList, farbtyp, True)



                                        ElseIf qualifier = "Milestones GR" Then
                                            Call awinDeleteMilestoneShapes(0)

                                            farbtyp = 2
                                            Call awinZeichneMilestones(nameList, farbtyp, False)
                                            farbtyp = 3
                                            Call awinZeichneMilestones(nameList, farbtyp, False)

                                        ElseIf qualifier = "Milestones GGR" Then
                                            Call awinDeleteMilestoneShapes(0)

                                            farbtyp = 1
                                            Call awinZeichneMilestones(nameList, farbtyp, False)
                                            farbtyp = 2
                                            Call awinZeichneMilestones(nameList, farbtyp, False)
                                            farbtyp = 3
                                            Call awinZeichneMilestones(nameList, farbtyp, False)

                                        ElseIf qualifier = "Milestones ALL" Then
                                            Call awinDeleteMilestoneShapes(0)

                                            farbtyp = 4
                                            Call awinZeichneMilestones(nameList, farbtyp, False)

                                        ElseIf qualifier = "Status" Then
                                            Call awinDeleteMilestoneShapes(0)
                                            Call awinZeichneStatus(True)
                                        End If


                                        rng.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
                                        colorrng.Interior.ColorIndex = -4142

                                        ' lösche alle Milestones wieder 
                                        If qualifier <> "" Then
                                            Call awinDeleteMilestoneShapes(0)
                                        End If
                                    End With


                                    ' set back 
                                    With appInstance.ActiveWindow
                                        .GridlineColor = RGB(220, 220, 220)
                                    End With


                                    newShape = pptSlide.Shapes.Paste
                                    Dim ratio As Double

                                    With newShape.Item(1)
                                        ratio = height / width
                                        If ratio < .Height / .Width Then
                                            ' orientieren an width 
                                            .Width = width * 0.96
                                            .Height = ratio * .Width
                                            ' left anpassen
                                            .Top = top + 0.02 * height
                                            .Left = left + 0.98 * (width - .Width) / 2

                                        Else
                                            .Height = height * 0.96
                                            .Width = .Height / ratio
                                            ' top anpassen 
                                            .Left = left + 0.02 * width
                                            .Top = top + 0.98 * (height - .Height) / 2
                                        End If

                                    End With


                                Else
                                    .TextFrame2.TextRange.Text = "Keine Projekte im angegebenen Zeitraum vorhanden"
                                End If


                            Case "Projekt-Tafel Phasen"

                                Dim rng As xlNS.Range
                                Dim colorrng As xlNS.Range
                                Dim selectionType As Integer = -1 ' keine Einschränkung
                                Dim ok As Boolean = True

                                von = showRangeLeft
                                bis = showRangeRight
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

                                    With appInstance.Worksheets(arrWsNames(3))
                                        rng = .range(.cells(1, minColumn), .cells(maxzeile, maxColumn))
                                        colorrng = .range(.cells(2, showRangeLeft), .cells(maxzeile, showRangeRight))

                                        Try
                                            colorrng.Interior.Color = showtimezone_color
                                        Catch ex1 As Exception

                                        End Try

                                        ' hier werden die Phasen gezeichnet 
                                        Call awinDeleteMilestoneShapes(0)

                                        Dim qstr(20) As String
                                        Dim phNameCollection As New Collection
                                        Dim phName As String = " "
                                        qstr = qualifier.Trim.Split(New Char() {"#"}, 18)

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

                                            Call awinZeichnePhasen(phNameCollection, False)
                                            rng.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)
                                            colorrng.Interior.ColorIndex = -4142

                                            ' lösche alle Phase Shapes wieder wieder 
                                            If qualifier <> "" Then
                                                Call awinDeleteMilestoneShapes(0)
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
                                        newShape = pptSlide.Shapes.Paste
                                        Dim ratio As Double

                                        With newShape
                                            ratio = height / width
                                            If ratio < .Height / .Width Then
                                                ' orientieren an width 
                                                .Width = width * 0.96
                                                .Height = ratio * .Width
                                                ' left anpassen
                                                .Top = top + 0.02 * height
                                                .Left = left + 0.98 * (width - .Width) / 2

                                            Else
                                                .Height = height * 0.96
                                                .Width = .Height / ratio
                                                ' top anpassen 
                                                .Left = left + 0.02 * width
                                                .Top = top + 0.98 * (height - .Height) / 2
                                            End If

                                        End With
                                    Else
                                        .TextFrame2.TextRange.Text = "es konnten keine Phasen erkannt werden ... "
                                    End If


                                Else
                                    .TextFrame2.TextRange.Text = "Keine Projekte im angegebenen Zeitraum vorhanden"
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
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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
                                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisiko, 0, False, True, True, htop, hleft, hwidth, hheight)


                                reportObj = obj

                                With reportObj
                                    .Chart.ChartTitle.Text = boxName
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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

                                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.FitRisikoVol, 0, False, True, True, htop, hleft, hwidth, hheight)

                                reportObj = obj

                                With reportObj
                                    .Chart.ChartTitle.Text = boxName
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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

                                Call awinCreatePortfolioDiagrams(myCollection, obj, False, PTpfdk.ComplexRisiko, 0, False, True, True, htop, hleft, hwidth, hheight)
                                'Call awinCreateZeitRiskVolumeDiagramm(myCollection, obj, False, False, True, True, htop, hleft, hwidth, hheight)

                                reportObj = obj

                                With reportObj
                                    .Chart.ChartTitle.Text = boxName
                                    .Chart.ChartTitle.Font.Size = pptSize
                                End With

                                reportObj.Copy()
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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
                                qstr = qualifier.Trim.Split(New Char() {"#"}, 18)

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
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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
                                newShape = pptSlide.Shapes.Paste

                                With newShape
                                    .Top = top + 0.02 * height
                                    .Left = left + 0.02 * width
                                    .Width = width * 0.96
                                    .Height = height * 0.96
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
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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

                                Dim qstr(20) As String
                                Dim phName As String = " "
                                qstr = qualifier.Trim.Split(New Char() {"#"}, 18)

                                ' Aufbau der Collection 
                                For i = 0 To qstr.Length - 1

                                    Try
                                        phName = qstr(i).Trim
                                        If PhaseDefinitions.Contains(phName) Then
                                            myCollection.Add(phName, phName)
                                        End If
                                    Catch ex As Exception
                                        Call MsgBox("Fehler: Phasen Name " & phName & " konnte nicht erkannt werden ...")
                                    End Try

                                Next


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
                                        Else
                                            .Chart.ChartTitle.Text = boxName
                                        End If

                                        .Chart.ChartTitle.Font.Size = pptSize
                                    End With

                                    reportObj.Copy()
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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
                                Dim MSnameList As SortedList(Of String, String)
                                MSnameList = ShowProjekte.getMilestoneNames

                                Dim qstr(20) As String
                                Dim msName As String = " "
                                qstr = qualifier.Trim.Split(New Char() {"#"}, 18)

                                ' Aufbau der Collection 
                                For i = 0 To qstr.Length - 1

                                    Try
                                        msName = qstr(i).Trim
                                        If MSnameList.ContainsKey(msName) Then
                                            myCollection.Add(msName, msName)
                                        End If
                                    Catch ex As Exception
                                        Call MsgBox("Fehler: Phasen Name " & msName & " konnte nicht erkannt werden ...")
                                    End Try

                                Next


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
                                        Else
                                            .Chart.ChartTitle.Text = boxName
                                        End If

                                        .Chart.ChartTitle.Font.Size = pptSize
                                    End With

                                    reportObj.Copy()
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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

                                Dim qstr(20) As String
                                Dim roleName As String = " "
                                qstr = qualifier.Trim.Split(New Char() {"#"}, 18)

                                ' Aufbau der Collection 
                                For i = 0 To qstr.Length - 1

                                    Try
                                        roleName = qstr(i).Trim
                                        If RoleDefinitions.Contains(roleName) Then
                                            myCollection.Add(roleName, roleName)
                                        End If
                                    Catch ex As Exception
                                        Call MsgBox("Fehler: Rolle " & roleName & " konnte nicht erkannt werden ...")
                                    End Try

                                Next


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
                                    ' jetzt wird die Überschrift neu bestimmt ...
                                    With reportObj
                                        Dim tmpTitle() As String = .Chart.ChartTitle.Text.Split(New Char() {"(", ")"}, 3)
                                        Try
                                            .Chart.ChartTitle.Text = qualifier & " (" & tmpTitle(1) & ")"
                                        Catch ex As Exception
                                            .Chart.ChartTitle.Text = qualifier
                                        End Try
                                        .Chart.ChartTitle.Font.Size = pptSize
                                    End With

                                    reportObj.Copy()
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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
                                Dim qstr(20) As String
                                Dim costName As String = " "
                                qstr = qualifier.Trim.Split(New Char() {"#"}, 18)

                                ' Aufbau der Collection 
                                For i = 0 To qstr.Length - 1

                                    Try
                                        costName = qstr(i).Trim
                                        If CostDefinitions.Contains(costName) Then
                                            myCollection.Add(costName, costName)
                                        End If
                                    Catch ex As Exception
                                        Call MsgBox("Fehler: Kostenart " & costName & " konnte nicht erkannt werden ...")
                                    End Try

                                Next


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
                                        Dim tmpTitle() As String = .Chart.ChartTitle.Text.Split(New Char() {"(", ")"}, 3)
                                        Try
                                            .Chart.ChartTitle.Text = qualifier & " (" & tmpTitle(1) & ")"
                                        Catch ex As Exception
                                            .Chart.ChartTitle.Text = qualifier
                                        End Try
                                        .Chart.ChartTitle.Font.Size = pptSize
                                    End With

                                    reportObj.Copy()
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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
                                .TextFrame2.TextRange.Text = boxName & " " & Date.Now.ToString("d")


                            Case "Zeitraum:"
                                .TextFrame2.TextRange.Text = boxName & " " & textZeitraum(showRangeLeft, showRangeRight)

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
                                    newShape = pptSlide.Shapes.Paste

                                    With newShape
                                        .Top = top + 0.02 * height
                                        .Left = left + 0.02 * width
                                        .Width = width * 0.96
                                        .Height = height * 0.96
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
            End If
            listofShapes.Clear()

        Next

        ' pptTemplate muss noch geschlossen werden

        If tatsErstellt = 1 Then
            e.Result = " Report mit " & tatsErstellt & " Seite erstellt !"
        Else
            e.Result = " Report mit " & tatsErstellt & " Seiten erstellt !"
        End If

        worker.ReportProgress(0, e)
        'frmSelectPPTTempl.statusNotification.Text = " Report mit " & tatsErstellt & " Seite erstellt !"


    End Sub



    Public Sub StoreAllProjectsinDB()

        Dim jetzt As Date = Now
        Dim zeitStempel As Date
        Dim request As New Request(awinSettings.databaseName)
        enableOnUpdate = False

        ' die aktuelle Konstellation wird unter dem Namen <Last> gespeichert ..
        Call awinStoreConstellation("Last")

        If request.pingMongoDb() Then

            Try
                ' jetzt werden die gezeigten Projekte in die Datenbank geschrieben 

                For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte

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

                zeitStempel = AlleProjekte.First.Value.timeStamp

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

    Public Function StoreSelectedProjectsinDB()

        Dim singleShp1 As Excel.Shape
        Dim hproj As clsProjekt
        Dim jetzt As Date = Now
        Dim zeitStempel As Date
        Dim anzSelectedProj As Integer = 0
        Dim anzStoredProj As Integer = 0

        Dim request As New Request(awinSettings.databaseName)

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
                        hproj = ShowProjekte.getProject(singleShp1.Name)
                    Catch ex As Exception
                        Throw New ArgumentException("Projekt nicht gefunden ...")
                        enableOnUpdate = True
                    End Try

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

                Next i

            Else
                'Call MsgBox("Es wurde kein Projekt selektiert")
                ' die Anzahl selektierter und auch gespeicherter Projekte ist damit = 0
                anzStoredProj = anzSelectedProj
                Return anzSelectedProj
            End If



            'historicDate = historicDate.AddMonths(1)

            '' jetzt werden alle definierten Constellations weggeschrieben

            'For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            '    Try
            '        If request.storeConstellationToDB(kvp.Value) Then
            '        Else
            '            Call MsgBox("Fehler in Schreiben Constellation " & kvp.Key)
            '        End If
            '    Catch ex As Exception
            '        Throw New ArgumentException("Fehler beim Speichern der Portfolios in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
            '        'Call MsgBox("Fehler beim Speichern der ProjekteConstellationen in die Datenbank. Datenbank nicht aktiviert?")
            '        'Exit Sub
            '    End Try

            'Next


            '' jetzt werden alle Abhängigkeiten weggeschreiben 

            'For Each kvp As KeyValuePair(Of String, clsDependenciesOfP) In allDependencies.getSortedList

            '    Try
            '        If request.storeDependencyofPToDB(kvp.Value) Then
            '        Else
            '            Call MsgBox("Fehler in Schreiben Dependency " & kvp.Key)
            '        End If
            '    Catch ex As Exception
            '        Throw New ArgumentException("Fehler beim Speichern der Abhängigkeiten in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
            '        'Call MsgBox("Fehler beim Speichern der Abhängigkeiten in die Datenbank. Datenbank nicht aktiviert?")
            '        'Exit Sub
            '    End Try


            'Next

            'zeitStempel = AlleProjekte.First.Value.timeStamp

            'Call MsgBox("ok, gespeichert!" & vbLf & zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString)

            '' Änderung 18.6 - wenn gespeichert wird, soll die Projekthistorie zurückgesetzt werden 
            'Try
            '    If projekthistorie.Count > 0 Then
            '        projekthistorie.clear()
            '    End If
            'Catch ex As Exception

            'End Try

        Else

            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen")

        End If


        enableOnUpdate = True

        zeitStempel = AlleProjekte.First.Value.timeStamp

        Call MsgBox("ok, " & anzStoredProj & " Projekte gespeichert!" & vbLf & zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString)
        Return anzStoredProj

    End Function


    Public Function RemoveSelectedProjectsfromDB()


        Dim hproj As New clsProjekt
        Dim jetzt As Date = Date.Now
        'Dim zeitStempel As Date
        Dim anzSelectedProj As Integer = 0
        Dim anzDeletedProj As Integer = 0
        Dim anzDeletedTS As Integer = 0
        Dim anzElements As Integer
        Dim found As Boolean = False
        Dim iSel As Integer = 0

        Dim selCollection As SortedList(Of Date, String)
        enableOnUpdate = False
        Dim tmpstr(4) As String

        Dim request As New Request(awinSettings.databaseName)
        Dim requestTrash As New Request(awinSettings.databaseName & "Trash")

        If request.pingMongoDb() Then

            If selectedToDelete.Count > 0 Then

                anzSelectedProj = selectedToDelete.Count


                For Each kvpSelToDel As KeyValuePair(Of String, SortedList(Of Date, String)) In selectedToDelete.Liste

                    selCollection = selectedToDelete.getTimeStamps(kvpSelToDel.Key)
                    anzElements = selCollection.Count

                    'If AlleProjekte.ContainsKey(kvpSelToDel.Key) Then
                    '    ' Projekt ist bereits im Hauptspeicher geladen
                    '    hproj = AlleProjekte(kvpSelToDel.Key)
                    'End If

                    If Not projekthistorie Is Nothing Then
                        projekthistorie.clear() ' alte Historie löschen
                    End If

                    'tmpstr = title.Trim.Split(New Char() {"#"}, 4)
                    tmpstr = kvpSelToDel.Key.Trim.Split(New Char() {"#"}, 4)   ' Projektnamen aus key separieren

                    projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=tmpstr(0), variantName:="", storedEarliest:=Date.MinValue, storedLatest:=Date.Now)

                    anzDeletedTS = 0    ' Anzahl gelöschter TimeStamps dieses Projekts

                    For i = 1 To anzElements  ' Schleife über die zu löschenden TimeStamps dieses Projekts

                        'Dim ms As Long = selCollection.ElementAt(i - 1).Key.Millisecond

                        found = False
                        iSel = 0

                        While Not found
                            hproj = projekthistorie.ElementAt(iSel)
                            If hproj.timeStamp = selCollection.ElementAt(i - 1).Key Then
                                found = True
                            End If
                            iSel = iSel + 1
                        End While

                        If requestTrash.storeProjectToDB(hproj) Then

                            If request.deleteProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                                                                         storedEarliest:=selCollection.ElementAt(i - 1).Key, storedLatest:=selCollection.ElementAt(i - 1).Key) Then
                                anzDeletedTS = anzDeletedTS + 1

                            Else
                                Call MsgBox("Fehler beim Löschen von " & hproj.name)
                            End If

                        Else
                            Call MsgBox("Fehler beim Speichern von " & hproj.name & " im Papierkorb")
                        End If

                    Next i      'nächsten TimeStamp holen


                    Call MsgBox("ok, " & anzDeletedTS & " TimeStamps zu Projekt " & hproj.name & " gelöscht")

                    If Not request.projectNameAlreadyExists(hproj.name, hproj.variantName) Then
                        If AlleProjekte.ContainsKey(hproj.name & "#" & hproj.variantName) Then
                            AlleProjekte.Remove(hproj.name & "#" & hproj.variantName)
                            Try
                                ShowProjekte.Remove(hproj.name)
                            Catch ex As Exception
                            End Try
                        End If
                    End If

                    anzDeletedProj = anzDeletedProj + 1

                Next

                'projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                '                                                 storedEarliest:=StartofCalendar, storedLatest:=Date.Now)

                'For Each kvpHist As KeyValuePair(Of Date, clsProjekt) In projekthistorie.liste

                '    If kvpHist.Value.timeStamp = kvpSelToDel.Value.timeStamp Then
                '        If requestTrash.storeProjectToDB(kvpHist.Value) Then

                '            If request.deleteProjectHistoryFromDB(projectname:=hproj.name, variantName:=hproj.variantName, _
                '                                                 storedEarliest:=kvpHist.Value.timeStamp, storedLatest:=kvpHist.Value.timeStamp) Then
                '                anzDeleted = anzDeleted + 1
                '                'Call MsgBox("ok, Projekt '" & hproj.name & "' gespeichert!" & vbLf & hproj.timeStamp.ToShortDateString)

                '            Else
                '                Call MsgBox("Fehler beim Löschen von Projekt " & kvpSelToDel.Value.name & vbLf & kvpHist.Value.timeStamp.ToShortDateString)
                '            End If

                '        Else

                '            Call MsgBox("Fehler in Löschen von Projekt " & hproj.name)
                '        End If
                '    Else
                '        ' Es ist nicht der richtige TimeStamp von hproj.name

                '    End If


                'Next kvpHist

                '    anzDeletedProj = anzDeletedProj + 1
                '    'Call MsgBox("ok, Projekt '" & hproj.name & "' gelöscht!" & vbLf & hproj.timeStamp.ToShortDateString)
                'End If

                '    Catch ex As Exception

                '    ' Call MsgBox("Fehler beim Speichern der Projekte in die Datenbank. Datenbank nicht aktiviert?")
                '    Throw New ArgumentException("Fehler beim Löschen der Projekte in die Datenbank." & vbLf & "Datenbank ist vermutlich nicht aktiviert?")
                '    'Exit Sub
                'End Try


            Else
                'Call MsgBox("Es wurde kein Projekt selektiert")
                ' die Anzahl selektierter und auch gespeicherter Projekte ist damit = 0
                anzDeletedProj = anzSelectedProj
                Return anzDeletedProj
            End If

        Else

            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen")

        End If


        enableOnUpdate = True

        Return anzDeletedProj

    End Function

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
    Sub awinCreateStatusDiagram1(ByRef ProjektListe As Collection, ByRef repChart As Object, ByVal compareToID As Integer, _
                                         ByVal auswahl As Integer, ByVal qualifier As String, _
                                         ByVal showLabels As Boolean, ByVal chartBorderVisible As Boolean, _
                                         ByVal top As Double, ByVal left As Double, ByVal width As Double, ByVal height As Double)

        Dim request As New Request(awinSettings.databaseName)
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

            pname = ProjektListe.Item(i)
            Try
                hproj = ShowProjekte.getProject(pname)
                variantName = hproj.variantName

                If Not projekthistorie Is Nothing Then
                    If projekthistorie.Count > 0 Then
                        vglName = projekthistorie.First.name
                    End If
                End If


                If vglName.Trim <> pname.Trim Then
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

                    ReDim werteH(hproj.Dauer - 1)
                    ReDim werteV(vglProj.Dauer - 1)
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



            With appInstance.Worksheets(arrWsNames(3))
                anzDiagrams = .ChartObjects.Count
                '
                ' um welches Diagramm handelt es sich ...
                '
                i = 1
                found = False
                While i <= anzDiagrams And Not found

                    Try
                        chtTitle = .ChartObjects(i).Chart.ChartTitle.text
                    Catch ex As Exception
                        chtTitle = " "
                    End Try

                    If chtTitle Like ("*" & diagramTitle & "*") Then
                        found = True
                        repChart = .ChartObjects(i)
                        Exit Sub
                    Else
                        i = i + 1
                    End If
                End While


                ReDim tempArray(anzBubbles - 1)


                With appInstance.Charts.Add

                    .SeriesCollection.NewSeries()
                    .SeriesCollection(1).name = diagramTitle
                    .SeriesCollection(1).ChartType = xlNS.XlChartType.xlXYScatter

                    For i = 1 To anzBubbles
                        tempArray(i - 1) = formerValues(i - 1)
                    Next i
                    .SeriesCollection(1).XValues = tempArray ' strategic

                    For i = 1 To anzBubbles
                        tempArray(i - 1) = currentValues(i - 1)
                    Next i
                    .SeriesCollection(1).Values = tempArray




                    'Dim series1 As xlNS.Series = _
                    '        CType(.SeriesCollection(1),  _
                    '                xlNS.Series)
                    'Dim point1 As xlNS.Point = _
                    '            CType(series1.Points(1), xlNS.Point)

                    'Dim testName As String
                    For i = 1 To anzBubbles

                        With .SeriesCollection(1).Points(i)

                            If showLabels Then
                                Try
                                    .HasDataLabel = True
                                    With .DataLabel
                                        .text = nameValues(i - 1)
                                        If singleProject Then
                                            .font.size = awinSettings.CPfontsizeItems + 4
                                        Else
                                            .font.size = awinSettings.CPfontsizeItems
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

                            .Interior.color = colorValues(i - 1)
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

                    With .Axes(xlNS.XlAxisType.xlCategory)
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
                            .Characters.text = "geplant"
                            .Characters.Font.Size = titlefontsize
                            .Characters.Font.Bold = False
                        End With
                        With .TickLabels.Font
                            .FontStyle = "Normal"
                            .Bold = True
                            .Size = awinSettings.fontsizeItems

                        End With

                    End With


                    With .Axes(xlNS.XlAxisType.xlValue)
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
                            .Characters.text = "tatsächlich"
                            .Characters.Font.Size = titlefontsize
                            .Characters.Font.Bold = False
                        End With

                        With .TickLabels.Font
                            .FontStyle = "Normal"
                            .bold = True
                            .Size = awinSettings.fontsizeItems
                        End With
                    End With
                    .HasLegend = False
                    .HasTitle = True
                    .ChartTitle.text = diagramTitle
                    .ChartTitle.Characters.Font.Size = awinSettings.fontsizeTitle
                    .Location(Where:=xlNS.XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                End With


                'appInstance.ShowChartTipNames = False
                'appInstance.ShowChartTipValues = False

                With .ChartObjects(anzDiagrams + 1)
                    .top = top
                    .left = left
                    .width = width
                    .height = height
                    .name = chtobjName
                End With



                With appInstance.ActiveSheet
                    Try
                        With appInstance.ActiveSheet
                            .Shapes(chtobjName).line.visible = chartBorderVisible
                        End With
                    Catch ex As Exception

                    End Try
                End With

                pfDiagram = New clsDiagramm

                'pfChart = New clsAwinEvent
                pfChart = New clsEventsPfCharts
                pfChart.PfChartEvents = .ChartObjects(anzDiagrams + 1).Chart

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

                repChart = .ChartObjects(anzDiagrams + 1)

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

        Dim request As New Request(awinSettings.databaseName)
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
            ReDim currentValues(hproj.Dauer - 1)

        Catch ex As Exception

            statusValue = 1.0
            statusColor = awinSettings.AmpelGruen
            Exit Sub

        End Try


        If Not projekthistorie Is Nothing Then
            If projekthistorie.Count > 0 Then
                vglName = projekthistorie.First.name
            End If
        End If


        If vglName.Trim <> pname.Trim Then
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


            ReDim formerValues(vglProj.Dauer - 1)
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
                pptShape.TextFrame2.TextRange.Text = boxName & "nicht vorhanden"
            End If

        Else
            pptShape.TextFrame2.TextRange.Text = "es gibt keine laufenden Projekte im betrachteten Zeitraum ... "
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
            farbTypenListe.Add(tmpfarbe, tmpfarbe)
            tmpfarbe = 3
            farbTypenListe.Add(tmpfarbe, tmpfarbe)
        End If

        Dim todoListe As New SortedList(Of Long, clsProjekt)
        Dim key As Long

        Dim selectionType As Integer = -1
        timeFrameProjekte = ShowProjekte.withinTimeFrame(selectionType, showRangeLeft, showRangeRight)

        For Each pname As String In timeFrameProjekte
            hproj = ShowProjekte.getProject(pname)
            key = 10000 * hproj.tfZeile + hproj.Start
            todoListe.Add(key, hproj)
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

                For r = 1 To cphase.CountResults
                    Dim cResult As clsResult
                    Dim cBewertung As clsBewertung

                    cResult = cphase.getResult(r)

                    cBewertung = cResult.getBewertung(1)

                    resultColumn = getColumnOfDate(cResult.getDate)

                    If farbTypenListe.Contains(cBewertung.colorIndex) Then
                        ' dann muss ein Eintrag in der Tabelle gemacht werden 

                        If (resultColumn < showRangeLeft Or resultColumn > showRangeRight) Then
                            ' nichts machen 
                        Else
                            ' hier die Tabellen-Einträge machen 

                            With tabelle

                                CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = msNumber.ToString
                                CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                                CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = cBewertung.color
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

    Sub zeichneProjektGrafik(ByRef pptslide As pptNS.Slide, ByRef pptShape As pptNS.Shape, ByVal hproj As clsProjekt)

        Dim rng As xlNS.Range
        Dim selectionType As Integer = -1 ' keine Einschränkung
        Dim pptSize As Integer
        Dim newshape As pptNS.ShapeRange


        pptSize = pptShape.TextFrame2.TextRange.Font.Size
        pptShape.TextFrame2.TextRange.Text = " "

        Dim minColumn As Integer, maxColumn As Integer
        minColumn = hproj.Start - 3
        If minColumn < 1 Then
            minColumn = 1
        End If

        maxColumn = hproj.Start + hproj.Dauer + 3

        ' set Gridlines to white 
        With appInstance.ActiveWindow
            .GridlineColor = RGB(255, 255, 255)
        End With

        Dim oldposition As Integer = hproj.tfZeile
        Dim projektShape As xlNS.Shape
        Dim allShapes As xlNS.Shapes
        Dim ptop As Double, pleft As Double, pwidth As Double, pheight As Double
        Dim number As Integer = 1
        Dim nameList As New SortedList(Of String, String)

        Call awinDeleteMilestoneShapes(0)

        With CType(appInstance.Worksheets(arrWsNames(3)), xlNS.Worksheet)

            allShapes = .Shapes
            projektShape = allShapes.Item(hproj.name)

            ' Projekt-Shape wird jetzt in neue Zeile geschoben 
            Dim newzeile As Integer = ShowProjekte.maxZeile + 4
            hproj.tfZeile = newzeile

            hproj.CalculateShapeCoord(ptop, pleft, pwidth, pheight)
            With projektShape
                .Top = ptop
                .Left = pleft
                .Height = pheight
                .Width = pwidth
            End With

            Call zeichneStatusSymbolInPlantafel(hproj, 0)
            Call zeichneResultMilestonesInProjekt(hproj, nameList, 4, False, True, number, True)


            rng = .Range(.Cells(newzeile, minColumn), .Cells(newzeile + 1, maxColumn))
            rng.CopyPicture(Microsoft.Office.Interop.Excel.XlPictureAppearance.xlScreen)


            Call awinDeleteMilestoneShapes(0)

            ' Shape wieder an die alte Position bringen 
            hproj.tfZeile = oldposition
            hproj.CalculateShapeCoord(ptop, pleft, pwidth, pheight)
            With projektShape
                .Top = ptop
                .Left = pleft
                .Height = pheight
                .Width = pwidth
            End With

        End With

        ' set back 
        With appInstance.ActiveWindow
            .GridlineColor = RGB(220, 220, 220)
        End With


        newshape = pptslide.Shapes.Paste
        Dim ratio As Double

        With newshape
            ratio = pptShape.Height / pptShape.Width
            If ratio < .Height / .Width Then
                ' orientieren an width 
                .Width = pptShape.Width * 0.96
                .Height = ratio * .Width
                ' left anpassen
                .Top = pptShape.Top + 0.02 * pptShape.Height
                .Left = pptShape.Left + 0.98 * (pptShape.Width - .Width) / 2

            Else
                .Height = pptShape.Height * 0.96
                .Width = .Height / ratio
                ' top anpassen 
                .Left = pptShape.Left + 0.02 * pptShape.Width
                .Top = pptShape.Top + 0.98 * (pptShape.Height - .Height) / 2
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
            tmpstr = title.Trim.Split(New Char() {"#"}, 4)
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
            tmpstr = title.Trim.Split(New Char() {"#"}, 4)
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
            tmpstr = title.Trim.Split(New Char() {"#"}, 4)
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
                Dim bdiff As Integer, ldiff As Integer
                Dim bphaseStart As Date
                Dim lphaseStart As Date


                Try
                    bphase = bproj.getPhase(cphase.name)
                    bphaseStart = bproj.startDate.AddMonths(bphase.relStart - 1)
                Catch ex As Exception
                    bphase = Nothing
                End Try

                Try
                    lphase = lproj.getPhase(cphase.name)
                    lphaseStart = lproj.startDate.AddMonths(lphase.relStart - 1)
                Catch ex As Exception
                    lphase = Nothing
                End Try



                For r = 1 To cphase.CountResults
                    Dim cResult As clsResult = Nothing
                    Dim cBewertung As clsBewertung = Nothing

                    Dim bResult As clsResult = Nothing
                    Dim bbewertung As clsBewertung = Nothing


                    Dim lResult As clsResult = Nothing
                    Dim lbewertung As clsBewertung = Nothing
                    Dim bDate As Date, lDate As Date
                    Dim currentDate As Date

                    cResult = cphase.getResult(r)
                    currentDate = cResult.getDate

                    If IsNothing(bphase) Then
                    Else

                    End If

                    bResult = bphase.getResult(cResult.name)
                    If IsNothing(bResult) Then
                        bdiff = -9999
                    Else
                        bDate = bResult.getDate
                        bdiff = DateDiff(DateInterval.Day, bDate, currentDate)
                    End If


                    lResult = lphase.getResult(cResult.name)
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
                            CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = cBewertung.color

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
                                CType(.Cell(tabellenzeile, 4), pptNS.Cell).Shape.Fill.ForeColor.RGB = bbewertung.color
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 4), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
                            End Try


                            ' Datum und Farbe für letzter Stand schreiben  
                            Try

                                CType(.Cell(tabellenzeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = lDate.ToShortDateString
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 5), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "---"
                            End Try

                            Try
                                CType(.Cell(tabellenzeile, 6), pptNS.Cell).Shape.Fill.ForeColor.RGB = lbewertung.color
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 6), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
                            End Try

                            ' Datum und Farbe für aktuellen Stand schreiben  
                            Try
                                CType(.Cell(tabellenzeile, 7), pptNS.Cell).Shape.TextFrame2.TextRange.Text = currentDate.ToShortDateString
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 7), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "---"
                            End Try

                            Try
                                CType(.Cell(tabellenzeile, 8), pptNS.Cell).Shape.Fill.ForeColor.RGB = cBewertung.color
                            Catch ex As Exception
                                CType(.Cell(tabellenzeile, 8), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
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

        If pptShape.HasTable Then
            tabelle = pptShape.Table
            anzZeilen = tabelle.Rows.Count
            If anzZeilen > 1 Then
                zeile = 1
                ' jetzt wird die Überschrift aktualisiert 
                With tabelle

                    CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Projekt" & vbLf & hproj.name

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
                                    If unterschiede.Contains(PThcc.phasen) Then
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
                                    If unterschiede.Contains(PThcc.resultdates) Or unterschiede.Contains(PThcc.resultampel) Then

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

        If pptShape.HasTable Then
            tabelle = pptShape.Table
            anzZeilen = tabelle.Rows.Count
            If anzZeilen > 1 Then
                zeile = 1
                ' jetzt wird die Überschrift aktualisiert 
                With tabelle

                    CType(.Cell(zeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = "Projekt" & vbLf & hproj.name

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
                                    If unterschiede.Contains(PThcc.phasen) Or unterschiede.Contains(PThcc.resultdates) Then
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
                .Width = korrFaktor * .Width
                .Height = korrFaktor * .Height
            End With

        End If

        ' jetzt bestimmen der Left , Top Koordinaten des Pfeils und setzen der Farbe

        With newZeichen(1)

            .Top = tabelle.Cell(tbZeile, tbSpalte).Shape.Top + (tabelle.Cell(tbZeile, tbSpalte).Shape.Height - .Height) / 2
            .Left = tabelle.Cell(tbZeile, tbSpalte).Shape.Left + (tabelle.Cell(tbZeile, tbSpalte).Shape.Width - .Width) / 2
            .Fill.ForeColor.RGB = farbkennung

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
                .Width = korrFaktor * .Width
                .Height = korrFaktor * .Height
            End With

        End If

        ' jetzt bestimmen der Left , Top Koordinaten des Pfeils und setzen der Farbe

        With newZeichen(1)

            .Top = tabelle.Cell(tbZeile, tbSpalte).Shape.Top + (tabelle.Cell(tbZeile, tbSpalte).Shape.Height - .Height) / 2
            .Left = tabelle.Cell(tbZeile, tbSpalte).Shape.Left + (tabelle.Cell(tbZeile, tbSpalte).Shape.Width - .Width) / 2
            .Fill.ForeColor.RGB = farbkennung
            .Line.ForeColor.RGB = lineColor
            .Line.Weight = 2

        End With



    End Sub

    Sub zeichneProjektTabelleZiele(ByRef pptShape As pptNS.Shape, ByVal hproj As clsProjekt)

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

        ' jetzt wird die todoListe abgearbeitet 
        Dim tabellenzeile As Integer = 2

        Try
            For p = 1 To hproj.CountPhases

                Dim cphase As clsPhase = hproj.getPhase(p)
                Dim phaseStart As Date = hproj.startDate.AddMonths(cphase.relStart - 1)

                For r = 1 To cphase.CountResults
                    Dim cResult As clsResult
                    Dim cBewertung As clsBewertung

                    cResult = cphase.getResult(r)
                    cBewertung = cResult.getBewertung(1)

                    'Try
                    '    cBewertung = cResult.getBewertung(1)
                    'Catch ex As Exception
                    '    cBewertung = New clsBewertung
                    'End Try


                    With tabelle

                        CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Text = msNumber.ToString
                        CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                        CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = cBewertung.color

                        CType(.Cell(tabellenzeile, 2), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cResult.name
                        CType(.Cell(tabellenzeile, 3), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cResult.getDate.ToShortDateString
                        CType(.Cell(tabellenzeile, 4), pptNS.Cell).Shape.TextFrame2.TextRange.Text = cBewertung.description

                    End With

                    msNumber = msNumber + 1
                    tabelle.Rows.Add()
                    tabellenzeile = tabellenzeile + 1

                Next

            Next

            Try
                tabelle.Rows(msNumber + 1).Delete()
            Catch ex1 As Exception

            End Try

        Catch ex As Exception
            Throw New Exception("Tabelle Projektziele hat evtl unzulässige Anzahl Zeilen / Spalten ...")
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
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
                ElseIf kvp.Value.ampelStatus = 1 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelGruen
                ElseIf kvp.Value.ampelStatus = 2 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelGelb
                ElseIf kvp.Value.ampelStatus = 3 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelRot
                Else
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
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
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
                ElseIf hproj.ampelStatus = 1 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelGruen
                ElseIf hproj.ampelStatus = 2 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelGelb
                ElseIf hproj.ampelStatus = 3 Then
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelRot
                Else
                    CType(.Cell(tabellenzeile, 1), pptNS.Cell).Shape.Fill.ForeColor.RGB = awinSettings.AmpelNichtBewertet
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
                hproj = ShowProjekte.getProject(pName)
                With hproj
                    If .Start < minColumn Then
                        minColumn = .Start
                    End If

                    If .Start + .Dauer - 1 > maxColumn Then
                        maxColumn = .Start + .Dauer - 1
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
    Sub awinCreateBetterWorsePortfolio(ByRef ProjektListe As Collection, ByRef repChart As Object, ByVal showAbsoluteDiff As Boolean, ByVal isTimeTimeVgl As Boolean, ByVal vglTyp As Integer, _
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
        Dim request As New Request(awinSettings.databaseName)
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
            pname = ProjektListe.Item(i)

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
                projekthistorie = Nothing
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


        chtobjName = getKennung("pf", charttype, ProjektListe)



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



        With appInstance.Worksheets(arrWsNames(3))
            anzDiagrams = .ChartObjects.Count
            '
            ' um welches Diagramm handelt es sich ...
            '
            i = 1
            found = False

            While i <= anzDiagrams And Not found
                If chtobjName = .chartObjects(i).name Then
                    found = True
                    repChart = .ChartObjects(i)
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

                .SeriesCollection.NewSeries()
                .SeriesCollection(1).name = diagramTitle


                .SeriesCollection(1).ChartType = xlNS.XlChartType.xlBubble3DEffect


                For i = 1 To anzBubbles
                    tempArray(i - 1) = xAchsenValues(i - 1)
                Next i
                .SeriesCollection(1).XValues = tempArray

                For i = 1 To anzBubbles
                    tempArray(i - 1) = yAchsenValues(i - 1)
                Next i
                .SeriesCollection(1).Values = tempArray

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


                .SeriesCollection(1).BubbleSizes = tempArray



                Dim series1 As xlNS.Series = _
                        CType(.SeriesCollection(1),  _
                                xlNS.Series)
                Dim point1 As xlNS.Point = _
                            CType(series1.Points(1), xlNS.Point)


                For i = 1 To anzBubbles

                    With CType(.SeriesCollection(1).Points(i), xlNS.Point)

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


                With .ChartGroups(1)

                    If singleProject Then
                        .BubbleScale = 20
                    Else
                        .BubbleScale = 20
                    End If

                    .SizeRepresents = xlNS.XlSizeRepresents.xlSizeIsArea
                    .shownegativeBubbles = True

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

                    .MajorTickMark = XlTickMark.xlTickMarkCross
                    .TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNextToAxis

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

                    .MajorTickMark = XlTickMark.xlTickMarkCross
                    .TickLabelPosition = XlTickLabelPosition.xlTickLabelPositionNextToAxis

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
                .ChartTitle.text = diagramTitle
                .ChartTitle.Characters.Font.Size = awinSettings.fontsizeTitle

                ' Events disablen, wegen Report erstellen
                appInstance.EnableEvents = False
                .Location(Where:=xlNS.XlChartLocation.xlLocationAsObject, Name:=appInstance.Worksheets(arrWsNames(3)).name)
                appInstance.EnableEvents = formerEE
                ' Events sind wieder zurückgesetzt
            End With


            'appInstance.ShowChartTipNames = False
            'appInstance.ShowChartTipValues = False

            With .ChartObjects(anzDiagrams + 1)
                .top = top
                .left = left
                .width = width
                .height = height
                .name = chtobjName
            End With



            With appInstance.ActiveSheet
                Try
                    With appInstance.ActiveSheet
                        .Shapes(chtobjName).line.visible = chartBorderVisible
                    End With
                Catch ex As Exception

                End Try
            End With

            pfDiagram = New clsDiagramm

            pfChart = New clsEventsPfCharts
            pfChart.PfChartEvents = .ChartObjects(anzDiagrams + 1).Chart

            pfDiagram.setDiagramEvent = pfChart

            With pfDiagram

                .kennung = getKennung("pf", charttype, ProjektListe)
                .DiagrammTitel = diagramTitle
                .diagrammTyp = DiagrammTypen(3)                     ' Portfolio
                .gsCollection = ProjektListe
                .isCockpitChart = False

            End With

            DiagramList.Add(pfDiagram)
            repChart = .ChartObjects(anzDiagrams + 1)

        End With

        appInstance.ScreenUpdating = formerSU

    End Sub  ' Ende Prozedur awinCreatePortfolioChartDiagramm

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="constellationName"></param>
    ''' <remarks></remarks>
    Public Sub awinStoreConstellation(ByVal constellationName As String)

        Dim request As New Request(awinSettings.databaseName)
        ' prüfen, ob diese Constellation bereits existiert ..
        If projectConstellations.Contains(constellationName) Then

            Try
                projectConstellations.Remove(constellationName)
            Catch ex As Exception

            End Try

        End If

        Dim newC As New clsConstellation
        With newC
            .constellationName = constellationName
        End With

        Dim newConstellationItem As clsConstellationItem
        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste
            newConstellationItem = New clsConstellationItem
            With newConstellationItem
                .projectName = kvp.Key
                .show = True
                .Start = kvp.Value.startDate
                .variantName = kvp.Value.variantName
                .zeile = kvp.Value.tfZeile
            End With
            newC.Add(newConstellationItem)
        Next


        Try
            projectConstellations.Add(newC)

        Catch ex As Exception
            Call MsgBox("Fehler bei Add projectConstellations in awinStoreConstellations")
        End Try

        ' Portfolio in die Datenbank speichern
        If request.pingMongoDb() Then
            If Not request.storeConstellationToDB(newC) Then
                Call MsgBox("Fehler beim Speichern der projektConstellation '" & newC.constellationName & "' in die Datenbank")
            End If
        Else
            Throw New ArgumentException("Datenbank-Verbindung ist unterbrochen!")
        End If

    End Sub
  
End Module
