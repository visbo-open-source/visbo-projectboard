
Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports Microsoft.Office.Core
Imports pptNS = Microsoft.Office.Interop.PowerPoint
Imports xlNS = Microsoft.Office.Interop.Excel

Public Module testModule

    ''' <summary>
    ''' erzeugt den Bericht Report auf Grundlage des Templates templatedossier.pptx
    ''' bei Aufruf ist sichergestellt, daß in Projekthistorie die Historie des Projektes steht 
    ''' </summary>
    ''' <param name="hproj"></param>
    ''' <remarks></remarks>
    Public Sub createPPTSlidesFromProject(ByRef hproj As clsProjekt)
        Dim pptApp As pptNS.Application = Nothing
        Dim pptPresentation As pptNS.Presentation = Nothing
        Dim pptSlide As pptNS.Slide = Nothing
        Dim shapeRange As pptNS.ShapeRange = Nothing
        Dim presentationFile As String = awinPath & requirementsOrdner & "projektdossier.pptx"
        Dim pptTemplate As String = awinPath & requirementsOrdner & "templatedossier.pptx"
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
                        kennzeichnung = "Vergleich mit letztem Stand" Or _
                        kennzeichnung = "Vergleich mit Vorlage" Or _
                        kennzeichnung = "Tabelle Projektziele" Or _
                        kennzeichnung = "Tabelle Projektstatus" Or _
                        kennzeichnung = "Tabelle Veränderungen" Or _
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

                                            tmpStr = qualifier.Trim.Split(New Char() {"(", ")"}, 20)
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
                                Call createPhasesBalken(noColorCollection, hproj, repObj1, scale, htop, hleft, hheight, hwidth, " ")

                                With repObj1
                                    htop = .Top + .Height + 3
                                End With


                                repObj2 = Nothing
                                Call createPhasesBalken(noColorCollection, cproj, repObj2, scale, htop, hleft, hheight, hwidth, "Vorlage")

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
                                Call createPhasesBalken(noColorCollection, hproj, repObj1, scale, htop, hleft, hheight, hwidth, " ")

                                With repObj1
                                    htop = .Top + .Height + 3
                                End With

                                repObj2 = Nothing
                                Call createPhasesBalken(noColorCollection, cproj, repObj2, scale, htop, hleft, hheight, hwidth, "letzter Stand")

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


                                'deleteStack.Add(.Name, .Name)
                                'Try
                                Call createProjektErgebnisCharakteristik2(hproj, obj)
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

                                Call awinCreateStratRisikMargeDiagramm(mycollection, obj, True, False, True, False, htop, hleft, hwidth, hheight)
                                reportObj = obj

                                notYetDone = True

                            Case "Teilprojekte"

                                Dim scale As Integer

                                Dim cproj As clsProjekt = Nothing
                                Dim vproj As clsProjektvorlage


                                scale = hproj.dauerInDays

                                If qualifier.Length > 0 Then
                                    If qualifier = "Vorlage" Then

                                        vproj = Projektvorlagen.getProject(hproj.VorlagenName)
                                        vproj.CopyTo(cproj)
                                        cproj.startDate = hproj.startDate


                                    ElseIf qualifier = "Beauftragung" Then
                                        cproj = bproj

                                    Else
                                        cproj = hproj

                                    End If
                                End If
                                Dim noColorCollection As New Collection
                                reportObj = Nothing
                                Call createPhasesBalken(noColorCollection, cproj, reportObj, scale, htop, hleft, hheight, hwidth, qualifier)


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
        Next





    End Sub
    '
    '
    '
    Public Sub createPPTSlidesFromConstellation()
        Dim pptApp As pptNS.Application = Nothing
        Dim pptPresentation As pptNS.Presentation = Nothing
        Dim pptSlide As pptNS.Slide = Nothing
        Dim shapeRange As pptNS.ShapeRange = Nothing
        Dim presentationFile As String = awinPath & requirementsOrdner & "boarddossier.pptx"
        Dim pptTemplate As String = awinPath & requirementsOrdner & "templateboarddossier.pptx"
        Dim pptShape As pptNS.Shape
        Dim portfolioName As String = "Multi Projekt Übersicht"
        Dim top As Double, left As Double, width As Double, height As Double
        Dim htop As Double, hleft As Double, hwidth As Double, hheight As Double
        Dim pptSize As Integer = 18
        Dim hproj As clsProjekt
        Dim pName As String
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

        For j = 1 To AnzAdded
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
                        kennzeichnung = "Fortschritt Personalkosten" Or _
                        kennzeichnung = "Fortschritt Sonstige Kosten" Or _
                        kennzeichnung = "Fortschritt Gesamtkosten" Or _
                        kennzeichnung = "Fortschritt Rolle" Or _
                        kennzeichnung = "Fortschritt Kostenart" Or _
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

                                            If .tfZeile > maxzeile Then
                                                maxzeile = .tfZeile
                                            End If
                                        End With
                                    Catch ex As Exception

                                    End Try

                                Next
                                maxzeile = maxzeile + 1
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

                                        Dim farbTyp As Integer = 3
                                        Call awinZeichneMilestones(nameList, farbTyp, True)
                                        

                                    ElseIf qualifier = "Milestones GR" Then
                                        Call awinDeleteMilestoneShapes(0)

                                        Dim farbTyp As Integer = 2
                                        Call awinZeichneMilestones(nameList, farbTyp, False)
                                        farbTyp = 3
                                        Call awinZeichneMilestones(nameList, farbTyp, False)
                                        
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

                                            If .tfZeile > maxzeile Then
                                                maxzeile = .tfZeile
                                            End If
                                        End With
                                    Catch ex As Exception

                                    End Try

                                Next
                                maxzeile = maxzeile + 1
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

                                If minColumn < von - 12 Then
                                    minColumn = von - 12
                                End If

                                If maxColumn > bis + 12 Then
                                    maxColumn = bis + 12
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

                                        Call awinZeichnePhasen(phNameCollection, 4, False)
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
                            Call awinCreateStratRisikMargeDiagramm(myCollection, obj, False, False, True, True, htop, hleft, hwidth, hheight)


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
                            'Call awinCreateComplexRiskVolumeDiagramm(myCollection, obj, False, False, True, True, htop, hleft, hwidth, hheight)
                            Call awinCreateZeitRiskVolumeDiagramm(myCollection, obj, False, False, True, True, htop, hleft, hwidth, hheight)


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
                            If PhaseDefinitions.Contains(qualifier) Then

                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "

                                myCollection.Add(qualifier, qualifier)

                                htop = 100
                                hleft = 100
                                hheight = miniHeight  ' height of all charts
                                hwidth = miniWidth   ' width of all charts
                                obj = Nothing
                                Call awinCreateprcCollectionDiagram(myCollection, obj, htop, hleft, hwidth, hheight, False, DiagrammTypen(0), True)

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
                                ' der Titel wird geändert im Report, deswegen wird das Diagramm  nicht gefunden in awinDeleteChart 

                                Try
                                    reportObj.Delete()
                                    'DiagramList.Remove(DiagramList.Count)
                                Catch ex As Exception

                                End Try

                            Else
                                .TextFrame2.TextRange.Text = "nicht definiert: " & qualifier
                            End If


                        Case "Rolle"


                            myCollection.Clear()
                            If RoleDefinitions.Contains(qualifier) Then

                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "

                                myCollection.Add(qualifier, qualifier)

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
                                    ' das diagramlist.remove darf nicht gemacht werden, weil sonst die gespeicherten Werte 
                                    ' für die top. left, .. Positionen verloren gehen  
                                    'DiagramList.Remove(DiagramList.Count)
                                Catch ex As Exception

                                End Try

                            Else
                                .TextFrame2.TextRange.Text = "nicht definiert: " & qualifier
                            End If

                        Case "Kostenart"


                            myCollection.Clear()
                            If CostDefinitions.Contains(qualifier) Then

                                pptSize = .TextFrame2.TextRange.Font.Size
                                .TextFrame2.TextRange.Text = " "

                                myCollection.Add(qualifier, qualifier)

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

            listofShapes.Clear()

        Next




    End Sub

    Public Sub StoreAllProjectsinDB()

        Dim jetzt As Date = Now
        Dim request As New Request(awinSettings.databaseName)
        enableOnUpdate = False
        ' die aktuelle Konstellation wird unter dem Namen <Last> gespeichert ..
        Call awinStoreConstellation("Last")

        ' jetzt werden die gezeigten Projekte in die Datenbank geschrieben 

        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte

            Try
                ' hier wird der Wert für kvp.Value.timeStamp = heute gesetzt 

                If demoModusHistory Then
                    kvp.Value.timeStamp = historicDate
                Else
                    kvp.Value.timeStamp = jetzt
                End If

                Call request.storeProjectToDB(kvp.Value)
            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

        Next

        historicDate = historicDate.AddMonths(1)

        ' jetzt werden alle definierten Constellations weggeschrieben

        For Each kvp As KeyValuePair(Of String, clsConstellation) In projectConstellations.Liste

            Try
                Call request.storeConstellationToDB(kvp.Value)
            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

        Next

        enableOnUpdate = True

    End Sub



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
                    ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde

                    projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pname, variantName:=variantName, _
                                                                        storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
                    projekthistorie.Add(Date.Now, hproj)
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
            ' projekthistorie muss nur dann neu bestimmt werden, wenn sie nicht bereits für dieses Projekt geholt wurde

            projekthistorie.liste = request.retrieveProjectHistoryFromDB(projectname:=pname, variantName:=variantName, _
                                                                storedEarliest:=StartofCalendar, storedLatest:=Date.Now)
            projekthistorie.Add(Date.Now, hproj)
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

        Try
            tabelle = pptShape.Table
        Catch ex As Exception
            Throw New Exception("Shape hat keine Tabelle")
        End Try




        Dim todoListe As New SortedList(Of Long, clsProjekt)
        Dim key As Long

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            key = 10000 * kvp.Value.tfZeile + kvp.Value.tfspalte
            todoListe.Add(key, kvp.Value)

        Next

        Dim msNumber As Integer = 1

        ' jetzt wird die todoListe abgearbeitet 
        Dim tabellenzeile As Integer = 2
        Dim hproj As clsProjekt
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
                    'Try
                    '    cBewertung = cResult.getBewertung(1)
                    'Catch ex As Exception
                    '    cBewertung = New clsBewertung
                    'End Try

                    resultColumn = getColumnOfDate(cResult.getDate)

                    If farbtyp = cBewertung.colorIndex Then
                        ' es muss nur etwas gemacht werden , wenn entweder alle Farben gezeichnet werden oder eben die übergebene

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


        newshape = pptSlide.Shapes.Paste
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


End Module
