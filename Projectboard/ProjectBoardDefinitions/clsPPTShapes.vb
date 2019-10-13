Imports pptNS = Microsoft.Office.Interop.PowerPoint
''' <summary>
''' nimmt die Powerpoint Shapes auf, die notwendig sind, 
''' um eine Einzelprojekt- oder Multiprojekt-Sicht zu erzeugen 
''' Es gibt Methoden, die überprüfen, ob alle notwendigen Hilfsshapes für eine konkrete Aufgabenstellung vorhanden sind   
''' </summary>
''' <remarks></remarks>
''' 

Public Class clsPPTShapes
    Private _pptSlide As pptNS.Slide
    Private _containerLeft As Double = 0.0
    Private _containerRight As Double = 0.0
    Private _containerTop As Double = 0.0
    Private _containerBottom As Double = 0.0

    Private _calendarLeft As Double = 0.0
    Private _calendarRight As Double = 0.0
    Private _calendarTop As Double = 0.0
    Private _calendarBottom As Double = 0.0

    Private _drawingAreaLeft As Double = 0.0
    Private _drawingAreaRight As Double = 0.0
    Private _drawingAreaTop As Double = 0.0
    Private _drawingAreaBottom As Double = 0.0

    Private _projectListLeft As Double = 0.0

    Private _legendAreaLeft As Double = 0.0
    Private _legendAreaRight As Double = 0.0
    Private _legendAreaTop As Double = 0.0
    Private _legendAreaBottom As Double = 0.0

    ' enthalten die relativen Abstände der Text-Shapes zu ihrem Phasen/Meilenstein Element 
    Private _yOffsetMsToText As Double = 0.0
    Private _yOffsetMsToDate As Double = 0.0

    Private _yOffsetPhToText As Double = 0.0
    Private _yOffsetPhToDate As Double = 0.0

    Private _containerShape As pptNS.Shape = Nothing
    Private _calendarLineShape As pptNS.Shape = Nothing
    Private _legendLineShape As pptNS.Shape = Nothing


    ' enthält das PPTStartofCalendar and PPTEndOfCalendar
    Private _PPTStartOFCalendar As Date = StartofCalendar
    Private _PPTEndOFCalendar As Date = StartofCalendar

    Private _anzahlTageImKalender As Integer = 0
    Private _tagesbreite As Double = 0

    ' was ist die Zeilenhöhe in der Zeichenarea
    Private _zeilenHoehe As Double = 0.0

    ' wo beginnen die Shapes relativ gesehen innerhalb einer Zeile, aufgeführt für Duration, ProjekctLine, Phase, Milestone
    Private _YDurationText As Double = 0.0
    Private _YDurationArrow As Double = 0.0

    Private _YProjectLine As Double = 0.0
    Private _YprojectName As Double = 0.0

    Private _YPhase As Double = 0.0
    Private _YPhasenText As Double = 0.0
    Private _YPhasenDatum As Double = 0.0

    Private _YMilestone As Double = 0.0
    Private _YMilestoneText As Double = 0.0
    Private _YMilestoneDate As Double = 0.0

    ''' <summary>
    ''' löscht das Shape inkl Try..catch Behandlung
    ''' </summary>
    ''' <param name="shape"></param>
    ''' <remarks></remarks>
    Private Sub makeShapeInvisible(ByRef shape As pptNS.Shape)

        If Not IsNothing(shape) Then

            Try
                'shape.Delete()
                shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
            Catch ex As Exception

            End Try

        End If
    End Sub

    ''' <summary>
    ''' alle Hilfsshapes, die auf der aktuellen Slide drauf sind, werden den entsprechenden 
    ''' Klassen-Variablen zugewiesen  
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub shapesZuweisen(ByVal pptslide As pptNS.Slide)

        Dim anzShapes As Integer = pptSlide.Shapes.Count
        Dim pptShape As pptNS.Shape
        ' rds = Report Defining Shapes nimmt alle Hilfsshapes auf, die für das Zeichnen des Reports notwendig sind 

        For i = 1 To anzShapes
            pptShape = pptSlide.Shapes(i)

            With pptShape

                ' jetzt muss geprüft werden, ob es sich um ein definierendes Element für die Multiprojekt-Sichten handelt
                If .Title.Length > 0 Then

                    ' Anmerkung: es ist wichtig den Properties die Zuweisung zu machen, andernfalls werden ggf die im Set Bereich definierten 
                    ' Aktionen nicht durchgeführt ...
                    Select Case .Title

                        Case "MilestoneDescription"
                            MsDescVorlagenShape = pptShape

                        Case "ProjectName"
                            projectNameVorlagenShape = pptShape

                        Case "CalendarLine"
                            calendarLineShape = pptShape

                        Case "QuarterMonthinCal"
                            quarterMonthVorlagenShape = pptShape

                        Case "YearInCal"
                            yearVorlagenShape = pptShape

                        Case "ProjectForm"
                            projectVorlagenShape = pptShape

                        Case "PhaseForm"
                            phaseVorlagenShape = pptShape

                        Case "MilestoneForm"
                            milestoneVorlagenShape = pptShape

                        Case "Ampel"
                            ampelVorlagenShape = pptShape

                        Case "Jahres-Trennstrich"
                            calendarYearSeparator = pptShape

                        Case "Quartals-Trennstrich"
                            calendarQuartalSeparator = pptShape

                        Case "Horizontale"
                            horizontalLineShape = pptShape

                        Case "TodayLine"
                            todayLineShape = pptShape

                        Case "LegendLine"
                            legendLineShape = pptShape

                        Case "LegendStart"
                            legendStartShape = pptShape

                        Case "LegendText"
                            legendTextVorlagenShape = pptShape

                        Case "LegendPhase"
                            legendPhaseVorlagenShape = pptShape

                        Case "LegendMilestone"
                            legendMilestoneVorlagenShape = pptShape

                        Case "Multiprojektsicht"
                            containerShape = pptShape

                        Case "Multivariantensicht"
                            containerShape = pptShape

                        Case "Einzelprojektsicht"
                            containerShape = pptShape

                        Case "AllePlanElemente"
                            containerShape = pptShape

                            ' alle Hierarchie-Stufe 1 Objekte sind Swimlanes
                        Case "Swimlanes"
                            containerShape = pptShape

                        Case "Swimlanes2"
                            containerShape = pptShape

                        Case "MilestoneCategories"
                            containerShape = pptShape

                        Case "CalendarHeight"
                            calendarHeightShape = pptShape

                        Case "MilestoneDate"
                            MsDateVorlagenShape = pptShape

                        Case "PhaseDescription"
                            PhDescVorlagenShape = pptShape

                        Case "PhaseDate"
                            PhDateVorlagenShape = pptShape

                        Case "CalendarStep"
                            ' optional
                            calendarStepShape = pptShape

                        Case "CalendarMark"
                            ' optional 
                            calendarMarkShape = pptShape

                        Case "Fehlermeldung"
                            ' optional 
                            errorVorlagenShape = pptShape

                        Case "LegendBuColor"
                            ' optional
                            legendBuColorShape = pptShape

                        Case "buColorShape"
                            ' optional
                            buColorShape = pptShape

                        Case "rowDifferentiator"
                            ' optional
                            rowDifferentiatorShape = pptShape

                        Case "PhaseDelimiter"
                            ' optional 
                            phaseDelimiterShape = pptShape

                        Case "durationArrow"
                            ' optional
                            durationArrowShape = pptShape

                        Case "durationText"
                            ' optional 
                            durationTextShape = pptShape

                        Case "SegmentText"
                            ' optional 
                            segmentVorlagenShape = pptShape

                        Case Else


                    End Select
                End If


            End With
        Next
    End Sub

    ''' <summary>
    ''' bestimme die relativen Abstände der Text-Shapes zu ihrem Phase/Milestone Element
    ''' yOffsetMsToText, yOffsetMsToDate
    ''' yOffsetPhToText, yOffsetPhToDate
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub calcRelDisTxtToElm()

        _yOffsetMsToText = _MsDescVorlagenShape.Top - _milestoneVorlagenShape.Top
        _yOffsetMsToDate = _MsDateVorlagenShape.Top - _milestoneVorlagenShape.Top

        _yOffsetPhToText = _PhDescVorlagenShape.Top - _phaseVorlagenShape.Top
        _yOffsetPhToDate = _PhDateVorlagenShape.Top - _phaseVorlagenShape.Top

    End Sub

    ''' <summary>
    ''' berechnet anhand der Daten des Startdatums, Ende-Datums die korrespondierenden x1, x2 Koordinaten
    ''' </summary>
    ''' <param name="startdate"></param>
    ''' <param name="enddate"></param>
    ''' <param name="x1Pos"></param>
    ''' <param name="x2Pos"></param>
    ''' <remarks></remarks>
    Public Sub calculatePPTx1x2(ByVal startdate As Date, ByVal enddate As Date, _
                                    ByRef x1Pos As Double, ByRef x2Pos As Double)


        Dim offset1 As Integer = CInt(DateDiff(DateInterval.Day, Me.PPTStartOFCalendar.Date, startdate.Date))

        If offset1 <= 0 Then
            x1Pos = Me.drawingAreaLeft
        Else
            x1Pos = Me.drawingAreaLeft + offset1 * Me.tagesbreite
        End If


        Dim offset2 As Integer = CInt(DateDiff(DateInterval.Day, Me.PPTStartOFCalendar.Date, enddate.Date))

        If offset2 >= Me.anzahlTageImKalender Then
            x2Pos = Me.drawingAreaRight
        Else
            x2Pos = Me.drawingAreaLeft + offset2 * Me.tagesbreite
            ' Änderung tk 27.10 , eine Phase geht von Anfang des Tages bis Ende des Tages... 
            'x2Pos = Me.drawingAreaLeft + (offset2 + 0.75) * Me.tagesbreite
        End If

        ' '' Test Funktionen , eingeführt um die Rückwärtsrechnung Koordinaten->Datum zu überprüfen ... 
        ''Dim tstStart As Date = calcXtoDate(x1Pos)
        ''If DateDiff(DateInterval.Day, startdate, tstStart) = 0 Then
        ''    ' alles in Ordnung 
        ''Else
        ''    Call MsgBox("Unterschied: " & startdate.ToString & " versus " & tstStart.ToString)
        ''End If

        ''Dim tstEnde As Date = calcEndDate(x1Pos, x2Pos - x1Pos)
        ''If DateDiff(DateInterval.Day, enddate, tstEnde) = 0 Then
        ''    ' alles in Ordnung 
        ''Else
        ''    Call MsgBox("Unterschied: " & enddate.ToString & " versus " & tstEnde.ToString)
        ''End If

        ''tstEnde = calcXtoDate(x2Pos)
        ''If DateDiff(DateInterval.Day, enddate, tstEnde) = 0 Then
        ''    ' alles in Ordnung 
        ''Else
        ''    Call MsgBox("Unterschied: " & enddate.ToString & " versus " & tstEnde.ToString)
        ''End If

    End Sub

    ''' <summary>
    ''' berechnet für die angegebene Koordinate das zugehörige Datum 
    ''' </summary>
    ''' <param name="xPos"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcXtoDate(ByVal xPos As Double) As Date

        Dim tmpDate As Date = Me.PPTStartOFCalendar

        If Me._tagesbreite > 0 Then
            tmpDate = Me.PPTStartOFCalendar.AddDays(CInt((xPos - Me._drawingAreaLeft) / Me._tagesbreite))
            ' Änderung tk 27.10.17, ein Phase die nur einen Tag dauert, soll auch so angezeigt werden ... 
            'tmpDate = Me.PPTStartOFCalendar.AddDays(CInt(System.Math.Truncate((xPos - Me._drawingAreaLeft) / Me._tagesbreite)))
        End If
        calcXtoDate = tmpDate

    End Function

    ''' <summary>
    ''' berechnet für die angegebene LEft Koordinate und Länge das zugehörige Datum 
    ''' </summary>
    ''' <param name="xPos"></param>
    ''' <param name="width"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function calcEndDate(ByVal xPos As Double, ByVal width As Double) As Date
        Dim tmpDate As Date = Me.PPTStartOFCalendar.AddDays(CInt((xPos + width - Me._drawingAreaLeft) / Me._tagesbreite))
        'Dim tmpDate As Date = Me.PPTStartOFCalendar.AddDays(CInt(System.Math.Truncate((xPos + width - Me._drawingAreaLeft) / Me._tagesbreite)))
        calcEndDate = tmpDate
    End Function

    Public Sub bestimmeZeilenHoehe(ByVal anzphasen As Integer, ByVal anzMeilensteine As Integer, _
                                   ByVal considerAll As Boolean)


        Dim minY As Double = _containerBottom, maxY As Double = _containerTop

        ' bestimme als erstes die maximale/minimale Y-Koordinate, die sich ergibt wenn man alle -relevanten- Shapes berücksichtigt 
        '
        '
        If Not IsNothing(projectNameVorlagenShape) Then

            With projectNameVorlagenShape
                minY = System.Math.Min(minY, .Top)
                maxY = System.Math.Max(maxY, .Top + .Height)
            End With

        End If

        ' soll überhaupt eine Dauer angezeigt werden ? 
        If awinSettings.mppSortiertDauer Then

            If IsNothing(durationTextShape) Then
                With durationTextShape
                    minY = System.Math.Min(minY, .Top)
                    maxY = System.Math.Max(maxY, .Top + .Height)
                End With
            End If

            If IsNothing(durationArrowShape) Then
                With durationArrowShape
                    minY = System.Math.Min(minY, .Top)
                    maxY = System.Math.Max(maxY, .Top + .Height)
                End With
            End If
        End If


        If awinSettings.mppShowProjectLine And Not IsNothing(projectVorlagenShape) Then
            With projectVorlagenShape
                minY = System.Math.Min(minY, .Top)
                maxY = System.Math.Max(maxY, .Top + .Height)
            End With
        End If

        ' Müssen Phasen überhaupt gezeichnet werden ? 
        If anzphasen > 0 Or considerAll Then
            If Not IsNothing(phaseVorlagenShape) Then
                With phaseVorlagenShape
                    minY = System.Math.Min(minY, .Top)
                    maxY = System.Math.Max(maxY, .Top + .Height)
                End With
            End If

            If Not awinSettings.mppUseInnerText Then
                If Not IsNothing(PhDescVorlagenShape) And awinSettings.mppShowPhName Then
                    With PhDescVorlagenShape
                        minY = System.Math.Min(minY, .Top)
                        maxY = System.Math.Max(maxY, .Top + .Height)
                    End With
                End If

                If Not IsNothing(PhDateVorlagenShape) And awinSettings.mppShowPhDate Then
                    With PhDateVorlagenShape
                        minY = System.Math.Min(minY, .Top)
                        maxY = System.Math.Max(maxY, .Top + .Height)
                    End With
                End If
            End If


        End If

        ' Müssen Meilensteine überhaupt gezeichnet werden ? 
        If anzMeilensteine > 0 Or considerAll Then
            If Not IsNothing(milestoneVorlagenShape) Then
                With milestoneVorlagenShape
                    minY = System.Math.Min(minY, .Top)
                    maxY = System.Math.Max(maxY, .Top + .Height)
                End With
            End If

            If Not IsNothing(MsDescVorlagenShape) And awinSettings.mppShowMsName Then
                With MsDescVorlagenShape
                    minY = System.Math.Min(minY, .Top)
                    maxY = System.Math.Max(maxY, .Top + .Height)
                End With
            End If

            If Not IsNothing(MsDateVorlagenShape) And awinSettings.mppShowMsDate Then
                With MsDateVorlagenShape
                    minY = System.Math.Min(minY, .Top)
                    maxY = System.Math.Max(maxY, .Top + .Height)
                End With
            End If

        End If

        '
        '
        ' jetzt ist die minimale/maximale Ausdehnung bestimmt 

        If minY <= maxY Then
            _zeilenHoehe = (maxY - minY) * 1.03
        End If


        ' und jetzt werden die relativen Offsets bestimmt 
        '
        '
        If Not IsNothing(projectNameVorlagenShape) Then

            With projectNameVorlagenShape
                _YprojectName = .Top - minY
            End With

        End If

        ' soll überhaupt eine Dauer angezeigt werden ? 
        If awinSettings.mppSortiertDauer Then

            If IsNothing(durationTextShape) Then
                With durationTextShape
                    _YDurationText = .Top - minY
                End With
            End If

            If IsNothing(durationArrowShape) Then
                With durationArrowShape
                    _YDurationArrow = .Top - minY
                End With
            End If
        End If


        If awinSettings.mppShowProjectLine And Not IsNothing(projectVorlagenShape) Then
            With projectVorlagenShape
                _YProjectLine = .Top - minY
            End With
        End If

        ' Müssen Phasen überhaupt gezeichnet werden ? 
        If anzphasen > 0 Or considerAll Then
            If Not IsNothing(phaseVorlagenShape) Then
                With phaseVorlagenShape
                    _YPhase = .Top - minY
                End With
            End If

            If Not awinSettings.mppUseInnerText Then
                If Not IsNothing(PhDescVorlagenShape) And awinSettings.mppShowPhName Then
                    With PhDescVorlagenShape
                        _YPhasenText = .Top - minY
                    End With
                End If

                If Not IsNothing(PhDateVorlagenShape) And awinSettings.mppShowPhDate Then
                    With PhDateVorlagenShape
                        _YPhasenDatum = .Top - minY
                    End With
                End If
            End If


        End If

        ' Müssen Meilensteine überhaupt gezeichnet werden ? 
        If anzMeilensteine > 0 Or considerAll Then
            If Not IsNothing(milestoneVorlagenShape) Then
                With milestoneVorlagenShape
                    _YMilestone = .Top - minY
                End With
            End If

            If Not IsNothing(MsDescVorlagenShape) And awinSettings.mppShowMsName Then
                With MsDescVorlagenShape
                    _YMilestoneText = .Top - minY
                End With
            End If

            If Not IsNothing(MsDateVorlagenShape) And awinSettings.mppShowMsDate Then
                With MsDateVorlagenShape
                    _YMilestoneDate = .Top - minY
                End With
            End If

        End If


        '
        '
        ' jetzt sind die relativen Offsets alle bestimmt; zumindest die, die aufgrund settings überhaupt relevant sind 




    End Sub

    ''' <summary>
    ''' gibt eine Liste zurück, die die fehlenden Hilfs-Shape Elemente für Epp bzw Mpp enthält 
    ''' wenn alles ok, dann ist die Liste leer
    ''' </summary>
    ''' <param name="kennzeichnung"></param>
    ''' <remarks></remarks>
    Public ReadOnly Property getMissingShpNames(ByVal kennzeichnung As String) As String

        Get
            Dim tmpErg As String = ""
            Dim tmpName As String = ""
            Dim firstTime As Boolean = True
            Dim ok As Boolean = True
            If kennzeichnung = "AllePlanElemente" Or _
                kennzeichnung = "Multivariantensicht" Or _
                kennzeichnung = "Multiprojektsicht" Or _
                kennzeichnung = "Einzelprojektsicht" Then

                If IsNothing(_MsDescVorlagenShape) Then
                    ok = False
                    tmpName = "Meilenstein-Beschriftung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_MsDateVorlagenShape) Then
                    ok = False
                    tmpName = "Meilenstein-Datum"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_PhDescVorlagenShape) Then
                    ok = False
                    tmpName = "Phasen-Beschriftung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_PhDateVorlagenShape) Then
                    ok = False
                    tmpName = "Phasen-Datum"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_projectNameVorlagenShape) Then
                    ok = False
                    tmpName = "Projekt-/Swimlane Name"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarLineShape) Then
                    ok = False
                    tmpName = "Kalender-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_quarterMonthVorlagenShape) Then
                    ok = False
                    tmpName = "Quartals/Monats/Kalenderwoche Beschriftung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_yearVorlagenShape) Then
                    ok = False
                    tmpName = "Jahres-Beschriftung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_projectVorlagenShape) Then
                    ok = False
                    tmpName = "Projekt-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_phaseVorlagenShape) Then
                    ok = False
                    tmpName = "Phasen-Form"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_milestoneVorlagenShape) Then
                    ok = False
                    tmpName = "Meilenstein-Form"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_ampelVorlagenShape) Then
                    ok = False
                    tmpName = "Ampel-Form"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarYearSeparator) Then
                    ok = False
                    tmpName = "Jahres-Trenn-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarQuartalSeparator) Then
                    ok = False
                    tmpName = "Q/M/KW-Trenn-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendLineShape) Then
                    ok = False
                    tmpName = "Legenden-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendStartShape) Then
                    ok = False
                    tmpName = "Legenden-Titel"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendTextVorlagenShape) Then
                    ok = False
                    tmpName = "Legenden-Textvorlage"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendPhaseVorlagenShape) Then
                    ok = False
                    tmpName = "Legenden-Phasen-Vorlage"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendMilestoneVorlagenShape) Then
                    ok = False
                    tmpName = "Legenden-Meilenstein-Vorlage"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_containerShape) Then
                    ok = False
                    tmpName = "Container Shape"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarHeightShape) Then
                    ok = False
                    tmpName = "Kalenderbegrenzung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarStepShape) Then
                    ok = False
                    tmpName = "Kalendestep-Begrenzung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarMarkShape) Then
                    ok = False
                    tmpName = "Kalender-Markierung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_errorVorlagenShape) Then
                    ok = False
                    tmpName = "Fehlertext-Vorlage"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_rowDifferentiatorShape) Then
                    ok = False
                    tmpName = "Zeilen-Hervorhebung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

            ElseIf kennzeichnung.StartsWith("Swimlane") Then

                If IsNothing(_MsDescVorlagenShape) Then
                    ok = False
                    tmpName = "Meilenstein-Beschriftung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_MsDateVorlagenShape) Then
                    ok = False
                    tmpName = "Meilenstein-Datum"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_PhDescVorlagenShape) Then
                    ok = False
                    tmpName = "Phasen-Beschriftung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_PhDateVorlagenShape) Then
                    ok = False
                    tmpName = "Phasen-Datum"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_projectNameVorlagenShape) Then
                    ok = False
                    tmpName = "Projekt-Name"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarLineShape) Then
                    ok = False
                    tmpName = "Kalender-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_quarterMonthVorlagenShape) Then
                    ok = False
                    tmpName = "Quartals/Monats/Kalenderwoche Beschriftung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_yearVorlagenShape) Then
                    ok = False
                    tmpName = "Jahres-Beschriftung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_projectVorlagenShape) Then
                    ok = False
                    tmpName = "Projekt-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_phaseVorlagenShape) Then
                    ok = False
                    tmpName = "Phasen-Form"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_milestoneVorlagenShape) Then
                    ok = False
                    tmpName = "Meilenstein-Form"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                ' muss hier nicht da sein 
                'If IsNothing(_ampelVorlagenShape) Then
                '    ok = False
                '    tmpName = "Ampel-Form"
                '    If firstTime Then
                '        tmpErg = "fehlende PPT-Shapes: " & vbLF & tmpName
                '        firstTime = False
                '    Else
                '        tmpErg = tmpErg & vbLF & tmpName
                '    End If
                'End If

                If IsNothing(_calendarYearSeparator) Then
                    ok = False
                    tmpName = "Jahres-Trenn-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If


                If IsNothing(_calendarQuartalSeparator) Then
                    ok = False
                    tmpName = "Q/M/KW-Trenn-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_horizontalLineShape) Then
                    ok = False
                    tmpName = "horizontale Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendLineShape) Then
                    ok = False
                    tmpName = "Legenden-Linie"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendStartShape) Then
                    ok = False
                    tmpName = "Legenden-Titel"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendTextVorlagenShape) Then
                    ok = False
                    tmpName = "Legenden-Textvorlage"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendPhaseVorlagenShape) Then
                    ok = False
                    tmpName = "Legenden-Phasen-Vorlage"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_legendMilestoneVorlagenShape) Then
                    ok = False
                    tmpName = "Legenden-Meilenstein-Vorlage"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_containerShape) Then
                    ok = False
                    tmpName = "Container Shape"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarHeightShape) Then
                    ok = False
                    tmpName = "Kalenderbegrenzung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarStepShape) Then
                    ok = False
                    tmpName = "Kalendestep-Begrenzung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_calendarMarkShape) Then
                    ok = False
                    tmpName = "Kalender-Markierung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_errorVorlagenShape) Then
                    ok = False
                    tmpName = "Fehlertext-Vorlage"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_rowDifferentiatorShape) Then
                    ok = False
                    tmpName = "Zeilen-Hervorhebung"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

                If IsNothing(_segmentVorlagenShape) Then
                    ok = False
                    tmpName = "Beschriftung Hierarchiestufe 1"
                    If firstTime Then
                        tmpErg = "fehlende PPT-Shapes: " & vbLf & tmpName
                        firstTime = False
                    Else
                        tmpErg = tmpErg & vbLf & tmpName
                    End If
                End If

            End If

            getMissingShpNames = tmpErg

        End Get
    End Property

    ''' <summary>
    ''' wenn Kalenderlinie oder Legend-Linie über Container Grenzen gehen, werden die Koordinaten der Lines entsprechend angepasst 
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub plausibilityAdjustments()

        If _calendarLineShape.Left < _containerLeft Then
            _calendarLineShape.Left = CSng(_containerLeft + 0.1 * (_containerRight - _containerLeft))
        End If

        If _calendarLineShape.Left + _calendarLineShape.Width > _containerRight Then
            _calendarLineShape.Width = CSng(0.9 * (_containerRight - _calendarLineShape.Left))
        End If

        If _legendLineShape.Left < _containerLeft Then
            _legendLineShape.Left = CSng(_containerLeft + 0.1 * (_containerRight - _containerLeft))
        End If

        If _legendLineShape.Left + _legendLineShape.Width > _containerRight Then
            _legendLineShape.Width = CSng(0.9 * (_containerRight - _legendLineShape.Left))
        End If

    End Sub


    ''' <summary>
    ''' wenn die Zuweisung gemacht wird, werden all die evtl auf dieser Slide vorhandenen Hilfsshapes
    ''' ausgelesen und entsprechend intern gesetzt  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property pptSlide As pptNS.Slide
        Get
            pptSlide = _pptSlide
        End Get
        Set(value As pptNS.Slide)
            _pptSlide = value

            If Not IsNothing(_pptSlide) Then
                Call shapesZuweisen(_pptSlide)
            End If

        End Set
    End Property

    Public ReadOnly Property containerleft As Double
        Get
            containerleft = _containerLeft
        End Get
    End Property

    Public ReadOnly Property containerRight As Double
        Get
            containerRight = _containerRight
        End Get
    End Property

    Public ReadOnly Property containerTop As Double
        Get
            containerTop = _containerTop
        End Get
    End Property

    Public ReadOnly Property containerBottom As Double
        Get
            containerBottom = _containerBottom
        End Get
    End Property

    Public ReadOnly Property calendarLeft As Double
        Get
            calendarLeft = _calendarLeft
        End Get
    End Property

    Public ReadOnly Property calendarRight As Double
        Get
            calendarRight = _calendarRight
        End Get
    End Property

    Public ReadOnly Property calendarTop As Double
        Get
            calendarTop = _calendarTop
        End Get

    End Property

    Public WriteOnly Property setCalendarTop As Double
        Set(value As Double)
            _calendarTop = value
        End Set
    End Property

    Public ReadOnly Property calendarBottom As Double
        Get
            calendarBottom = _calendarBottom
        End Get
    End Property

    Public ReadOnly Property drawingAreaWidth As Double
        Get
            drawingAreaWidth = _drawingAreaRight - _drawingAreaLeft
        End Get
    End Property

    Public ReadOnly Property drawingAreaLeft As Double
        Get
            drawingAreaLeft = _drawingAreaLeft
        End Get
    End Property

    Public ReadOnly Property drawingAreaRight As Double
        Get
            drawingAreaRight = _drawingAreaRight
        End Get
    End Property


    Public ReadOnly Property drawingAreaTop As Double
        Get
            drawingAreaTop = _drawingAreaTop
        End Get
    End Property

    Public ReadOnly Property availableSpace As Double
        Get
            availableSpace = _drawingAreaBottom - _drawingAreaTop
        End Get
    End Property

    Public ReadOnly Property drawingAreaBottom As Double
        Get
            drawingAreaBottom = _drawingAreaBottom
        End Get
    End Property

    Public ReadOnly Property projectListLeft As Double
        Get
            projectListLeft = _projectListLeft
        End Get
    End Property

    Public ReadOnly Property legendAreaLeft As Double
        Get
            legendAreaLeft = _legendAreaLeft
        End Get
    End Property


    Public ReadOnly Property legendAreaRight As Double
        Get
            legendAreaRight = _legendAreaRight
        End Get
    End Property

    Public ReadOnly Property legendAreaTop As Double
        Get
            legendAreaTop = _legendAreaTop
        End Get
    End Property

    Public ReadOnly Property legendAreaBottom As Double
        Get
            legendAreaBottom = _legendAreaBottom
        End Get
    End Property

    ''' <summary>
    ''' Readonly, wird gesetzt in Methode calcRelDisTxtToElem
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property yOffsetMsToText As Double
        Get
            yOffsetMsToText = _yOffsetMsToText
        End Get
    End Property

    ''' <summary>
    ''' Readonly, wird gesetzt in Methode calcRelDisTxtToElem
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property yOffsetMsToDate As Double
        Get
            yOffsetMsToDate = _yOffsetMsToDate
        End Get
    End Property


    ''' <summary>
    ''' Readonly, wird gesetzt in Methode calcRelDisTxtToElem
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property yOffsetPhToText As Double
        Get
            yOffsetPhToText = _yOffsetPhToText
        End Get
    End Property

    ''' <summary>
    ''' Readonly, wird gesetzt in Methode calcRelDisTxtToElem
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property yOffsetPhToDate As Double
        Get
            yOffsetPhToDate = _yOffsetPhToDate
        End Get
    End Property

    Public ReadOnly Property PPTStartOFCalendar As Date
        Get
            PPTStartOFCalendar = _PPTStartOFCalendar
        End Get
    End Property


    Public ReadOnly Property PPTEndOFCalendar As Date
        Get
            PPTEndOFCalendar = _PPTEndOFCalendar
        End Get
    End Property

    Public Sub setCalendarDates(ByVal pptStartOfCalendar As Date, ByVal pptEndOfCalendar As Date)

        If pptEndOfCalendar > pptStartOfCalendar Then
            If pptStartOfCalendar >= StartofCalendar Then
                _PPTStartOFCalendar = pptStartOfCalendar
                _PPTEndOFCalendar = pptEndOfCalendar
                _anzahlTageImKalender = CInt(DateDiff(DateInterval.Day, _PPTStartOFCalendar, _PPTEndOFCalendar))

                ' falls die CalendarlineShape existiert 
                If Not IsNothing(calendarLineShape) Then
                    If _anzahlTageImKalender > 0 Then
                        _tagesbreite = calendarLineShape.Width / _anzahlTageImKalender
                    End If
                End If
            Else
                Throw New ArgumentException("Das Startdatum in der Konfigurations-Datei muss vor dem gewählten Start-Datum liegen" & vbLF _
                                            & "bitte ändern Sie das Datum ggf. in der Konfigurations-Datei")
            End If
        Else
            Throw New ArgumentException("Ende Datum kann nicht vor Start-Datum liegen!")
        End If

    End Sub

    ''' <summary>
    ''' Readonly, gibt die Anzahl tage im Kalender zuürkc 
    ''' wird gesetzt in setCalendarDates
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property anzahlTageImKalender As Integer
        Get
            anzahlTageImKalender = _anzahlTageImKalender
        End Get
    End Property

    ''' <summary>
    ''' Readonly, gibt die Tagesbreite im Kalender zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property tagesbreite As Double
        Get
            tagesbreite = _tagesbreite
        End Get
    End Property

    ''' <summary>
    ''' Readonly, gibt die 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property segmentHoehe As Double
        Get
            If IsNothing(segmentVorlagenShape) Then
                segmentHoehe = 0.0
            Else
                segmentHoehe = segmentVorlagenShape.Height
            End If
        End Get
    End Property

    ''' <summary>
    ''' Readonly, gibt die Zeilenhöhe zurück 
    ''' wird gesetzt in Methode bestimmeZeilenhoehe 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property zeilenHoehe As Double
        Get
            zeilenHoehe = _zeilenHoehe
        End Get
    End Property


    ''' <summary>
    ''' Readonly: relativer Top von DurationText
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YDurationText As Double
        Get
            YDurationText = _YDurationText
        End Get
    End Property

    ''' <summary>
    ''' Readonly: relativer Top von DurationArrow
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YDurationArrow As Double
        Get
            YDurationArrow = _YDurationArrow
        End Get
    End Property

    ''' <summary>
    ''' Readonly: relativer Top von projectLine
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YProjectLine As Double
        Get
            YProjectLine = _YProjectLine
        End Get
    End Property


    ''' <summary>
    ''' Readonly: relativer Top von projectLine
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YprojectName As Double
        Get
            YprojectName = _YprojectName
        End Get
    End Property

    ''' <summary>
    ''' Readonly: relativer Top des Phasen-Balkens
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YPhase As Double
        Get
            YPhase = _YPhase
        End Get
    End Property

    ''' <summary>
    ''' Readonly: relativer Top der Phasen-Beschriftung 
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YPhasenText As Double
        Get
            YPhasenText = _YPhasenText
        End Get
    End Property


    ''' <summary>
    ''' Readonly: relativer Top des Phasen-Datums  
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YPhasenDatum As Double
        Get
            YPhasenDatum = _YPhasenDatum
        End Get
    End Property


    ''' <summary>
    ''' Readonly: relativer Top des Meilenstein-Symbols  
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YMilestone As Double
        Get
            YMilestone = _YMilestone
        End Get
    End Property

    ''' <summary>
    ''' Readonly: relativer Top der Meilenstein-Beschriftung  
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YMilestoneText As Double
        Get
            YMilestoneText = _YMilestoneText
        End Get
    End Property

    ''' <summary>
    ''' Readonly: relativer Top des Meilenstein-Datums  
    ''' wird gesetzt in Methode bestimmeZeilenhoehe
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property YMilestoneDate As Double
        Get
            YMilestoneDate = _YMilestoneDate
        End Get
    End Property


    ''' <summary>
    ''' ermittelt die Koordinaten für Kalender, linker Rand Projektbeschriftung, Projekt-Fläche, Legenden-Fläche
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub bestimmeZeichenKoordinaten()

        ' bestimme Container Area und linker Rand der Projektliste
        If Not IsNothing(_containerShape) Then
            With _containerShape
                _containerLeft = .Left
                _containerRight = .Left + .Width
                _containerTop = .Top
                _containerBottom = .Top + .Height
                _projectListLeft = .Left + 10
            End With
        End If


        ' bestimme KalenderArea
        If Not IsNothing(_calendarLineShape) Then
            _calendarLeft = _calendarLineShape.Left
            _calendarRight = _calendarLineShape.Left + _calendarLineShape.Width

            If _PPTEndOFCalendar > _PPTStartOFCalendar Then
                _anzahlTageImKalender = CInt(DateDiff(DateInterval.Day, _PPTStartOFCalendar, _PPTEndOFCalendar))
                If _anzahlTageImKalender > 0 Then
                    _tagesbreite = calendarLineShape.Width / _anzahlTageImKalender
                End If
            End If

        Else

        End If

        ' _calendarTop = _containerTop + 5

        If Not IsNothing(_calendarHeightShape) Then
            _calendarTop = _calendarLineShape.Top - _calendarHeightShape.Height
        End If


        If Not IsNothing(_calendarHeightShape) Then
            _calendarBottom = _calendarLineShape.Top + _calendarLineShape.Height
        End If


        ' bestimme Drawing Area
        _drawingAreaLeft = _calendarLeft
        _drawingAreaRight = _calendarRight
        '_drawingAreaTop = _calendarBottom + 2
        _drawingAreaTop = _calendarLineShape.Top + _calendarLineShape.Height + 2


        If awinSettings.mppShowLegend And Not IsNothing(_legendLineShape) Then
            _drawingAreaBottom = _legendLineShape.Top - 5
        Else
            _drawingAreaBottom = _containerBottom - 10
        End If



        ' bestimme Legend Drawing Area 
        If awinSettings.mppShowLegend And Not IsNothing(_legendLineShape) Then
            _legendAreaTop = _legendLineShape.Top + (_containerBottom - _legendLineShape.Top) * 0.05
            _legendAreaBottom = _containerBottom - (_containerBottom - _legendLineShape.Top) * 0.1
            _legendAreaRight = System.Math.Min(_legendLineShape.Left + _legendLineShape.Width, _containerRight - 5)
        Else
            _legendAreaTop = _containerBottom - 5
            _legendAreaBottom = _containerBottom
        End If


        _legendAreaLeft = _drawingAreaLeft

    End Sub
    ' bestimmt Schriftart, Farbe, Größe der Projekt-Namen bzw. Swimlane-Beschriftung 
    ' bestimmt ausserdem den linken Rand der Text-Beschriftung
    Public Property projectNameVorlagenShape As pptNS.Shape

    ' Kalenderlinie; bestimmt Dicke, Farbe und Strichtyp der Kalenderbegrenzung; 
    ' bestimmt ausserdem den oberen sowie linken und rechten Rand der Zeichenfläche 
    Public Property calendarLineShape As pptNS.Shape
        Get
            calendarLineShape = _calendarLineShape
        End Get
        Set(value As pptNS.Shape)

            If Not IsNothing(value) Then
                _calendarLineShape = value

                ' dadurch werden die Koordinaten der Zeichenarea bestimmt 
                ' wenn die Daten für den Kalender bereits bekannt sind, wird dort auch die Tagesbreite gesetzt 
                Call bestimmeZeichenKoordinaten()


            End If
        End Set
    End Property

    ' bestimmt Schriftart, Farbe, Größe der Quartals/Monats-Beschriftung im Kalender  
    Public Property quarterMonthVorlagenShape As pptNS.Shape

    ' bestimmt Schriftart, Farbe, Größe der Jahres-Beschriftung im Kalender  
    Public Property yearVorlagenShape As pptNS.Shape

    ' bestimmt Dicke, Farbe und Strichtyp der Projekt-Linie bzw. Swimlane-Linie
    Public Property projectVorlagenShape As pptNS.Shape

    ' bestimmt Höhe eines Balkens für die Darstellung auf einer PPT
    Public Property phaseVorlagenShape As pptNS.Shape

    ' bestimmt Höhe eines Meilensteins für die Darstellung auf einer PPT
    Public Property milestoneVorlagenShape As pptNS.Shape

    ' bestimmt Form der Ampel, die dann mit der entsprechenden Farbe ausgefüllt wird 
    Public Property ampelVorlagenShape As pptNS.Shape

    ' bestimmt  Dicke, Farbe und Strichtyp der vertikalen Jahres-Linie auf dem Kalender 
    Public Property calendarYearSeparator As pptNS.Shape

    ' bestimmt  Dicke, Farbe und Strichtyp der vertikalen Quartals/Monats-/KW-Linie auf dem Kalender 
    Public Property calendarQuartalSeparator As pptNS.Shape

    ' bestimmt  Dicke, Farbe und Strichtyp der horizontalen Begrenzung einer Swimlane / Projektes 
    Public Property horizontalLineShape As pptNS.Shape

    ' bestimmt  Dicke, Farbe und Strichtyp der Legendenlinie
    ' markiert ausserdem das untere Ende der Zeichenfläche  
    Public Property legendLineShape As pptNS.Shape
        Get
            legendLineShape = _legendLineShape
        End Get
        Set(value As pptNS.Shape)
            If Not IsNothing(value) Then

                _legendLineShape = value

                If awinSettings.mppShowLegend Then
                    _drawingAreaBottom = _legendLineShape.Top - 5

                    _legendAreaTop = _legendLineShape.Top + (_containerBottom - _legendLineShape.Top) * 0.05
                    _legendAreaBottom = _containerBottom - (_containerBottom - _legendLineShape.Top) * 0.1
                    _legendAreaRight = System.Math.Min(_legendLineShape.Left + _legendLineShape.Width, _containerRight - 5)
                Else
                    _drawingAreaBottom = _containerBottom - 10

                    _legendAreaTop = _containerBottom - 5
                    _legendAreaBottom = _containerBottom
                    _legendAreaRight = _containerRight - 5
                End If

                _legendAreaLeft = _drawingAreaLeft


            End If
        End Set
    End Property

    ' bestimmt Schriftart, Farbe, Größe des Legenden-Titels 
    Public Property legendStartShape As pptNS.Shape

    ' bestimmt Schriftart, Farbe, Größe der Legenden Beschriftung 
    Public Property legendTextVorlagenShape As pptNS.Shape

    ' bestimmt Höhe und Breite eines Legenden-Balkens für die Darstellung auf einer PPT
    Public Property legendPhaseVorlagenShape As pptNS.Shape

    ' bestimmt Höhe und Breite eines Legenden-Balkens für die Darstellung auf einer PPT
    Public Property legendMilestoneVorlagenShape As pptNS.Shape

    ' bestimmt Höhe und Breite des Containers, in dem alles (Kalender, Projekt(e), Legende) gezeichnet wird 
    Public Property containerShape As pptNS.Shape
        Get
            containerShape = _containerShape
        End Get
        Set(value As pptNS.Shape)

            If Not IsNothing(value) Then

                _containerShape = value
                With _containerShape
                    _containerLeft = .Left
                    _containerRight = .Left + .Width
                    _containerTop = .Top
                    _containerBottom = .Top + .Height
                End With

            End If

        End Set
    End Property

    ' bestimmt Dicke, Farbe und Strichtyp der vertikalen Kalender- und Jahresbegrenzungen 
    Public Property calendarHeightShape As pptNS.Shape

    ' bestimmt Schriftart, Farbe, Größe der Milestone-Beschriftung 
    Public Property MsDescVorlagenShape As pptNS.Shape

    ' bestimmt Schriftart, Farbe, Größe des Meilenstein-Datums 
    Public Property MsDateVorlagenShape As pptNS.Shape

    ' bestimmt Schriftart, Farbe, Größe des Phasen-Namens 
    Public Property PhDescVorlagenShape As pptNS.Shape

    ' bestimmt Schriftart, Farbe, Größe des Phasen-Datums 
    Public Property PhDateVorlagenShape As pptNS.Shape

    ' bestimmt Stärke, Farbe, Strichart und Höhe der Kalenderunterteilung  
    Public Property calendarStepShape As pptNS.Shape

    ' bestimmt Farbe, Transparenz der Markierung von ShowrangeLeft/showrangeRight im Kalender 
    Public Property calendarMarkShape As pptNS.Shape

    ' bestimmt die Heute Linie, die in das PPT gezeichnet wird 
    Public Property todayLineShape As pptNS.Shape

    ' bestimmt Farbe, Schriftart und -Größe der Fehlermeldung  
    Public Property errorVorlagenShape As pptNS.Shape

    ' bestimmt Farbe, Höhe und Breite des Shapes, das zur Darstellung der Business Unit in der Legende verwendet wird 
    Public Property legendBuColorShape As pptNS.Shape

    ' bestimmt Farbe und Breite des Shapes, das zur Darstellung der Business Unit verwendet wird 
    Public Property buColorShape As pptNS.Shape

    ' bestimmt Farbe und Transparenz des Shapes, das zur Projekt-/Swimlane-Differenzierung verwendet werden soll 
    Public Property rowDifferentiatorShape As pptNS.Shape

    ' bestimmt Strichstärke, -Art und Farbe der Linie, die zur Markierung Phasen-Anfang/Ende verwendet werden soll 
    Public Property phaseDelimiterShape As pptNS.Shape

    ' bestimmt Form, Farbe, Pfeilspitzen der Linie, die zur Markierung der Dauer verwendet werden soll 
    Public Property durationArrowShape As pptNS.Shape

    ' bestimmt Schriftart, Farbe, Größe der Dauer-Beschriftung 
    Public Property durationTextShape As pptNS.Shape

    ' bestimmt Schriftart, Farbe, Größe der Segment Beschriftung (=Phasen der Hierarchie-Stude 1 eines Projektes)
    Public Property segmentVorlagenShape As pptNS.Shape

    Public Sub setShapesInvisible(Optional ByVal whichOnes As String = "")

        If whichOnes = "" Then
            Call makeShapeInvisible(_projectNameVorlagenShape)
            Call makeShapeInvisible(_calendarLineShape)
            Call makeShapeInvisible(_quarterMonthVorlagenShape)
            Call makeShapeInvisible(_yearVorlagenShape)
            Call makeShapeInvisible(_projectVorlagenShape)
            Call makeShapeInvisible(_phaseVorlagenShape)
            Call makeShapeInvisible(_milestoneVorlagenShape)
            Call makeShapeInvisible(_ampelVorlagenShape)
            Call makeShapeInvisible(_calendarYearSeparator)
            Call makeShapeInvisible(_calendarQuartalSeparator)
            Call makeShapeInvisible(_horizontalLineShape)
            Call makeShapeInvisible(_legendLineShape)
            Call makeShapeInvisible(_legendStartShape)
            Call makeShapeInvisible(_legendTextVorlagenShape)
            Call makeShapeInvisible(_legendPhaseVorlagenShape)
            Call makeShapeInvisible(_legendMilestoneVorlagenShape)

            If Not IsNothing(_containerShape) Then
                _containerShape.TextFrame2.TextRange.Text = ""
            End If

            Call makeShapeInvisible(_calendarHeightShape)
            Call makeShapeInvisible(_MsDescVorlagenShape)
            Call makeShapeInvisible(_MsDateVorlagenShape)
            Call makeShapeInvisible(_PhDescVorlagenShape)
            Call makeShapeInvisible(_PhDateVorlagenShape)
            Call makeShapeInvisible(_todayLineShape)
            Call makeShapeInvisible(_calendarStepShape)
            Call makeShapeInvisible(_calendarMarkShape)
            Call makeShapeInvisible(_errorVorlagenShape)
            Call makeShapeInvisible(_legendBuColorShape)
            Call makeShapeInvisible(_buColorShape)
            Call makeShapeInvisible(_rowDifferentiatorShape)
            Call makeShapeInvisible(_phaseDelimiterShape)
            Call makeShapeInvisible(_durationArrowShape)
            Call makeShapeInvisible(_durationTextShape)
            Call makeShapeInvisible(_segmentVorlagenShape)

            'If Not IsNothing(_quarterMonthVorlagenShape) Then
            '    _quarterMonthVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_yearVorlagenShape) Then
            '    _yearVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_projectVorlagenShape) Then
            '    _projectVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_phaseVorlagenShape) Then
            '    _phaseVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_milestoneVorlagenShape) Then
            '    _milestoneVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_ampelVorlagenShape) Then
            '    _ampelVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_calendarYearSeparator) Then
            '    _calendarYearSeparator.Delete()
            'End If

            'If Not IsNothing(_calendarQuartalSeparator) Then
            '    _calendarQuartalSeparator.Delete()
            'End If

            'If Not IsNothing(_horizontalLineShape) Then
            '    _horizontalLineShape.Delete()
            'End If

            'If Not IsNothing(_legendLineShape) Then
            '    _legendLineShape.Delete()
            'End If

            'If Not IsNothing(_legendStartShape) Then
            '    _legendStartShape.Delete()
            'End If

            'If Not IsNothing(_legendTextVorlagenShape) Then
            '    _legendTextVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_legendPhaseVorlagenShape) Then
            '    _legendPhaseVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_legendMilestoneVorlagenShape) Then
            '    _legendMilestoneVorlagenShape.Delete()
            'End If



            'If Not IsNothing(_calendarHeightShape) Then
            '    _calendarHeightShape.Delete()
            'End If

            'If Not IsNothing(_MsDescVorlagenShape) Then
            '    _MsDescVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_MsDateVorlagenShape) Then
            '    _MsDateVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_PhDescVorlagenShape) Then
            '    _PhDescVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_PhDateVorlagenShape) Then
            '    _PhDateVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_calendarStepShape) Then
            '    _calendarStepShape.Delete()
            'End If

            'If Not IsNothing(_calendarMarkShape) Then
            '    _calendarMarkShape.Delete()
            'End If

            'If Not IsNothing(_errorVorlagenShape) Then
            '    _errorVorlagenShape.Delete()
            'End If

            'If Not IsNothing(_legendBuColorShape) Then
            '    _legendBuColorShape.Delete()
            'End If

            'If Not IsNothing(_buColorShape) Then
            '    _buColorShape.Delete()
            'End If

            'If Not IsNothing(_rowDifferentiatorShape) Then
            '    _rowDifferentiatorShape.Delete()
            'End If

            'If Not IsNothing(_phaseDelimiterShape) Then
            '    _phaseDelimiterShape.Delete()
            'End If

            'If Not IsNothing(_durationArrowShape) Then
            '    _durationArrowShape.Delete()
            'End If

            'If Not IsNothing(_durationTextShape) Then
            '    _durationTextShape.Delete()
            'End If

            'If Not IsNothing(_segmentVorlagenShape) Then
            '    _segmentVorlagenShape.Delete()
            'End If

        End If
    End Sub


    Public Sub New()

        _projectNameVorlagenShape = Nothing
        _calendarLineShape = Nothing
        _quarterMonthVorlagenShape = Nothing
        _yearVorlagenShape = Nothing
        _projectVorlagenShape = Nothing
        _phaseVorlagenShape = Nothing
        _milestoneVorlagenShape = Nothing
        _ampelVorlagenShape = Nothing
        _calendarYearSeparator = Nothing
        _calendarQuartalSeparator = Nothing
        _horizontalLineShape = Nothing
        _legendLineShape = Nothing
        _legendStartShape = Nothing
        _legendTextVorlagenShape = Nothing
        _legendPhaseVorlagenShape = Nothing
        _legendMilestoneVorlagenShape = Nothing
        _containerShape = Nothing
        _calendarHeightShape = Nothing
        _MsDescVorlagenShape = Nothing
        _MsDateVorlagenShape = Nothing
        _PhDescVorlagenShape = Nothing
        _PhDateVorlagenShape = Nothing
        _calendarStepShape = Nothing
        _calendarMarkShape = Nothing
        _errorVorlagenShape = Nothing
        _todayLineShape = Nothing
        _legendBuColorShape = Nothing
        _buColorShape = Nothing
        _rowDifferentiatorShape = Nothing
        _phaseDelimiterShape = Nothing
        _durationArrowShape = Nothing
        _durationTextShape = Nothing
        _segmentVorlagenShape = Nothing

    End Sub
End Class
