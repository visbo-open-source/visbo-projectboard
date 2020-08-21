Imports System
Imports System.Math

Public Class clsProjekt
    Inherits clsProjektvorlage

    ' diese Variable würde die Variable aus der inherited Klasse clsProjektvorlage überschatten .. 
    ' deshalb auskommentiert 
    'Private _Dauer As Integer


    'Private AllPhases As List(Of clsPhase)
    'Private _relStart As Integer
    Private _imarge As Double
    Private _uuid As Long

    Private _StartOffset As Integer
    Private _Start As Integer
    Private _earliestStart As Integer
    Private _latestStart As Integer
    Private _Status As String
    Private _earliestStartDate As Date
    Private _startDate As Date
    Private _latestStartDate As Date
    ' Änderung tk: ist jetzt in der Phase 1 , Bewertung (last) abgespeichert 
    'Private _ampelStatus As Integer
    'Private _ampelErlaeuterung As String
    Private _name As String = "Project Dummy Name"
    Private _variantName As String = ""
    Private _variantDescription As String = ""
    ' Projektbeschreibung 
    Private _description As String = ""

    ' geändert 07.04.2014: Damit jedes Projekt auf der Projekttafel angezeigt werden kann.
    Private NullDatum As Date = StartofCalendar

    ' tk ergänzt am 12.6.18 eine Kundeninterne Nummer 
    ' kann als eine vom Kunden vergebene Kundenspezifische Projekt-Nummer benutzt werden
    Private _kundenNummer As String = ""
    Public Property kundenNummer As String
        Get
            kundenNummer = _kundenNummer
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _kundenNummer = value
            Else
                _kundenNummer = ""
            End If
        End Set
    End Property

    Public Overrides Property projectType As Integer
        Get
            projectType = _projectType
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                If value >= 0 And value <= 2 Then
                    _projectType = value
                Else
                    _projectType = ptPRPFType.project
                End If
            Else
                _projectType = ptPRPFType.project
            End If
        End Set
    End Property

    ' tk ergänzt am 9.6.18 actualDataUntil 
    ' gibt an, bis zu welchem Monat einschließlich die Ressourcen und Kostenbedarfs-Werte den Ist-Daten aus der Zeiterfassung entsprechen 
    Private _actualDataUntil As Date
    Public Property actualDataUntil As Date
        Get
            actualDataUntil = _actualDataUntil
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                If value > StartofCalendar Then
                    ' Actual Data Until kann aber nicht größer werden als das Projektende ...
                    ' das gilt aber nur, wenn das Projekt nicht gerade erst aufgebaut wird, es also schon eine Dauer gibt 
                    If dauerInDays > 0 Then
                        If endeDate < value Then
                            _actualDataUntil = endeDate
                        Else
                            _actualDataUntil = value
                        End If
                    Else
                        _actualDataUntil = value
                    End If

                Else
                    _actualDataUntil = Date.MinValue
                End If

            Else
                _actualDataUntil = Date.MinValue
            End If
        End Set
    End Property


    ' ergänzt am 20.8.17 
    ' Marker für Projekte, um anzuzeigen, dass es zu einer bestimmten Menge gehört ; wird nicht in der Datenbank gespeichert, kommt deshalb nicht in clsProjektDB vor
    Private _marker As Boolean = False
    Public Property marker As Boolean
        Get
            marker = _marker
        End Get
        Set(value As Boolean)
            _marker = value
        End Set
    End Property

    ' Kennzeichnung, ob ein Projekt manuell verschoben werden kann; wird nicht in der Datenbank gespeichert, kommt deshalb nicht in clsProjektDB vor
    Private _movable As Boolean = False
    Public Property movable As Boolean
        Get
            movable = _movable
        End Get
        Set(value As Boolean)
            If value = True Then
                If _Status = ProjektStatus(PTProjektStati.geplant) Or
                _Status = ProjektStatus(PTProjektStati.ChangeRequest) Or
                (_Status = ProjektStatus(PTProjektStati.beauftragt) And _variantName <> "") Or
                value = False Then
                    _movable = value

                Else
                    Dim errmsg As String
                    If awinSettings.englishLanguage Then
                        errmsg = "project status does not allow movement!"
                    Else
                        errmsg = "Projekt Status erlaubt keine Verschiebung / Dehnung / Kürzung"
                    End If
                    Throw New ArgumentException(errmsg)
                End If

            Else
                _movable = value
            End If

        End Set
    End Property


    ' die ShapeUID des Projektes  
    Private _shpUID As String = ""
    Public Property shpUID As String
        Get
            If Not IsNothing(_shpUID) Then
                shpUID = _shpUID
            Else
                shpUID = ""
            End If
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _shpUID = value
            Else
                _shpUID = ""
            End If

        End Set
    End Property

    Private _Risiko As Double = 0.0
    Public Property Risiko As Double
        Get
            If Not IsNothing(_Risiko) Then
                Risiko = _Risiko
            Else
                Risiko = 0.0
            End If
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                If value >= 0 And value < 10 Then
                    _Risiko = value
                Else
                    _Risiko = 5
                End If

            Else
                _Risiko = 0.0
            End If

        End Set
    End Property


    Private _StrategicFit As Double = 0.0
    Public Property StrategicFit As Double
        Get
            If Not IsNothing(_StrategicFit) Then
                StrategicFit = _StrategicFit
            Else
                StrategicFit = 0.0
            End If
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                If value >= 0 And value < 10 Then
                    _StrategicFit = value
                Else
                    _StrategicFit = 5
                End If

            Else
                _StrategicFit = 0.0
            End If

        End Set
    End Property

    ' hier war vorher die Property Erloes - wurde in clsProjektVorlage verschoben

    ''' <summary>
    ''' gibt die Budgetwerte des Projekts zurück
    ''' die werden 
    ''' beim Laden aus der Datenbank bestimmt oder 
    ''' beim Ändern des Erlös Werts 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property budgetWerte As Double()
        Get
            Dim costvalues() As Double = Me.getGesamtKostenBedarf()
            Dim gK As Double = costvalues.Sum
            Dim _budgetWerte() As Double
            ReDim _budgetWerte(_Dauer - 1)
            Dim avgBudget As Double = Me.Erloes / _Dauer
            Dim pMarge As Double = Me.ProjectMarge
            Dim riskCost As Double = Me.risikoKosten

            ' ProjectMarge = (Me.Erloes - gk) / gk

            For i As Integer = 0 To _Dauer - 1
                If gK > 0 Then
                    _budgetWerte(i) = costvalues(i) * (1 + pMarge)
                Else
                    _budgetWerte(i) = avgBudget
                End If
            Next

            budgetWerte = _budgetWerte
        End Get
    End Property

    Private _leadPerson As String = ""
    Public Property leadPerson As String
        Get
            If Not IsNothing(_leadPerson) Then
                leadPerson = _leadPerson
            Else
                leadPerson = ""
            End If
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _leadPerson = value
            Else
                _leadPerson = ""
            End If

        End Set
    End Property
    'Public Property tfSpalte As Integer

    Private _tfZeile As Integer = 2
    ''' <summary>
    ''' muss immer richtig gesetzt sein; wird verwendet um Projekt, Phasne und Meilensteine zu zeichnen 
    ''' wenn es neu gesetzt wird, werden auch die aktuelle Constellation und die "Hintergrund-Constellation" entsprechend gesetzt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property tfZeile As Integer
        Get
            If Not IsNothing(_tfZeile) Then
                tfZeile = _tfZeile
            Else
                tfZeile = 2
            End If
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                If value >= 2 Then
                    _tfZeile = value
                Else
                    _tfZeile = 2
                End If

            Else
                _tfZeile = 2
            End If

            ' die tfzeile wird immer aufgrund der constellationItem.zeile gesetzt 
            ' '' jetzt werden die currentConstellationsession und von der currentConstellation Werte entsprechend gesetzt 
            Dim key As String = calcProjektKey(Me.name, Me.variantName)
            currentSessionConstellation.updateTFzeile(key, _tfZeile)

            Dim tmpConst As clsConstellation =
                projectConstellations.getConstellation(currentConstellationName)
            If Not IsNothing(tmpConst) Then
                tmpConst.updateTFzeile(key, _tfZeile)
            End If

        End Set
    End Property

    Private _Id As String = ""
    Public Property Id As String
        Get
            If Not IsNothing(_Id) Then
                Id = _Id
            Else
                Id = ""
            End If
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _Id = value
            Else
                _Id = ""
            End If
        End Set
    End Property

    Private _timeStamp As Date = Date.Now
    Public Property timeStamp As Date
        Get
            If Not IsNothing(_timeStamp) Then
                timeStamp = _timeStamp
            Else
                timeStamp = Date.Now
            End If
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                _timeStamp = value
            Else
                _timeStamp = Date.Now
            End If
        End Set
    End Property

    ' ergänzt am 26.10.13 - nicht in Vorlage aufgenommen, da es für jedes Projekt individuell ist 

    Private _volume As Double = 0.0
    Public Property volume As Double
        Get
            If Not IsNothing(_volume) Then
                volume = _volume
            Else
                volume = 0.0
            End If
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                _volume = value
            Else
                _volume = 0.0
            End If
        End Set
    End Property

    Private _complexity As Double = 0.0
    Public Property complexity As Double
        Get
            If Not IsNothing(_complexity) Then
                complexity = _complexity
            Else
                complexity = 0.0
            End If
        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                _complexity = value
            Else
                _complexity = 0.0
            End If
        End Set
    End Property

    Private _businessUnit As String = ""
    Public Property businessUnit As String
        Get
            If Not IsNothing(_businessUnit) Then
                businessUnit = _businessUnit
            Else
                businessUnit = ""
            End If
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _businessUnit = value
            Else
                _businessUnit = ""
            End If
        End Set
    End Property


    Private _updatedAt As String
    Public Property updatedAt As String
        Get
            updatedAt = _updatedAt
        End Get
        Set(value As String)
            _updatedAt = value
        End Set
    End Property


    ''''  Definitionen zu einem Projekt, die nicht in der DB abgespeichert werden

    ' ergänzt am 30.1.14 - diffToPrev , wird benutzt, um zu kennzeichnen , welches Projekt sich im Vergleich zu vorher verändert hat 

    'Private _diffToPrev As Boolean = False
    'Public Property diffToPrev As Boolean
    '    Get
    '        If Not IsNothing(_diffToPrev) Then
    '            diffToPrev = _diffToPrev
    '        Else
    '            diffToPrev = False
    '        End If
    '    End Get
    '    Set(value As Boolean)
    '        If Not IsNothing(value) Then
    '            _diffToPrev = value
    '        Else
    '            _diffToPrev = False
    '        End If
    '    End Set
    'End Property

    ' ergänzt am 16.09.2015 - extendedView , wird benutzt, um zu kennzeichnen , welches Projekt in extended View dargestellt werden soll
    Private _extendedView As Boolean = False
    Public Property extendedView As Boolean
        Get
            If Not IsNothing(_extendedView) Then
                extendedView = _extendedView
            Else
                extendedView = False
            End If
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _extendedView = value
            Else
                _extendedView = False
            End If
        End Set
    End Property


    ' ergänzt am 14.10.2019: keyMetrics wird beim Speichern berechnet und in der ProjektVersion immer gespeichert
    'Private _keyMetrics As clsKeyMetrics
    Public Property keyMetrics As clsKeyMetrics
    '    Get
    '        If Not IsNothing(_keyMetrics) Then
    '            keyMetrics = _keyMetrics
    '        Else
    '            keyMetrics = New clsKeyMetrics
    '        End If
    '    End Get
    '    Set(value As clsKeyMetrics)
    '        If Not IsNothing(value) Then
    '            _keyMetrics = value
    '        Else
    '            _keyMetrics = Nothing
    '        End If
    '    End Set
    'End Property

    ''' <summary>
    ''' prüft, ob zwei Projekte die identische Kunden-Nummer haben 
    ''' </summary>
    ''' <param name="vProj"></param>
    ''' <returns></returns>
    Public ReadOnly Property hasIdenticalKdNr(ByVal vProj As clsProjekt) As Boolean
        Get
            Dim tmpResult As Boolean = False

            If Not IsNothing(vProj) Then
                tmpResult = (Me.kundenNummer = vProj.kundenNummer)
            End If

            hasIdenticalKdNr = tmpResult
        End Get
    End Property
    ''' <summary>
    ''' prüft, ob ein Projekt in allen Belangen genau identisch mit einem anderen Projekt ist
    ''' wird benutzt, um zu prüfen, ob gespeichert werden soll oder nicht ... 
    ''' </summary>
    ''' <param name="vProj"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vProj As clsProjekt) As Boolean
        Get
            Dim stillOK As Boolean = False

            Try
                If Not IsNothing(vProj) Then

                    If Me.name = vProj.name Then
                        If Me.kundenNummer <> vProj.kundenNummer Then

                        End If
                    End If

                    If Me.name = vProj.name And
                           Me.variantName = vProj.variantName And
                           Me.variantDescription = vProj.variantDescription And
                           Me.description = vProj.description And
                           Me.projectType = vProj.projectType And
                           DateDiff(DateInterval.Month, Me.actualDataUntil, vProj.actualDataUntil) = 0 And
                           Me.kundenNummer = vProj.kundenNummer Then

                        If Me.startDate.Date = vProj.startDate.Date And
                               Me.endeDate.Date = vProj.endeDate.Date Then

                            If Me.ampelStatus = vProj.ampelStatus And
                                    Me.ampelErlaeuterung = vProj.ampelErlaeuterung Then

                                ' es soll nur auf Budget Gelichheit geprüft werden , die Verteilun g macht doch an der Stelle gar keinen Sinn .. . 
                                ' If (Not arraysAreDifferent(Me.budgetWerte, .budgetWerte) Or IsNothing(Me.budgetWerte) Or IsNothing(.budgetWerte)) And
                                If Me.Erloes = vProj.Erloes Then
                                    ' If (Not arraysAreDifferent(Me.budgetWerte, .budgetWerte) Or IsNothing(Me.budgetWerte) Or IsNothing(.budgetWerte)) And
                                    'Me.Erloes = .Erloes Then

                                    If Me.businessUnit = vProj.businessUnit And
                                            Me.complexity = vProj.complexity And
                                            Me.Status = vProj.Status And
                                            Me.StrategicFit = vProj.StrategicFit And
                                            Me.Risiko = vProj.Risiko And
                                            Me.VorlagenName = vProj.VorlagenName And
                                            Me.volume = vProj.volume And
                                            Me.leadPerson = vProj.leadPerson Then

                                        stillOK = True

                                        ' tk, 30.12.16 das wurde jetzt rausgenommen ... das wird ja bis auf weiteres überhaupt nicht gebraucht 
                                        'Me.earliestStartDate = .earliestStartDate And _
                                        'Me.latestStartDate = .latestStartDate And _

                                    End If


                                End If

                            End If

                        End If

                    End If


                    ' jetzt die Phasen prüfen, dann die Meilensteine 
                    If stillOK And Me.CountPhases = vProj.CountPhases Then

                        Dim pNr As Integer = 1
                        Do While stillOK And pNr <= Me.CountPhases
                            Dim cPhase As clsPhase = Me.getPhase(pNr)
                            Dim vPhase As clsPhase = vProj.getPhase(pNr)
                            If cPhase.isIdenticalTo(vPhase) Then
                                ' alles ok 
                                pNr = pNr + 1
                            Else
                                stillOK = False
                            End If
                        Loop

                    Else
                        stillOK = False
                    End If

                    ' jetzt die Custom Fields prüfen 
                    If stillOK And
                            Me.customBoolFields.Count = vProj.customBoolFields.Count And
                            Me.customDblFields.Count = vProj.customDblFields.Count And
                            Me.customStringFields.Count = vProj.customStringFields.Count Then
                        ' alle sind gleich , detaillierte Überprüfung lohnt 


                        ' String CustomFields
                        Dim ix As Integer = 0
                        Do While stillOK And ix <= Me.customStringFields.Count - 1

                            Dim cFMe As KeyValuePair(Of Integer, String) = Me.customStringFields.ElementAt(ix)
                            Dim cFVgl As KeyValuePair(Of Integer, String) = vProj.customStringFields.ElementAt(ix)

                            If Not IsNothing(cFMe.Value) And Not IsNothing(cFVgl.Value) Then
                                Try
                                    If cFMe.Key = cFVgl.Key And cFMe.Value.Trim = cFVgl.Value.Trim Then
                                        ix = ix + 1
                                    Else
                                        stillOK = False
                                    End If
                                Catch ex As Exception
                                    stillOK = False
                                End Try
                            ElseIf IsNothing(cFMe.Value) And IsNothing(cFVgl.Value) Then
                                ' unverändert lassen 
                            Else
                                ' eines is Nothing , eines ist nicht Nothing, also unterscheidelich 
                                stillOK = False
                            End If

                        Loop


                        If stillOK Then
                            ' prüfe Double Custom Fields
                            ix = 0
                            Do While stillOK And ix <= Me.customDblFields.Count - 1
                                Dim cFMe As KeyValuePair(Of Integer, Double) = Me.customDblFields.ElementAt(ix)
                                Dim cFVgl As KeyValuePair(Of Integer, Double) = vProj.customDblFields.ElementAt(ix)

                                If cFMe.Key = cFVgl.Key And cFMe.Value = cFVgl.Value Then
                                    ix = ix + 1
                                Else
                                    stillOK = False
                                End If
                            Loop

                            If stillOK Then
                                ' prüfe Bool Custom fields
                                ix = 0
                                Do While stillOK And ix <= Me.customBoolFields.Count - 1
                                    Dim cFMe As KeyValuePair(Of Integer, Boolean) = Me.customBoolFields.ElementAt(ix)
                                    Dim cFVgl As KeyValuePair(Of Integer, Boolean) = vProj.customBoolFields.ElementAt(ix)

                                    If cFMe.Key = cFVgl.Key And cFMe.Value = cFVgl.Value Then
                                        ix = ix + 1
                                    Else
                                        stillOK = False
                                    End If
                                Loop
                            End If
                        End If


                    Else
                        stillOK = False
                    End If



                Else    ' vproj = nothing
                    stillOK = False
                End If

            Catch ex As Exception

                stillOK = False

            End Try


            isIdenticalTo = stillOK

        End Get
    End Property

    ''' <summary>
    ''' gibt für die übergebenen Listen an Phasen und Meilensteinen das früheste bzw. späteste Datum zurück, das in den 
    ''' aufgeführten Phasen bzw. Meilensteinen existiert; 
    ''' ausserdem wird die Dauer in Tagen zwischen minDate und maxDate zurückgegeben 
    ''' wenn nicht wenigstens zwei unterschiedliche Daten existieren , wird 0 als Länge zurückgegeben  
    ''' </summary>
    ''' <param name="selPhases">Liste der Phasen Namen</param>
    ''' <param name="selMilestones">Liste der Meilenstein Namen</param>
    ''' <param name="minDate"></param>
    ''' <param name="maxDate"></param>
    ''' <param name="durationInDays"></param>
    ''' <remarks></remarks>
    Public Sub getMinMaxDatesAndDuration(ByVal selPhases As Collection, ByVal selMilestones As Collection,
                                             ByRef minDate As Date, ByRef maxDate As Date, ByRef durationInDays As Long)

        Dim earliestDate As Date = Me.endeDate.AddMonths(1)
        Dim latestDate As Date = Me.startDate.AddMonths(-1)
        Dim earliestfound As Boolean = False
        Dim latestfound As Boolean = False
        Dim tmpStartDate As Date
        Dim tmpEndDate As Date
        Dim phaseName As String = ""
        Dim fullPhaseName As String
        Dim cphase As clsPhase

        ' Phasen Information untersuchen 


        For ix As Integer = 1 To selPhases.Count

            fullPhaseName = CStr(selPhases.Item(ix))

            Dim breadcrumb As String = ""
            Dim type As Integer = -1
            Dim pvName As String = ""
            Call splitHryFullnameTo2(fullPhaseName, phaseName, breadcrumb, type, pvName)

            If type = -1 Or
                (type = PTItemType.projekt And pvName = Me.name) Or
                (type = PTItemType.vorlage And pvName = Me.VorlagenName) Then

                Dim phaseIndices() As Integer = Me.hierarchy.getPhaseIndices(phaseName, breadcrumb)

                For px As Integer = 0 To phaseIndices.Length - 1

                    cphase = Me.getPhase(phaseIndices(px))

                    If Not IsNothing(cphase) Then
                        Try
                            tmpStartDate = cphase.getStartDate
                            tmpEndDate = cphase.getEndDate

                            If DateDiff(DateInterval.Day, tmpStartDate, earliestDate) > 0 Then
                                earliestDate = tmpStartDate
                                earliestfound = True
                            End If

                            If DateDiff(DateInterval.Day, latestDate, tmpEndDate) > 0 Then
                                latestDate = tmpEndDate
                                latestfound = True
                            End If

                        Catch ex As Exception
                            ' nichts tun 
                        End Try
                    Else
                        ' nichts tun
                    End If


                Next
            ElseIf type = PTItemType.categoryList Then

                ' alle IDs von Phasen holen, die die Darstellungsklasse haben .. wird in splithryfullnameto2 rausgeholt 
                Dim idCollection As Collection = Me.getPhaseIDsWithCat(pvName)

                For Each tmpID As String In idCollection
                    cphase = Me.getPhaseByID(tmpID)

                    If Not IsNothing(cphase) Then
                        Try
                            tmpStartDate = cphase.getStartDate
                            tmpEndDate = cphase.getEndDate

                            If DateDiff(DateInterval.Day, tmpStartDate, earliestDate) > 0 Then
                                earliestDate = tmpStartDate
                                earliestfound = True
                            End If

                            If DateDiff(DateInterval.Day, latestDate, tmpEndDate) > 0 Then
                                latestDate = tmpEndDate
                                latestfound = True
                            End If

                        Catch ex As Exception
                            ' nichts tun 
                        End Try
                    Else
                        ' nichts tun
                    End If
                Next

            End If


        Next


        ' Meilensteine schreiben 
        Dim fullMsName As String
        Dim msName As String = ""
        Dim milestone As clsMeilenstein = Nothing

        For ix As Integer = 1 To selMilestones.Count
            fullMsName = CStr(selMilestones.Item(ix))

            Dim breadcrumb As String = ""
            Dim type As Integer = -1
            Dim pvName As String = ""
            Call splitHryFullnameTo2(fullMsName, msName, breadcrumb, type, pvName)

            If type = -1 Or
                (type = PTItemType.projekt And pvName = Me.name) Or
                (type = PTItemType.vorlage And pvName = Me.VorlagenName) Then

                Dim milestoneIndices(,) As Integer = Me.hierarchy.getMilestoneIndices(msName, breadcrumb)
                ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste

                For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                    milestone = Me.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))

                    If Not IsNothing(milestone) Then
                        Try
                            tmpStartDate = milestone.getDate

                            If DateDiff(DateInterval.Day, tmpStartDate, earliestDate) > 0 Then
                                earliestDate = tmpStartDate
                                earliestfound = True
                            End If

                            If DateDiff(DateInterval.Day, latestDate, tmpStartDate) > 0 Then
                                latestDate = tmpStartDate
                                latestfound = True
                            End If

                        Catch ex As Exception
                            ' nichts tun
                        End Try
                    Else
                        ' nichts tun 

                    End If

                Next

            ElseIf type = PTItemType.categoryList Then

                ' alle IDs von Phasen holen, die die Darstellungsklasse haben .. wird in splithryfullnameto2 rausgeholt 
                Dim idCollection As Collection = Me.getMilestoneIDsWithCat(pvName)

                For Each tmpID As String In idCollection

                    milestone = Me.getMilestoneByID(tmpID)

                    If Not IsNothing(milestone) Then
                        Try
                            tmpStartDate = milestone.getDate

                            If DateDiff(DateInterval.Day, tmpStartDate, earliestDate) > 0 Then
                                earliestDate = tmpStartDate
                                earliestfound = True
                            End If

                            If DateDiff(DateInterval.Day, latestDate, tmpStartDate) > 0 Then
                                latestDate = tmpStartDate
                                latestfound = True
                            End If

                        Catch ex As Exception
                            ' nichts tun
                        End Try
                    Else
                        ' nichts tun 

                    End If

                Next

            End If


        Next


        If earliestfound And latestfound Then
            durationInDays = DateDiff(DateInterval.Day, earliestDate, latestDate)
        Else
            durationInDays = 0
        End If

        minDate = earliestDate
        maxDate = latestDate


    End Sub

    ''' <summary>
    ''' filtert die übergebene Liste an IDs so , dass hinterher nur Elemente enthalten sind, die auch im Zeitraum liegen  
    ''' </summary>
    ''' <param name="todoCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property filterbyZeitraum(ByVal todoCollection As Collection) As Collection
        Get
            Dim tmpCollection As New Collection

            ' prüfen, ob Showranges gültige Werte haben, wenn nein, wird die todoCollection gar nicht gefiltert
            If showRangeLeft > 0 And showRangeRight > showRangeLeft Then

                For Each tmpID As String In todoCollection

                    If elemIDIstMeilenstein(tmpID) Then
                        ' es geht um einen Meilenstein 
                        Dim milestone As clsMeilenstein = Me.getMilestoneByID(tmpID)
                        If Not IsNothing(milestone) Then
                            If milestoneWithinTimeFrame(milestone.getDate, showRangeLeft, showRangeRight) Then
                                Try
                                    ' da es eigentlich gar nicht vorkommen kann, dass es bereits enthalten ist, wird auf den contains Aufruf verzichtet
                                    ' in diesem Fall wäre das langsamer, da contains jedesmal aufgerufen wird, der Try aber nur im eigentlich 
                                    ' gar nicht vorkommenden Fehlerfall zuschlägt
                                    tmpCollection.Add(tmpID, tmpID)
                                Catch ex As Exception

                                End Try

                            End If
                        End If

                    Else
                        ' es handelt sich um eine Phase
                        Dim cPhase As clsPhase = Me.getPhaseByID(tmpID)
                        If Not IsNothing(cPhase) Then
                            If phaseWithinTimeFrame(Me.Start, cPhase.relStart, cPhase.relEnde,
                                                     showRangeLeft, showRangeRight) Then
                                Try
                                    ' da es eigentlich gar nicht vorkommen kann, dass es bereits enthalten ist, wird auf den contains Aufruf verzichtet
                                    ' in diesem Fall wäre das langsamer, da contains jedesmal aufgerufen wird, der Try aber nur im eigentlich 
                                    ' gar nicht vorkommenden Fehlerfall zuschlägt
                                    tmpCollection.Add(tmpID, tmpID)
                                Catch ex As Exception

                                End Try
                            End If
                        End If
                    End If
                Next

            Else
                For Each tmpID As String In todoCollection
                    Try
                        ' da es eigentlich gar nicht vorkommen kann, dass es bereits enthalten ist, wird auf den contains Aufruf verzichtet
                        ' in diesem Fall wäre das langsamer, da contains jedesmal aufgerufen wird, der Try aber nur im eigentlich 
                        ' gar nicht vorkommenden Fehlerfall zuschlägt
                        tmpCollection.Add(tmpID, tmpID)
                    Catch ex As Exception

                    End Try
                Next
            End If

            filterbyZeitraum = tmpCollection

        End Get
    End Property

    ''' <summary>
    ''' synchronisiert die Arrays mit der evtl veränderten Array Länge durch eine Verschiebung des Projekts 
    ''' berechnet und bestimmt die XWerte der Rollen und Kostenarten für die Phasen neu
    ''' wird aus set Startdate heraus aufgerufen; dadurch kann es sein, daß sich die 
    ''' Dimension der Arrays für die Rollen und kostenarten verändert 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub syncXWertePhases()
        Dim tmpValue As Boolean = True
        Dim cphase As clsPhase
        Dim dimension As Integer
        Dim phaseStart As Date, phaseEnd As Date


        ' prüfen, ob die Gesamtlänge übereinstimmt  
        For p = 1 To Me.CountPhases
            cphase = Me.getPhase(p)
            phaseEnd = cphase.getEndDate
            phaseStart = cphase.getStartDate

            dimension = getColumnOfDate(phaseEnd) - getColumnOfDate(phaseStart)

            If cphase.countRoles > 0 Or cphase.countCosts > 0 Then

                ' hier müssen jetzt die Xwerte neu gesetzt werden 
                Call cphase.calcNewXwerte(dimension, 1)

            End If




        Next




    End Sub
    ''' <summary>
    ''' liest / schreibt die Description eines Projektes
    ''' stellt sicher, dass es niemals Null sein kann 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property description As String
        Get
            If IsNothing(_description) Then
                _description = ""
            End If
            description = _description
        End Get

        Set(value As String)
            If IsNothing(value) Then
                _description = ""
            Else
                Try
                    If value.Trim.Length > 0 Then
                        _description = value.Trim

                    Else
                        _description = ""
                    End If

                Catch ex As Exception
                    _description = ""
                End Try
            End If
        End Set
    End Property

    ''' <summary>
    ''' gibt als Erläuterung die volle Description, Projekt- plus Varianten - Beschreibung zurück; 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property fullDescription As String
        Get
            If IsNothing(_description) Then
                _description = ""
            End If
            If IsNothing(_variantDescription) Or _variantName = "" Then
                fullDescription = _description
            Else
                fullDescription = _description & "; [" & _variantDescription & "]"
            End If

        End Get
    End Property

    ''' <summary>
    ''' stellt sicher, daß variantName niemals Nothing sein kann
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property variantName As String
        Get
            If IsNothing(_variantName) Then
                _variantName = ""
            End If
            variantName = _variantName
        End Get

        Set(value As String)

            If IsNothing(value) Then
                _variantName = ""
            Else
                Try
                    If value.Trim.Length > 0 Then
                        _variantName = value.Trim

                    Else
                        _variantName = ""
                    End If

                Catch ex As Exception
                    _variantName = ""
                End Try
            End If


        End Set
    End Property



    ''' <summary>
    ''' stellt sicher, daß variantDescription niemals Nothing sein kann
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property variantDescription As String
        Get

            If IsNothing(_variantDescription) Then
                _variantDescription = ""
            End If
            variantDescription = _variantDescription


        End Get

        Set(value As String)

            If IsNothing(value) Then
                _variantDescription = ""
            Else
                Try
                    If value.Trim.Length > 0 Then
                        _variantDescription = value.Trim

                    Else
                        _variantDescription = ""
                    End If

                Catch ex As Exception
                    _variantDescription = ""
                End Try
            End If


        End Set
    End Property

    ''' <summary>
    ''' gibt den Text für das Shape zurück; 
    ''' ist entweder nur der Projektname, oder aber der Projektname ( Varianten-Name ) 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShapeText() As String
        Get

            If Not IsNothing(Me.variantName) Then
                If Me.variantName.Length > 0 Then
                    getShapeText = Me.name & "[ " & Me.variantName & " ]"
                Else
                    getShapeText = Me.name
                End If
            Else
                getShapeText = Me.name
            End If

        End Get
    End Property

    ''' <summary>
    ''' setzt den Namen des Projektes fest oder gibt ihn zurück
    ''' gleichzeitig wird auch der Name der Phase(1),  auf den Namen "rootPhaseName" festgesetzt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property name As String
        Get
            name = _name
        End Get

        Set(value As String)

            Try
                If value.Trim.Length > 0 Then
                    If isValidProjectName(value.Trim) Then
                        _name = value.Trim
                    Else
                        Dim msgTxt As String = ""
                        If awinSettings.englishLanguage Then
                            msgTxt = "name must not contain any #, (, or )-characters"
                        Else
                            msgTxt = "Name darf keine #, (, or )-Zeichen enthalten"
                        End If
                        Throw New ArgumentException(msgTxt)
                    End If
                Else
                    _name = ""
                End If

            Catch ex As Exception
                _name = ""
            End Try

            If Me.CountPhases > 0 Then
                ' Änderung 13.4.15 Root Phasen Namen heisst immer so, nicht mehr wie Projekt: 
                'Me.getPhase(1).name = _name
                Me.getPhase(1).nameID = rootPhaseName
            End If


        End Set
    End Property


    ''' <summary>
    ''' prüft , ob das Projekt in seinen Werten konsistent ist
    ''' es ist nicht konsistent, wenn 
    ''' Dauer nicht gleich Monat(Ende)-Monat(Start +1 
    ''' die Dimensionen der Rollen/Kosten Xwerte nicht gleich Dauer der Phase in Monaten ist 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isConsistent As Boolean

        Get

            Dim tmpValue As Boolean = True
            Dim p As Integer = 1

            ' prüfen, ob die Gesamtlänge übereinstimmt  
            If Me.anzahlRasterElemente <> getColumnOfDate(Me.endeDate) - getColumnOfDate(Me.startDate) + 1 Then
                tmpValue = False
            End If

            ' prüfen, ob die Xwerte der Kosten und Rollen zu der Phasenlänge passt   

            While tmpValue And p <= Me.CountPhases
                tmpValue = Me.getPhase(p).isConsistent
                p = p + 1
            End While

            isConsistent = tmpValue

        End Get

    End Property

    Public Overrides Sub AddPhase(ByVal phase As clsPhase,
                                  Optional ByVal origName As String = "",
                                  Optional ByVal parentID As String = "")

        Dim phaseEnde As Double
        Dim maxM As Integer

        ' wenn der Origname gesetzt werden soll ...
        If origName <> "" Then
            If phase.originalName <> origName Then
                phase.originalName = origName
            End If
        End If

        With phase

            phaseEnde = .startOffsetinDays + .dauerInDays - 1

            For m = 1 To .countMilestones
                If phaseEnde < .startOffsetinDays + .getMilestone(m).offset Then
                    phaseEnde = .startOffsetinDays + .getMilestone(m).offset
                End If
            Next

        End With

        If phaseEnde > 0 Then

            maxM = CInt(DateDiff(DateInterval.Month, Me.startDate, Me.startDate.AddDays(phaseEnde)) + 1)
            If maxM <> _Dauer And maxM > 0 Then
                _Dauer = maxM
                ' hier muss jetzt die Dauer der Allgemeinen Phase angepasst werden ... 
            End If
        End If


        AllPhases.Add(phase)

        ' jetzt muss die Phase in die Projekt-Hierarchie aufgenommen werden 
        Dim currentElementNode As New clsHierarchyNode
        With currentElementNode

            If Me.CountPhases = 1 Then
                .elemName = "."
            Else
                .elemName = phase.name
            End If

            ' Änderung tk 29.5.16 origName ist nicht mehr Bestandteil von HierarchyNode, 
            ''If origName = "" Then
            ''    .origName = .elemName
            ''Else
            ''    .origName = origName
            ''End If

            .indexOfElem = Me.CountPhases

            If parentID = "" Then
                If .indexOfElem = 1 Then
                    .parentNodeKey = ""
                Else
                    .parentNodeKey = rootPhaseName
                End If
            Else
                .parentNodeKey = parentID
            End If

        End With

        With Me.hierarchy
            .addNode(currentElementNode, phase.nameID)
        End With

        ' jetzt müssen noch alle bereits in der Phase existierenden Meilensteine aufgenommen werden 
        For m As Integer = 1 To phase.countMilestones
            Dim cmilestone As clsMeilenstein = phase.getMilestone(m)
            currentElementNode = New clsHierarchyNode

            With currentElementNode

                .elemName = elemNameOfElemID(cmilestone.nameID)
                '.origName = .elemName
                .indexOfElem = m
                .parentNodeKey = phase.nameID

            End With

            With Me.hierarchy
                .addNode(currentElementNode, cmilestone.nameID)
            End With

        Next

    End Sub

    ''' <summary>
    ''' trägt die Liste von clsCustomFields, die in der Collection sind, in das Projekt ein
    ''' </summary>
    ''' <param name="listOfCustomFields"></param>
    Public Sub addListOfCustomFields(ByVal listOfCustomFields As Collection)
        If Not IsNothing(listOfCustomFields) Then

            If listOfCustomFields.Count > 0 Then

                For Each cfObj As clsCustomField In listOfCustomFields

                    Try
                        Dim uniqueID As Integer = CInt(cfObj.uid)
                        Dim cfType As Integer = customFieldDefinitions.getTyp(uniqueID)

                        Select Case cfType

                            Case ptCustomFields.Str
                                Me.addSetCustomSField(uniqueID, CStr(cfObj.wert))
                            Case ptCustomFields.Dbl
                                Me.addSetCustomDField(uniqueID, CDbl(cfObj.wert))
                            Case ptCustomFields.bool
                                Me.addSetCustomBField(uniqueID, CBool(cfObj.wert))
                            Case Else

                        End Select
                    Catch ex As Exception

                    End Try

                Next

            End If

        End If
    End Sub

    ''' <summary>
    ''' setzt das Budget auf den Wert, der sich aus den Ressourcen- und Kostenbedarfen ergibt
    ''' </summary>
    Public Sub setBudgetAsNeeded()

        Try
            Dim a As Integer = Me.dauerInDays
            Dim neededBudget As Double = 0.0, tmpERL As Double, tmpPK As Double, tmpOK As Double, tmpRK As Double, tmpERG As Double

            Call Me.calculateRoundedKPI(tmpERL, tmpPK, tmpOK, tmpRK, tmpERG)

            neededBudget = tmpPK + tmpOK
            ' tk 15.6.2020 die folgende Zeile setzt doch nicht das richtige Budget !? 
            'If tmpERG < 0 Then
            '    neededBudget = -1 * tmpERG
            'End If
            Erloes = neededBudget
        Catch ex As Exception

            If awinSettings.visboDebug Then
                Call MsgBox("Fehler in Projekt anlegen, Name: " & Me.name)
            End If

        End Try
    End Sub

    '''' <summary>
    '''' fügt dem aktuellen Projekt Me  , der existierenden Phase nameID die Rolle bzw Kostenart zu;
    '''' wenn addWhenexisting true, wird addiert, andernfalls replaced 
    '''' Vorbedingung: alle Plausibilitätsbedingungen wurden im Vorfeld abgeklärt, also Phase existiert, Rolle/Kostenart existiert und Summe ist positiv 
    '''' </summary>
    '''' <param name="phaseNameID"></param>
    '''' <param name="rcNameID">wenn Rolle: uid.tostring; teamID.tostring oder roleUid.tostring</param>(
    '''' <param name="summe"></param>
    '''' <param name="addWhenExisting"></param>
    'Public Sub addCostRoleToPhase(ByVal phaseNameID As String, ByVal rcNameID As String, ByVal summe As Double,
    '                          ByVal isrole As Boolean,
    '                          ByVal addWhenExisting As Boolean)

    '    ' es werden die Plausibilitätsprüfungen gemacht 
    '    Dim cphase As clsPhase = Me.getPhaseByID(phaseNameID)

    '    If Not IsNothing(cphase) Then
    '        If isrole Then
    '            ' eine Rolle wird hinzugefügt 
    '            Call cphase.AddRole(rcNameID, summe, addWhenExisting)

    '        Else
    '            ' eine Kostenart wird hinzugefügt
    '            Call cphase.AddCost(rcNameID, summe, addWhenExisting)
    '        End If
    '    Else

    '    End If

    'End Sub

    ''' <summary>
    ''' löscht in allen Phasen alle vorkommenden Rollen und Kosten
    ''' 
    ''' </summary>
    Public Sub deleteAllRolesCosts()

        For ip As Integer = 1 To Me.CountPhases
            Dim cPhase As clsPhase = Me.getPhase(ip)
            With cPhase

            End With
        Next

    End Sub

    ''' <summary>
    ''' gibt ein PRojekt zurück, wo die angegebenen Rollen , ggf. inkl Kinder, und die Kosten gelöscht werden 
    ''' </summary>
    ''' <param name="rolesToBeDeleted">werden in der Form string: uid;teamID übergeben</param>
    ''' <param name="costsToBedeleted">werden in der Form uid übergeben</param>
    ''' <param name="includingChilds"></param>
    ''' <returns></returns>
    Public Function deleteRolesAndCosts(ByVal rolesToBeDeleted As Collection,
                                        ByVal costsToBedeleted As Collection,
                                        ByVal includingChilds As Boolean) As clsProjekt

        Dim newProj As clsProjekt = Me.createVariant("$delete$", "")

        ' hier passiert das jetzt 
        Dim roleNameIDs As New SortedList(Of String, Double)

        If Not IsNothing(rolesToBeDeleted) Then
            For Each roleName As String In rolesToBeDeleted

                Dim teamID As Integer = -1
                Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(roleName, teamID)
                If Not IsNothing(tmpRole) Then

                    Dim curRoleNameID As String = RoleDefinitions.bestimmeRoleNameID(tmpRole.UID, teamID)

                    If includingChilds Then
                        Dim tmpRoleIDS As SortedList(Of String, Double) = RoleDefinitions.getSubRoleNameIDsOf(curRoleNameID, type:=PTcbr.all)
                        For Each srKvP As KeyValuePair(Of String, Double) In tmpRoleIDS
                            If roleNameIDs.ContainsKey(srKvP.Key) Then
                                ' muss nichts getan werden - ist schon in der Liste  
                            Else
                                ' der Value entspricht dem Anteil der Kapa der Subrole in der übergeordneten Sammelrolle, 
                                ' das ist hier aber irrerelevant .. deswegen immer auf 1 setzen 
                                roleNameIDs.Add(srKvP.Key, 1.0)
                            End If
                        Next
                    Else

                        If Not roleNameIDs.ContainsKey(curRoleNameID) Then
                            roleNameIDs.Add(curRoleNameID, 1.0)
                        End If

                    End If

                End If



            Next
        End If


        ' jetzt sind alle RoleIds, die gelöscht werden sollen in der Liste  roleIDs 
        ' jetzt werden einfach alle Phasen durchgegangen, ob sie eine der  Rollen enthalten 
        For ip = 1 To newProj.CountPhases
            Dim cPhase As clsPhase = newProj.getPhase(ip)

            If Not IsNothing(rolesToBeDeleted) Then
                If roleNameIDs.Count > 0 Then
                    Dim delCollection As New Collection
                    For dx As Integer = 1 To cPhase.countRoles
                        Dim tmpRole As clsRolle = cPhase.getRole(dx)
                        Dim tmpKey As String = RoleDefinitions.bestimmeRoleNameID(tmpRole.uid, tmpRole.teamID)
                        If roleNameIDs.ContainsKey(tmpKey) Then
                            ' löschen 
                            If Not delCollection.Contains(tmpKey) Then
                                delCollection.Add(tmpKey, tmpKey)
                            End If
                        End If
                    Next

                    ' jetzt müssen alle delCollection Einträge gelöscht werden 
                    For Each item As String In delCollection
                        If item <> "" Then
                            cPhase.removeRoleByNameID(item)
                        End If
                    Next

                End If
            End If

            If Not IsNothing(costsToBedeleted) Then
                ' jetzt kommen die Kosten dran 
                If costsToBedeleted.Count > 0 Then
                    Dim delCollection As New Collection
                    For cx As Integer = 1 To cPhase.countCosts
                        Dim tmpCost As clsKostenart = cPhase.getCost(cx)
                        If costsToBedeleted.Contains(tmpCost.name) Then
                            If Not delCollection.Contains(tmpCost.name) Then
                                delCollection.Add(tmpCost.name, tmpCost.name)
                            End If
                        End If
                    Next

                    ' jetzt müssen alle delCollection Einträge gelöscht werden 
                    For Each item As String In delCollection
                        If item <> "" Then
                            cPhase.removeCostByName(item)
                        End If
                    Next
                End If
            End If


        Next

        ' ende Aktionen
        newProj.variantName = Me.variantName
        deleteRolesAndCosts = newProj

    End Function
    ''' <summary>
    ''' Methode prüft auf Identität mit einem Vergleichsprojekt 
    ''' es wird verglichen: Startdatum, Endedatum (nur type=0), Phasen, Milestones, Personalkosten, Sonstige Kosten, Ergebnis, Attribute, Projekt-Ampel, Milestone-Ampeln, 
    ''' Deliverables, CustomFields, Projekt-Typ verglichen  
    ''' type 0: Vergleich eines Projektes mit einer seiner Projekt-Varianten bzw. einem anderen zeitlichen Stand; der Start/das Ende des Projektes macht einen Unterschied !
    ''' type 1: Vergleich eines Projektes mit einem anderen Projekt; der Start des Projektes macht keinen Unterschied !  
    ''' type 2: Vergleich eines Projektes mit seiner Vorlage: Startdatum, Ende-Datum, Ergebnis werden nicht miteinander verglichen; bei den CustomFields werden nur die keys miteinander verglichen   
    ''' in beiden Typen werden neben Startdatum (abhängig von type) die Phasen, Milestones, Personalkosten, Sonstige Kosten, Ergebnis, Attribute, Projekt-Ampel, Milestone-Ampeln, 
    ''' Deliverables, CustomFields, Projekt-Typ verglichen  
    ''' </summary>
    ''' <param name="vglproj">Projekt vom Typ clsProjekt</param>
    ''' <param name="absolut">soll absolut verglichen werden oder relativ; nur relevant bei Overview</param>
    ''' <param name="type">gibt den Vergleichstyp an</param>
    ''' <param name="strongRoleIdentity" >true: unterschiede werden ausgewiesen, wenn ein einzelner Monat einen unterschiedlichen Wert aufweist
    ''' false: Unterschied wird ausgewiesen, wenn Summe unterschiedlich ist; egal wie sich die einzelnen Werte verteilen</param>
    ''' ''' <param name="strongCostIdentity" >true: unterschiede werden ausgewiesen, wenn ein einzelner Monat einen unterschiedlichen Wert aufweist
    ''' false: Unterschied wird ausgewiesen, wenn Summe unterschiedlich ist; egal wie sich die einzelnen Werte verteilen</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property listOfDifferences(ByVal vglproj As clsProjekt, ByVal absolut As Boolean, ByVal type As Integer,
                                               Optional strongRoleIdentity As Boolean = False,
                                               Optional strongCostIdentity As Boolean = False) As Collection
        Get

            ' im Folgenden sind viele Try .. Catch drin
            ' ein ..contains wird extra nicht gemacht, weil der Eintrag eigentlich gar nicht vorkommen kann
            ' wenn die Prüfung jedesmal gemacht wird, verlangsamt es die Sache unnötig. 
            ' 
            Dim isDifferent As Boolean = False
            Dim tmpCollection As New Collection
            Dim hValues() As Double, cValues() As Double
            'Dim hdates As SortedList(Of Date, String)
            'Dim cdates As SortedList(Of Date, String)

            If Not IsNothing(vglproj) Then


                Dim verify As Integer = Me.dauerInDays
                verify = vglproj.dauerInDays

                Dim istVorlage As Boolean
                If type = 2 Then
                    ' Vorlage 
                    istVorlage = True
                Else
                    istVorlage = False
                End If


                ' Vergleich eines Projektes mit einer seiner Projekt-Varianten bzw. einem anderen zeitlichen Stand

                If type = 0 Then
                    ' Ist das startdatum unterschiedlich?
                    If Me.startDate.Date <> vglproj.startDate.Date Then
                        Try
                            tmpCollection.Add(CInt(PThcc.startdatum).ToString, CInt(PThcc.startdatum).ToString)
                        Catch ex As Exception

                        End Try

                    End If

                    ' Ist das Ende-Datum unterschiedlich?
                    If Me.endeDate.Date <> vglproj.endeDate.Date Then
                        Try
                            tmpCollection.Add(CInt(PThcc.endedatum).ToString, CInt(PThcc.endedatum).ToString)
                        Catch ex As Exception

                        End Try

                    End If
                End If


                ' prüfen, ob die Phasen identisch sind bzgl (StartOffset, Dauer)
                hValues = Me.getPhaseInfos
                cValues = vglproj.getPhaseInfos
                If arraysAreDifferent(hValues, cValues) Then
                    Try
                        tmpCollection.Add(CInt(PThcc.phasen).ToString, CInt(PThcc.phasen).ToString)
                    Catch ex As Exception

                    End Try

                End If

                ' prüfen, ob die Milestones identisch sind 
                ' muss bei allen Vergleichs-Typen projekt / version ./variante , ./vorlage, ./projekt2 gemacht werden
                hValues = Me.getMilestoneOffsets.Keys.ToArray
                cValues = vglproj.getMilestoneOffsets.Keys.ToArray
                If arraysAreDifferent(hValues, cValues) Then
                    Try
                        tmpCollection.Add(CInt(PThcc.resultdates).ToString, CInt(PThcc.resultdates).ToString)
                    Catch ex As Exception

                    End Try

                End If
                'End If


                If Not istVorlage Then
                    ' bei einer Vorlage macht es wenig Sinn, gegen Personalkosten, Andere Kosten, Ergebnis zu prüfen 

                    ' prüfen , ob die Personalkosten identisch sind 
                    ' muss bei allen Vergleichs-Typen projekt / version ./variante , ./vorlage, ./projekt2 gemacht werden
                    hValues = Me.getAllPersonalKosten
                    cValues = vglproj.getAllPersonalKosten

                    If strongCostIdentity Then
                        If arraysAreDifferent(hValues, cValues) And (hValues.Sum > 0 Or cValues.Sum > 0) Then
                            Try
                                tmpCollection.Add(CInt(PThcc.perscost).ToString, CInt(PThcc.perscost).ToString)
                            Catch ex As Exception

                            End Try

                        End If
                    Else
                        If hValues.Sum <> cValues.Sum Then
                            Try
                                tmpCollection.Add(CInt(PThcc.perscost).ToString, CInt(PThcc.perscost).ToString)
                            Catch ex As Exception

                            End Try
                        End If
                    End If


                    ' prüfen, ob sonstige Kosten identisch sind 
                    ' muss bei allen Vergleichs-Typen projekt / version ./variante , ./vorlage, ./projekt2 gemacht werden
                    hValues = Me.getGesamtAndereKosten
                    cValues = vglproj.getGesamtAndereKosten
                    If strongCostIdentity Then
                        If arraysAreDifferent(hValues, cValues) And (hValues.Sum > 0 Or cValues.Sum > 0) Then
                            Try
                                tmpCollection.Add(CInt(PThcc.othercost).ToString, CInt(PThcc.othercost).ToString)
                            Catch ex As Exception

                            End Try

                        End If

                    Else
                        If hValues.Sum <> cValues.Sum Then
                            Try
                                tmpCollection.Add(CInt(PThcc.othercost).ToString, CInt(PThcc.othercost).ToString)
                            Catch ex As Exception

                            End Try
                        End If
                    End If


                    ' prüfen, ob das Ergebnis identisch ist 
                    ' muss nicht bei Vergleichs-Typ 2 (Vorlage) gemacht werden 
                    Dim aktBudget As Double, aktPCost As Double, aktSCost As Double, aktRCost As Double, aktErg As Double
                    Dim vglBudget As Double, vglPCost As Double, vglSCost As Double, vglRCost As Double, vglErg As Double

                    With Me
                        .calculateRoundedKPI(aktBudget, aktPCost, aktSCost, aktRCost, aktErg)
                    End With

                    With vglproj
                        .calculateRoundedKPI(vglBudget, vglPCost, vglSCost, vglRCost, vglErg)
                    End With

                    If aktErg <> vglErg Then
                        Try
                            tmpCollection.Add(CInt(PThcc.ergebnis).ToString, CInt(PThcc.ergebnis).ToString)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob die Attribute identisch sind
                    If Me.StrategicFit <> vglproj.StrategicFit Or
                                Me.Risiko <> vglproj.Risiko Then
                        Try
                            tmpCollection.Add(CInt(PThcc.fitrisk).ToString, CInt(PThcc.fitrisk).ToString)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob die Projekt Ampel unterschiedlich ist 
                    If Me.ampelStatus <> vglproj.ampelStatus Then
                        Try
                            tmpCollection.Add(CInt(PThcc.projektampel).ToString, CInt(PThcc.projektampel).ToString)
                        Catch ex As Exception

                        End Try

                    End If

                    ' prüfen, ob die Meilenstein Ampeln unterschiedlich sind 
                    hValues = Me.getMilestoneColors
                    cValues = vglproj.getMilestoneColors
                    If arraysAreDifferent(hValues, cValues) Then
                        Try
                            tmpCollection.Add(CInt(PThcc.resultampel).ToString, CInt(PThcc.resultampel).ToString)
                        Catch ex As Exception

                        End Try

                    End If


                End If

                ' prüfen, ob die Deliverables identisch sind 

                Try
                    Dim hsortedList As SortedList(Of String, String) = Me.getDeliverables
                    Dim cSortedList As SortedList(Of String, String) = vglproj.getDeliverables
                    If sortedListsAreDifferent(hsortedList, cSortedList, 0) Then

                        Try
                            tmpCollection.Add(CInt(PThcc.deliverables).ToString, CInt(PThcc.deliverables).ToString)
                        Catch ex As Exception

                        End Try

                    End If

                Catch ex As Exception

                End Try

                ' prüfen, ob die Custom-Fields identisch sind 
                Dim verschieden As Boolean = False
                ' die String Custom Fields

                Try
                    Dim hsortedList As SortedList(Of Integer, String) = Me.customStringFields
                    Dim cSortedList As SortedList(Of Integer, String) = vglproj.customStringFields


                    If sortedListsAreDifferent(hsortedList, cSortedList, 1, istVorlage) Then

                        verschieden = True
                        Try
                            tmpCollection.Add(CInt(PThcc.customfields).ToString, CInt(PThcc.customfields).ToString)
                        Catch ex As Exception

                        End Try

                    End If

                Catch ex As Exception

                End Try

                ' die Double Custom Fields
                If Not verschieden Then
                    Try
                        Dim hsortedList As SortedList(Of Integer, Double) = Me.customDblFields
                        Dim cSortedList As SortedList(Of Integer, Double) = vglproj.customDblFields

                        If sortedListsAreDifferent(hsortedList, cSortedList, 2, istVorlage) Then

                            verschieden = True
                            Try
                                tmpCollection.Add(CInt(PThcc.customfields).ToString, CInt(PThcc.customfields).ToString)
                            Catch ex As Exception

                            End Try

                        End If

                    Catch ex As Exception

                    End Try
                End If

                ' die Bool Fields
                If Not verschieden Then
                    Try
                        Dim hsortedList As SortedList(Of Integer, Boolean) = Me.customBoolFields
                        Dim cSortedList As SortedList(Of Integer, Boolean) = vglproj.customBoolFields

                        If sortedListsAreDifferent(hsortedList, cSortedList, 3, istVorlage) Then

                            verschieden = True
                            Try
                                tmpCollection.Add(CInt(PThcc.customfields).ToString, CInt(PThcc.customfields).ToString)
                            Catch ex As Exception

                            End Try

                        End If

                    Catch ex As Exception

                    End Try
                End If


                ' prüfen, ob der Projekt-Typ der gleiche ist 
                If Not istVorlage Then
                    Dim hvalue As String = Me.VorlagenName
                    Dim cvalue As String = vglproj.VorlagenName

                    If hvalue <> cvalue Then
                        Try
                            tmpCollection.Add(CInt(PThcc.projecttype).ToString, CInt(PThcc.projecttype).ToString)
                        Catch ex As Exception

                        End Try
                    End If
                End If


            End If       ' Ende von if not isnothing(vglproj)


            listOfDifferences = tmpCollection
        End Get
    End Property



    ''' <summary>
    ''' liefert zu einem gegebenen Meilenstein das definierte Datum zurück
    ''' die Ampelfarbe wird ebenfalls in das Datum als Ablauf von Sekunden nach Mitternacht integriert
    ''' 0-nicht bewertet, 1-grün, 2-gelb, 3-rot
    ''' Nothing, wenn Meilenstein nicht existiert
    ''' Existieren mehrere Meilensteine desselben Namens so wird nur der erste zurückgebracht 
    ''' </summary>
    ''' <param name="milestoneName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneDate(ByVal milestoneName As String,
                                              Optional breadCrumb As String = "",
                                              Optional lfdNr As Integer = 1) As Date
        Get
            Dim found As Boolean = False
            'Dim cphase As clsPhase
            Dim cresult As clsMeilenstein
            Dim tmpDate As Date = Nothing
            Dim p As Integer = 1
            Dim colorIndex As Integer

            ' neu
            Dim hryindices(,) As Integer = Me.hierarchy.getMilestoneIndices(milestoneName)
            Dim milestoneIndices(,) As Integer = Me.hierarchy.getMilestoneIndices(milestoneName, breadCrumb)


            For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                If milestoneIndices(0, mx) > 0 And milestoneIndices(1, mx) > 0 _
                    And mx = lfdNr - 1 Then

                    Try
                        cresult = Me.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))
                        If Not IsNothing(cresult) Then

                            colorIndex = cresult.getBewertung(1).colorIndex
                            tmpDate = cresult.getDate.Date          ' hier wird der Zeit-Teil des MS-Datums abgeschnitten und wird nach tmpdate gespeichert

                            ' jetzt wird die Ampelfarbe ins Datum kodiert 
                            tmpDate = tmpDate.AddSeconds(colorIndex)
                            found = True

                            ' jetzt wird in das Datum kodiert, ob der Meilenstein abgeschlossen sein sollte
                            ' wenn timestamp nach dem Meilenstein-Datum steht, sollte der Meilenstein abgeschlossen sein 
                            If DateDiff(DateInterval.Day, Me.timeStamp, tmpDate) < 0 Then

                                ' Meilenstein Datum liegt vor dem Datum, an dem dieser Planungs-Stand abgegeben wurde
                                tmpDate = tmpDate.AddHours(6)

                            End If

                        End If

                    Catch ex As Exception

                    End Try


                End If

            Next

            ' neu Ende 

            ' alt: bis 20.9.2016
            ''Do While p <= Me.CountPhases And Not found

            ''    cphase = Me.getPhase(p)

            ''    cresult = cphase.getMilestone(milestoneName)

            ''    If Not IsNothing(cresult) Then

            ''        colorIndex = cresult.getBewertung(1).colorIndex
            ''        tmpDate = cresult.getDate.Date          ' hier wird der Zeit-Teil des MS-Datums abgeschnitten und wird nach tmpdate gespeichert

            ''        ' jetzt wird die Ampelfarbe ins Datum kodiert 
            ''        tmpDate = tmpDate.AddSeconds(colorIndex)
            ''        found = True

            ''        ' jetzt wird in das Datum kodiert, ob der Meilenstein abgeschlossen sein sollte
            ''        ' wenn timestamp nach dem Meilenstein-Datum steht, sollte der Meilenstein abgeschlossen sein 
            ''        If DateDiff(DateInterval.Day, Me.timeStamp, tmpDate) < 0 Then

            ''            ' Meilenstein Datum liegt vor dem Datum, an dem dieser Planungs-Stand abgegeben wurde
            ''            tmpDate = tmpDate.AddHours(6)

            ''        End If

            ''    End If

            ''    p = p + 1

            ''Loop

            If found Then
                getMilestoneDate = tmpDate
            Else
                getMilestoneDate = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' diese Routine berücksichtigt, wieviel von der phase im Start- bzw End Monat liegt; 
    ''' es wird für Start und Ende Monat nicht automatisch 1 gesetzt, sondern ein anteiliger Wert, der sich daran bemisst, 
    ''' wieviel Phase im Start bzw End Monat liegt; 
    '''   
    ''' </summary>
    ''' <param name="phaseName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>da eine Vorlage kein StartDatum kennt, muss diese Methode als overridable/overrides definiert werden   
    ''' </remarks>
    Public Overrides ReadOnly Property getPhasenBedarf(phaseName As String) As Double()

        Get
            Dim phaseValues() As Double
            Dim anzPhasen As Integer
            Dim i As Integer, p As Integer
            Dim phase As clsPhase
            Dim phaseStart As Date, phaseEnd As Date
            'Dim numberOfDays As Integer
            Dim anteil As Double
            Dim daysPMonth(12) As Integer
            Dim anzTage As Integer

            daysPMonth(0) = 0
            daysPMonth(1) = 31
            daysPMonth(2) = 28
            daysPMonth(3) = 31
            daysPMonth(4) = 30
            daysPMonth(5) = 31
            daysPMonth(6) = 30
            daysPMonth(7) = 31
            daysPMonth(8) = 31
            daysPMonth(9) = 30
            daysPMonth(10) = 31
            daysPMonth(11) = 30
            daysPMonth(12) = 31



            ReDim phaseValues(_Dauer - 1)

            If _Dauer > 0 Then



                anzPhasen = AllPhases.Count
                If anzPhasen > 0 Then

                    For p = 0 To anzPhasen - 1
                        phase = AllPhases.Item(p)


                        If phase.name = phaseName Then


                            phaseStart = phase.getStartDate
                            phaseEnd = phase.getEndDate


                            With phase
                                For i = 0 To .relEnde - .relStart

                                    If i = 0 Then

                                        If awinSettings.phasesProzentual Then

                                            If .relEnde = .relStart Then
                                                anzTage = CInt(DateDiff(DateInterval.Day, phaseStart, phaseEnd)) + 1
                                            Else
                                                anzTage = daysPMonth(phaseStart.Month) - phaseStart.Day + 1
                                            End If

                                            anteil = (daysPMonth(phaseStart.Month) - phaseStart.Day + 1) / daysPMonth(phaseStart.Month)
                                            phaseValues(.relStart - 1 + i) = anteil
                                        Else
                                            phaseValues(.relStart - 1 + i) = 1
                                        End If

                                    ElseIf i = .relEnde - .relStart Then

                                        If awinSettings.phasesProzentual Then
                                            anteil = phaseEnd.Day / daysPMonth(phaseEnd.Month)
                                            phaseValues(.relStart - 1 + i) = anteil
                                        Else
                                            phaseValues(.relStart - 1 + i) = 1
                                        End If



                                    Else

                                        phaseValues(.relStart - 1 + i) = 1

                                    End If

                                Next
                            End With

                        End If

                    Next p ' Loop über alle Phasen
                Else
                    Throw New ArgumentException("Projekt hat keine Phasen")
                End If


                getPhasenBedarf = phaseValues

            Else
                Throw New ArgumentException("Projekt hat keine Dauer")
                getPhasenBedarf = phaseValues
            End If
        End Get

    End Property


    Public Overrides ReadOnly Property dauerInDays As Integer

        Get
            Dim i As Integer
            Dim max As Double = 0
            Dim offsetProjStart As Integer = CInt(DateDiff(DateInterval.Day, StartofCalendar, Me.startDate))

            ' Bestimmung der Dauer 

            For i = 1 To Me.CountPhases

                With Me.getPhase(i)

                    If max < .startOffsetinDays + .dauerInDays Then
                        max = .startOffsetinDays + .dauerInDays
                    End If

                End With

            Next i

            ' jetzt aus Konsistenzgründen die Dauer in Monaten setzen 
            '_Dauer = getColumnOfDate(StartofCalendar.AddDays(offsetProjStart + max - 1)) - getColumnOfDate(StartofCalendar.AddDays(offsetProjStart)) + 1

            If Me.CountPhases > 0 Then

                _Dauer = Me.anzahlRasterElemente


            End If

            dauerInDays = CInt(max)


        End Get
    End Property



    Public ReadOnly Property tfspalte As Integer
        Get
            tfspalte = _Start
        End Get
    End Property

    ''' <summary>
    ''' ist für das Projekt jetzt in der Rootphase gespeichert 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ampelStatus As Integer
        Get
            'ampelStatus = _ampelStatus
            If Me.CountPhases > 0 Then
                ampelStatus = Me.getPhase(1).ampelStatus
            Else
                ampelStatus = 0
            End If

        End Get

        Set(value As Integer)
            If Not (IsNothing(value)) Then
                If IsNumeric(value) Then
                    If value >= 0 And value <= 3 Then
                        If Me.CountPhases > 0 Then
                            Me.getPhase(1).ampelStatus = value
                        End If
                    Else
                        Throw New ArgumentException("unzulässiger Ampel-Wert")
                    End If
                Else
                    Throw New ArgumentException("nicht-numerischer Ampel-Wert")
                End If
            Else
                ' ohne Bewertung
                If Me.CountPhases > 0 Then
                    Me.getPhase(1).ampelStatus = 0
                End If
            End If

        End Set
    End Property

    ''' <summary>
    ''' ist für das Projekt jetzt in der RootPhase gespeichert 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ampelErlaeuterung As String
        Get
            'ampelErlaeuterung = _ampelErlaeuterung
            If Me.CountPhases > 0 Then
                ampelErlaeuterung = Me.getPhase(1).ampelErlaeuterung
            Else
                ampelErlaeuterung = ""
            End If
        End Get
        Set(value As String)
            If Not (IsNothing(value)) Then
                If Me.CountPhases > 0 Then
                    Me.getPhase(1).ampelErlaeuterung = value
                End If
            Else
                ' nichts tun 
            End If
        End Set
    End Property

    ''' <summary>
    ''' liefert das Ende-Datum des Projektes zurück - Readonly 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property endeDate As Date
        Get
            endeDate = Me.startDate.AddDays(Me.dauerInDays - 1).Date
        End Get
    End Property


    ''' <summary>
    ''' das Startdatum darf nur verschoben werden, wenn gilt: 
    ''' es handelt sich um eine Variante und das Flag movable = true 
    ''' 
    ''' oder 
    ''' das Projekt ist noch im Planungs-Stadium und das Flag movable = true 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property startDate As Date
        Get
            startDate = _startDate
        End Get

        Set(newValue As Date)
            Dim value As Date = newValue.Date
            Dim olddate As Date = _startDate
            Dim differenzInTagen As Integer = CInt(DateDiff(DateInterval.Day, olddate, value))
            Dim updatePhases As Boolean = False

            ' Änderung am 25.5.14: es ist nicht mehr erlaubt, das Startdatum innerhalb des gleichen Monats zu verschieben 
            ' es muss geprüft werden, ob es noch im Planungs-Stadium ist: nur dann darf noch verschoben werden ...
            If (differenzInTagen <> 0 And Me.movable) And
                (_Status = ProjektStatus(0) Or _variantName <> "") Then
                _startDate = value
                _Start = CInt(DateDiff(DateInterval.Month, StartofCalendar, value) + 1)
                ' Änderung 25.5 die Xwerte müssen jetzt synchronisiert werden 
                'If Not currentConstellationName.EndsWith("(*)") And currentConstellationName <> "Last" Then
                '    currentConstellationName = currentConstellationName & "(*)"
                'End If

            ElseIf _startDate = NullDatum Then
                _startDate = value
                _Start = CInt(DateDiff(DateInterval.Month, StartofCalendar, value) + 1)


            ElseIf _Status <> ProjektStatus(0) And _variantName = "" And Not Me.movable Then
                Throw New ArgumentException("der Startzeitpunkt kann nicht mehr verändert werden ... ")


            ElseIf DateDiff(DateInterval.Day, StartofCalendar, newValue) < 0 Then
                Throw New ArgumentException("der Startzeitpunkt liegt vor dem Kalenderstart  ... ")

            End If

            ' um _Dauer neu zu berechnen ; ergänzt am 12.5.2014
            If differenzInTagen <> 0 Then
                Dim anzahlTage As Integer = Me.dauerInDays
            End If

        End Set
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn das Projekt zum angegebenen Zeitpunkt in dieser Phase ist 
    ''' pHname kann der Elem-Name oder der Name mit (Tel-)Breadcrumb sein
    ''' wenn nicht eindeutig, dann wird die erste Phase verwendet 
    ''' </summary>
    ''' <param name="phName"></param>
    ''' <returns></returns>
    Public Function isInPhase(ByVal phName As String, ByVal timestamp As Date, ByRef difference As Integer) As Boolean

        Dim tmpResult As Boolean = False
        Dim elemName As String = ""
        Dim breadCrumb As String = ""
        Dim type As Integer = -1
        Dim pvname As String = ""

        Call splitHryFullnameTo2(phName, elemName, breadCrumb, type, pvname)
        Dim cphase As clsPhase = Me.getPhase(elemName, breadCrumb)

        If Not IsNothing(cphase) Then
            If cphase.getStartDate <= timestamp And cphase.getEndDate >= timestamp Then
                tmpResult = True
                difference = 0
            Else
                ' wird benötigt, um die Phase zu bestimmen die gerade eben fertig geworden ist, wenn keine einzige passt ...
                difference = CInt(DateDiff(DateInterval.Day, cphase.getEndDate, timestamp))
            End If
        End If

        isInPhase = tmpResult

    End Function

    ''' <summary>
    ''' gibt an, wieviel Prozent der Phase zum angegebenen Zeitpunkt verstrichen ist 
    ''' </summary>
    ''' <param name="phName"></param>
    ''' <param name="timestamp"></param>
    ''' <returns></returns>
    Public Function przTimeSpend(ByVal phName As String, ByVal timestamp As Date) As Double
        Dim tmpResult As Double = 0.0

        Dim elemName As String = ""
        Dim breadCrumb As String = ""
        Dim type As Integer = -1
        Dim pvname As String = ""

        Call splitHryFullnameTo2(phName, elemName, breadCrumb, type, pvname)
        Dim cphase As clsPhase = Me.getPhase(elemName, breadCrumb)

        If Not IsNothing(cphase) Then

            If cphase.getStartDate >= timestamp Then
                tmpResult = 0.0
            ElseIf cphase.getEndDate <= timestamp Then
                tmpResult = 1.0
            Else
                Dim gesamtdauer As Long = DateDiff(DateInterval.Day,
                                                      cphase.getStartDate.Date,
                                                      cphase.getEndDate.Date) + 1
                Dim goneTime As Long = DateDiff(DateInterval.Day,
                                                   cphase.getStartDate,
                                                   timestamp) + 1
                If gesamtdauer > 0 And goneTime > 0 Then
                    tmpResult = CDbl(goneTime / gesamtdauer)
                Else
                    tmpResult = 0.0
                End If
            End If

        End If

        przTimeSpend = tmpResult
    End Function

    Public Property earliestStartDate As Date
        Get
            earliestStartDate = _earliestStartDate
        End Get
        Set(value As Date)

            _earliestStartDate = value.Date


        End Set
    End Property
    ''' <summary>
    ''' wird benutzt beim Einlesen vom File, Konsistenz mit Status vorausgesetzt
    ''' </summary>
    ''' <param name="anyway"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property earliestStartDate(anyway As Boolean) As Date
        Get
            earliestStartDate = _earliestStartDate
        End Get
        Set(value As Date)
            Dim Heute As Date = Now

            _earliestStartDate = value.Date


        End Set
    End Property

    Public Property latestStartDate As Date
        Get
            latestStartDate = _latestStartDate
        End Get
        Set(value As Date)
            Dim heute As Date = Now

            _latestStartDate = value.Date

        End Set
    End Property

    ''' <summary>
    ''' wird eingesetzt beim Einlesen vom File - Konsistenz vorausgesetzt 
    ''' </summary>
    ''' <param name="anyway"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property latestStartDate(anyway As Boolean) As Date
        Get
            latestStartDate = _latestStartDate
        End Get
        Set(value As Date)
            Dim heute As Date = Now

            _latestStartDate = value.Date


        End Set
    End Property

    ''' <summary>
    ''' gibt eine Liste von Phasen zurück, die für das gegebene Projekt im angegebenen Zeitrahmen liegen
    ''' wenn namenliste leer ist, werden alle Projekte des Projekts betrachtet 
    ''' </summary>
    ''' <param name="areMilestones">gibt an, ob Meilensteine geuscht werden, oder Phasen</param>
    ''' <param name="von">linker Rand des Zeitraums</param>
    ''' <param name="bis">rechter Rand des Zeitraums</param>
    ''' <param name="namenListe" >gibt an, welche elemIDs nur betrachtet werden sollen; wenn namenListe leer ist, dann werden alle Phasen/Meilensteine betrachtet </param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property phasesWithinTimeFrame(ByVal areMilestones As Boolean, von As Integer, bis As Integer, ByVal namenListe As Collection) As Collection
        Get
            Dim tmpListe As New Collection
            ' selection type wird aktuell noch ignoriert .... 
            Dim elemID As String
            Dim considerAllNames As Boolean
            Dim startIX As Integer, endIX As Integer

            ' ein Zeitraum muss definiert sein 
            If von <= 0 Or bis <= 0 Or bis - von < 0 Then
                phasesWithinTimeFrame = tmpListe
            Else
                Dim ix As Integer
                Dim anzElements As Integer

                If namenListe.Count = 0 Then
                    considerAllNames = True
                    If areMilestones Then
                        startIX = Me.hierarchy.getIndexOf1stMilestone
                        endIX = Me.hierarchy.count
                        anzElements = endIX - startIX + 1
                    Else
                        startIX = 1
                        endIX = Me.hierarchy.getIndexOf1stMilestone - 1
                        anzElements = endIX - startIX + 1
                    End If
                Else
                    considerAllNames = False
                    startIX = 1
                    endIX = namenListe.Count
                End If

                ' jetzt muss die Schleife kommen 
                ix = startIX
                While ix <= endIX

                    If considerAllNames Then
                        elemID = Me.hierarchy.getIDAtIndex(ix)
                    Else
                        elemID = CStr(namenListe.Item(ix))
                    End If

                    If areMilestones Then
                        ' Behandlung von Meilensteinen
                        Dim cMilestone As clsMeilenstein = Me.getMilestoneByID(elemID)
                        Dim milestoneColumn As Integer = getColumnOfDate(cMilestone.getDate)
                        If milestoneColumn < von Or milestoneColumn > bis Then
                            ' nichts machen 
                        Else
                            ' Milestone ist im Zeitraum 
                            If tmpListe.Contains(cMilestone.nameID) Then
                                ' nichts tun, denn jede Phase wird nur einmal eingetragen ....
                            Else
                                tmpListe.Add(cMilestone.nameID, cMilestone.nameID)
                            End If
                        End If

                    Else
                        ' Behandlung von Phasen
                        Dim cphase As clsPhase = Me.getPhaseByID(elemID)
                        If Me._Start + cphase.relStart - 1 > bis Or
                            Me._Start + cphase.relEnde - 1 < von Then
                            ' nichts tun 
                        Else
                            ' ist innerhalb des Zeitrahmens
                            If tmpListe.Contains(cphase.nameID) Then
                                ' nichts tun, denn jede Phase wird nur einmal eingetragen ....
                            Else
                                tmpListe.Add(cphase.nameID, cphase.nameID)
                            End If
                        End If
                    End If

                    ix = ix + 1

                End While

            End If

            phasesWithinTimeFrame = tmpListe

        End Get
    End Property

    ''' <summary>
    ''' gibt zuück, ob das Projekt in dem TimeFrame, definiert durch Moant von und Monat bis, liegt 
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isWithinTimeFrame(ByVal von As Integer, ByVal bis As Integer) As Boolean
        Get
            Dim tmpResult As Boolean
            If bis - von < 1 Then
                tmpResult = True
            Else
                tmpResult = Not (getColumnOfDate(Me.startDate) > bis Or
                    getColumnOfDate(Me.endeDate) < von)
            End If
            isWithinTimeFrame = tmpResult
        End Get
    End Property


    Public Sub clearPhases()

        Try
            AllPhases.Clear()
        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

    End Sub



    ''' <summary>
    ''' stellt sicher, dass die Phase1 immer das gesamte Projekt umfasst 
    ''' und dass die Projektlaenge richtig kalkuliert ist 
    ''' Me.dauerindays setzt die interne privat Variable 
    ''' </summary>
    ''' <param name="phasenEnde"></param>
    ''' <remarks></remarks>
    Public Sub keepPhase1consistent(ByVal phasenEnde As Integer)

        Try
            Dim phase1 As clsPhase = Me.getPhase(1)
            If Not IsNothing(phase1) Then
                If phase1.dauerInDays < phasenEnde Then
                    phase1.changeStartandDauerPhase1(0, phasenEnde)
                    ' im Nebeneffekt wird ausserdem _Dauer aktualisiert  
                    Dim projektLaengeInDays As Integer = Me.dauerInDays
                End If
            End If

        Catch ex As Exception

        End Try



    End Sub



    ''' <summary>
    ''' es werden alle Bewertungen gelöscht, auch die Phase(1), also Projekt-Bewertung
    ''' </summary>
    Public Sub clearBewertungen()
        Dim cPhase As clsPhase


        For p = 1 To Me.CountPhases
            cPhase = Me.getPhase(p)
            cPhase.clearBewertungen()

            For r = 1 To cPhase.countMilestones
                With cPhase.getMilestone(r)
                    .clearBewertungen()
                End With
            Next
        Next

    End Sub

    Public ReadOnly Property risikoKostenfaktor As Double
        Get
            Dim tmp As Double = 0.0

            If awinSettings.considerRiskFee Then
                tmp = Me.Risiko / 100
                If tmp < 0 Then
                    tmp = 0
                End If

                If DateDiff(DateInterval.Day, Me.endeDate, Date.Now) >= 0 Then
                    tmp = 0
                End If
            End If

            risikoKostenfaktor = tmp
        End Get
    End Property

    ''' <summary>
    ''' gibt die Risikokosten zurück
    ''' pro Risiko Punkt 1% vom Erloes
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property risikoKosten As Double
        Get

            Dim tmp As Double
            tmp = Me.Erloes * risikoKostenfaktor
            If tmp < 0 Then
                tmp = 0
            End If

            risikoKosten = tmp

        End Get
    End Property

    ''' <summary>
    ''' gibt die KeyMetric für TimeCompletion zurück auf die einzelnen Monate des Projektes
    ''' Es werden nur die Phasen- und Meilenstein-EndDates betrachtet, die in der Beauftragung des Projektes enthalten sind
    ''' </summary>
    ''' <returns>Array mit den monatlichen TimeCompletion-Werten</returns>
    Public ReadOnly Property getTimeCompletionMetric(ByVal baseMs As SortedList(Of Date, String),
                                                     ByVal basePhases As SortedList(Of Date, String),
                                                     ByVal bezugsdatum As Date,
                                                     Optional ByVal Total As Boolean = False) As Double()
        Get
            Dim timeCompletionValues() As Double
            Dim phase As clsPhase
            Dim ms As clsMeilenstein

            Dim bezugsOffset As Long = DateDiff(DateInterval.Day, Me.startDate, bezugsdatum)
            Dim currentEndIndex As Integer
            Dim reversePrzDone As Double
            Dim isElemOfPast As Boolean

            Dim phaseEndOffset As Integer
            Dim msOffset As Integer


            If _Dauer > 0 Then

                ReDim timeCompletionValues(_Dauer - 1)

                ' Loop über alle Phasen der Beauftragung (basePhases)

                For Each kvp As KeyValuePair(Of Date, String) In basePhases

                    Dim phaseId As String = kvp.Value

                    phase = Me.getPhaseByID(phaseId)

                    If Not IsNothing(phase) Then

                        phaseEndOffset = phase.startOffsetinDays + phase.dauerInDays - 1
                        currentEndIndex = phase.relEnde - 1
                        If phase.percentDone <> 0 Then
                            reversePrzDone = 1 / phase.percentDone
                        Else
                            reversePrzDone = 1 / 0.1                      ' da Division durch 0 nicht möglich
                        End If

                        'isElemOfPast = (phase.getEndDate < bezugsdatum)
                        isElemOfPast = phaseEndOffset < bezugsOffset

                        If Not Me.variantName = ptVariantFixNames.pfv.ToString Then
                            ' Planungsstand des Projektes
                            If Total Then
                                If isElemOfPast Then
                                    timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + phaseEndOffset * reversePrzDone
                                Else
                                    timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + phaseEndOffset
                                End If
                            Else
                                If isElemOfPast Then
                                    timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + phaseEndOffset * reversePrzDone
                                End If
                            End If
                        Else
                            ' beauftragtes Projekt
                            If Total Then
                                timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + phaseEndOffset
                            Else
                                If isElemOfPast Then
                                    timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + phaseEndOffset
                                End If
                            End If

                        End If

                    Else
                        ' Phase ist nicht mehr im Projekt d.h. timeCompletion künstlich erhöhen um 10000
                        timeCompletionValues(0) = timeCompletionValues(0) + 10000
                    End If

                Next


                ' Loop über alle Meilensteine der Beauftragung (baseMS)

                For Each kvpMS As KeyValuePair(Of Date, String) In baseMs

                    Dim msId As String = kvpMS.Value                         ' Id des Meilenstein des beauftragten Projektes holen

                    ms = Me.getMilestoneByID(msId)                           ' mit der Id den Meilenstein dieses Projektes holen 


                    If Not IsNothing(ms) Then

                        msOffset = ms.Parent.startOffsetinDays + CInt(ms.offset) 'Offset des ms in einem Projekt
                        currentEndIndex = getColumnOfDate(ms.getDate) - Me.Start
                        If ms.percentDone <> 0 Then
                            reversePrzDone = 1 / ms.percentDone
                        Else
                            reversePrzDone = 1 / 0.1                            ' da Division durch 0 nicht möglich
                        End If

                        'isElemOfPast = (ms.getDate < bezugsdatum)
                        isElemOfPast = msOffset < bezugsOffset

                        If Not Me.variantName = ptVariantFixNames.pfv.ToString Then
                            If Total Then
                                If isElemOfPast Then
                                    timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + msOffset * reversePrzDone
                                Else
                                    timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + msOffset
                                End If
                            Else
                                If isElemOfPast Then
                                    timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + msOffset * reversePrzDone
                                End If
                            End If
                        Else
                            ' beauftragtes Projekt
                            If Total Then
                                timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + msOffset
                            Else
                                If isElemOfPast Then
                                    timeCompletionValues(currentEndIndex) = timeCompletionValues(currentEndIndex) + msOffset
                                End If
                            End If

                        End If
                    Else
                        ' Meilenstein ist nicht mehr in dem Projekt d.h. timeCompletion künstlich erhöhen um 10000
                        timeCompletionValues(0) = timeCompletionValues(0) + 10000
                    End If
                Next


                getTimeCompletionMetric = timeCompletionValues

            Else
                Throw New ArgumentException("Projekt hat keine Dauer")
            End If
        End Get
    End Property


    ''' <summary>
    ''' gibt die KeyMetric für TimeCompletion zurück auf die einzelnen Monate des Projektes
    ''' Es werden nur die Phasen- und Meilenstein-EndDates betrachtet, die in der Beauftragung des Projektes enthalten sind
    ''' </summary>
    ''' <param name="baseDeliverables"></param>
    ''' <returns></returns>
    Public ReadOnly Property getDeliverableCompletionMetric(ByVal baseDeliverables As SortedList(Of String, String),
                                                            ByVal bezugsdatum As Date,
                                                            Optional ByVal Total As Boolean = False) As Double()
        Get
            Dim deliverableCompletionValues() As Double
            Dim phase As clsPhase
            Dim ms As clsMeilenstein

            Dim startIndex As Integer = Me.Start
            Dim currentEndIndex As Integer
            Dim currentPrzDone As Double
            Dim isElemOfPast As Boolean

            If _Dauer > 0 Then

                ReDim deliverableCompletionValues(_Dauer - 1)

                ' Loop über alle Deliverables in der Beauftragung

                For Each kvp As KeyValuePair(Of String, String) In baseDeliverables

                    Dim elemId As String = kvp.Value
                    Dim baseDeliv As String = kvp.Key

                    If elemIDIstMeilenstein(elemId) Then
                        ' Id des Meilenstein des beauftragten Projektes holen

                        ms = Me.getMilestoneByID(elemId)                ' mit der Id den Meilenstein dieses Projektes holen 

                        If Not IsNothing(ms) Then
                            Dim baseDelivName As String = ""
                            Dim tmpstr() As String = Split(baseDeliv, "(", 2)

                            If tmpstr.Length > 0 Then
                                baseDelivName = tmpstr(0)
                            End If
                            currentEndIndex = getColumnOfDate(ms.getDate) - Me.Start
                            currentPrzDone = ms.percentDone
                            isElemOfPast = (ms.getDate < bezugsdatum)

                            If Not Me.variantName = ptVariantFixNames.pfv.ToString Then
                                If Total Then
                                    If isElemOfPast Then
                                        If baseDelivName <> "" And ms.containsDeliverable(baseDelivName) Then
                                            deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1 * currentPrzDone
                                        End If
                                    Else
                                        If baseDelivName <> "" And ms.containsDeliverable(baseDelivName) Then
                                            deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1
                                        End If
                                    End If
                                Else
                                    If isElemOfPast Then
                                        If baseDelivName <> "" And ms.containsDeliverable(baseDelivName) Then
                                            deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1 * currentPrzDone
                                        End If
                                    End If
                                End If

                            Else
                                If Total Then

                                    If baseDelivName <> "" And ms.containsDeliverable(baseDelivName) Then
                                        deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1
                                    End If

                                Else
                                    If isElemOfPast Then
                                        If baseDelivName <> "" And ms.containsDeliverable(baseDelivName) Then
                                            deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1
                                        End If
                                    End If
                                End If

                            End If
                        Else

                        End If

                    Else
                        ' mit der elemID die Phase holen

                        phase = Me.getPhaseByID(elemId)

                        If Not IsNothing(phase) Then

                            Dim baseDelivName As String = ""
                            Dim tmpstr() As String = Split(baseDeliv, "(", 2)

                            If tmpstr.Length > 0 Then
                                baseDelivName = tmpstr(0)
                            End If

                            currentEndIndex = phase.relEnde - 1
                            currentPrzDone = phase.percentDone
                            isElemOfPast = (phase.getEndDate < bezugsdatum)

                            If Not Me.variantName = ptVariantFixNames.pfv.ToString Then

                                If Total Then
                                    If isElemOfPast Then
                                        If baseDelivName <> "" And phase.containsDeliverable(baseDelivName) Then
                                            deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1 * currentPrzDone
                                        End If
                                    Else
                                        If baseDelivName <> "" And phase.containsDeliverable(baseDelivName) Then
                                            deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1
                                        End If
                                    End If
                                Else
                                    If isElemOfPast Then
                                        If baseDelivName <> "" And phase.containsDeliverable(baseDelivName) Then
                                            deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1 * currentPrzDone
                                        End If
                                    End If
                                End If

                            Else
                                If Total Then

                                    If baseDelivName <> "" And phase.containsDeliverable(baseDelivName) Then
                                        deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1
                                    End If

                                Else
                                    If isElemOfPast Then
                                        If baseDelivName <> "" And phase.containsDeliverable(baseDelivName) Then
                                            deliverableCompletionValues(currentEndIndex) = deliverableCompletionValues(currentEndIndex) + 1
                                        End If
                                    End If
                                End If

                            End If

                        End If
                    End If




                Next

                getDeliverableCompletionMetric = deliverableCompletionValues

            Else
                Throw New ArgumentException("Projekt hat keine Dauer")
            End If
        End Get
    End Property

    ''' <summary>
    ''' kopiert die Attribute eines Projektes in newproject;  bei der Quelle handelt es sich um 
    ''' ein normales Projekt 
    ''' </summary>
    ''' <param name="newproject"></param>
    ''' <remarks></remarks>
    Public Overrides Sub copyAttrTo(ByRef newproject As clsProjekt)

        With newproject

            .farbe = farbe
            .Schrift = Schrift
            .Schriftfarbe = Schriftfarbe
            .VorlagenName = VorlagenName
            .Risiko = _Risiko
            .StrategicFit = _StrategicFit
            .Erloes = Erloes
            .description = _description
            .variantName = _variantName
            .variantDescription = _variantDescription
            .volume = _volume
            .complexity = _complexity
            .businessUnit = _businessUnit
            .projectType = _projectType
            .StartOffset = _StartOffset
            .startDate = _startDate
            .earliestStartDate = _earliestStartDate
            .latestStartDate = _latestStartDate
            .earliestStart = _earliestStart
            .latestStart = _latestStart
            .leadPerson = _leadPerson
            .Status = _Status
            .extendedView = Me.extendedView
            .actualDataUntil = Me.actualDataUntil
            .kundenNummer = Me.kundenNummer
            .vpID = Me.vpID

            Try
                .movable = _movable
            Catch ex As Exception
                .movable = False
            End Try


        End With

        ' jetzt wird die Hierarchie kopiert 
        Call copyHryTo(newproject)

        ' jetzt werden die CustomFields kopiert, so fern es welche gibt ... 
        Try
            With newproject
                For Each kvp As KeyValuePair(Of Integer, String) In customStringFields
                    .customStringFields.Add(kvp.Key, kvp.Value)
                Next

                For Each kvp As KeyValuePair(Of Integer, Double) In customDblFields
                    .customDblFields.Add(kvp.Key, kvp.Value)
                Next

                For Each kvp As KeyValuePair(Of Integer, Boolean) In customBoolFields
                    .customBoolFields.Add(kvp.Key, kvp.Value)
                Next

            End With
        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' sogenannte Heil-Methode, um Varianten, die beim Erzeugen ihre CustomFields nicht mitbekommen haben (der Fehler ist inzwischen behoben) 
    ''' diese CustomFileds wieder mitzugeben
    ''' </summary>
    ''' <param name="baseProject"></param>
    ''' <remarks></remarks>
    Public Sub copyCustomFieldsFrom(ByVal baseProject As clsProjekt)

        ' jetzt werden die CustomFields kopiert, so fern es welche gibt ... 
        Try

            ' wenn das Projekt keine Custom-Fields hat 
            If Me.customStringFields.Count = 0 And
                Me.customDblFields.Count = 0 And
                Me.customBoolFields.Count = 0 Then

                For Each kvp As KeyValuePair(Of Integer, String) In baseProject.customStringFields
                    Me.customStringFields.Add(kvp.Key, kvp.Value)
                Next

                For Each kvp As KeyValuePair(Of Integer, Double) In baseProject.customDblFields
                    Me.customDblFields.Add(kvp.Key, kvp.Value)
                Next

                For Each kvp As KeyValuePair(Of Integer, Boolean) In baseProject.customBoolFields
                    Me.customBoolFields.Add(kvp.Key, kvp.Value)
                Next

            End If


        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' liefert den Sortierungs-Key für das das angegebene Sort-Kriterium 
    ''' dient zur Verwendung in der Constellation
    ''' </summary>
    ''' <param name="sortType"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSortKeyForConstellation(ByVal sortType As Integer,
                                                              Optional ByVal lfdNr As Integer = 99999) As String
        Get
            Dim formatStr As String = "00000000"
            Dim tmpResult As String = "xxx"
            Select Case sortType

                Case ptSortCriteria.alphabet
                    ' das ist die Default-Lösung 
                    tmpResult = Me.name

                Case ptSortCriteria.buStartName
                    Dim offsetDays As Long = DateDiff(DateInterval.Day, StartofCalendar, Me.startDate)
                    tmpResult = Me.businessUnit & offsetDays.ToString(formatStr) & Me.name

                Case ptSortCriteria.customFields12
                    ' nimm aktuell die Default- Lösung 
                    tmpResult = Me.name

                Case ptSortCriteria.customListe
                    ' in diesem Fall muss die Sortier-Kennung aus einer Excel-Liste kommen 
                    tmpResult = calcSortKeyCustomTF(lfdNr)

                Case ptSortCriteria.customTF
                    tmpResult = calcSortKeyCustomTF(lfdNr)

                Case ptSortCriteria.formel
                    ' nimm aktuell die Default- Lösung 
                    tmpResult = Me.name

                Case ptSortCriteria.strategyProfitLossRisk
                    Dim tmp(4) As Double
                    Call Me.calculateRoundedKPI(tmp(0), tmp(1), tmp(2), tmp(3), tmp(4))
                    tmpResult = CInt(Me.StrategicFit * 10000 + tmp(4) * 1000 - Me.Risiko * 800).ToString(formatStr) & Me.name

                Case ptSortCriteria.strategyRiskProfitLoss
                    Dim tmp(4) As Double
                    Call Me.calculateRoundedKPI(tmp(0), tmp(1), tmp(2), tmp(3), tmp(4))
                    tmpResult = CInt(Me.StrategicFit * 10000 - Me.Risiko * 1000 + tmp(4) * 10).ToString(formatStr) & Me.name


                Case Else
                    ' nimm die Default- Lösung 
                    tmpResult = Me.name

            End Select

            getSortKeyForConstellation = tmpResult

        End Get
    End Property


    ''' <summary>
    ''' gibt die Anzahl insgesamt definierter CustomFields zurück  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getCustomFieldsCount() As Integer
        Get

            Dim tmpResult As Integer = Me.customStringFields.Count +
                                        Me.customDblFields.Count +
                                        Me.customBoolFields.Count

            getCustomFieldsCount = tmpResult

        End Get
    End Property

    ''' <summary>
    ''' übernimmt für die eigenen Phasen die entsprechenden Ressourcenbedarfe aus dem anderen Projekt 
    ''' berücksichtigt ausschließlich Phasen, die im eigenen Projekt vorkommen .. 
    ''' wennactualDataUntil > projektstart, dann muss dafür gesorgt werden , dass die Ist-Werte erhalten bleiben ...  
    ''' wenn das eigene PRojekt bereits Ressourcen enthält , dann wird nichts gemacht !  
    ''' es wird auch das Budget übernommen, wenn es nicht schon existiert 
    ''' </summary>
    ''' <param name="pRessources">das Projekt mit Ressourcen </param>
    ''' <returns></returns>
    Public Function updateProjectWithRessourcesFrom(ByVal pRessources As clsProjekt) As clsProjekt
        Dim ergebnisProjekt As clsProjekt = createVariant("$temp$", "Ressource-Merge")

        ' Budget übernehmen ? 
        If ergebnisProjekt.Erloes = 0 And pRessources.Erloes > 0 Then
            ergebnisProjekt.Erloes = pRessources.Erloes
        End If


        If getGesamtKostenBedarf.Sum > 0 Or pRessources.getGesamtKostenBedarf.Sum = 0 Then
            ' es gibt bereits Ressourcen oder AndereKosten .. nichts machen oder das andere Projekt hat gar keine Ressourcen
            ergebnisProjekt = Nothing
        Else

            Dim considerActualData As Boolean = pRessources.getPhase(1).hasActualData
            Dim actualDataIX As Integer = -1

            If considerActualData Then
                actualDataIX = getColumnOfDate(pRessources.actualDataUntil) - getColumnOfDate(pRessources.startDate)
            End If

            For Each cphase As clsPhase In ergebnisProjekt.AllPhases


                Dim NameID As String = cphase.nameID
                Dim otherPhase As clsPhase = pRessources.getPhaseByID(NameID)



                If Not IsNothing(otherPhase) Then
                    ' die Phase wurde gefunden, jetzt muss überprüft werden, ob es actualdataUntil gibt ... 
                    Dim considerActualPhaseData As Boolean = considerActualData And actualDataIX >= otherPhase.relStart - 1
                    Dim phaseActualIX As Integer = actualDataIX - otherPhase.relStart + 1

                    For Each role As clsRolle In otherPhase.rollenListe
                        ' gibt es für diese Rolle ActualData 
                        Dim sumActualValues As Double = 0
                        Dim sumForecastValues As Double = 0

                        If considerActualPhaseData Then
                            sumActualValues = role.getSumUntil(phaseActualIX)
                            sumForecastValues = role.getSumFrom(phaseActualIX)
                        Else
                            sumForecastValues = role.Xwerte.Sum
                        End If

                        ' jetzt muss das ggf in dem neuen PRojekt angelegt werden 
                        ' die Rolle hat bisher noch nicht existiert ...
                        Dim dimension As Integer = cphase.relEnde - cphase.relStart
                        Dim tmpRole As clsRolle = New clsRolle(dimension)

                        Dim newXwerte(dimension) As Double

                        If considerActualPhaseData Then
                            ' sind sumActualValues > 0 ? 
                            If sumActualValues > 0 Then

                                ' dann muss Start der beiden Phasen identisch sein .. 
                                If getColumnOfDate(cphase.getStartDate) = getColumnOfDate(otherPhase.getStartDate) Then
                                    ' alles ok 
                                    For i As Integer = 0 To phaseActualIX
                                        newXwerte(i) = role.Xwerte(i)
                                    Next
                                Else
                                    ' nicht zugelassen, weil ja beim alten Projekt schon Ist-Daten Zuordnungen da waren ... 
                                    Throw New ArgumentException("Phasen haben Ist-Daten, aber unterschiedliche Start-Daten ... " & cphase.name)
                                End If
                            Else
                                For i As Integer = 0 To phaseActualIX
                                    newXwerte(i) = 0
                                Next
                            End If

                            Dim csum(0) As Double
                            csum(0) = sumForecastValues
                            Dim forecastWerte() As Double = calcVerteilungAufMonate(cphase.getStartDate.Date.AddMonths(phaseActualIX + 1), cphase.getEndDate.Date, csum, 1.0)
                            For i As Integer = 0 To forecastWerte.Length - 1
                                newXwerte(i + phaseActualIX + 1) = forecastWerte(i)
                            Next

                        Else
                            Dim csum(0) As Double
                            csum(0) = sumForecastValues
                            newXwerte = calcVerteilungAufMonate(cphase.getStartDate.Date, cphase.getEndDate.Date, csum, 1.0)
                            'Dim chcknewXwerte() As Double = cphase.berechneBedarfeNew(cphase.getStartDate.Date, cphase.getEndDate.Date, csum, 1.0)
                        End If


                        With tmpRole
                            .uid = role.uid
                            .teamID = role.teamID
                            .Xwerte = newXwerte
                        End With

                        cphase.addRole(tmpRole)

                    Next

                End If
            Next

            ergebnisProjekt.variantName = variantName
            ergebnisProjekt.variantDescription = variantDescription
        End If


        updateProjectWithRessourcesFrom = ergebnisProjekt
    End Function

    ''' <summary>
    ''' wird benötigt für ensureStableIDs, erstellt die entsprechenden Input-Listen
    ''' </summary>
    ''' <returns></returns>
    Public Function getBreadCrumbIDList() As SortedList(Of String, String)

        Dim resultList As New SortedList(Of String, String)

        For Each kvp As KeyValuePair(Of String, clsHierarchyNode) In hierarchy.hryListe
            Dim fullbreadCrumb As String = Me.getBcElemName(kvp.Key)

            If fullbreadCrumb <> "" Then
                resultList.Add(fullbreadCrumb, kvp.Key)
            End If

        Next

        getBreadCrumbIDList = resultList

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="breadCrumbIDList"></param>
    ''' <returns></returns>
    Public Function anyDifferentIDs(ByVal breadCrumbIDList As SortedList(Of String, String)) As Boolean
        Dim tmpResult = False
        Dim ix As Integer = 0

        If breadCrumbIDList.Count = 0 Then
        Else
            Do While tmpResult = False And ix < breadCrumbIDList.Count
                Dim vglBreadCrumb As String = breadCrumbIDList.ElementAt(ix).Key
                Dim vglNameID As String = breadCrumbIDList.ElementAt(ix).Value

                If vglBreadCrumb = getBcElemName(vglNameID) Then
                    ix = ix + 1
                Else
                    tmpResult = True
                End If
            Loop

        End If

        anyDifferentIDs = tmpResult
    End Function

    ''' <summary>
    ''' stellt anhand der sortierten Listen mit jeweils Key=breadcrumb, value=nameID sicher, dass ein und derselbe BreadCrumb immer die gleiche NameID hat
    ''' wenn keine Anpasusngen notwendig waren: kommt Nothing zurück ... 
    ''' </summary>
    ''' <param name="baseLineList"></param>
    ''' <param name="lastProjList"></param>
    ''' <returns></returns>
    Public Function ensureStableIDs(ByVal baseLineList As SortedList(Of String, String), ByVal lastProjList As SortedList(Of String, String)) As clsProjekt
        Dim newproj As clsProjekt = Nothing

        ' nur wenn es bereits Unterschiede gbt, etwas machen
        If anyDifferentIDs(baseLineList) Or anyDifferentIDs(lastProjList) Then
            newproj = New clsProjekt
            Try
                With newproj
                    .name = Me.name
                    .farbe = Me.farbe
                    .Schrift = Schrift
                    .Schriftfarbe = Schriftfarbe
                    .VorlagenName = VorlagenName
                    .Risiko = _Risiko
                    .StrategicFit = _StrategicFit
                    .Erloes = Erloes
                    .description = _description
                    .variantName = _variantName
                    .variantDescription = _variantDescription
                    .volume = _volume
                    .complexity = _complexity
                    .businessUnit = _businessUnit
                    .projectType = _projectType
                    .StartOffset = _StartOffset
                    .startDate = _startDate
                    .earliestStartDate = _earliestStartDate
                    .latestStartDate = _latestStartDate
                    .earliestStart = _earliestStart
                    .latestStart = _latestStart
                    .leadPerson = _leadPerson
                    .Status = _Status
                    .extendedView = Me.extendedView
                    .actualDataUntil = Me.actualDataUntil
                    .kundenNummer = Me.kundenNummer
                    .vpID = Me.vpID

                End With

                ' jetzt müssen die Phasen komplett neu aufgebaut werden ...
                ' analog dem Import aus awinImportProjectmitHry
                Dim curBreadCrumb As String = ""
                Dim curElemID As String = rootPhaseName
                Dim newElemID As String = ""
                Dim curName As String = ""
                Dim curIndentLevel As Integer = 0
                Dim nextNameID As String = ""
                Dim abbruchLevel As Integer = 0
                Dim lastElemID As String = rootPhaseName

                Dim listOFHryPhases As New List(Of clsPhase)


                Do While curElemID <> ""

                    Dim cPhase As clsPhase = Nothing
                    Dim newPhase As clsPhase = Nothing
                    Dim cMilestone As clsMeilenstein = Nothing
                    Dim newMilestone As clsMeilenstein = Nothing
                    Dim isMilestone As Boolean = elemIDIstMeilenstein(curElemID)
                    Dim myFullBreadCrumb As String = getBcElemName(curElemID)
                    Dim myNameID As String = ""

                    If Not isMilestone Then
                        ' es handelt sich um eine Phase
                        cPhase = getPhaseByID(curElemID)
                        curIndentLevel = hierarchy.getIndentLevel(curElemID)
                        Dim parentID As String = hierarchy.getParentIDOfID(curElemID)

                        ' jetzt muss die NameID ggf neu betimmt werden 
                        If baseLineList.ContainsKey(myFullBreadCrumb) Then
                            myNameID = baseLineList.Item(myFullBreadCrumb)
                        ElseIf lastProjList.ContainsKey(myFullBreadCrumb) Then
                            myNameID = lastProjList.Item(myFullBreadCrumb)
                        Else
                            If Not baseLineList.ContainsValue(curElemID) And Not lastProjList.ContainsValue(curElemID) Then
                                myNameID = curElemID
                            Else
                                ' bestimme die neue ElemID 
                                Dim tmpElemName As String = elemNameOfElemID(curElemID)
                                ' bestimme die größte auftretende LfdNr in baseline und lastProj

                                Dim maxLfdNr As Integer = bestimmeLfdNrMax(baseLineList, lastProjList, tmpElemName, isMilestone)
                                If maxLfdNr = -1 Then
                                    myNameID = curElemID
                                Else
                                    myNameID = calcHryElemKey(tmpElemName, isMilestone, maxLfdNr + 1)
                                End If

                            End If

                        End If

                        ' jetzt muss sichergestellt sein, dass es sich immer noch um die ID einer Phase handelt ... 
                        If Not elemIDIstMeilenstein(myNameID) Then
                            ' alles in Ordnung 
                        Else
                            myNameID = curElemID
                        End If

                        newPhase = New clsPhase(parent:=newproj)
                        newPhase.nameID = myNameID

                        ' ohne Meilensteine kopieren, die werden ja separat übernommen  
                        cPhase.copyTo(newPhase, withoutNameID:=True, withoutMS:=True)

                        If parentID <> "" Then
                            If Not IsNothing(newproj.getPhaseByID(parentID)) Then
                                ' alles ok 
                            Else
                                Dim found As Boolean = False
                                Do While parentID <> "" And Not found
                                    parentID = hierarchy.getParentIDOfID(parentID)

                                    If parentID <> "" Then
                                        found = Not IsNothing(newproj.getPhaseByID(parentID))
                                    End If

                                Loop
                            End If
                        End If

                        newproj.AddPhase(newPhase, parentID:=parentID)

                        curIndentLevel = newproj.hierarchy.getIndentLevel(newPhase.nameID)

                        If curIndentLevel <= listOFHryPhases.Count - 1 Then
                            listOFHryPhases.Item(curIndentLevel) = newPhase

                            Do While curIndentLevel < listOFHryPhases.Count - 1
                                listOFHryPhases.RemoveAt(listOFHryPhases.Count - 1)
                            Loop

                        Else
                            listOFHryPhases.Add(newPhase)
                        End If


                    Else
                        ' es handelt sich um einen Meilenstein 
                        cMilestone = getMilestoneByID(curElemID)
                        curIndentLevel = hierarchy.getIndentLevel(curElemID)
                        Dim parentPhase As clsPhase = listOFHryPhases.ElementAt(curIndentLevel - 1)

                        If Not IsNothing(parentPhase) Then
                            newMilestone = New clsMeilenstein(parent:=listOFHryPhases.ElementAt(curIndentLevel - 1))


                            ' bestimme ggf die neue ID 
                            ' jetzt muss die NameID ggf neu betimmt werden 
                            If baseLineList.ContainsKey(myFullBreadCrumb) Then
                                myNameID = baseLineList.Item(myFullBreadCrumb)
                            ElseIf lastProjList.ContainsKey(myFullBreadCrumb) Then
                                myNameID = lastProjList.Item(myFullBreadCrumb)
                            Else
                                If Not baseLineList.ContainsValue(curElemID) And Not lastProjList.ContainsValue(curElemID) Then
                                    myNameID = curElemID
                                Else
                                    ' bestimme die neue ElemID 
                                    Dim tmpElemName As String = elemNameOfElemID(curElemID)
                                    ' bestimme die größte auftretende LfdNr in baseline und lastProj

                                    Dim maxLfdNr As Integer = bestimmeLfdNrMax(baseLineList, lastProjList, tmpElemName, isMilestone)
                                    If maxLfdNr = -1 Then
                                        myNameID = curElemID
                                    Else
                                        myNameID = calcHryElemKey(tmpElemName, isMilestone, maxLfdNr + 1)
                                    End If
                                End If

                            End If

                            ' jetzt muss sichergestellt sein, dass es sich immer noch um die ID eines Meilensteinshandelt  ... 
                            If elemIDIstMeilenstein(myNameID) Then
                                ' alles in Ordnung 
                            Else
                                myNameID = curElemID
                            End If

                            cMilestone.copyTo(newMilestone, optNameID:=myNameID)

                            ' jetzt den Meilenstein zur Phase hinzufügen  
                            parentPhase.addMilestone(newMilestone)

                        Else
                            ' darf eigentlich nicht sein .. 
                            Call MsgBox("Error in ensureStableIDS ... cancelled")
                        End If



                    End If

                    curElemID = hierarchy.getNextIdOfId(curElemID, curIndentLevel, abbruchLevel)

                Loop

                ' customFields übernehmen
                Call newproj.copyCustomFieldsFrom(Me)
            Catch ex As Exception
                newproj = Nothing
            End Try


        Else
            newproj = Nothing
        End If



        ensureStableIDs = newproj
    End Function


    Public Function createVariant(ByVal variantName As String, ByVal variantDescription As String) As clsProjekt

        Dim newproj As New clsProjekt
        Me.copyTo(newproj)

        If newproj.dauerInDays <> Me.dauerInDays Then
            Throw New ArgumentException("Variant-Creation failed ...")
        End If

        With newproj
            .name = Me.name
            ' tk das muss der gleiche Timestamp sein ... nicht das heutige Datum ...
            .timeStamp = Me.timeStamp
            .actualDataUntil = Me.actualDataUntil
            .shpUID = Me.shpUID
            .tfZeile = Me.tfZeile
            .variantName = variantName
            .variantDescription = variantDescription
        End With

        createVariant = newproj

    End Function

    '''' <summary>
    '''' gibt ein Projekt zurück, das um die Ressourcen und Kostenbedarfe des otherproj reduziert wurde 
    '''' das otherproj darf nicht vor dem Me-Projekt starten od er enden. Ein Fehler wird mit Exception beendet  
    '''' </summary>
    '''' <param name="otherproj"></param>
    '''' <returns></returns>
    'Public Function subtractProject(ByVal otherproj As clsProjekt) As clsProjekt

    '    Dim ok As Boolean = True
    '    Dim myLength As Integer = Me.anzahlRasterElemente
    '    Dim otherLength As Integer = otherproj.anzahlRasterElemente

    '    Dim myStartColumn As Integer = getColumnOfDate(Me.startDate)
    '    Dim otherStartColumn As Integer = getColumnOfDate(otherproj.startDate)
    '    Dim otherIndexStart As Integer = 0

    '    ' ist es überhaupt zulässig ? 

    '    Dim newProj As clsProjekt = Me.createVariant("$Subtract$", "")

    '    If myStartColumn <= otherStartColumn Then
    '        otherIndexStart = otherStartColumn - myStartColumn
    '    Else
    '        ok = False
    '    End If

    '    If myStartColumn + myLength < otherStartColumn + otherLength Then
    '        ok = False
    '    End If

    '    If Not ok Then
    '        Throw New ArgumentException("hier kann keine Subtraktion vorgenommen werden ...")
    '    Else
    '        ' jetzt kann die Subtraktion beginnen ..

    '        ' es wird nur in der cphase(1) abgezogen
    '        Dim mycPhase As clsPhase = newProj.getPhase(1)
    '        If IsNothing(mycPhase) Then
    '            Throw New ArgumentException("es gibt keine Projekt-Phase im Projekt ...")
    '        Else
    '            Dim myRoleNames As Collection = newProj.getRoleNames
    '            Dim otherRoleNames As Collection = otherproj.getRoleNames

    '            For Each tmpRoleName As String In otherRoleNames
    '                If myRoleNames.Contains(tmpRoleName) Then
    '                    Dim myTmpRole As clsRolle = mycPhase.getRole(tmpRoleName)
    '                    Dim myValues() As Double = myTmpRole.Xwerte
    '                    Dim tmpValues() As Double = otherproj.getRessourcenBedarf(tmpRoleName)
    '                    For i As Integer = otherStartColumn To otherLength - 1
    '                        myValues(otherStartColumn) = myValues(otherStartColumn) - tmpValues(i - otherStartColumn)
    '                    Next
    '                Else
    '                    Throw New ArgumentException("Rolle existiert nicht : " & tmpRoleName)
    '                End If

    '            Next

    '        End If

    '    End If

    '    ' jetzt wieder den Me-Variant-Name einstellen 
    '    newProj.variantName = Me.variantName
    '    subtractProject = newProj

    'End Function

    ''' <summary>
    ''' macht für den Portfolio Manager aus einem Projekt mit Detail-Ressourcen-Zuordnungen  ein Projekt mit den aggregierten Werten für die 
    ''' in der summaryRoleIDs angegebenen Sammelrollen
    ''' </summary>
    ''' <param name="summaryRoleIDs"></param>
    ''' <returns></returns>
    Public Function aggregateForPortfolioMgr(ByVal summaryRoleIDs() As Integer) As clsProjekt


        Dim newProj As clsProjekt = Me.createVariant("$aggregate$", "")


        For i As Integer = 1 To CountPhases

            Dim cphase As clsPhase = getPhase(i)
            Dim newPhase As clsPhase = newProj.getPhase(i)
            Dim toDoList As New SortedList(Of String, clsRolle)
            Dim toDoListSR As New SortedList(Of String, Integer)

            For Each curRole As clsRolle In cphase.rollenListe

                ' tk 20.08.20 copiedRole eingeführt, weil ansonsten curRole verändert wird - und damit das Ausgangs-Projekt
                ' jetzt muss curRole kopiert werden - andernfalls gibt es Fehler und Seiteneffekte ..
                Dim copiedRole As New clsRolle
                Call curRole.CopyTo(copiedRole)

                Dim roleNameID As String = copiedRole.getNameID

                ' tk ist es einer Skill/Team zugeordnet 
                Dim teamID As Integer = -1
                Dim roleID As Integer = RoleDefinitions.parseRoleNameID(roleNameID, teamID)

                Dim found As Boolean = False
                Dim ix As Integer = 1

                ' tk 19.1.20 um rollen mit Skills der Skill-Gruppe zuzuordnen ... 
                ' Anfang ... 
                If teamID > 0 And RoleDefinitions.containsUid(teamID) Then

                    Do While ix <= summaryRoleIDs.Length And Not found
                        If teamID = summaryRoleIDs(ix - 1) Then
                            found = True
                        ElseIf RoleDefinitions.hasAnyChildParentRelationsship(teamID, summaryRoleIDs(ix - 1)) Then
                            found = True
                        Else
                            ix = ix + 1
                        End If
                    Loop

                End If

                If Not found Then
                    ix = 1
                End If
                ' Ende ...

                Do While ix <= summaryRoleIDs.Length And Not found

                    If copiedRole.uid <> summaryRoleIDs(ix - 1) Then
                        ' darauf achten, dass nicht unnötigerweise Rolle1 durch Rolle1 ersetzt wird 
                        If RoleDefinitions.hasAnyChildParentRelationsship(roleNameID, summaryRoleIDs(ix - 1), includingVirtualChilds:=True) Then
                            found = True

                        ElseIf RoleDefinitions.hasAnyChildParentRelationsship(copiedRole.uid, summaryRoleIDs(ix - 1)) Then
                            found = True

                        Else
                            ix = ix + 1
                        End If
                    Else
                        ix = ix + 1
                    End If

                Loop

                If found Then
                    ' in toDoList eintragen 
                    toDoList.Add(roleNameID, copiedRole)
                    toDoListSR.Add(roleNameID, summaryRoleIDs(ix - 1))
                End If

            Next

            ' jetzt müssen diese rollen gelöscht und in aggregierter Form neu aufgenommen werden 
            ' dass soll kostenneutral erfolgen ...
            If toDoList.Count > 0 Then

                ' löschen der alten, detaillierten Rollen ..
                For Each kvp As KeyValuePair(Of String, clsRolle) In toDoList

                    Dim teamID As Integer = -1
                    Dim sRoleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByID(toDoListSR.Item(kvp.Key))

                    ' jetzt wird die alte Rolle removed ..
                    newPhase.removeRoleByNameID(kvp.Key)

                    ' jetzt wird der Umrechnungsfaktor bestimmt 
                    Dim curTagessatz As Double = kvp.Value.tagessatzIntern
                    Dim sRoleTagessatz As Double = sRoleDef.tagessatzIntern
                    Dim ptFaktor As Double = 1.0
                    If curTagessatz > 0 And sRoleTagessatz > 0 Then
                        ptFaktor = curTagessatz / sRoleTagessatz
                    End If

                    ' jetzt werden die PT Werte der current Role umgerechnet ... damit die Kosten gleich bleiben: PT * tagessatz
                    For x As Integer = 0 To kvp.Value.Xwerte.Length - 1
                        kvp.Value.Xwerte(x) = ptFaktor * kvp.Value.Xwerte(x)
                    Next

                    ' jetzt wird die curRole "umdefiniert" 
                    kvp.Value.uid = sRoleDef.UID
                    kvp.Value.teamID = -1

                    ' jetzt wird sie in die Phase aufgenommen ..
                    newPhase.addRole(kvp.Value)

                Next

            End If

        Next

        newProj.variantName = variantName
        aggregateForPortfolioMgr = newProj

    End Function

    ''' <summary>
    ''' setzt den Varianten-Namen entsprechend der customUSerRole
    ''' wird üblicherweise vor dem Speichern aufgerufen:
    ''' - ein Portfolio Manager schreibt nur mit Varianten-Name "pfv" oder einem anderen Varianten-Namen
    ''' - ein Resource Manager schreibt nur mit Varianten-NAme "" oder einem anderen Varianten-Namen 
    ''' </summary>
    Public Sub setVariantNameAccordingUserRole()

        If myCustomUserRole.customUserRole = ptCustomUserRoles.PortfolioManager Then
            If variantName = "" Then
                variantName = getDefaultVariantNameAccordingUserRole()
                Status = ProjektStatus(PTProjektStati.beauftragt)          ' ur: 2.11.2019: Projekt das Vorgaben bekommt ist autom. beauftragt
            End If
        Else
            ' tk 7.7 bei allen anderen darf der Varianten-Name nicht pfv sein 
            'ElseIf myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager Or myCustomUserRole.customUserRole = ptCustomUserRoles.TeamManager Then
            If variantName = ptVariantFixNames.pfv.ToString Then
                variantName = getDefaultVariantNameAccordingUserRole()
            End If
        End If

    End Sub

    ''' <summary>
    ''' checks whether or not project has roles with resource needs where role has already left the company or is not yet part of the company
    ''' returns a collection with names of Roles ; if empty there are no such roles
    ''' </summary>
    ''' <returns></returns>
    Public Function hasRolesWithInvalidNeeds() As Collection
        Dim allInvalidNames As New Collection

        For Each cphase As clsPhase In AllPhases
            Dim invalidPhaseRoles As Collection = cphase.hasRolesWithInvalidNeeds

            If invalidPhaseRoles.Count > 0 Then

                For Each invalidName As String In invalidPhaseRoles
                    If Not allInvalidNames.Contains(invalidName) Then
                        allInvalidNames.Add(invalidName, invalidName)
                    End If
                Next

            End If
        Next


        hasRolesWithInvalidNeeds = allInvalidNames
    End Function

    ''' <summary>
    ''' true, wenn die Anzahl Phase und die einzelnen PhaseNameIDs identisch sind und ebenso die Start- und Endezeitpunkte 
    ''' </summary>
    ''' <param name="vglProj"></param>
    ''' <returns></returns>
    Public Function areHavingSamePSP(ByVal vglProj As clsProjekt) As Boolean
        Dim tmpResult As Boolean = True

        If CountPhases = vglProj.CountPhases Then
            For Each cPhase As clsPhase In AllPhases

                Dim vglPhase As clsPhase = vglProj.getPhaseByID(cPhase.nameID)
                If Not IsNothing(vglPhase) Then
                    ' erstmal alles ok
                    tmpResult = tmpResult And cPhase.getStartDate.Date = vglPhase.getStartDate.Date And
                                              cPhase.getEndDate.Date = vglPhase.getEndDate.Date
                    If tmpResult = False Then
                        Exit For
                    End If
                Else
                    tmpResult = False
                    Exit For
                End If
            Next
        Else
            tmpResult = False
        End If

        areHavingSamePSP = tmpResult
    End Function


    ''' <summary>
    ''' merged zwei Projekte; dabei werden die in RoleNameID Collction und CostCollection angegebenen Rollen / Kosten im aktuellen Projekt gelöscht und 
    ''' und durch die Rollen / Kosten im anderen ersetzt; aber nur die Rollen + Kinder, die in den Collections angegeben sind 
    ''' </summary>
    ''' <param name="summaryRoleIDCollection">specifics aus myCustomUserRole</param>
    ''' <param name="costCollection"></param>
    ''' <param name="mProj"></param>
    ''' <returns></returns>
    Public Function deleteAndMerge(ByVal summaryRoleIDCollection As Collection,
                                   ByVal costCollection As Collection,
                                   ByVal mProj As clsProjekt) As clsProjekt

        Dim newProj As clsProjekt = Nothing

        ' es wird überprüft , ob die beiden identische Struktur haben ... 
        Dim areHavingSameStructure As Boolean = areHavingSamePSP(mProj)


        If areHavingSameStructure Then
            newProj = Me.deleteRolesAndCosts(summaryRoleIDCollection, costCollection, True)

            ' jetzt wird für jede Phase / Rolle überprüft, ob sie übernommen werden muss 
            For Each cPhase As clsPhase In mProj.AllPhases

                Dim newprojPhase As clsPhase = newProj.getPhaseByID(cPhase.nameID)

                ' jetzt alle Kosten übernehmen 
                For Each tmpRole As clsRolle In cPhase.rollenListe

                    Dim found As Boolean = False
                    For Each srNameID As String In summaryRoleIDCollection

                        Dim teamID As Integer = -9
                        Dim summaryRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(srNameID, teamID)

                        If RoleDefinitions.hasAnyChildParentRelationsship(tmpRole.getNameID, summaryRole.UID) = True Then
                            ' merken, dass die Rolle übernommen werden muss 
                            found = True
                            Exit For
                        End If
                    Next

                    If found Then
                        newprojPhase.addRole(tmpRole)
                    End If

                Next


                If Not IsNothing(costCollection) Then
                    For Each tmpCost As clsKostenart In cPhase.kostenListe

                        If costCollection.Contains(tmpCost.name) Then
                            newprojPhase.AddCost(tmpCost)
                        End If

                    Next
                End If

            Next

            newProj.variantName = Me.variantName
            newProj.timeStamp = mProj.timeStamp

        Else
            deleteAndMerge = Nothing
        End If


        deleteAndMerge = newProj
    End Function

    ''' <summary>
    ''' für AllianzImport 2
    ''' Dimension double() und phaseNameIDs() muss gleich sein
    ''' </summary>
    ''' <param name="rolePhaseValues">enthält rcNameID als Key, dann die Summenwerte für die angegebenen Phasen</param>
    ''' <param name="costPhaseValues">enthält CostName als key, dann die Summenwerte für die angegebenen Phasen</param>
    ''' <param name="phaseNameIDs"></param>
    ''' <returns></returns>
    Public Function merge(ByVal rolePhaseValues As SortedList(Of String, Double()),
                          ByVal costPhaseValues As SortedList(Of String, Double()),
                          ByVal phaseNameIDs As String(),
                          ByVal addWhenexisting As Boolean) As clsProjekt

        Dim newProj As clsProjekt = Me.createVariant("$merge$", "")
        ' hier passiert das jetzt 
        Dim anzPhasen As Integer = phaseNameIDs.Length

        For ip = 1 To anzPhasen

            Dim cphase As clsPhase = newProj.getPhaseByID(phaseNameIDs(ip - 1))
            If Not IsNothing(cphase) Then

                ' jetzt dieser Phase die Rollen und entsprechenden Werte zuordnen 
                If Not IsNothing(rolePhaseValues) Then
                    For Each kvp As KeyValuePair(Of String, Double()) In rolePhaseValues

                        Dim roleSumme As Double = kvp.Value(ip - 1)
                        If roleSumme > 0 Then
                            Call cphase.addCostRole(kvp.Key, roleSumme, True, addWhenexisting)
                        End If

                    Next
                End If

                ' jetzt die Kostenarten aufnehmen 
                If Not IsNothing(costPhaseValues) Then
                    For Each kvp As KeyValuePair(Of String, Double()) In costPhaseValues

                        Dim costSum As Double = kvp.Value(ip - 1)
                        If costSum > 0 Then
                            Call cphase.addCostRole(kvp.Key, costSum, False, addWhenexisting)
                        End If

                    Next
                End If

            End If
        Next

        ' ende Aktionen
        newProj.variantName = Me.variantName
        merge = newProj
    End Function


    ''' <summary>
    ''' gibt ein Projekt zurück, das die Vereinigung der beiden Projekte darstellt. 
    ''' Das Me Projekt muss ein Union Projekt sein
    ''' Es werden die Ressourcenbedarfe vereinigt. Wenn die Projekte zu unterschiedlichen Zeiten beginnen und unterschiedlich lang sind, so wird das 
    ''' ebenfalls berücksichtigt - im Vergleich zu addProject. Das neue Projekt hat keinerlei Phasen-Struktur
    ''' </summary>
    ''' <param name="otherProj"></param>
    ''' <returns></returns>
    Public Function unionizeWith(ByVal otherProj As clsProjekt) As clsProjekt

        Dim newStartdate As Date
        Dim newEndeDate As Date


        If Me.startDate <= otherProj.startDate Then
            newStartdate = Me.startDate
        Else
            newStartdate = otherProj.startDate
        End If

        If Me.endeDate >= otherProj.endeDate Then
            newEndeDate = Me.endeDate
        Else
            newEndeDate = otherProj.endeDate
        End If

        Dim newProj As New clsProjekt(Me.name, True, newStartdate, newEndeDate)

        ' jetzt werden die Attribute neu gesetzt ...
        With newProj

            .farbe = Me.farbe
            .Schrift = Me.Schrift
            .Schriftfarbe = Me.Schriftfarbe
            .VorlagenName = ""
            .Risiko = Me.Risiko
            .StrategicFit = Me.StrategicFit
            .Erloes = Me.Erloes + otherProj.Erloes
            .description = Me.description
            .variantName = Me.variantName
            .variantDescription = Me.variantDescription

            .businessUnit = Me.businessUnit

            .startDate = newStartdate
            '.earliestStartDate = _earliestStartDate
            '.latestStartDate = _latestStartDate
            '.earliestStart = _earliestStart
            '.latestStart = _latestStart
            .leadPerson = Me.leadPerson
            .Status = Me.Status
            .extendedView = Me.extendedView
            .movable = Me.movable




        End With

        ' ------------------------------------------------------------------------------------------------------
        ' Holen der Rootphase - die wurde ja bereits mit dem New oben angelegt ... 
        ' ------------------------------------------------------------------------------------------------------
        Dim cPhase As clsPhase = newProj.getPhase(1)

        Dim myLength As Integer = Me.anzahlRasterElemente
        Dim otherLength As Integer = otherProj.anzahlRasterElemente
        Dim newLength As Integer = newProj.anzahlRasterElemente

        Dim myStartColumn As Integer = getColumnOfDate(Me.startDate)
        Dim otherStartColumn As Integer = getColumnOfDate(otherProj.startDate)
        Dim myIndexStart As Integer, otherIndexStart As Integer

        If myStartColumn <= otherStartColumn Then
            myIndexStart = 0
            otherIndexStart = otherStartColumn - myStartColumn
        Else
            otherIndexStart = 0
            myIndexStart = myStartColumn - otherStartColumn
        End If

        ' jetzt werden die Role-Values von me übertragen , an dieser stelle gibt es noch keine Rollen im neuen Projekt!
        Dim tmpRoleNameIDs As Collection = Me.getRoleNameIDs()
        Dim newValues() As Double

        For Each tmpRoleNameID As String In tmpRoleNameIDs
            ' zurücksetzen 
            ReDim newValues(newLength - 1)

            Dim myValues() As Double = Me.getRessourcenBedarf(tmpRoleNameID)
            Dim newRole As New clsRolle(newLength - 1)

            With newRole
                Dim teamID As Integer = -1
                .uid = RoleDefinitions.parseRoleNameID(tmpRoleNameID, teamID)
                .teamID = teamID

                For ix As Integer = myIndexStart To myIndexStart + myLength - 1
                    newValues(ix) = myValues(ix - myIndexStart)
                Next
                .Xwerte = newValues
            End With

            cPhase.addRole(newRole)

        Next

        ' jetzt werden die Cost-Values von me übertragen , an dieser stelle gibt es noch keine Kosten im neuen Projekt!
        Dim tmpCosts As Collection = Me.getCostNames

        For Each tmpCost As String In tmpCosts
            ' zurücksetzen 
            ReDim newValues(newLength - 1)

            Dim myValues() As Double = Me.getKostenBedarf(tmpCost)
            Dim newCost As New clsKostenart(newLength - 1)

            With newCost
                .KostenTyp = CostDefinitions.getCostdef(tmpCost).UID
                For ix As Integer = myIndexStart To myIndexStart + myLength - 1
                    newValues(ix) = myValues(ix - myIndexStart)
                Next
                .Xwerte = newValues
            End With

            cPhase.AddCost(newCost)

        Next

        ' jetzt werden die Role-Values von otherProj übertragen , an dieser stelle gibt es evtl bereits diese Rolle im neuen Projekt!
        tmpRoleNameIDs = otherProj.getRoleNameIDs

        For Each tmpRoleNameID As String In tmpRoleNameIDs
            ' zurücksetzen 
            Dim newRole As clsRolle = newProj.getPhase(1).getRoleByRoleNameID(tmpRoleNameID)
            Dim roleDidExist As Boolean = True

            If IsNothing(newRole) Then
                roleDidExist = False
                newRole = New clsRolle(newLength - 1)
                Dim teamID As Integer = -1
                newRole.uid = RoleDefinitions.parseRoleNameID(tmpRoleNameID, teamID)
                newRole.teamID = teamID
            End If

            newValues = newRole.Xwerte

            Dim otherValues() As Double = otherProj.getRessourcenBedarf(tmpRoleNameID)


            With newRole
                For ix As Integer = otherIndexStart To otherIndexStart + otherLength - 1
                    newValues(ix) = newValues(ix) + otherValues(ix - otherIndexStart)
                Next
                .Xwerte = newValues
            End With

            If Not roleDidExist Then
                cPhase.addRole(newRole)
            End If

        Next

        ' jetzt werden die Cost-Values von me übertragen , an dieser stelle gibt es noch keine Kosten im neuen Projekt!
        tmpCosts = otherProj.getCostNames

        For Each tmpCost As String In tmpCosts
            ' zurücksetzen
            Dim newCost As clsKostenart = newProj.getPhase(1).getCost(tmpCost)
            Dim costDidExist As Boolean = True

            If IsNothing(newCost) Then
                costDidExist = False
                newCost = New clsKostenart(newLength - 1)
                newCost.KostenTyp = CostDefinitions.getCostdef(tmpCost).UID
            End If

            newValues = newCost.Xwerte

            Dim otherValues() As Double = otherProj.getKostenBedarf(tmpCost)

            With newCost
                .KostenTyp = CostDefinitions.getCostdef(tmpCost).UID
                For ix As Integer = otherIndexStart To otherIndexStart + otherLength - 1
                    newValues(ix) = newValues(ix) + otherValues(ix - otherIndexStart)
                Next
                .Xwerte = newValues
            End With

            If Not costDidExist Then
                cPhase.AddCost(newCost)
            End If


        Next

        unionizeWith = newProj
    End Function

    ''' <summary>
    ''' wird nur in AllianzIstDaten Import benötigt- soll die bereits durch vorherige Istdaten Importe entstandenen Team-MEmber Werte, die jetzt auf Nul gesetzt sind 
    ''' rausnehmen, um das nicht extrem unübersichtlich werden zu lassen 
    ''' </summary>
    Public Sub deleteTeamMembersWithNull()

        For Each cphase As clsPhase In AllPhases
            Dim delTeamRoles As New Collection
            For Each role As clsRolle In cphase.rollenListe
                Try
                    If role.teamID > -1 Or RoleDefinitions.getRoleDefByID(role.uid).getSubRoleCount = 0 Then
                        If role.Xwerte.Sum = 0 Then
                            delTeamRoles.Add(role.getNameID)
                        End If
                    End If
                Catch ex As Exception

                End Try

            Next

            If delTeamRoles.Count > 0 Then
                For Each nameID As String In delTeamRoles
                    cphase.removeRoleByNameID(nameID)
                Next
            End If
        Next

        ' jetzt müssen die 
    End Sub

    ''' <summary>
    ''' wird für alle Projekte aufgerufen, die im aktuellen Portfolio vorkommen, für die es aber keine Ist-Daten gab. 
    ''' setzt alle Werte zwischen actualDatauntil des Projektes und newActualDataUntil  auf Null
    ''' aber nur die Rollen, die in awinsettings.ActualdataOrgaUnits aufgeführt sind. 
    ''' restFromScratch steuert, ob alles von anfang an zurückgesetzt wird oder nur der Anteil, der seit old_actualDatauntil ist 
    ''' gibt false zurück, wenn nicht erfolgreich
    ''' </summary>
    ''' <param name="newActualDataUntil"></param>
    Public Function setNewActualValuesToNull(ByVal newActualDataUntil As Date, ByVal resetfromScratch As Boolean) As Boolean
        Dim result As Boolean = True

        Try
            Dim columnOFActualData As Integer = getColumnOfDate(actualDataUntil)
            Dim columnOFNewActualData As Integer = getColumnOfDate(newActualDataUntil)

            Dim considerAllRoles As Boolean = (awinSettings.ActualdataOrgaUnits = "")
            Dim actualDataParentIDs As Integer() = Nothing

            If Not considerAllRoles Then
                actualDataParentIDs = RoleDefinitions.getIDArray(awinSettings.ActualdataOrgaUnits)
            End If



            If columnOFNewActualData > columnOFActualData Then
                ' nur dann muss etwas gemacht werden 

                ' jetzt alle Werte im hproj, deren Rollen zu ActualDataOrgaUnits gehören auf Null setzen 
                For p As Integer = 1 To CountPhases
                    Dim curPhase As clsPhase = getPhase(p)
                    Dim columnOfPhaseStart As Integer = getColumnOfDate(curPhase.getStartDate)
                    Dim columnOfPhaseEnd As Integer = getColumnOfDate(curPhase.getEndDate)

                    If columnOfPhaseStart <= columnOFNewActualData And columnOfPhaseEnd > columnOFActualData Then
                        ' dann gibt es was zu tun 
                        For r = 1 To curPhase.countRoles
                            Dim curRole As clsRolle = curPhase.getRole(r)

                            ' darf nur zurückgesetzt werden, wenn auch zu externen Ressourcen Istdaten eingelesen werden ... 
                            Dim criteriaFulfilled As Boolean = True

                            If awinSettings.ExternRessourcesWithActualData = True Then
                                criteriaFulfilled = True
                            Else
                                criteriaFulfilled = Not curRole.isExtern
                            End If

                            If criteriaFulfilled Then

                                If RoleDefinitions.hasAnyChildParentRelationsship(roleNameID:=curRole.getNameID, summaryRoleIDs:=actualDataParentIDs, includingVirtualChilds:=True) Then
                                    Dim startIndex As Integer = 0
                                    If Not resetfromScratch Then
                                        startIndex = System.Math.Max(0, columnOFActualData - columnOfPhaseStart + 1)
                                    End If

                                    Dim endIndex As Integer = System.Math.Min(columnOFNewActualData - columnOfPhaseStart, curRole.getDimension)

                                    For ix As Integer = startIndex To endIndex
                                        curRole.Xwerte(ix) = 0
                                    Next
                                End If


                            End If
                        Next
                    End If

                Next
            End If

        Catch ex As Exception
            result = False
        End Try

        setNewActualValuesToNull = result

    End Function
    ''' <summary>
    ''' wird in ImportProjekteEintragen aufgerufen: stellt sicher, dass die existierenden Ist-Daten nicht durch einen Import von steckbrief überschrieben werden ... 
    ''' alle Ist-Daten aus oldProj werden dem aktuellen Projekt übertragen
    ''' es bleiben nur die Werte ab Monat ActualDataUntil +1 erhalten, alle anderen werden entweder auf Null gesetzt oder aber aus oldProj übernommen 
    ''' dabei werden nur die Werte  awinsettings.Istdaten berücksichtigt. 
    ''' Vorbedingung: vPID müssen gleich sein, VPID darf nicht leer sein
    ''' alle Termine, die in der Vergangenheit liegen und wo Ist-Daten liegen sind im jetzigen Projekt gleich 
    ''' </summary>
    ''' <param name="oldProj"></param>
    Public Sub mergeActualValues(ByVal oldProj As clsProjekt)

        Dim columnOFActualData As Integer = getColumnOfDate(oldProj.actualDataUntil)
        Dim considerAllRoles As Boolean = (awinSettings.ActualdataOrgaUnits = "")
        Dim actualDataParentIDs As Integer() = Nothing
        Dim doneKeyValues As New List(Of String)

        If Not considerAllRoles Then
            actualDataParentIDs = RoleDefinitions.getIDArray(awinSettings.ActualdataOrgaUnits)
        End If


        If vpID <> oldProj.vpID Or vpID = "" Then
            Throw New ArgumentException("VPIDs need to be identical and must not be empty string")
            Exit Sub
        End If

        If columnOFActualData < Start Then
            ' dann ist er schon fertig - es gibt nichts zu tun 
            Exit Sub
        End If

        ' jetzt alle Werte im hproj, deren Rollen zu ActualDataOrgaUnits gehören auf Null setzen 
        For p As Integer = 1 To CountPhases
            Dim curPhase As clsPhase = getPhase(p)
            Dim oldPhase As clsPhase = oldProj.getPhaseByID(curPhase.nameID)
            Dim columnOfPhaseStart As Integer = getColumnOfDate(curPhase.getStartDate)

            If columnOfPhaseStart <= columnOFActualData Then

                For r = 1 To curPhase.countRoles
                    Dim curRole As clsRolle = curPhase.getRole(r)


                    If RoleDefinitions.hasAnyChildParentRelationsship(roleNameID:=curRole.getNameID, summaryRoleIDs:=actualDataParentIDs, includingVirtualChilds:=True) Then
                        Dim endIndex As Integer = System.Math.Min(columnOFActualData - columnOfPhaseStart, curRole.getDimension)

                        For ix As Integer = 0 To endIndex
                            curRole.Xwerte(ix) = 0
                        Next

                        ' jetzt aus dem alten Stand oldProj die entsprechenden Werte holen 
                        If Not IsNothing(oldPhase) Then
                            Dim oldRole As clsRolle = oldPhase.getRole(curRole.name, curRole.teamID)
                            If Not IsNothing(oldRole) Then
                                endIndex = System.Math.Min(columnOFActualData - columnOfPhaseStart, oldRole.getDimension)
                                For ix As Integer = 0 To endIndex
                                    curRole.Xwerte(ix) = oldRole.Xwerte(ix)
                                Next

                                ' übernehme in die Liste der abgearbeiteten Paare: Phase-NameID, RoleNameID 
                                Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(curRole.uid, curRole.teamID)
                                Dim keyValue As String = curPhase.nameID & "#" & roleNameID
                                If Not doneKeyValues.Contains(keyValue) Then
                                    doneKeyValues.Add(keyValue)
                                End If

                            End If

                        Else
                            ' do nothing - curPhase remains as it is ... 
                        End If

                    End If
                Next
            End If

        Next

        ' jetzt müssen die Ist-Werte aller alten Rollen, die nicht mehr im neuen existieren, in das neue übernommen werden 
        ' wenn die zugehörige Phase nicht mehr existiert, werden die Werte in der rootphase aufgesammelt 
        For p As Integer = 1 To oldProj.CountPhases
            Dim oldPhase As clsPhase = oldProj.getPhase(p)
            Dim columnOfPhaseStart As Integer = getColumnOfDate(oldPhase.getStartDate)
            ' wird benötigt, wenn eine Phase gar nicht mehr existiert und sie in der RootPhase aufgenommen werden muss 
            Dim deltaIndex As Integer = 0

            If columnOfPhaseStart <= columnOFActualData Then

                For r As Integer = 1 To oldPhase.countRoles
                    Dim oldRole As clsRolle = oldPhase.getRole(r)
                    Dim roleNameID As String = RoleDefinitions.bestimmeRoleNameID(oldRole.uid, oldRole.teamID)
                    Dim keyValue As String = oldPhase.nameID & "#" & roleNameID
                    Dim curPhase As clsPhase = getPhaseByID(oldPhase.nameID)

                    If IsNothing(curPhase) Then
                        ' es gibt demnach im alten Projekt Ist-Daten zu Phasen, die jetzt gar nicht mehr existieren ... 
                        ' tk 13.8.2020 das darf nicht sein ! Abbruch !! 
                        Dim errMsg As String = "Phase des alten PRojektes enthält Ist-Daten und existiert nicht mehr: " & oldPhase.nameID
                        Throw New ArgumentException(errMsg)
                        Exit Sub
                        'curPhase = getPhaseByID(rootPhaseName)
                        'deltaIndex = columnOfPhaseStart - getColumnOfDate(curPhase.getStartDate)
                    End If

                    Dim laenge As Integer = curPhase.relEnde - curPhase.relStart + 1
                    Dim endIndex As Integer = System.Math.Min(columnOFActualData - getColumnOfDate(oldPhase.getStartDate), oldRole.getDimension)
                    If endIndex < 0 Then
                        endIndex = 0
                    End If

                    If Not doneKeyValues.Contains(keyValue) Then

                        Dim newRole As New clsRolle(laenge - 1)
                        With newRole
                            .uid = oldRole.uid
                            .teamID = oldRole.teamID
                            For ix As Integer = 0 To endIndex
                                .Xwerte(ix + deltaIndex) = oldRole.Xwerte(ix)
                            Next
                        End With

                        curPhase.addRole(newRole)
                        doneKeyValues.Add(keyValue)

                    End If

                Next

            End If

        Next






    End Sub

    ''' <summary>
    ''' merged die angegebenen Ist-Values für die Rolle in das Projekt 
    ''' Werte werden ersetzt ; Rahmenbedingung: die actualValues werden von vorne in die Rolle reingeschrieben 
    ''' wird in ImportAllianzIstDaten verwendet und in ImportIstdaten
    ''' </summary>
    ''' <param name="phNameID"></param>
    ''' <param name="actualValues"></param>
    Public Sub mergeActualValues(ByVal phNameID As String, ByVal actualValues As SortedList(Of String, Double()))

        Dim cPhase As clsPhase = Me.getPhaseByID(phNameID)

        If Not IsNothing(cPhase) Then

            Dim roleXwerte() As Double
            Dim dimension As Integer = cPhase.relEnde - cPhase.relStart


            For Each rvkvp As KeyValuePair(Of String, Double()) In actualValues

                Dim teamID As Integer = -1
                Dim hroleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(rvkvp.Key, teamID)
                ReDim roleXwerte(dimension)

                ' nur wenn die Rolle existiert und ausserdem Werte von größer Null hat, soll sie angelegt werden ..
                If Not IsNothing(hroleDef) And rvkvp.Value.Sum > 0 Then

                    Dim ixEnde As Integer = System.Math.Min(rvkvp.Value.Length - 1, dimension)
                    For ix As Integer = 0 To ixEnde
                        roleXwerte(ix) = rvkvp.Value(ix)
                    Next

                    Dim curRoleName As String = hroleDef.name
                    Dim curRole As clsRolle = New clsRolle(cPhase.relEnde - cPhase.relStart)

                    With curRole
                        .uid = hroleDef.UID
                        .teamID = teamID
                        .Xwerte = roleXwerte
                    End With
                    ' wenn es schon existiert, werden die Werte addiert ...
                    cPhase.addRole(curRole)

                End If

            Next

        Else
            Throw New ArgumentException("Merge Failed: Phase does not exist " & phNameID)
        End If
    End Sub
    ''' <summary>
    ''' liest den geldwerten Betrag der Rollen bis zum Monat , ggf werden sie in Abhängigkeit von resetValuesToNull auf Null gesetzt 
    ''' setzt die Werte all der Rollen / SammelRollen bis einschließlich untilMonthIncl auf Null, die in der roleCostCollection verzeichnet sind   
    ''' </summary>
    ''' <param name="roleCostCollection"></param>
    ''' <param name="relMonthCol"></param>
    ''' <param name="resetValuesToNull">gibt an, ob die entsprechenden Werte dann auf Null gesetzt werden sollen</param>
    ''' <returns></returns>
    Public Function getSetRoleCostUntil(ByVal roleCostCollection As Collection, ByVal relMonthCol As Integer, ByVal resetValuesToNull As Boolean) As Double

        Dim usedRoleNameIDs As Collection = Me.getRoleNameIDs
        Dim usedCosts As Collection = Me.getCostNames

        Dim actualValue As Double = 0.0

        For Each roleNameID As String In usedRoleNameIDs
            If isRelevantForNulling(roleNameID, roleCostCollection) Then
                actualValue = actualValue + Me.getSetRoleValuesUntil(roleNameID, relMonthCol, resetValuesToNull)
            End If
        Next

        getSetRoleCostUntil = actualValue

    End Function

    ''' <summary>
    ''' roleCostName kann eine ID sein 
    ''' </summary>
    ''' <param name="roleCostName"></param>
    ''' <param name="roleCostCollection"></param>
    ''' <returns></returns>
    Private Function isRelevantForNulling(ByVal roleCostName As String, ByVal roleCostCollection As Collection) As Boolean
        Dim tmpResult As Boolean = False

        Try

            Dim isCost As Boolean = False
            Dim teamID As Integer = -1
            Dim tmpRole As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(roleCostName, teamID)
            Dim isRole As Boolean = Not IsNothing(tmpRole)

            If Not isRole Then
                isCost = (CostDefinitions.containsName(roleCostName))
            End If

            If isRole Then

                ' darf nur zurückgesetzt werden, wenn auch zu externen Ressourcen Istdaten eingelesen werden ... 
                Dim criteriaFulfilled As Boolean = True

                If awinSettings.ExternRessourcesWithActualData = True Then
                    criteriaFulfilled = True
                Else
                    criteriaFulfilled = Not tmpRole.isExternRole
                End If

                Dim roleName As String = tmpRole.name

                If criteriaFulfilled Then
                    If RoleDefinitions.hasAnyChildParentRelationsship(roleName, roleCostCollection) Then
                        tmpResult = True
                    End If
                End If

            Else
                ' ist Kostenart - Vergleich auf Namensgleichheit reicht; es gibt noch keine Hierarchien
                tmpResult = roleCostCollection.Contains(roleCostName)
            End If
        Catch ex As Exception

        End Try


        isRelevantForNulling = tmpResult
    End Function

    ''' <summary>
    ''' liefert zurück, ob die angegegebene Phase des Projekts überhaupt noch Forecast Monate hat 
    ''' das steuert z.Bsp ob überhaupt noch editiert werden darf
    ''' </summary>
    ''' <param name="phaseNameID"></param>
    ''' <returns></returns>
    Public ReadOnly Property isPhaseWithForecastMonths(ByVal phaseNameID As String) As Boolean
        Get
            Dim tmpResult As Boolean = False
            Dim cphase As clsPhase = getPhaseByID(phaseNameID)
            If Not IsNothing(cphase) Then
                tmpResult = DateDiff(DateInterval.Month, actualDataUntil, cphase.getEndDate) > 0
            End If

            isPhaseWithForecastMonths = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' liefert für eine ausgewählte Phase zurück, wieviel für eine bestimmte Rolle, angegeben als roleID;teamID oder eine 
    ''' Kostenart, angegeben als costID bis zum angegebenen Datum aufgelaufen ist 
    ''' </summary>
    ''' <param name="phaseNameID"></param>
    ''' <param name="rcNameID"></param>
    ''' <param name="isRole"></param>
    ''' <param name="outputInEuro">true, wenn Rolle in Euro ausgegeben werden soll; in PT sonst
    ''' Kosten werden immer in Euro ausgegeben</param>
    ''' <returns></returns>
    Public Function getPhaseRCActualValues(ByVal phaseNameID As String,
                                          ByVal rcNameID As String,
                                          ByVal isRole As Boolean,
                                          ByVal outputInEuro As Boolean) As Double()

        ' die Ist-Werte sind immer die Werte vom anfang der Phase bis atualDatauntil einschließlich

        Dim tmpResult() As Double = Nothing

        Dim xWerte() As Double = Nothing
        Dim cphase As clsPhase = Me.getPhaseByID(phaseNameID)
        Dim notYetDone As Boolean = True
        Dim tagessatz As Double = 800

        If Not IsNothing(cphase) Then

            Dim pstart As Integer = getColumnOfDate(cphase.getStartDate)
            Dim pEnde As Integer = getColumnOfDate(cphase.getEndDate)
            Dim actualIX As Integer
            Dim arrayEnde As Integer

            If DateDiff(DateInterval.Month, StartofCalendar, actualDataUntil) > 0 Then
                actualIX = getColumnOfDate(actualDataUntil)
                arrayEnde = System.Math.Min(pEnde, actualIX)
            Else
                ' das ist das Abbruch-Kriterium, es gibt keine Ist-Daten
                arrayEnde = pstart - 1
            End If


            If pstart > arrayEnde Then
                ' es kann noch keine Ist-Daten geben 
                ReDim tmpResult(0)
                tmpResult(0) = 0

            ElseIf pstart <= arrayEnde Then
                ReDim tmpResult(arrayEnde - pstart)
                If isRole Then
                    ' enthält diese Phase überhaupt diese Rolle ?
                    Dim teamID As Integer = -1
                    Dim roleID As Integer = RoleDefinitions.parseRoleNameID(rcNameID, teamID)
                    If rcLists.phaseContainsRoleID(phaseNameID, roleID, teamID) Then

                        cphase = getPhaseByID(phaseNameID)
                        Dim tmpRole As clsRolle = cphase.getRoleByRoleNameID(rcNameID)
                        If Not IsNothing(tmpRole) Then
                            tagessatz = tmpRole.tagessatzIntern
                            xWerte = tmpRole.Xwerte
                        Else
                            ReDim tmpResult(0)
                            tmpResult(0) = 0
                            notYetDone = False
                        End If
                    Else
                        ReDim tmpResult(0)
                        tmpResult(0) = 0
                        notYetDone = False
                    End If
                ElseIf rcNameID <> "" Then
                    If CostDefinitions.containsName(rcNameID) Then
                        Dim costID As Integer = CostDefinitions.getCostdef(rcNameID).UID
                        If rcLists.phaseContainsCost(phaseNameID, costID) Then

                            cphase = getPhaseByID(phaseNameID)
                            Dim tmpCost As clsKostenart = cphase.getCost(rcNameID)
                            If Not IsNothing(tmpCost) Then
                                xWerte = tmpCost.Xwerte
                            Else
                                ReDim tmpResult(0)
                                tmpResult(0) = 0
                                notYetDone = False
                            End If
                        Else
                            ReDim tmpResult(0)
                            tmpResult(0) = 0
                            notYetDone = False
                        End If

                    Else
                        notYetDone = False
                    End If


                Else
                    notYetDone = False
                End If

                If notYetDone Then

                    For i As Integer = 0 To arrayEnde - pstart
                        If isRole And outputInEuro Then
                            ' mit Tagessatz multiplizieren 
                            tmpResult(i) = xWerte(i) * tagessatz
                        Else
                            tmpResult(i) = xWerte(i)
                        End If

                    Next
                Else
                    ReDim tmpResult(0)
                    tmpResult(0) = 0
                End If

            End If

        End If

        getPhaseRCActualValues = tmpResult

    End Function

    ''' <summary>
    ''' wenn für Externe Rollen keine Istdaten eingelesen werden: passiert nur für Rollen, die nicht als Extern gekennzeichnet sind 
    ''' setzt die Werte all der Rollen / Kostenarten bis einschließlich untilMonth auf Null
    ''' der geldwerte Betrag all der Werte, die auf Null gesetzt werden, wird im Return zurückgegeben
    ''' </summary>
    ''' <param name="roleNameID"></param>
    ''' <param name="relMonthCol"></param>
    ''' <returns></returns>
    Public Function getSetRoleValuesUntil(ByVal roleNameID As String, ByVal relMonthCol As Integer, ByVal resetValuesToNull As Boolean) As Double

        Dim tmpValue As Double = 0.0
        Dim teamID As Integer = -1
        Dim currentRoleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByIDKennung(roleNameID, teamID)


        If Not IsNothing(currentRoleDef) Then

            ' darf nur zurückgesetzt werden, wenn auch zu externen Ressourcen Istdaten eingelesen werden ... 
            Dim criteriaFulfilled As Boolean = True

            If awinSettings.ExternRessourcesWithActualData = True Then
                criteriaFulfilled = True
            Else
                criteriaFulfilled = Not currentRoleDef.isExternRole
            End If

            If criteriaFulfilled Then
                Dim roleUID As Integer = currentRoleDef.UID
                Dim tagessatz As Double = currentRoleDef.tagessatzIntern

                Dim listOfPhases As Collection = Me.rcLists.getPhasesWithRole(currentRoleDef.name)

                For Each phNameID As String In listOfPhases

                    Dim cPhase As clsPhase = Me.getPhaseByID(phNameID)
                    If Not IsNothing(cPhase) Then
                        With cPhase

                            If .relStart <= relMonthCol Then
                                ' jetzt die Werte auslesen und ggf. auf Null setzen 
                                'Dim cRole As clsRolle = .getRole(currentRoleDef.name)
                                Dim cRole As clsRolle = .getRoleByRoleNameID(roleNameID)

                                If Not IsNothing(cRole) Then
                                    Dim oldSum As Double = 0.0
                                    Dim ende As Integer = System.Math.Min(.relEnde, relMonthCol)

                                    For ix As Integer = 0 To ende - .relStart
                                        oldSum = oldSum + cRole.Xwerte(ix)

                                        ' hier werden ggf die Werte zurückgesetzt 
                                        If resetValuesToNull Then
                                            cRole.Xwerte(ix) = 0
                                        End If

                                    Next

                                    tmpValue = tmpValue + oldSum * tagessatz
                                End If

                            End If

                        End With

                    End If
                Next
            End If

        End If


        getSetRoleValuesUntil = tmpValue

    End Function


    ''' <summary>
    ''' gibt die Bedarfe (Phasen / Rollen / Kostenarten / Ergebnis pro Monat zurück 
    ''' </summary>
    ''' <param name="mycollection">ist eine Liste mit Namen der zu betrachtenden Phasen-, Rollen-, Kosten bzw. Ergebnisse </param>
    ''' <param name="type">gibt an , worum es sich handelt; Phase, Rolle, Kostenart, Ergebnis</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBedarfeInMonths(ByVal mycollection As Collection, ByVal type As String, Optional ByVal inclSubRoles As Boolean = False) As Double()
        Get
            Dim i As Integer, k As Integer, projektDauer As Integer = Me.anzahlRasterElemente
            Dim valueArray() As Double
            Dim tempArray() As Double
            Dim riskarray() As Double
            Dim itemName As String
            Dim projektMarge As Double = Me.ProjectMarge

            If mycollection.Count = 0 Then
                Throw New ArgumentException("interner Fehler in getBedarfeinMonths: myCollection ist leer ")
            Else
                If projektDauer > 0 Then

                    ReDim valueArray(projektDauer - 1)
                    ReDim tempArray(projektDauer - 1)
                    ReDim riskarray(projektDauer - 1)

                    Select Case type
                        Case DiagrammTypen(0)

                            itemName = CStr(mycollection.Item(1))
                            ' jetzt wird der Wert berechnet ...
                            valueArray = Me.getPhasenBedarf(itemName)

                            For i = 2 To mycollection.Count
                                itemName = CStr(mycollection.Item(i))
                                tempArray = Me.getPhasenBedarf(itemName)
                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) + tempArray(k)
                                Next
                            Next

                        Case DiagrammTypen(1)

                            itemName = CStr(mycollection.Item(1))
                            ' jetzt wird der Wert berechnet ...
                            valueArray = Me.getRessourcenBedarf(itemName, inclSubRoles:=inclSubRoles)

                            For i = 2 To mycollection.Count
                                itemName = CStr(mycollection.Item(i))
                                tempArray = Me.getRessourcenBedarf(itemName, inclSubRoles:=inclSubRoles)
                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) + tempArray(k)
                                Next
                            Next

                        Case DiagrammTypen(2)

                            itemName = CStr(mycollection.Item(1))
                            ' jetzt wird der Wert berechnet ...
                            valueArray = Me.getKostenBedarf(itemName)


                            For i = 2 To mycollection.Count
                                itemName = CStr(mycollection.Item(i))
                                tempArray = Me.getKostenBedarf(itemName)
                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) + tempArray(k)
                                Next
                            Next

                        Case DiagrammTypen(4)
                            Dim riskShare As Double
                            itemName = CStr(mycollection.Item(1))
                            ' jetzt wird der Wert berechnet ...
                            valueArray = Me.getGesamtKostenBedarf

                            If itemName = ergebnisChartName(0) Then
                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) * projektMarge
                                Next

                            ElseIf itemName = ergebnisChartName(1) Then

                                riskShare = Me.risikoKostenfaktor

                                If riskShare < 0 Then
                                    riskShare = 0
                                End If

                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) * (projektMarge - riskShare)
                                Next

                            ElseIf itemName = ergebnisChartName(3) Then

                                riskShare = Me.risikoKostenfaktor

                                If riskShare < 0 Then
                                    riskShare = 0
                                End If

                                For k = 0 To projektDauer - 1
                                    valueArray(k) = valueArray(k) * riskShare
                                Next
                            Else
                                Throw New ArgumentException("unbekannter Ergebnis-Typ ...")
                            End If

                        Case Else
                            Throw New ArgumentException("unbekannter Diagramm-Typ ...")

                    End Select
                Else
                    Throw New ArgumentException("Projekt " & Me.name & " hat keine Dauer ...")
                End If
            End If

            getBedarfeInMonths = valueArray

        End Get
    End Property

    ''' <summary>
    ''' gibt die Zahl der grünen/gelben/roten/grauen Bewertungen der Vergangenheit, der Zukunft oder beides an 
    ''' </summary>
    ''' <param name="colorIndex">0: keine Bewertung 1: grün 2: gelb 3: rot</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getNrColorIndexes(ByVal colorIndex As Integer) As Integer()
        Get
            Dim resultValues() As Integer
            Dim anzResults As Integer
            Dim anzPhasen As Integer
            Dim p As Integer, r As Integer
            Dim phase As clsPhase
            Dim result As clsMeilenstein
            Dim phasenStart As Integer, phasenEnde As Integer
            Dim monatsIndex As Integer


            If Me.anzahlRasterElemente > 0 Then

                ReDim resultValues(Me.anzahlRasterElemente - 1)


                'anzPhasen = Me.AllPhases.Count
                anzPhasen = MyBase.CountPhases

                For p = 1 To anzPhasen
                    phase = MyBase.getPhase(p)
                    With phase
                        ' Off1
                        anzResults = .countMilestones
                        phasenStart = .relStart - 1
                        phasenEnde = .relEnde - 1


                        For r = 1 To anzResults

                            Try
                                result = .getMilestone(r)
                                monatsIndex = CInt(DateDiff(DateInterval.Month, Me.startDate, result.getDate))

                                ' Sicherstellen, daß Ergebnisse, die vor oder auch nach dem Projekt erreicht werden sollen, richtig behandelt werden 

                                If monatsIndex < 0 Then
                                    monatsIndex = 0
                                ElseIf monatsIndex > Me.anzahlRasterElemente - 1 Then
                                    monatsIndex = Me.anzahlRasterElemente - 1
                                End If

                                ' hier muss noch unterschieden werden, ob der ColorIndex = 0 ist: dann muss auch mitgezählt werden, wenn ein Result ohne Bewertung da ist ...

                                If result.getBewertung(1).colorIndex = colorIndex Then
                                    resultValues(monatsIndex) = resultValues(monatsIndex) + 1
                                End If

                                'Try
                                '    If result.getBewertung(1).colorIndex = colorIndex Then
                                '        resultValues(monatsIndex) = resultValues(monatsIndex) + 1
                                '    End If
                                'Catch ex1 As Exception
                                '    ' hierher kommt er, wenn es ein Result, aber keine Bewertung gibt 
                                '    If colorIndex = 0 Then
                                '        resultValues(monatsIndex) = resultValues(monatsIndex) + 1
                                '    End If
                                'End Try



                            Catch ex As Exception

                            End Try



                        Next r

                    End With ' phase

                Next p ' Loop über alle Phasen



            Else

                ReDim resultValues(0)
                resultValues(0) = 0

            End If

            getNrColorIndexes = resultValues

        End Get
    End Property


    Public ReadOnly Property getAllResults() As String()

        Get


            Dim ResultValues() As String
            Dim anzResults As Integer
            Dim anzPhasen As Integer
            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim result As clsMeilenstein
            Dim monatsIndex As Integer



            If Me.anzahlRasterElemente > 0 Then

                ReDim ResultValues(Me.anzahlRasterElemente - 1)
                For i = 0 To Me.anzahlRasterElemente - 1
                    ResultValues(i) = ""
                Next

                anzPhasen = AllPhases.Count

                For p = 0 To anzPhasen - 1
                    phase = AllPhases.Item(p)
                    With phase
                        ' Off1
                        anzResults = .countMilestones


                        For r = 1 To anzResults

                            result = .getMilestone(r)
                            monatsIndex = CInt(DateDiff(DateInterval.Month, Me.startDate, result.getDate))
                            ' Sicherstellen, daß Ergebnisse, die vor oder auch nach dem Projekt erreicht werden sollen, richtig behandelt werden 

                            If monatsIndex >= 0 And monatsIndex <= Me.anzahlRasterElemente - 1 Then

                                ResultValues(monatsIndex) = ResultValues(monatsIndex) & vbLf & result.name &
                                                        " (" & result.getDate.ToShortDateString & ")"

                            End If





                        Next r

                    End With ' phase

                Next p ' Loop über alle Phasen



            Else

                ReDim ResultValues(0)
                ResultValues(0) = ""

            End If

            getAllResults = ResultValues

        End Get

    End Property

    ''' <summary>
    ''' liefert einen Array zurück, der die prognostizierten Zahlungseingänge für den Cash-Flow enthält; d.h der Array kann länger sein als das Projekt ... 
    ''' es werden dabei auch die Vertrags-Strafen berücksichtigt  
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getInvoicesPenalties() As Double()
        Get
            Dim invoiceValues() As Double = Nothing
            Dim penaltyValues() As Double = Nothing

            ' dieser Array nimmt die Werte auf, der key steht dabei für den relativen Monat des Projektes, key=0 im ersten Monat des Projektes 
            Dim invoiceArray As New SortedList(Of Integer, Double)
            Dim penaltyArray As New SortedList(Of Integer, Double)

            ' bestimme die Länge des Invoice-Arrays


            If _Dauer > 0 Then

                ReDim invoiceValues(_Dauer - 1)
                ReDim penaltyValues(_Dauer - 1)
                ' bestimme die Länge des Invoice-Arrays

                Dim projectStartCol As Integer = getColumnOfDate(startDate)
                Dim invoiceRelCol As Integer = -1
                Dim penaltyRelCol As Integer = -1

                For p = 1 To CountPhases
                    Dim curPhase As clsPhase = getPhase(p)

                    ' kann Rechnung gesteltl werden ? 
                    If curPhase.invoice.Key > 0 Then

                        invoiceRelCol = getColumnOfDate(curPhase.getEndDate.AddDays(curPhase.invoice.Value)) - projectStartCol

                        If invoiceRelCol > 0 Then
                            If invoiceArray.ContainsKey(invoiceRelCol) Then
                                invoiceArray.Item(invoiceRelCol) = invoiceArray.Item(invoiceRelCol) + curPhase.invoice.Key
                            Else
                                invoiceArray.Add(invoiceRelCol, curPhase.invoice.Key)
                            End If
                        End If

                    End If

                    ' Penalty relevant ?
                    If curPhase.penalty.Key > StartofCalendar Then
                        If curPhase.penalty.Key <= curPhase.getEndDate Then
                            penaltyRelCol = getColumnOfDate(curPhase.penalty.Key.AddDays(30)) - projectStartCol

                            If penaltyRelCol > 0 Then
                                If penaltyArray.ContainsKey(penaltyRelCol) Then
                                    penaltyArray.Item(penaltyRelCol) = penaltyArray.Item(penaltyRelCol) + curPhase.penalty.Value
                                Else
                                    penaltyArray.Add(penaltyRelCol, curPhase.penalty.Value)
                                End If
                            End If

                        End If
                    End If


                    For msIx As Integer = 1 To curPhase.countMilestones

                        Dim curMilestone As clsMeilenstein = curPhase.getMilestone(msIx)

                        If curMilestone.invoice.Key > 0 Then
                            invoiceRelCol = getColumnOfDate(curMilestone.getDate.AddDays(curMilestone.invoice.Value)) - projectStartCol

                            If invoiceRelCol > 0 Then
                                If invoiceArray.ContainsKey(invoiceRelCol) Then
                                    invoiceArray.Item(invoiceRelCol) = invoiceArray.Item(invoiceRelCol) + curMilestone.invoice.Key
                                Else
                                    invoiceArray.Add(invoiceRelCol, curMilestone.invoice.Key)
                                End If
                            End If

                        End If

                        ' Penalty relevant ?
                        If curMilestone.penalty.Key > StartofCalendar Then
                            If curMilestone.penalty.Key <= curMilestone.getDate Then
                                penaltyRelCol = getColumnOfDate(curMilestone.penalty.Key.AddDays(30)) - projectStartCol

                                If penaltyRelCol > 0 Then
                                    If penaltyArray.ContainsKey(penaltyRelCol) Then
                                        penaltyArray.Item(penaltyRelCol) = penaltyArray.Item(penaltyRelCol) + curMilestone.penalty.Value
                                    Else
                                        penaltyArray.Add(penaltyRelCol, curMilestone.penalty.Value)
                                    End If
                                End If

                            End If
                        End If


                    Next msIx

                Next p

                If invoiceArray.Count > 0 Or penaltyArray.Count > 0 Then

                    Dim lenInvoices As Integer = 0
                    Dim lenPenalties As Integer = 0
                    If invoiceArray.Count > 0 Then
                        lenInvoices = invoiceArray.Last.Key
                    End If
                    If penaltyArray.Count > 0 Then
                        lenPenalties = penaltyArray.Last.Key
                    End If
                    Dim invPenMax As Integer = System.Math.Max(lenInvoices, lenPenalties)
                    Dim resultDimension As Integer = System.Math.Max(_Dauer - 1, invPenMax)
                    ReDim invoiceValues(resultDimension)

                    For Each kvp As KeyValuePair(Of Integer, Double) In invoiceArray
                        invoiceValues(kvp.Key) = kvp.Value
                    Next

                    For Each kvp As KeyValuePair(Of Integer, Double) In penaltyArray
                        invoiceValues(kvp.Key) = invoiceValues(kvp.Key) - kvp.Value
                    Next

                Else
                    invoiceValues = Nothing
                End If

            End If

            getInvoicesPenalties = invoiceValues

        End Get
    End Property

    ''' <summary>
    ''' setzt alle Traffic-Lights, %Done, Verantwortlichkeiten etc. zurück 
    ''' </summary>
    Public Sub resetTrafficLightsEtc()

        For p As Integer = 1 To CountPhases

            Dim curPhase As clsPhase = getPhase(p)

            curPhase.verantwortlich = ""
            curPhase.ampelStatus = 0
            curPhase.ampelErlaeuterung = ""
            curPhase.percentDone = 0
            curPhase.invoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
            curPhase.penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0)

            For mIX As Integer = 1 To curPhase.countMilestones
                Dim curMS As clsMeilenstein = curPhase.getMilestone(mIX)
                curMS.verantwortlich = ""
                curMS.ampelStatus = 0
                curMS.ampelErlaeuterung = ""
                curMS.percentDone = 0
                curMS.invoice = New KeyValuePair(Of Double, Integer)(0.0, 0)
                curMS.penalty = New KeyValuePair(Of Date, Double)(Date.MaxValue, 0)
            Next

        Next


    End Sub

    '''' <summary>
    '''' nur temporär: setzt bei Meilensteinen des Namens gleichverteilt eine Rechnung an .. 
    '''' </summary>
    '''' <param name="name"></param>
    'Public Sub setMilestoneInvoices(ByVal name As String)

    '    Dim msNameIndices() As Integer
    '    msNameIndices = Me.hierarchy.getMilestoneHryIndices(name)
    '    Dim anzMilestones As Integer = msNameIndices.Length

    '    If anzMilestones = 0 Or msNameIndices(0) = 0 Then
    '        Exit Sub
    '    End If


    '    Try
    '        Dim singleInvoice As Double = budgetWerte.Sum / anzMilestones * 1.15
    '        For Each msID As Integer In msNameIndices
    '            Dim msnameID As String = Me.hierarchy.getIDAtIndex(msID)
    '            Dim curMS As clsMeilenstein = Me.getMilestoneByID(msnameID)

    '            curMS.invoice = New KeyValuePair(Of Double, Integer)(singleInvoice, 30)
    '        Next
    '    Catch ex As Exception

    '    End Try


    'End Sub

    ''' <summary>
    ''' gibt den Bedarf der Rolle in dem Monat X an; X=1 entspricht StartofCalendar usw.
    '''   
    ''' </summary>
    ''' <param name="mycollection"></param>
    ''' <param name="type"></param>
    ''' <param name="monat"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getBedarfeInMonth(mycollection As Collection, type As String, monat As Integer, Optional ByVal inclSubRoles As Boolean = False) As Double


        Get
            Dim valueArray() As Double
            Dim tmpValue As Double
            Dim projektDauer As Integer = Me.anzahlRasterElemente
            Dim start As Integer = Me.Start

            If mycollection.Count = 0 Then
                Throw New ArgumentException("interner Fehler in getBedarfeinMonth: myCollection ist leer ")
            Else
                If projektDauer > 0 Then
                    ReDim valueArray(projektDauer - 1)
                    valueArray = Me.getBedarfeInMonths(mycollection, type, inclSubRoles)
                    If monat >= start And monat <= start + projektDauer - 1 Then
                        tmpValue = valueArray(monat - start)
                    Else
                        tmpValue = 0.0
                    End If
                Else
                    Throw New ArgumentException("getBedarfeinMonth: Projekt hat keine Dauer, " & Me.name)
                End If

            End If

            getBedarfeInMonth = tmpValue

        End Get
    End Property

    ''' <summary>
    ''' gibt die Ressourcenbedarfe des Projektes im Zeitraum für die angegebene Rolle zurück , bewertet in PT oder in T€ 
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="roleNameID"></param>
    ''' <param name="inclSubRoles">optional: incl all subroles, Default: false</param>
    ''' <param name="outPutInEuro">wenn true, werden die Werte in Euro umgerechnet</param> 
    ''' <returns></returns>
    Public ReadOnly Property getResourceValuesInTimeFrame(ByVal von As Integer, ByVal bis As Integer,
                                                          ByVal roleNameID As String,
                                                          Optional ByVal inclSubRoles As Boolean = False,
                                                          Optional ByVal outPutInEuro As Boolean = False) As Double()
        Get
            Dim ergebnisValues() As Double
            Dim timeFrame As Integer = bis - von
            ReDim ergebnisValues(timeFrame)
            Dim projektDauer As Integer = Me.anzahlRasterElemente
            Dim start As Integer = Me.Start
            Dim valueArray() As Double

            If projektDauer > 0 Then

                ReDim valueArray(projektDauer - 1)
                valueArray = Me.getRessourcenBedarf(roleNameID, inclSubRoles:=inclSubRoles, outPutInEuro:=outPutInEuro)
                ' hier werden die Werte übernommen, die in den Zeitraum fallen ...
                ergebnisValues = calcArrayIntersection(von, bis, start, start + projektDauer - 1, valueArray)

            End If

            getResourceValuesInTimeFrame = ergebnisValues
        End Get
    End Property

    ''' <summary>
    ''' gibt die Kostenbedarfe des Projektes im Zeitraum für die angegebene Kostenart zurück , immer in T€ 
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <param name="costName"></param>
    ''' <returns></returns>
    Public ReadOnly Property getCostValuesInTimeFrame(ByVal von As Integer, ByVal bis As Integer,
                                                          ByVal costName As String) As Double()
        Get
            Dim ergebnisValues() As Double
            Dim timeFrame As Integer = bis - von
            ReDim ergebnisValues(timeFrame)
            Dim projektDauer As Integer = Me.anzahlRasterElemente
            Dim start As Integer = Me.Start
            Dim valueArray() As Double

            If projektDauer > 0 Then

                ReDim valueArray(projektDauer - 1)
                valueArray = Me.getKostenBedarf(costName)
                ' hier werden die Werte übernommen, die in den Zeitraum fallen ...
                ergebnisValues = calcArrayIntersection(von, bis, start, start + projektDauer - 1, valueArray)

            End If

            getCostValuesInTimeFrame = ergebnisValues
        End Get
    End Property

    ''' <summary>
    ''' gibt die Summe aller Ressourcen des Projektes im angegebenen Zeitraum zurück  
    ''' </summary>
    ''' <param name="von"></param>
    ''' <param name="bis"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllResBedarfimZeitraum(ByVal von As Integer, ByVal bis As Integer) As Double
        Get
            Dim valueArray() As Double
            Dim ergArray() As Double
            Dim tmpValue As Double = 0.0
            Dim projektDauer As Integer = Me.anzahlRasterElemente
            Dim start As Integer = Me.Start


            If projektDauer > 0 Then
                ReDim valueArray(projektDauer - 1)
                valueArray = Me.getAlleRessourcen

                ergArray = calcArrayIntersection(von, bis, start, start + projektDauer - 1, valueArray)
                tmpValue = ergArray.Sum
            Else
                tmpValue = 0.0
            End If

            getAllResBedarfimZeitraum = tmpValue

        End Get
    End Property


    Public ReadOnly Property hasDifferentRoleNeeds(ByVal compareProj As clsProjekt, roleName As String) As Boolean
        Get
            Dim myArray() As Double
            Dim hisArray() As Double
            Dim istIdentisch As Boolean = True
            Dim i As Integer


            Try
                myArray = Me.getRessourcenBedarf(roleName)
                hisArray = compareProj.getRessourcenBedarf(roleName)
                If myArray.Length <> hisArray.Length Then
                    istIdentisch = False
                End If
                i = 0
                While i <= myArray.Length - 1 And istIdentisch
                    If myArray(i) <> hisArray(i) Then
                        istIdentisch = False
                    Else
                        i = i + 1
                    End If
                End While
            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

            hasDifferentRoleNeeds = Not istIdentisch

        End Get
    End Property

    Public ReadOnly Property hasActualValues As Boolean
        Get
            hasActualValues = startDate < actualDataUntil
        End Get
    End Property

    Public ReadOnly Property hasDifferentCostNeeds(ByVal compareProj As clsProjekt, costName As String) As Boolean
        Get
            Dim myArray() As Double
            Dim hisArray() As Double
            Dim istIdentisch As Boolean = True
            Dim i As Integer

            Try
                myArray = Me.getKostenBedarf(costName)
                hisArray = compareProj.getKostenBedarf(costName)
                If myArray.Length <> hisArray.Length Then
                    istIdentisch = False
                End If
                i = 0
                While i <= myArray.Length - 1 And istIdentisch
                    If myArray(i) <> hisArray(i) Then
                        istIdentisch = False
                    Else
                        i = i + 1
                    End If
                End While

            Catch ex As Exception
                Call MsgBox(ex.Message)
            End Try

            hasDifferentCostNeeds = Not istIdentisch

        End Get
    End Property
    ' wird wohl überhaupt nicht mehr benötigt - es gibt keine Aufrufe !? 
    ' ''' <summary>
    ' ''' kopiert alle Meilensteine, aber ohne Bewertung 
    ' ''' </summary>
    ' ''' <param name="newproj"></param>
    ' ''' <remarks></remarks>
    'Public Sub copyMilestonesTo(ByRef newproj As clsProjekt)

    '    Dim newresult As clsMeilenstein
    '    Dim newphase As clsPhase

    '    ' Kopiere die Ampel - und die Ampel-Bewertung
    '    With newproj
    '        .ampelStatus = Me.ampelStatus
    '        .ampelErlaeuterung = Me.ampelErlaeuterung
    '    End With

    '    For Each cphase In MyBase.Liste

    '        Try
    '            newphase = newproj.getPhase(cphase.name)
    '            ' wenn gefunden dann alle Results kopieren 
    '            For r = 1 To cphase.countMilestones
    '                newresult = New clsMeilenstein(parent:=newphase)
    '                cphase.getMilestone(r).CopyToWithoutBewertung(newresult)

    '                Try
    '                    newphase.addMilestone(newresult)
    '                Catch ex As Exception

    '                End Try

    '            Next

    '        Catch ex As Exception
    '            ' in diesem Falle gibt es die komplette Phase in dem Projekt nicht mehr 
    '            ' dann muss auch nichts gemacht werden 
    '        End Try


    '    Next

    'End Sub



    Public Sub copyBewertungenTo(ByRef newproj As clsProjekt)

        Dim newresult As clsMeilenstein
        Dim newphase As clsPhase

        ' Kopiere die Ampel - und die Ampel-Bewertung
        With newproj
            .ampelStatus = Me.ampelStatus
            .ampelErlaeuterung = Me.ampelErlaeuterung
        End With

        For Each cphase In MyBase.Liste

            Try
                newphase = newproj.getPhaseByID(cphase.nameID)
                ' wenn gefunden dann alle Results kopieren 
                For r = 1 To cphase.countMilestones
                    newresult = New clsMeilenstein(parent:=newphase)
                    cphase.getMilestone(r).copyTo(newresult)

                    Try
                        newphase.addMilestone(newresult)
                    Catch ex1 As Exception

                    End Try


                Next

            Catch ex As Exception
                ' in diesem Falle gibt es die komplette Phase in dem Projekt nicht mehr 
                ' dann muss auch nichts gemacht werden 
            End Try


        Next

    End Sub


    Public Overrides Sub copyTo(ByRef newproject As clsProjekt)

        Dim newphase As clsPhase
        'Dim parentID As String
        Dim origName As String = ""

        Call copyAttrTo(newproject)

        For Each hphase In MyBase.Liste
            newphase = New clsPhase(newproject)

            'parentID = Me.hierarchy.getParentIDOfID(hphase.nameID)

            hphase.copyTo(newphase)
            newproject.AddPhase(newphase)
            'newproject.AddPhase(newphase, origName:="", parentID:=parentID)
        Next

        ' Besonderheit: 17.11.15 erst durch den Aufruf con dauerindays wird die _Dauer nochmal explizit gesetzt .. 
        If Me.dauerInDays <> newproject.dauerInDays Then
            'Throw New ArgumentException("Dauern der beiden Projekte sind unterschiedlich ...")
        End If

    End Sub


    Public Overrides Sub korrCopyTo(ByRef newproject As clsProjekt, ByVal startdate As Date, ByVal endedate As Date,
                                      Optional ByVal zielRenditenVorgabe As Double = -99999.0)
        Dim p As Integer
        Dim newphase As clsPhase
        Dim oldphase As clsPhase
        Dim ProjectDauerInDays As Integer
        Dim CorrectFactor As Double
        Dim newPhaseNameID As String = ""

        Call copyAttrTo(newproject)

        Try
            With newproject
                .startDate = startdate
                .earliestStart = _earliestStart
                .latestStart = _latestStart

                ProjectDauerInDays = calcDauerIndays(startdate, endedate)
                CorrectFactor = ProjectDauerInDays / Me.dauerInDays


                For p = 1 To Me.CountPhases

                    oldphase = Me.getPhase(p)
                    newphase = New clsPhase(newproject)

                    oldphase.korrCopyTo(newphase, CorrectFactor, newPhaseNameID, zielRenditenVorgabe)

                    .AddPhase(newphase)

                Next p


            End With
        Catch ex As Exception

        End Try




    End Sub

    ''' <summary>
    ''' gibt zurück, 
    ''' in gettimeCostColor(0): ob das Projekt schneller oder langsamer als das Vergleichsprojekt ist ;
    ''' in gettimeCostColor(1): ob das Projekt günstiger oder teurer als das Vergleichsprojekt ist ;
    ''' in gettimeCostColor(2): welche Bewertung der nächste bzw. letzte  Meilenstein (in Abhängigkeit von Auswahl) hat 
    ''' 
    ''' </summary>
    ''' <param name="vproj"></param>
    ''' meist der Planungs-Stand zur Zeit der Beauftragung, oder aber der letzte Stand
    ''' <param name="auswahl">
    ''' 0: Vergleiche Projektende
    ''' 1: vergleiche nächsten Meilenstein 
    ''' </param>
    ''' <param name="refDate">
    ''' bestimmt das Datum, ab dem der nächstgelegene nächste Meilenstein gesucht wird</param>
    ''' <value>
    ''' gibt die Time Kennzahl zurück: "kleiner 1": ist schneller; "größer 1": ist langsamer
    ''' </value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTimeCostColor(ByVal vproj As clsProjekt, ByVal auswahl As Integer, ByVal showAbsoluteDiff As Boolean, ByVal refDate As Date) As Double()
        Get
            Dim tmpValue(2) As Double
            Dim curMsName As String = ""
            Dim curPhNameID As String = ""
            Dim curAbstand As Integer = 10000
            Dim tmpAbstand As Integer
            Dim tmpPhase As clsPhase
            Dim tmpColor As Integer = -1
            Dim anzResults As Integer
            Dim relMonat1 As Integer, relMonat2 As Integer
            Dim chkDate1 As Date, chkDate2 As Date, tmpDate As Date

            If auswahl = 0 Then
                If showAbsoluteDiff Then
                    tmpValue(0) = Me.dauerInDays - vproj.dauerInDays
                    tmpValue(1) = Me.getSummeKosten - vproj.getSummeKosten
                    tmpValue(2) = Me.ampelStatus
                Else
                    tmpValue(0) = Me.dauerInDays / vproj.dauerInDays
                    tmpValue(1) = Me.getSummeKosten / vproj.getSummeKosten
                    tmpValue(2) = Me.ampelStatus
                End If


            ElseIf auswahl = 1 Then

                Dim nullWert As Integer = CInt(DateDiff(DateInterval.Day, Me.startDate, refDate) + 1)

                If nullWert > Me.dauerInDays Then
                    ' Projekt ist bereits beendet ...
                    If showAbsoluteDiff Then
                        tmpValue(0) = Me.dauerInDays - vproj.dauerInDays
                        tmpValue(1) = Me.getSummeKosten - vproj.getSummeKosten
                        tmpValue(2) = Me.ampelStatus
                    Else
                        tmpValue(0) = Me.dauerInDays / vproj.dauerInDays
                        tmpValue(1) = Me.getSummeKosten / vproj.getSummeKosten
                        tmpValue(2) = Me.ampelStatus
                    End If
                Else

                    Dim vglWert1 As Integer = -1
                    Dim vglWert2 As Integer = -1

                    ' bestimme die Phase und Meilenstein , der als nächstes nach refdate kommt 
                    For p = 1 To Me.CountPhases

                        tmpPhase = Me.getPhase(p)
                        anzResults = tmpPhase.countMilestones


                        For r = 1 To anzResults
                            tmpDate = tmpPhase.getMilestone(r).getDate
                            tmpAbstand = CInt(DateDiff(DateInterval.Day, refDate, tmpDate))
                            If tmpAbstand > 0 And tmpAbstand < curAbstand Then
                                curMsName = tmpPhase.getMilestone(r).nameID
                                curPhNameID = tmpPhase.nameID
                                curAbstand = tmpAbstand
                                chkDate1 = tmpDate
                                tmpColor = tmpPhase.getMilestone(r).getBewertung(1).colorIndex
                            End If
                        Next

                        tmpDate = tmpPhase.getEndDate
                        ' falls es in dieser Phase keinen Meilenstein gab ... oder falls das Phasen Ende noch vor dem Meilenstein lag
                        If tmpPhase.dauerInDays > nullWert And tmpPhase.dauerInDays - nullWert < curAbstand Then
                            curMsName = ""
                            curPhNameID = tmpPhase.nameID
                            curAbstand = tmpPhase.dauerInDays - nullWert
                            chkDate1 = tmpDate
                            If tmpColor = -1 Then
                                tmpColor = Me.ampelStatus
                            End If
                        End If

                    Next

                    ' jetzt ist sichergestellt , daß es zumindest curPhName (current PhaseName) gibt, evtl auch curMsName (current MilestoneName)
                    If curPhNameID <> "" Then
                        vglWert1 = curAbstand + nullWert
                        ' jetzt muss der Vergleichswert2 bestimmt werden ...
                        tmpPhase = vproj.getPhaseByID(curPhNameID)

                        If IsNothing(tmpPhase) Then
                            ' im vergleichsprojekt gibt es die Phase gar nicht , also muss auf das Gesamtprojekt verglichen werden 
                            vglWert1 = Me.dauerInDays
                            vglWert2 = vproj.dauerInDays
                            chkDate1 = Me.endeDate
                            chkDate2 = vproj.endeDate
                        Else

                            If curMsName <> "" Then
                                Dim tmpResult As clsMeilenstein
                                tmpResult = tmpPhase.getMilestone(curMsName)
                                ' gibt es den Meilenstein in der Phase ? 
                                If IsNothing(tmpResult) Then

                                    ' die beiden Phasen-Ende als Vergleichskriterien nehmen 
                                    With Me.getPhaseByID(curPhNameID)
                                        vglWert1 = .startOffsetinDays + .dauerInDays
                                        chkDate1 = .getEndDate
                                    End With

                                    With tmpPhase
                                        vglWert2 = .startOffsetinDays + .dauerInDays
                                        chkDate2 = .getEndDate
                                    End With

                                Else

                                    With tmpPhase
                                        vglWert2 = CInt(.startOffsetinDays + tmpResult.offset)
                                        chkDate2 = tmpResult.getDate
                                    End With

                                End If

                            Else
                                With Me.getPhaseByID(curPhNameID)
                                    vglWert1 = .startOffsetinDays + .dauerInDays
                                    chkDate1 = .getEndDate
                                End With

                                With tmpPhase
                                    vglWert2 = .startOffsetinDays + .dauerInDays
                                    chkDate2 = .getEndDate
                                End With

                            End If

                        End If

                        relMonat1 = getColumnOfDate(chkDate1) - Me.Start
                        relMonat2 = getColumnOfDate(chkDate2) - vproj.Start

                        If showAbsoluteDiff Then
                            tmpValue(0) = vglWert1 - vglWert2

                            ' nun jeweils die Summen bis zum angegebenen Monat aufsummieren ....
                            ' ... und die Kennzahl berechnen 
                            tmpValue(1) = Me.getSummeKosten(relMonat1) - vproj.getSummeKosten(relMonat2)
                            tmpValue(2) = tmpColor
                        Else
                            tmpValue(0) = vglWert1 / vglWert2
                            ' nun jeweils die Summen bis zum angegebenen Monat aufsummieren ....
                            ' ... und die Kennzahl berechnen 
                            tmpValue(1) = Me.getSummeKosten(relMonat1) / vproj.getSummeKosten(relMonat2)
                            tmpValue(2) = tmpColor
                        End If



                    Else
                        If showAbsoluteDiff Then
                            tmpValue(0) = Me.dauerInDays - vproj.dauerInDays
                            tmpValue(1) = Me.getSummeKosten - vproj.getSummeKosten
                            tmpValue(2) = Me.ampelStatus
                        Else
                            tmpValue(0) = Me.dauerInDays / vproj.dauerInDays
                            tmpValue(1) = Me.getSummeKosten / vproj.getSummeKosten
                            tmpValue(2) = Me.ampelStatus
                        End If
                    End If

                End If



            Else
                ' Fehler: Auswahl nicht definiert 
                Throw New Exception("Fehler in getTimeIndex")
            End If


            ' Sicherstellen Konsistenzbedingung: Farbe kann nicht negativ sein  
            If tmpValue(2) < 0 Then
                tmpValue(2) = 0
            End If

            getTimeCostColor = tmpValue


        End Get
    End Property

    Public ReadOnly Property getTimeTimeColor(ByVal vproj As clsProjekt, ByVal showAbsoluteDiff As Boolean, ByVal refDate As Date) As Double()
        Get
            Dim tmpValue(2) As Double
            Dim curMsName As String = ""
            Dim curPhNameID As String = ""
            Dim curAbstand As Integer = 10000
            Dim tmpAbstand As Integer
            Dim tmpPhase As clsPhase
            Dim tmpColor As Integer = -1
            Dim anzResults As Integer
            Dim chkDate1 As Date, chkDate2 As Date, tmpDate As Date

            ' hier ist jetzt klar, was die Unterschiede im Vergleich Projektende/Projektende sind
            If showAbsoluteDiff Then
                tmpValue(0) = Me.dauerInDays - vproj.dauerInDays

            Else
                tmpValue(0) = Me.dauerInDays / vproj.dauerInDays

            End If



            Dim nullWert As Integer = CInt(DateDiff(DateInterval.Day, Me.startDate, refDate) + 1)

            If nullWert > Me.dauerInDays Then
                ' Projekt ist bereits beendet ...
                If showAbsoluteDiff Then
                    tmpValue(1) = tmpValue(0)
                    tmpValue(2) = Me.ampelStatus
                Else
                    tmpValue(1) = tmpValue(0)
                    tmpValue(2) = Me.ampelStatus
                End If
            Else

                Dim vglWert1 As Integer = -1
                Dim vglWert2 As Integer = -1

                ' bestimme die Phase und Meilenstein , der als nächstes nach refdate kommt 
                For p = 1 To Me.CountPhases

                    tmpPhase = Me.getPhase(p)
                    anzResults = tmpPhase.countMilestones


                    For r = 1 To anzResults
                        tmpDate = tmpPhase.getMilestone(r).getDate
                        tmpAbstand = CInt(DateDiff(DateInterval.Day, refDate, tmpDate))
                        If tmpAbstand > 0 And tmpAbstand < curAbstand Then
                            curMsName = tmpPhase.getMilestone(r).nameID
                            curPhNameID = tmpPhase.nameID
                            curAbstand = tmpAbstand
                            chkDate1 = tmpDate
                            tmpColor = tmpPhase.getMilestone(r).getBewertung(1).colorIndex
                        End If
                    Next

                    tmpDate = tmpPhase.getEndDate
                    ' falls es in dieser Phase keinen Meilenstein gab ... oder falls das Phasen Ende noch vor dem Meilenstein lag
                    If tmpPhase.dauerInDays > nullWert And tmpPhase.dauerInDays - nullWert < curAbstand Then
                        curMsName = ""
                        curPhNameID = tmpPhase.nameID
                        curAbstand = tmpPhase.dauerInDays - nullWert
                        chkDate1 = tmpDate
                        If tmpColor = -1 Then
                            tmpColor = Me.ampelStatus
                        End If
                    End If

                Next

                ' jetzt ist sichergestellt , daß es zumindest curPhName (current PhaseName) gibt, evtl auch curMsName (current MilestoneName)
                If curPhNameID <> "" Then
                    vglWert1 = curAbstand + nullWert
                    ' jetzt muss der Vergleichswert2 bestimmt werden ...
                    tmpPhase = vproj.getPhaseByID(curPhNameID)

                    If IsNothing(tmpPhase) Then
                        ' im vergleichsprojekt gibt es die Phase gar nicht , also muss auf das Gesamtprojekt verglichen werden 
                        vglWert1 = Me.dauerInDays
                        vglWert2 = vproj.dauerInDays
                        chkDate1 = Me.endeDate
                        chkDate2 = vproj.endeDate
                    Else

                        If curMsName <> "" Then
                            Dim tmpResult As clsMeilenstein
                            tmpResult = tmpPhase.getMilestone(curMsName)
                            ' gibt es den Meilenstein in der Phase ? 
                            If IsNothing(tmpResult) Then

                                ' die beiden Phasen-Ende als Vergleichskriterien nehmen 
                                With Me.getPhaseByID(curPhNameID)
                                    vglWert1 = .startOffsetinDays + .dauerInDays
                                    chkDate1 = .getEndDate
                                End With

                                With tmpPhase
                                    vglWert2 = .startOffsetinDays + .dauerInDays
                                    chkDate2 = .getEndDate
                                End With

                            Else

                                With tmpPhase
                                    vglWert2 = CInt(.startOffsetinDays + tmpResult.offset)
                                    chkDate2 = tmpResult.getDate
                                End With

                            End If

                        Else
                            With Me.getPhaseByID(curPhNameID)
                                vglWert1 = .startOffsetinDays + .dauerInDays
                                chkDate1 = .getEndDate
                            End With

                            With tmpPhase
                                vglWert2 = .startOffsetinDays + .dauerInDays
                                chkDate2 = .getEndDate
                            End With

                        End If

                    End If


                    If showAbsoluteDiff Then
                        tmpValue(1) = vglWert1 - vglWert2
                        tmpValue(2) = tmpColor
                    Else

                        tmpValue(1) = vglWert1 / vglWert2
                        tmpValue(2) = tmpColor
                    End If



                Else
                    If showAbsoluteDiff Then
                        tmpValue(1) = Me.dauerInDays - vproj.dauerInDays
                        tmpValue(2) = Me.ampelStatus
                    Else
                        tmpValue(1) = Me.dauerInDays / vproj.dauerInDays
                        tmpValue(2) = Me.ampelStatus
                    End If
                End If

            End If




            ' Sicherstellen Konsistenzbedingung: Farbe kann nicht negativ sein  
            If tmpValue(2) < 0 Then
                tmpValue(2) = 0
            End If

            getTimeTimeColor = tmpValue


        End Get
    End Property

    '
    ' übergibt in Project Marge die berechnete Marge: Erloes - Kosten
    '
    Public ReadOnly Property ProjectMarge() As Double


        Get
            Dim gk As Double = 10.0
            Try
                gk = Me.getGesamtKostenBedarf.Sum
                ' prüfen , ob die Marge konsistent ist mit Verhältnis Erlös und Kosten  ... 

                If gk > 0 Then
                    ProjectMarge = (Me.Erloes - gk) / gk
                Else
                    ProjectMarge = 0
                End If

            Catch ex As Exception
                'Call MsgBox("Projekt: " & Me.name & vbLf & "gk: " & gk.ToString)
                ProjectMarge = 0
            End Try


        End Get

    End Property

    Public Overrides Property earliestStart() As Integer

        Get
            earliestStart = _earliestStart
        End Get

        Set(value As Integer)
            Dim heuteColumn As Integer = getColumnOfDate(Date.Today)
            Dim reasonableValue As Integer

            If value <= 0 Then
                If Me.Start + value > heuteColumn Then
                    ' es ist zugelassen 
                    _earliestStart = value
                Else
                    ' das Projekt kann frühestens im Folge Monat beginnen  
                    reasonableValue = heuteColumn + 1 - Me.Start
                    If reasonableValue > 0 Then
                        reasonableValue = 0
                    End If
                    _earliestStart = reasonableValue
                End If

            End If
        End Set
    End Property


    Public Overrides Property latestStart() As Integer

        Get
            latestStart = _latestStart
        End Get

        Set(value As Integer)
            If value > 0 Then
                _latestStart = value
            End If

        End Set

    End Property

    Public ReadOnly Property Start() As Integer

        Get
            Start = _Start
        End Get


    End Property

    Public Property Status() As String
        Get
            Status = _Status
        End Get
        Set(value As String)
            If value = ProjektStatus(0) Or
                value = ProjektStatus(1) Or
                value = ProjektStatus(2) Or
                value = ProjektStatus(3) Or
                value = ProjektStatus(4) Or
                value = ProjektStatus(5) Or
                value = ProjektStatus(6) Then

                ' bis auf weiteres soll es keinen ChangeRequest mehr geben ..
                ' muss erst noch weiter spezifiziert werden 
                If value = ProjektStatus(PTProjektStati.ChangeRequest) Then
                    value = ProjektStatus(PTProjektStati.beauftragt)
                End If

                _Status = value
            Else
                Call MsgBox("Wert als Status nicht zugelassen: " & value)
            End If
        End Set
    End Property

    Public Property StartOffset As Integer
        Get
            StartOffset = _StartOffset
        End Get

        Set(value As Integer)
            If value >= _earliestStart And value <= _latestStart Then
                _StartOffset = value
            Else
                Call MsgBox("unzulässiger Wert für StartOffset ...")
            End If
        End Set

    End Property

    ''' <summary>
    ''' errechnet  die Position (top, left) und Größe (width, height) des Shapes, das das Projekt repräsentieren soll 
    ''' Voraussetzung: tfzeile und tfspalte haben den für das Projekt richtigen Wert  
    ''' </summary>
    ''' <param name="top"></param>
    ''' <param name="Left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub CalculateShapeCoord(ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)


        Dim projektStartdate As Date = Me.startDate
        Dim startpunkt As Integer = CInt(DateDiff(DateInterval.Day, StartofCalendar, projektStartdate))

        If startpunkt < 0 Then
            Throw New Exception("calculate Shape Coord für Phase: Projektstart liegt vor Start of Calendar ...")
        End If

        Dim projektlaenge As Integer = Me.dauerInDays

        If Me.tfZeile <= 1 Then
            Me.tfZeile = 2
        End If

        If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.anzahlRasterElemente > 0 Then
            If awinSettings.drawProjectLine Then
                top = topOfMagicBoard + (Me.tfZeile - 0.6) * boxHeight
            Else
                top = topOfMagicBoard + (Me.tfZeile - 0.95) * boxHeight
            End If

            left = (startpunkt / 365) * boxWidth * 12
            width = ((projektlaenge) / 365) * boxWidth * 12
            height = 0.8 * boxHeight
        Else
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
        End If


    End Sub

    ' ''' <summary>
    ' ''' berechnet die Koordinaten der Phase mit Nummer  phaseNr. 
    ' ''' </summary>
    ' ''' <param name="phaseNr"></param>
    ' ''' <param name="top"></param>
    ' ''' <param name="left"></param>
    ' ''' <param name="width"></param>
    ' ''' <param name="height"></param>
    ' ''' <remarks></remarks>
    'Public Sub CalculateShapeCoord(ByVal phaseNr As Integer, ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)

    '    Dim cphase As clsPhase

    '    Try

    '        Dim projektStartdate As Date = Me.startDate
    '        Dim startpunkt As Integer = CInt(DateDiff(DateInterval.Day, StartofCalendar, projektStartdate))

    '        If startpunkt < 0 Then
    '            Throw New Exception("calculate Shape Coord für Phase: Projektstart liegt vor Start of Calendar ...")
    '        End If

    '        cphase = Me.getPhase(phaseNr)
    '        Dim phasenStart As Integer = startpunkt + cphase.startOffsetinDays
    '        Dim phasenDauer As Integer = cphase.dauerInDays



    '        If Me.tfZeile > 1 And phasenStart >= 1 And phasenDauer > 0 Then


    '            If phaseNr = 1 Then
    '                Me.CalculateShapeCoord(top, left, width, height)

    '                top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight
    '                ' Änderung 28.11 jetzt wird tagesgenau positioniert 
    '                left = (phasenStart / 365) * boxWidth * 12
    '                width = ((phasenDauer) / 365) * boxWidth * 12
    '                height = 0.8 * boxHeight
    '            Else
    '                If top <= 0 Then
    '                    top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight + 0.1 * boxHeight
    '                Else
    '                    ' nichts tun : top wird an der Aufrufenden Stelle gesetzt
    '                    ' zeichneProjektinPlantafel2 Änderung 18.3.14 
    '                End If

    '                left = (phasenStart / 365) * boxWidth * 12
    '                width = ((phasenDauer) / 365) * boxWidth * 12
    '                height = 0.6 * boxHeight
    '            End If


    '        Else
    '            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & cphase.name)
    '        End If

    '    Catch ex As Exception
    '        Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name & "Phase: " & phaseNr.ToString)
    '    End Try


    'End Sub

    ''' <summary>
    ''' gibt für die angegebene Phasen-Nummer den zeilenoffset zurück sowie die 
    ''' Werte für top, left, width, height des Phasen-Shapes
    ''' </summary>
    ''' <param name="phaseNr"></param>
    ''' <param name="zeilenOffset"></param>
    ''' <param name="top"></param>
    ''' <param name="left"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <remarks></remarks>
    Public Sub calculateShapeCoord(ByVal phaseNr As Integer, ByRef zeilenOffset As Integer,
                                       ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)
        Dim cphase As clsPhase
        Dim lastEndDate As Date = StartofCalendar.AddDays(-1)


        If phaseNr > Me.CountPhases Then
            Throw New ArgumentException("es gibt diese Phasen-Numer nicht: " & phaseNr & vbLf &
                                         "Projekt: " & Me.name & ", Anzahl Phasen: " & Me.CountPhases)
        End If

        For i = 1 To phaseNr

            With Me.getPhase(i)

                'phasenNameID = .nameID
                If DateDiff(DateInterval.Day, lastEndDate, .getStartDate) < 0 Then
                    zeilenOffset = zeilenOffset + 1
                    lastEndDate = StartofCalendar.AddDays(-1)
                End If

                If DateDiff(DateInterval.Day, lastEndDate, .getEndDate) > 0 Then
                    lastEndDate = .getEndDate
                End If

            End With
        Next


        Try

            Dim projektStartdate As Date = Me.startDate
            Dim startpunkt As Integer = CInt(DateDiff(DateInterval.Day, StartofCalendar, projektStartdate))

            If startpunkt < 0 Then
                Throw New Exception("calculate Shape Coord für Phase: Projektstart liegt vor Start of Calendar ...")
            End If

            cphase = Me.getPhase(phaseNr)
            Dim phasenStart As Integer = startpunkt + cphase.startOffsetinDays
            Dim phasenDauer As Integer = cphase.dauerInDays



            If Me.tfZeile > 1 And phasenStart >= 1 And phasenDauer > 0 Then

                ' Änderung 18.3.14 Zeilenoffset gibt an, in die wievielte Zeile das geschrieben werden soll 
                If phaseNr = 1 Then
                    Me.CalculateShapeCoord(top, left, width, height)
                Else
                    cphase.calculatePhaseShapeCoord(top, left, width, height)
                    top = top + (zeilenOffset) * boxHeight
                End If


            Else
                Throw New ArgumentException("es kann kein Shape berechnet werden für : " & cphase.nameID)
            End If

        Catch ex As Exception
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name & "Phase: " & phaseNr.ToString)
        End Try


    End Sub


    'Public Sub calculateResultCoord(ByVal resultDate As Date, ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)



    '    Dim msStart As Integer = DateDiff(DateInterval.Day, StartofCalendar, resultDate)
    '    Dim faktor As Double = 0.66

    '    'Dim tagebisResult As Integer = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(Me.Start - 1), resultDate)
    '    'Dim ratio As Double = tagebisResult / anzahlTage

    '    If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.anzahlRasterElemente > 0 Then
    '        top = topOfMagicBoard + (Me.tfZeile - 1.0) * boxHeight + boxHeight * 0.05
    '        left = (msStart / 365) * boxWidth * 12 - boxHeight * 0.5 * faktor
    '        'width = boxWidth
    '        'height = boxWidth
    '        width = boxHeight * faktor
    '        height = boxHeight * faktor
    '    Else
    '        Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
    '    End If


    'End Sub

    Public Sub calculateMilestoneCoord(ByVal resultDate As Date, ByVal zeilenOffset As Integer, ByVal b2h As Double,
                                    ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)


        'Dim endDatum As Date = StartofCalendar.AddMonths(Me.Start - 1 + Dauer).AddDays(-1)
        Dim diffMonths As Integer = CInt(DateDiff(DateInterval.Month, StartofCalendar, resultDate))
        Dim dayOfMilestone As Integer = resultDate.Day
        Dim monthOfMilestone As Integer = resultDate.Month
        Dim msStart As Integer = CInt(DateDiff(DateInterval.Day, StartofCalendar, resultDate))

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


        Dim faktor As Double = 0.6

        If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.anzahlRasterElemente > 0 Then

            ' Änderung 18.3.14 Zeilenoffset gibt an, in die wievielte Zeile das geschrieben werden soll 
            ' Änderung 26.11 eine Unterscheidung zeilenoffset ist nicht notwendig 
            ' Änderung 3.1.15 es wird das Verhältnis Breite/Höhe = b2h mitübergeben, um die relative Größe der Vorlagenshapes zu erhalten 
            top = topOfMagicBoard + (Me.tfZeile - 1) * boxHeight + 0.5 * (0.8 - faktor) * boxHeight + (zeilenOffset) * boxHeight
            height = boxHeight * faktor
            width = height * b2h
            left = (diffMonths + dayOfMilestone / tageProMonat(monthOfMilestone)) * boxWidth - width / 2

        Else
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
        End If


    End Sub

    Public Sub calculateRoundedKPI(ByRef budget As Double, ByRef personalKosten As Double, ByRef sonstKosten As Double, ByRef risikoKosten As Double, ByRef ergebnis As Double,
                                   Optional roundIT As Boolean = True)

        With Me
            Dim gk As Double = .getSummeKosten

            If roundIT Then
                budget = System.Math.Round(.Erloes, mode:=MidpointRounding.ToEven)
                risikoKosten = System.Math.Round(.risikoKostenfaktor * gk, mode:=MidpointRounding.ToEven)
                personalKosten = System.Math.Round(.getAllPersonalKosten.Sum, mode:=MidpointRounding.ToEven)
                sonstKosten = System.Math.Round(.getGesamtAndereKosten.Sum, mode:=MidpointRounding.ToEven)
                ergebnis = budget - (risikoKosten + personalKosten + sonstKosten)
            Else
                budget = .Erloes
                risikoKosten = .risikoKostenfaktor * gk
                personalKosten = .getAllPersonalKosten.Sum
                sonstKosten = .getGesamtAndereKosten.Sum
                ergebnis = budget - (risikoKosten + personalKosten + sonstKosten)
            End If

        End With

    End Sub



    Public Sub calculateStatusCoord(ByVal resultDate As Date, ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)


        ' es wird geprüft, ob das Projekt nicht schon beendet ist oder noch gar nicht angefangen hat 
        Dim endDatum As Date = Me.startDate.AddDays(Me.dauerInDays - 1)

        If DateDiff(DateInterval.Month, Me.startDate, resultDate) < 0 Then
            ' Projekt-Start hat noch gar nicht stattgefunden 
            resultDate = Me.startDate
        ElseIf DateDiff(DateInterval.Month, resultDate, endDatum) < 0 Then
            resultDate = endDatum
        End If



        Dim diffMonths As Integer = CInt(DateDiff(DateInterval.Month, StartofCalendar, resultDate))
        'Dim dayOfResult As Integer = resultDate.Day
        Dim dayOfResult As Integer = 15 ' wähle die Mitte des Monats

        'Dim tagebisResult As Integer = DateDiff(DateInterval.Day, StartofCalendar.AddMonths(Me.Start - 1), resultDate)
        'Dim ratio As Double = tagebisResult / anzahlTage

        If Me.tfZeile > 1 And Me.tfspalte >= 1 And Me.anzahlRasterElemente > 0 Then
            top = topOfMagicBoard + (Me.tfZeile - 1.0) * boxHeight
            left = diffMonths * boxWidth + dayOfResult * (boxWidth / 30) - 0.5 * boxWidth

            width = boxWidth
            height = boxWidth
        Else
            Throw New ArgumentException("es kann kein Shape berechnet werden für : " & Me.name)
        End If


    End Sub
    ' '' '' '' ''' <summary>
    ' '' '' '' ''' gibt die Anzahl Zeilen zurück, die das aktuelle Projekt im "Extended Drawing Mode" benötigt 
    ' '' '' '' ''' </summary>
    ' '' '' '' ''' <returns></returns>
    ' '' '' '' ''' <remarks></remarks>
    ' '' '' ''Public ReadOnly Property calcNeededLines(ByVal selectedPhases As Collection, ByVal extended As Boolean, ByVal considerTimespace As Boolean) As Integer
    ' '' '' ''    Get

    ' '' '' ''        Dim phasenName As String
    ' '' '' ''        Dim zeilenOffset As Integer = 1
    ' '' '' ''        Dim lastEndDate As Date = StartofCalendar.AddDays(-1)
    ' '' '' ''        Dim tmpValue As Integer

    ' '' '' ''        Dim selPhaseName As String = ""
    ' '' '' ''        Dim breadcrumb As String = ""



    ' '' '' ''        If extended And selectedPhases.Count > 0 Then ' extended Sicht bzw. Report mit selektierte Phasen

    ' '' '' ''            Dim anzPhases As Integer = 0
    ' '' '' ''            Dim cphase As clsPhase = Nothing

    ' '' '' ''            For i = 1 To Me.CountPhases ' Schleife über alle Phasen eines Projektes
    ' '' '' ''                Try
    ' '' '' ''                    cphase = Me.getPhase(i)
    ' '' '' ''                    If Not IsNothing(cphase) Then

    ' '' '' ''                        ' herausfinden, ob cphase in den selektierten Phasen enthalten ist
    ' '' '' ''                        Dim found As Boolean = False
    ' '' '' ''                        Dim j As Integer = 1
    ' '' '' ''                        While j <= selectedPhases.Count And Not found

    ' '' '' ''                            Call splitHryFullnameTo2(CStr(selectedPhases(j)), selPhaseName, breadcrumb)

    ' '' '' ''                            If cphase.name = selPhaseName Then
    ' '' '' ''                                found = True
    ' '' '' ''                            End If
    ' '' '' ''                            j = j + 1
    ' '' '' ''                        End While

    ' '' '' ''                        If found Then           ' cphase ist eine der selektierten Phasen

    ' '' '' ''                            If Not considerTimespace _
    ' '' '' ''                                Or _
    ' '' '' ''                                (considerTimespace And phaseWithinTimeFrame(Me.Start, cphase.relStart, cphase.relEnde, showRangeLeft, showRangeRight)) Then

    ' '' '' ''                                With cphase

    ' '' '' ''                                    'phasenName = .name
    ' '' '' ''                                    If DateDiff(DateInterval.Day, lastEndDate, .getStartDate) < 0 Then
    ' '' '' ''                                        zeilenOffset = zeilenOffset + 1
    ' '' '' ''                                        lastEndDate = StartofCalendar.AddDays(-1)
    ' '' '' ''                                    End If

    ' '' '' ''                                    If DateDiff(DateInterval.Day, lastEndDate, .getEndDate) > 0 Then
    ' '' '' ''                                        lastEndDate = .getEndDate
    ' '' '' ''                                    End If

    ' '' '' ''                                End With
    ' '' '' ''                                anzPhases = anzPhases + 1
    ' '' '' ''                            Else

    ' '' '' ''                            End If
    ' '' '' ''                        End If
    ' '' '' ''                    End If

    ' '' '' ''                Catch ex As Exception

    ' '' '' ''                End Try



    ' '' '' ''            Next i      ' nächste Phase im Projekt betrachten



    ' '' '' ''            If anzPhases > 1 Then
    ' '' '' ''                tmpValue = zeilenOffset + 1     'ur: 17.04.2015:  +1 für die übrigen Meilensteine
    ' '' '' ''            Else
    ' '' '' ''                tmpValue = 1 + 1                ' ur: 17.04.2015: +1 für die übrigen Meilensteine
    ' '' '' ''            End If


    ' '' '' ''        ElseIf extended And selectedPhases.Count < 1 Then   ' extended Sicht bzw. Report ohne selektierte Phasen


    ' '' '' ''            For i = 1 To Me.CountPhases ' Schleife über alle Phasen eines Projektes

    ' '' '' ''                With Me.getPhase(i)

    ' '' '' ''                    phasenName = .name
    ' '' '' ''                    If DateDiff(DateInterval.Day, lastEndDate, .getStartDate) < 0 Then
    ' '' '' ''                        zeilenOffset = zeilenOffset + 1
    ' '' '' ''                        lastEndDate = StartofCalendar.AddDays(-1)
    ' '' '' ''                    End If

    ' '' '' ''                    If DateDiff(DateInterval.Day, lastEndDate, .getEndDate) > 0 Then
    ' '' '' ''                        lastEndDate = .getEndDate
    ' '' '' ''                    End If

    ' '' '' ''                End With
    ' '' '' ''            Next

    ' '' '' ''            If Me.CountPhases > 1 Then
    ' '' '' ''                tmpValue = zeilenOffset + 1      ' ur: 17.04.2015: +1 für die übrigen Meilensteine
    ' '' '' ''            Else
    ' '' '' ''                tmpValue = 1 + 1                 ' ur: 17.04.2015: +1 für die übrigen Meilensteine
    ' '' '' ''            End If

    ' '' '' ''        Else    ' keine extended Sicht (bzw. Report) 
    ' '' '' ''            tmpValue = 1
    ' '' '' ''        End If


    ' '' '' ''        calcNeededLines = tmpValue

    ' '' '' ''    End Get

    ' '' '' ''End Property

    ''' <summary>
    ''' gibt die Anzahl Zeilen zurück, die das aktuelle Projekt im "Extended Drawing Mode" benötigt 
    ''' Neu: im extendedMode wird noch nachsehen, ob selektierte Meilensteine einen Parent oder Parent/Parent ... haben
    ''' </summary>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="extended"></param>
    ''' <param name="considerTimespace"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property calcNeededLines(ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, ByVal extended As Boolean, ByVal considerTimespace As Boolean) As Integer
        Get

            Dim phasenName As String
            Dim zeilenOffset As Integer = 1
            Dim lastEndDate As Date = StartofCalendar.AddDays(-1)
            Dim tmpValue As Integer

            Dim selPhaseName As String = ""
            Dim breadcrumb As String = ""



            If extended And selectedPhases.Count > 0 Then ' extended Sicht bzw. Report mit selektierte Phasen

                Dim anzPhases As Integer = 0
                Dim cphase As clsPhase = Nothing

                For i = 1 To Me.CountPhases ' Schleife über alle Phasen eines Projektes
                    Try
                        cphase = Me.getPhase(i)
                        If Not IsNothing(cphase) Then

                            ' herausfinden, ob cphase in den selektierten Phasen enthalten ist
                            Dim found As Boolean = False
                            Dim j As Integer = 1
                            While j <= selectedPhases.Count And Not found

                                Dim type As Integer = -1
                                Dim pvName As String = ""
                                Call splitHryFullnameTo2(CStr(selectedPhases(j)), selPhaseName, breadcrumb, type, pvName)

                                If type = -1 Or
                                    (type = PTItemType.projekt And pvName = Me.name) Or
                                    (type = PTItemType.vorlage And pvName = Me.VorlagenName) Then

                                    If cphase.name = selPhaseName Then
                                        found = True
                                    End If
                                ElseIf type = PTItemType.categoryList Then

                                    Try
                                        Dim categoryItem As String = calcHryCategoryName(cphase.appearance, False)
                                        If selectedPhases.Contains(categoryItem) Then
                                            found = True
                                        End If
                                    Catch ex As Exception

                                    End Try

                                End If

                                j = j + 1
                            End While

                            If found Then           ' cphase ist eine der selektierten Phasen

                                If Not considerTimespace _
                                    Or
                                    (considerTimespace And phaseWithinTimeFrame(Me.Start, cphase.relStart, cphase.relEnde, showRangeLeft, showRangeRight)) Then

                                    With cphase

                                        'phasenName = .name
                                        If DateDiff(DateInterval.Day, lastEndDate, .getStartDate) < 0 Then
                                            zeilenOffset = zeilenOffset + 1
                                            lastEndDate = StartofCalendar.AddDays(-1)
                                        End If

                                        If DateDiff(DateInterval.Day, lastEndDate, .getEndDate) > 0 Then
                                            lastEndDate = .getEndDate
                                        End If

                                    End With
                                    anzPhases = anzPhases + 1
                                Else

                                End If
                            End If
                        End If

                    Catch ex As Exception

                    End Try



                Next i      ' nächste Phase im Projekt betrachten

                ' ur: 28.09.2015
                ' Bestimmen, zu welcher Phase die selektieren Meilenstein jeweils gezeichnet werden sollen und mitzählen, wieviele zusätzliche
                ' Zeilen benötigt werden dazu.

                Dim drawliste As New SortedList(Of String, SortedList)
                Dim addLines As Integer = 1

                If selectedMilestones.Count > 0 Then


                    Call selMilestonesToselPhase(selectedPhases, selectedMilestones, False, addLines, drawliste)

                End If


                If anzPhases > 1 Then
                    tmpValue = zeilenOffset + addLines    'ur: 17.04.2015:  +addlines für die übrigen Meilensteine
                Else
                    tmpValue = 1 + addLines              ' ur: 17.04.2015: + für die übrigen Meilensteine
                End If


            ElseIf extended And selectedPhases.Count < 1 Then   ' extended Sicht bzw. Report ohne selektierte Phasen


                For i = 1 To Me.CountPhases ' Schleife über alle Phasen eines Projektes

                    With Me.getPhase(i)

                        phasenName = .name
                        If DateDiff(DateInterval.Day, lastEndDate, .getStartDate) < 0 Then
                            zeilenOffset = zeilenOffset + 1
                            lastEndDate = StartofCalendar.AddDays(-1)
                        End If

                        If DateDiff(DateInterval.Day, lastEndDate, .getEndDate) > 0 Then
                            lastEndDate = .getEndDate
                        End If

                    End With
                Next

                If Me.CountPhases > 1 Then
                    tmpValue = zeilenOffset      ' ur: 17.04.2015: +1 für die übrigen Meilensteine
                Else
                    tmpValue = 1                ' ur: 17.04.2015: +1 für die übrigen Meilensteine
                End If

            Else    ' keine extended Sicht (bzw. Report) 
                tmpValue = 1
            End If


            calcNeededLines = tmpValue

        End Get

    End Property

    ''' <summary>
    ''' gibt die Anzahl Zeilen zurück, die die Swimlane phaseID im aktuellen Projekt im "Extended Drawing Mode" benötigt 
    ''' Aktuell ist es so, dass nur Phasen Zeilenvorschub triggern, Meilensteine werden in der obersten Phase oder in der Phase gezeichnet, 
    ''' die ihr Großvater, Ur-Großvater, etc ist     ''' 
    ''' </summary>
    ''' <param name="selectedPhaseIDs">die Liste mit den PhaseIDs, die gezeichnet werden sollen</param>
    ''' <param name="selectedMilestoneIDs">die Liste mit den MilestoneIDs, die gezeichnet werden sollen</param>
    ''' <param name="extended">wenn </param>
    ''' <param name="considerTimespace">ist ein Zeitraum zu berücksichtigen? dann triggern Phasen nur dann einen Zeilenvorschub, wenn sie im Zeitraum liegen </param>
    ''' <param name="zeitraumGrenzeL" >der linke Rand des Zeitraums; kann showRangeL sein, muss aber nicht wenn showallIfOne gesetzt ist</param>
    ''' <param name="zeitraumGrenzeR" >der rechte Rand des Zeitraums; kann showRangeR sein, muss aber nicht wenn showallIfOne gesetzt ist</param>
    ''' <param name="considerAll"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>segment-id hat nur dann einen Wert anders als "", wenn es sich um den Swimlane2 Modus handelt; wenn segmentID = swimlaneID , dann sollen die 
    ''' Meilensteine des segments gezeichnet werden ujnd die brauchen nur eine Zeile ... ; genauso, wenn der Modus extended = false ist 
    ''' </remarks>
    Public ReadOnly Property calcNeededLinesSwl(ByVal swimlaneID As String,
                                                ByVal selectedPhaseIDs As Collection, ByVal selectedMilestoneIDs As Collection,
                                                ByVal extended As Boolean, ByVal considerTimespace As Boolean,
                                                ByVal zeitraumGrenzeL As Integer, ByVal zeitraumGrenzeR As Integer,
                                                ByVal considerAll As Boolean,
                                                ByVal segmentID As String) As Integer
        Get

            Dim tmpValue As Integer

            If Not extended Or swimlaneID = segmentID Then
                tmpValue = 1
            Else


                ' jetzt wird erst mal bestimmt, von welcher Phase bis zu welcher Phase die Kind-Phasen der swimlaneID liegen
                ' dabei wird der Umstand ausgenutzt, dass in der PhasenListe 1..PhasesCount alle Kind-Phasen 
                ' unmittelbar nach der Eltern-Phase kommen ;
                ' generell können Kind-Elemente, egal ob Meilensteine oder Phasen nur in den PhasenNummern start .. ende vorkommen

                Dim startNr As Integer = 0
                Dim endNr As Integer = 0


                ' in startNr ist nachher die Phasen-Nummer der swimlane, in startNr +1 die Phasen-Nummer des ersten Kindes 
                ' in endNr ist die Phasen-Nummer des letzten Kindes 
                Call Me.calcStartEndChildNrs(swimlaneID, startNr, endNr)

                ' zum Bestimmen der optimierten Zeilenanzahl 
                ' es kann in dieser Swimlane nicht mehr als endNr-startNr Zeilen geben 
                Dim dimension As Integer = endNr - startNr
                Dim lastEndDates(dimension) As Date
                ' list of Phases dient dazu, die IDs der Phasen, die in dieser Zeile gezeichnet wurden aufzunehmen
                ' damit wird ein Cap eingeführt, das heisst keine Phase wird in der Swimlane über ihrer Eltern-Phase gezeichnet 
                Dim listOfPhases(dimension) As Collection

                For i As Integer = 0 To dimension
                    lastEndDates(i) = StartofCalendar.AddDays(-1)
                    listOfPhases(i) = New Collection
                Next

                Dim maxOffsetZeile As Integer = 1
                Dim curOffsetZeile As Integer = 1

                ' jetzt wird bestimmt, wieviele der selectedPhaseIDs, selectedMilestoneIDs denn überhaupt (Kindes-)Kinder der betrachteten Swimlane sind 
                ' es ist nicht notwendig, das bei considerAll zu machen 

                Dim childPhaseIDs As New Collection
                Dim childMilestoneIDs As New Collection

                If Not considerAll Then
                    If swimlaneID = rootPhaseName Then
                        childMilestoneIDs = Me.getPhaseByID(rootPhaseName).getMilestoneIDs
                    Else
                        childPhaseIDs = Me.schnittmengeChilds(swimlaneID, selectedPhaseIDs)
                        childMilestoneIDs = Me.schnittmengeChilds(swimlaneID, selectedMilestoneIDs)
                    End If

                End If

                Dim zeilenOffset As Integer = 1

                If Not extended Then
                    ' es wird grundsätzlich nur eine Zeile benötigt 
                    tmpValue = 1

                ElseIf childPhaseIDs.Count <= 1 And Not considerAll Then
                    ' es wird nur eine Zeile benötigt 
                    tmpValue = 1

                ElseIf swimlaneID = rootPhaseName Then
                    ' für die Darstellung der Meilensteine einer Swimlane wird nur eine Zeile benötigt  
                    tmpValue = 1

                Else
                    ' Schleife über alle Kind Phasen der Swimlane (startnr+1 bis zu endNr)
                    ' muss erst ab startnr + 1 beginnen, da phase(startNr) ja die swimlane selber ist ... 
                    Dim bestStartAtLevel As New SortedList(Of Integer, Integer)
                    Dim currentLevel As Integer = -1
                    Dim previousLevel As Integer = -1
                    Dim minBestStart As Integer = 0

                    For i = startNr + 1 To endNr
                        Try
                            Dim cPhase As clsPhase = Me.getPhase(i)
                            Dim relevant As Boolean = False
                            If Not IsNothing(cPhase) Then
                                If considerAll Then
                                    relevant = True
                                Else
                                    If childPhaseIDs.Contains(cPhase.nameID) Then
                                        relevant = True
                                    Else
                                        relevant = False
                                    End If
                                End If

                                If relevant Then           ' cphase ist eine der selektierten Phasen

                                    If Not considerTimespace _
                                        Or
                                        (considerTimespace And phaseWithinTimeFrame(Me.Start, cPhase.relStart, cPhase.relEnde,
                                                                                    zeitraumGrenzeL, zeitraumGrenzeR)) Then


                                        currentLevel = Me.hierarchy.getIndentLevel(cPhase.nameID)
                                        ' wenn es sich um ein Element handelt, das in der Hierarchie höher als das vorhergehende war , dann wird das die neue Start-Zeile 
                                        If currentLevel < previousLevel Then
                                            If bestStartAtLevel.ContainsKey(currentLevel) Then
                                                minBestStart = bestStartAtLevel(currentLevel) - 1
                                                If minBestStart < 0 Then
                                                    minBestStart = 0
                                                End If
                                            End If
                                        End If

                                        Dim requiredZeilen As Integer = Me.calcNeededLinesSwl(cPhase.nameID,
                                                                                               selectedPhaseIDs,
                                                                                               selectedMilestoneIDs,
                                                                                               extended,
                                                                                               considerTimespace, zeitraumGrenzeL, zeitraumGrenzeR,
                                                                                               considerAll, "")

                                        Dim bestStart As Integer = minBestStart
                                        ' Dim bestStart As Integer = 0
                                        ' von unten her beginnend: enthält eine der Zeilen ein Eltern- oder Großeltern-Teil 
                                        ' das ist dann der Fall, wenn der BreadCrumb der aktuellen Phase den Breadcrumb einer der Zeilen-Phasen vollständig enthält 

                                        ' tk 15.5.18
                                        ''Dim parentFound As Boolean = False
                                        ''Dim curBreadCrumb As String = Me.hierarchy.getBreadCrumb(cPhase.nameID)
                                        ''Dim ix As Integer = maxOffsetZeile

                                        ''While ix > 0 And Not parentFound

                                        ''    If listOfPhases(ix - 1).Count > 0 Then
                                        ''        Dim kx As Integer = 1
                                        ''        While kx <= listOfPhases(ix - 1).Count And Not parentFound
                                        ''            Dim vglBreadCrumb As String = Me.hierarchy.getBreadCrumb(CStr(listOfPhases(ix - 1).Item(kx)))
                                        ''            If curBreadCrumb.StartsWith(vglBreadCrumb) And curBreadCrumb.Length > vglBreadCrumb.Length Then
                                        ''                parentFound = True
                                        ''            Else
                                        ''                kx = kx + 1
                                        ''            End If
                                        ''        End While

                                        ''        If Not parentFound Then
                                        ''            ix = ix - 1
                                        ''        End If

                                        ''    Else
                                        ''        ix = ix - 1
                                        ''    End If
                                        ''End While

                                        ''If parentFound Then
                                        ''    bestStart = ix
                                        ''Else
                                        ''    bestStart = 0
                                        ''End If

                                        With cPhase

                                            'zeilenOffset = findeBesteZeile(lastEndDates, bestStart, maxOffsetZeile, .getStartDate, requiredZeilen)
                                            zeilenOffset = findeBesteZeile(lastEndDates, minBestStart, maxOffsetZeile, .getStartDate, requiredZeilen)
                                            ' wenn das aktuelle Element in der Hierarchie höher steht als das zuvor behandelte , dann wird die jetzt ermittelte Zeile als Start verwendet 
                                            If currentLevel < previousLevel Then
                                                minBestStart = zeilenOffset
                                            End If

                                            maxOffsetZeile = System.Math.Max(zeilenOffset + requiredZeilen - 1, maxOffsetZeile)


                                            ' jetzt vermerken, welche Phase in der Zeile gezeichnet wurde ...
                                            If Not listOfPhases(zeilenOffset - 1).Contains(cPhase.nameID) Then
                                                listOfPhases(zeilenOffset - 1).Add(cPhase.nameID, cPhase.nameID)
                                            End If

                                            If DateDiff(DateInterval.Day, lastEndDates(zeilenOffset - 1), .getEndDate) > 0 Then
                                                lastEndDates(zeilenOffset - 1) = .getEndDate
                                            End If


                                            'End If

                                        End With



                                    Else
                                        ' Phase ist nicht im Zeitraum, also kein Zeilenoffset notwendig, kein lastEndDate notwendig 
                                    End If
                                End If


                            End If
                        Catch ex As Exception

                        End Try

                    Next

                    tmpValue = maxOffsetZeile

                End If
            End If

            calcNeededLinesSwl = tmpValue

        End Get

    End Property

    ''' <summary>
    ''' berechnet für die gegebene Phasen-ID die Start und End-Nummer der Kind-Phasen
    ''' in der Liste der Phasen in einem Projekt sind alle Kind-Phasen unmittelbar nach der Eltern-Phase
    ''' </summary>
    ''' <param name="phaseID"></param>
    ''' <param name="startNr"></param>
    ''' <param name="endNr"></param>
    ''' <remarks></remarks>
    Public Sub calcStartEndChildNrs(ByVal phaseID As String,
                                         ByRef startNr As Integer, ByRef endNr As Integer)

        ' jetzt wird erst mal bestimmt, von welcher Phase bis zu welcher Phase die Kind-Phasen der swimlaneID liegen
        ' dabei wird der Umstand ausgenutzt, dass in der PhasenListe 1..PhasesCount alle Kind-Phasen 
        ' unmittelbar nach der Eltern-Phase kommen ;
        ' generell können Kind-Elemente, egal ob Meilensteine oder Phasen nur in den PhasenNummern start .. ende vorkommen

        Dim stillChild As Boolean = True
        Dim fullSwlBreadCrumb As String = Me.getBcElemName(phaseID)

        startNr = Me.hierarchy.getPMIndexOfID(phaseID)
        endNr = startNr

        Do While endNr + 1 <= Me.CountPhases And stillChild
            Dim cPhase As clsPhase = Me.getPhase(endNr + 1)

            If Not IsNothing(cPhase) Then
                Dim curFullBreadCrumb As String = Me.getBcElemName(cPhase.nameID)
                If curFullBreadCrumb.StartsWith(fullSwlBreadCrumb) Then
                    ' is still Child
                    endNr = endNr + 1
                Else
                    stillChild = False
                End If
            Else
                stillChild = False
            End If
        Loop


    End Sub

    ''' <summary>
    ''' gibt eine Collection zurück, die die IDs der Elemente enthält, die in IDCollection enthalten sind 
    ''' und ausserdem Kinder bzw Kindes-Kinder des Elements mit ID=phaseID  sind 
    ''' </summary>
    ''' <param name="phaseID"></param>
    ''' <param name="IDCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property schnittmengeChilds(ByVal phaseID As String, ByVal IDCollection As Collection) As Collection
        Get
            Dim fullSwlBreadCrumb As String = Me.getBcElemName(phaseID)
            Dim childCollection As New Collection

            For Each item As Object In IDCollection
                If CStr(item) <> phaseID Then
                    ' sich selber ausschließen ...
                    Dim curFullBreadCrumb As String = Me.getBcElemName(CStr(item))

                    If curFullBreadCrumb.StartsWith(fullSwlBreadCrumb) Then
                        ' ist Kind Element, daher aufnehmen 
                        childCollection.Add(CStr(item), CStr(item))
                    End If
                End If
            Next

            schnittmengeChilds = childCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt zu einer als als voller Name (Breadcrumb + Elemename) übergebenen Phase zurück, ob die so im Projekt existiert 
    ''' wenn strict = false: true , wenn der ElemName vorkommt, unabhängig wo in der Hierarchie
    ''' wenn strict = true: true, wenn der ElemName genau in der angegebenen Hierarchie-Stufe vorkommt  
    '''  
    ''' </summary>
    ''' <param name="fullName">der volle Name, das heisst Breadcrum plus Name</param>
    ''' <param name="strict">gibt an, ob der volle Breadcrumb berücksichtigt werden soll oder nur der Name</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsPhase(ByVal fullName As String, ByVal strict As Boolean) As Boolean
        Get
            Dim elemName As String = ""
            Dim breadcrumb As String = ""
            Dim type As Integer = -1
            Dim pvName As String = ""
            Call splitHryFullnameTo2(fullName, elemName, breadcrumb, type, pvName)
            If type = -1 Or
                (type = PTItemType.projekt And Me.name = pvName) Or
                (type = PTItemType.vorlage And Me.VorlagenName = pvName) Then

                If strict Then
                    ' breadcrumb soll unverändert beachtet werden 
                Else
                    breadcrumb = ""
                End If

                Dim cphase As clsPhase = Me.getPhase(elemName, breadcrumb, 1)
                If IsNothing(cphase) Then
                    containsPhase = False
                Else
                    containsPhase = True
                End If
            Else
                containsPhase = False
            End If


        End Get
    End Property

    ''' <summary>
    ''' bekommt eine RoleName ID der Form roleID;TeamID übergeben
    ''' gibt true zurück, wenn diese RoleNameID im Projekt vorkommt 
    ''' </summary>
    ''' <param name="roleNameID"></param>
    ''' <returns></returns>
    Public ReadOnly Property containsRoleNameID(ByVal roleNameID As String) As Boolean
        Get
            containsRoleNameID = getRoleNameIDs.Contains(roleNameID)
        End Get
    End Property

    ''' <summary>
    ''' gibt zu einem als als voller Name (Breadcrumb + Elemename) übergebenen Meilenstein zurück, ob der so im Projekt existiert 
    ''' wenn strict = false: true , wenn der ElemName vorkommt, unabhängig wo in der Hierarchie
    ''' wenn strict = true: true, wenn der ElemName genau in der angegebenen Hierarchie-Stufe vorkommt  
    ''' </summary>
    ''' <param name="fullName"></param>
    ''' <param name="strict"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsMilestone(ByVal fullName As String, ByVal strict As Boolean) As Boolean
        Get
            Dim elemName As String = ""
            Dim breadcrumb As String = ""
            Dim type As Integer = -1
            Dim pvName As String = ""
            Call splitHryFullnameTo2(fullName, elemName, breadcrumb, type, pvName)

            If type = -1 Or
                (type = PTItemType.projekt And Me.name = pvName) Or
                (type = PTItemType.vorlage And Me.VorlagenName = pvName) Then

                If strict Then
                    ' breadcrumb soll unverändert beachtet werden 
                Else
                    breadcrumb = ""
                End If

                Dim cMilestone As clsMeilenstein = Me.getMilestone(elemName, breadcrumb, 1)
                If IsNothing(cMilestone) Then
                    containsMilestone = False
                Else
                    containsMilestone = True
                End If
            Else
                containsMilestone = False
            End If


        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn in der Vorlage irgendeiner der Meilensteine, entweder über BreadCrumb oder nur als Name angegeben, vorhanden ist
    ''' </summary>
    ''' <param name="msCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property containsAnyMilestonesOfCollection(ByVal msCollection As Collection) As Boolean
        Get
            Dim ix As Integer
            Dim fullName As String
            Dim tmpResult As Boolean = False
            Dim containsMS As Boolean = False
            Dim tmpMilestone As clsMeilenstein

            If msCollection.Count = 0 Then
                tmpResult = True
            Else
                While ix <= msCollection.Count And Not containsMS

                    fullName = CStr(msCollection.Item(ix))
                    Dim curMsName As String = ""
                    Dim breadcrumb As String = ""
                    Dim pvName As String = ""
                    Dim type As Integer = -1

                    ' hier wird der Eintrag in filterMilestone aufgesplittet in curMsName und breadcrumb) 
                    Call splitHryFullnameTo2(fullName, curMsName, breadcrumb, type, pvName)

                    If type = -1 Or
                        (type = PTItemType.projekt And pvName = Me.name) Or
                        (type = PTItemType.vorlage And pvName = Me.VorlagenName) Then

                        Dim milestoneIndices(,) As Integer = Me.hierarchy.getMilestoneIndices(curMsName, breadcrumb)
                        ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste

                        For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1

                            tmpMilestone = Me.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx))
                            If IsNothing(tmpMilestone) Then

                            Else
                                containsMS = True
                                Exit For
                            End If

                        Next
                    ElseIf type = PTItemType.categoryList Then
                        containsMS = Me.containsMilestoneCategory(pvName)
                    End If

                    ix = ix + 1

                End While
                tmpResult = containsMS
            End If

            containsAnyMilestonesOfCollection = tmpResult

        End Get
    End Property

    ''' <summary>
    ''' gibt true zurück, wenn in der Vorlage irgendeiner der Meilensteine, entweder über BreadCrumb oder nur als Name angegeben, vorhanden ist
    ''' </summary>
    ''' <param name="phCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overrides ReadOnly Property containsAnyPhasesOfCollection(ByVal phCollection As Collection) As Boolean
        Get
            Dim ix As Integer
            Dim fullName As String
            Dim tmpResult As Boolean = False
            Dim containsPH As Boolean = False
            Dim tmpPhase As clsPhase

            If phCollection.Count = 0 Then
                tmpResult = True
            Else
                While ix <= phCollection.Count And Not containsPH

                    fullName = CStr(phCollection.Item(ix))
                    Dim curPhName As String = ""
                    Dim breadcrumb As String = ""
                    Dim pvName As String = ""
                    Dim type As Integer = -1

                    ' hier wird der Eintrag in filterMilestone aufgesplittet in curMsName und breadcrumb) 
                    Call splitHryFullnameTo2(fullName, curPhName, breadcrumb, type, pvName)

                    If type = -1 Or
                        (type = PTItemType.projekt And pvName = Me.name) Or
                        (type = PTItemType.vorlage And pvName = Me.VorlagenName) Then

                        Dim phaseIndices() As Integer = Me.hierarchy.getPhaseIndices(curPhName, breadcrumb)
                        ' in milestoneIndices sind jetzt die Phasen- und Meilenstein Index der Phasen bzw Meilenstein Liste

                        For mx As Integer = 0 To CInt(phaseIndices.Length) - 1

                            tmpPhase = Me.getPhase(phaseIndices(mx))
                            If IsNothing(tmpPhase) Then

                            Else
                                containsPH = True
                                Exit For
                            End If

                        Next

                    ElseIf type = PTItemType.categoryList Then
                        containsPH = Me.containsPhaseCategory(pvName)
                    End If

                    ix = ix + 1

                End While
                tmpResult = containsPH
            End If

            containsAnyPhasesOfCollection = tmpResult

        End Get
    End Property

    Public ReadOnly Property getPhaseIdsWithRoleCost(ByVal namenListe As Collection, ByVal namesAreRoleIDs As Boolean) As Collection
        Get
            Dim iDCollection As New Collection
            Dim teamID As Integer = -1

            For i As Integer = 1 To Me.CountPhases

                Dim cphase As clsPhase = Me.getPhase(i)
                Dim phaseNameID As String = cphase.nameID

                If namesAreRoleIDs Then
                    For Each tmpItem As String In namenListe
                        Dim roleID As Integer = RoleDefinitions.getRoleDefByIDKennung(tmpItem, teamID).UID
                        If Me.rcLists.phaseContainsRoleID(phaseNameID, roleID, teamID) Then
                            If Not iDCollection.Contains(phaseNameID) Then
                                iDCollection.Add(phaseNameID, phaseNameID)
                            End If
                        End If
                    Next
                Else
                    ' Kosten 
                    For Each tmpItem As String In namenListe
                        Dim costID As Integer = CostDefinitions.getCostdef(tmpItem).UID
                        If Me.rcLists.phaseContainsCost(phaseNameID, costID) Then
                            If Not iDCollection.Contains(phaseNameID) Then
                                iDCollection.Add(phaseNameID, phaseNameID)
                            End If
                        End If
                    Next
                End If


            Next

            getPhaseIdsWithRoleCost = iDCollection
        End Get
    End Property

    ''' <summary>
    ''' gibt zu der angegebenen elemID alle Kind und Kindes-KinderIDs zurück
    ''' das umfasst IDs von Phasen und IDs von Meilensteinen 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <returns></returns>
    Public ReadOnly Property getAllChildIDsOf(ByVal elemID As String) As Collection
        Get
            Dim tmpCollection As New Collection

            ' hole alle Phasen Ids von elemID
            For Each childPhaseNameID As String In Me.hierarchy.getChildIDsOf(elemID, False)
                If Not tmpCollection.Contains(childPhaseNameID) Then
                    tmpCollection.Add(childPhaseNameID, childPhaseNameID)
                End If

                Dim childCollection As Collection = getAllChildIDsOf(childPhaseNameID)
                If Not IsNothing(childCollection) Then
                    For Each tmpNameID As String In childCollection
                        If Not tmpCollection.Contains(tmpNameID) Then
                            tmpCollection.Add(tmpNameID, tmpNameID)
                        End If
                    Next
                End If

            Next

            ' hole alle Meilenstein IDs von elemID 
            For Each childMilestoneNameID As String In Me.hierarchy.getChildIDsOf(elemID, True)
                If Not tmpCollection.Contains(childMilestoneNameID) Then
                    tmpCollection.Add(childMilestoneNameID, childMilestoneNameID)
                End If
            Next

            getAllChildIDsOf = tmpCollection
        End Get
    End Property

    ''' <summary>
    ''' in der namenListe können Elem-Namen oder Elem-IDs sein; wenn ein Elem-NAme gefunden wird, 
    ''' so wird er ersetzt durch alle Elem-IDs, die diesen Namen tragen 
    ''' es wird sichergestellt, dass jede ID tatsächlich nur einmal aufgeführt ist 
    ''' </summary>
    ''' <param name="namenListe"></param>
    ''' <param name="namesAreMilestones"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getElemIdsOf(ByVal namenListe As Collection, ByVal namesAreMilestones As Boolean) As Collection
        Get
            Dim iDCollection As New Collection
            Dim tmpSortList As New SortedList(Of DateTime, String)
            Dim sortDate As DateTime
            Dim itemName As String = ""
            Dim itemBreadcrumb As String = ""
            Dim iDItem As String
            Dim phaseIndices() As Integer
            Dim milestoneIndices(,) As Integer

            For i As Integer = 1 To namenListe.Count

                itemName = CStr(namenListe.Item(i))

                If istElemID(itemName) Then

                    Dim ok As Boolean = True
                    If namesAreMilestones Then
                        Dim cMilestone As clsMeilenstein = Me.getMilestoneByID(itemName)
                        If Not IsNothing(cMilestone) Then
                            sortDate = cMilestone.getDate
                        Else
                            ok = False
                        End If

                    Else
                        Dim cphase As clsPhase = Me.getPhaseByID(itemName)
                        If Not IsNothing(cphase) Then
                            sortDate = cphase.getStartDate
                        Else
                            ok = False
                        End If

                    End If

                    If ok And Not tmpSortList.ContainsValue(itemName) Then

                        Do While tmpSortList.ContainsKey(sortDate)
                            sortDate = sortDate.AddMilliseconds(1)
                        Loop

                        tmpSortList.Add(sortDate, itemName)

                    End If


                Else
                    Dim type As Integer = -1
                    Dim pvName As String = ""
                    Call splitHryFullnameTo2(CStr(namenListe.Item(i)), itemName, itemBreadcrumb, type, pvName)

                    If type = -1 Or
                        (type = PTItemType.projekt And pvName = Me.name) Or
                        (type = PTItemType.vorlage And pvName = Me.VorlagenName) Then

                        If namesAreMilestones Then
                            milestoneIndices = Me.hierarchy.getMilestoneIndices(itemName, itemBreadcrumb)

                            For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1
                                ' wenn der Wert Null ist , so existiert der Wert nicht 
                                If milestoneIndices(0, mx) > 0 And milestoneIndices(1, mx) > 0 Then

                                    Try
                                        iDItem = Me.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx)).nameID
                                        sortDate = Me.getMilestoneByID(iDItem).getDate

                                        If Not tmpSortList.ContainsValue(iDItem) Then

                                            Do While tmpSortList.ContainsKey(sortDate)
                                                sortDate = sortDate.AddMilliseconds(1)
                                            Loop


                                            tmpSortList.Add(sortDate, iDItem)


                                        End If


                                    Catch ex As Exception

                                    End Try

                                End If

                            Next
                        Else
                            phaseIndices = Me.hierarchy.getPhaseIndices(itemName, itemBreadcrumb)
                            For px As Integer = 0 To phaseIndices.Length - 1

                                If phaseIndices(px) > 0 And phaseIndices(px) <= Me.CountPhases Then
                                    iDItem = Me.getPhase(phaseIndices(px)).nameID

                                    sortDate = Me.getPhaseByID(iDItem).getStartDate

                                    If Not tmpSortList.ContainsValue(iDItem) Then

                                        Do While tmpSortList.ContainsKey(sortDate)
                                            sortDate = sortDate.AddMilliseconds(1)
                                        Loop

                                        tmpSortList.Add(sortDate, iDItem)

                                    End If

                                    'If Not iDCollection.Contains(iDItem) Then
                                    '    iDCollection.Add(iDItem, iDItem)
                                    'End If
                                End If

                            Next
                        End If

                    ElseIf type = PTItemType.categoryList Then

                        If namesAreMilestones Then
                            ' im pvName steht jetzt der Meilenstein Category Name  ... 
                            Dim tmpCollection As Collection = Me.getMilestoneIDsWithCat(pvName)

                            For Each tmpID As String In tmpCollection

                                If Not tmpSortList.ContainsValue(tmpID) Then
                                    sortDate = Me.getMilestoneByID(tmpID).getDate

                                    Do While tmpSortList.ContainsKey(sortDate)
                                        sortDate = sortDate.AddMilliseconds(1)
                                    Loop


                                    tmpSortList.Add(sortDate, tmpID)

                                End If

                            Next
                        Else
                            ' im pvName steht jetzt der Phasen Category Name ... 
                            Dim tmpCollection As Collection = Me.getPhaseIDsWithCat(pvName)

                            For Each tmpID As String In tmpCollection

                                If Not tmpSortList.ContainsValue(tmpID) Then
                                    sortDate = Me.getPhaseByID(tmpID).getStartDate

                                    Do While tmpSortList.ContainsKey(sortDate)
                                        sortDate = sortDate.AddMilliseconds(1)
                                    Loop


                                    tmpSortList.Add(sortDate, tmpID)

                                End If

                            Next
                        End If


                    End If



                End If

            Next


            ' jetzt muss umkopiert werden 
            For Each kvp As KeyValuePair(Of DateTime, String) In tmpSortList
                iDkey = kvp.Value
                iDCollection.Add(kvp.Value, kvp.Value)
            Next

            getElemIdsOf = iDCollection

        End Get
    End Property


    ''' <summary>
    ''' findet für das aktuelle Projekt heraus, wieviele zusätzliche Zeilen für die selektierten Meilensteine
    '''  (gezeichnet zur nächst höheren aber auch selektierten Phase) beim Report benötigt werden
    ''' außerdem werden in drawMStoPhaseListe die selektierten Meilensteine zu der passenden selektierten Phase gemerkt
    ''' </summary>
    ''' <param name="selectedPhases"></param>
    ''' <param name="selectedMilestones"></param>
    ''' <param name="considerTimespace"></param>
    ''' <param name="anzLines"></param>
    ''' <param name="drawMStoPhaseListe"></param>
    ''' <remarks></remarks>
    Public Sub selMilestonesToselPhase(ByVal selectedPhases As Collection, ByVal selectedMilestones As Collection, ByVal considerTimespace As Boolean,
                                       ByRef anzLines As Integer, ByRef drawMStoPhaseListe As SortedList(Of String, SortedList))


        If selectedMilestones.Count > 0 Then

            Dim drawMSinPhase As New SortedList(Of String, SortedList)
            ' Phasen die zusätzliche MS einzuzeichnen haben
            Dim listMS As New SortedList
            Dim found As Boolean = False
            Dim parentID As String = ""
            Dim selMSName As String = ""
            Dim selPHName As String = ""
            Dim msnameID As String = ""
            Dim mx As Integer, j As Integer
            Dim breadcrumb As String = ""

            For mx = 1 To selectedMilestones.Count  ' Schleife über alle selektierten Meilensteine
                found = False

                ' Herausfinden der UniqueID der selektierten Meilensteine
                Dim mstype As Integer = -1
                Dim pvname As String = ""
                Call splitHryFullnameTo2(CStr(selectedMilestones(mx)), selMSName, breadcrumb, mstype, pvname)

                If mstype = -1 Or
                    (mstype = PTItemType.projekt And pvname = Me.name) Or
                    (mstype = PTItemType.vorlage And pvname = Me.VorlagenName) Then

                    Dim msNameIndices() As Integer
                    msNameIndices = Me.hierarchy.getMilestoneHryIndices(selMSName, breadcrumb)

                    If msNameIndices(0) = 0 Then
                        ' Änderung tk: in diesem Fall gibt es den Meilenstein gar nicht 
                        ' einfach in der Schleife weitermachen ...
                    Else
                        For j = 0 To msNameIndices.Length - 1

                            msnameID = Me.hierarchy.getIDAtIndex(msNameIndices(j))

                            parentID = Me.hierarchy.getParentIDOfID(msnameID)
                            'While Not (x = rootPhaseName Or found)
                            Dim zaehler As Integer = 0

                            ' -------------------------------------------
                            ' Änderung tk 9.4.16 wenn found hier nicht auf False gesetzt wird, dann kann der nächste msNAmeID nicht mehr aufgenommen werden ... 
                            ' das found = false hat vorher gefehlt ... 
                            found = False
                            ' Ende Änderung tk 9.4.16 -------------------

                            While Not found
                                zaehler = zaehler + 1
                                ' nachsehen, ob diese Phase in den selektierten Phasen enthalten ist
                                Dim phind As Integer = 1
                                While Not found And phind <= selectedPhases.Count

                                    Dim phtype As Integer = -1
                                    pvname = ""
                                    Call splitHryFullnameTo2(CStr(selectedPhases(phind)), selPHName, breadcrumb, phtype, pvname)

                                    If phtype = -1 Or
                                        (phtype = PTItemType.projekt And pvname = Me.name) Or
                                        (phtype = PTItemType.vorlage And pvname = Me.VorlagenName) Then

                                        Dim phNameIndices() As Integer
                                        phNameIndices = Me.hierarchy.getPhaseHryIndices(selPHName, breadcrumb)
                                        If phNameIndices.Contains(Me.hierarchy.getIndexOfID(parentID)) Then
                                            found = True
                                        End If

                                    ElseIf phtype = PTItemType.categoryList Then
                                        ' 3.12.17 Behandlung muss noch überprüft / gemacht werden 
                                        Try
                                            found = selectedPhases.Contains(calcHryCategoryName(pvname, False))
                                        Catch ex As Exception
                                            found = False
                                        End Try
                                    End If

                                    phind = phind + 1

                                End While
                                If Not found Then
                                    parentID = Me.hierarchy.getParentIDOfID(parentID) 'Parent eine Stufe höher finden
                                    If parentID = Nothing Or parentID = "" Then
                                        parentID = rootPhaseName
                                        found = True
                                    End If
                                End If

                            End While

                            If zaehler > 1 Or parentID = rootPhaseName Then ' Parent des Meilenstein soll nicht angezeigt werden, ist also nicht selektiert
                                ' oder letzte Stufe ist erreicht, nämlich Phase rootPhaseName

                                If drawMSinPhase.ContainsKey(parentID) Then
                                    listMS = drawMSinPhase(parentID)
                                Else
                                    listMS = New SortedList
                                    drawMSinPhase.Add(parentID, listMS)
                                End If

                                If Not listMS.Contains(msnameID) Then
                                    listMS.Add(msnameID, msnameID)

                                End If

                            End If

                        Next j
                    End If
                ElseIf mstype = PTItemType.categoryList Then

                    Dim idCollection As Collection = Me.getMilestoneIDsWithCat(pvname)
                    For Each tmpID As String In idCollection
                        msnameID = tmpID
                        parentID = Me.hierarchy.getParentIDOfID(msnameID)
                        'While Not (x = rootPhaseName Or found)
                        Dim zaehler As Integer = 0

                        ' -------------------------------------------
                        ' Änderung tk 9.4.16 wenn found hier nicht auf False gesetzt wird, dann kann der nächste msNAmeID nicht mehr aufgenommen werden ... 
                        ' das found = false hat vorher gefehlt ... 
                        found = False
                        ' Ende Änderung tk 9.4.16 -------------------

                        While Not found
                            zaehler = zaehler + 1
                            ' nachsehen, ob diese Phase in den selektierten Phasen enthalten ist
                            Dim phind As Integer = 1
                            While Not found And phind <= selectedPhases.Count

                                Dim phtype As Integer = -1
                                pvname = ""
                                Call splitHryFullnameTo2(CStr(selectedPhases(phind)), selPHName, breadcrumb, phtype, pvname)

                                If phtype = -1 Or
                                    (phtype = PTItemType.projekt And pvname = Me.name) Or
                                    (phtype = PTItemType.vorlage And pvname = Me.VorlagenName) Then

                                    Dim phNameIndices() As Integer
                                    phNameIndices = Me.hierarchy.getPhaseHryIndices(selPHName, breadcrumb)
                                    If phNameIndices.Contains(Me.hierarchy.getIndexOfID(parentID)) Then
                                        found = True
                                    End If
                                ElseIf phtype = PTItemType.categoryList Then
                                    ' 3.12.17 Behandlung muss noch überprüft / gemacht werden 
                                    Try
                                        found = selectedPhases.Contains(calcHryCategoryName(pvname, False))
                                    Catch ex As Exception
                                        found = False
                                    End Try
                                End If

                                phind = phind + 1

                            End While
                            If Not found Then
                                parentID = Me.hierarchy.getParentIDOfID(parentID) 'Parent eine Stufe höher finden
                                If parentID = Nothing Or parentID = "" Then
                                    parentID = rootPhaseName
                                    found = True
                                End If
                            End If

                        End While

                        If zaehler > 1 Or parentID = rootPhaseName Then ' Parent des Meilenstein soll nicht angezeigt werden, ist also nicht selektiert
                            ' oder letzte Stufe ist erreicht, nämlich Phase rootPhaseName

                            If drawMSinPhase.ContainsKey(parentID) Then
                                listMS = drawMSinPhase(parentID)
                            Else
                                listMS = New SortedList
                                drawMSinPhase.Add(parentID, listMS)
                            End If

                            If Not listMS.Contains(msnameID) Then
                                listMS.Add(msnameID, msnameID)

                            End If

                        End If

                    Next

                End If




            Next mx

            drawMStoPhaseListe = drawMSinPhase
            anzLines = drawMStoPhaseListe.Count
        End If

    End Sub
    ''' <summary>
    ''' gibt die Anzahl Zeilen zurück, die das aktuelle Projekt im "Extended Drawing Mode" benötigt, wenn alle zughörigen Phasen gezeichnet werden
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property calcNeededLines() As Integer
        Get

            Dim phasenName As String = ""
            Dim zeilenOffset As Integer = 1
            Dim lastEndDate As Date = StartofCalendar.AddDays(-1)
            Dim tmpValue As Integer
            Dim breadcrumb As String = ""

            Dim anzPhases As Integer = 0
            Dim cphase As clsPhase = Nothing

            For i = 1 To Me.CountPhases ' Schleife über alle Phasen eines Projektes
                Try
                    cphase = Me.getPhase(i)
                    If Not IsNothing(cphase) Then

                        'Call splitHryFullnameTo2(CStr(cphase.nameID), phasenName, breadcrumb)

                        With Me.getPhase(i)

                            'phasenName = .name
                            If DateDiff(DateInterval.Day, lastEndDate, .getStartDate) < 0 Then
                                zeilenOffset = zeilenOffset + 1
                                lastEndDate = StartofCalendar.AddDays(-1)
                            End If

                            If DateDiff(DateInterval.Day, lastEndDate, .getEndDate) > 0 Then
                                lastEndDate = .getEndDate
                            End If

                        End With

                        anzPhases = anzPhases + 1

                    End If


                Catch ex As Exception

                End Try


            Next i      ' nächste Phase im Projekt betrachten

            If anzPhases > 1 Then
                tmpValue = zeilenOffset
            Else
                tmpValue = 1
            End If


            calcNeededLines = tmpValue

        End Get

    End Property

    Public Sub New()

        AllPhases = New List(Of clsPhase)
        _extendedView = False
        '_relStart = 1
        _leadPerson = ""

        _StartOffset = 0
        _Start = 0
        _startDate = NullDatum
        _earliestStart = 0
        _latestStart = 0
        _Status = ProjektStatus(PTProjektStati.geplant)
        _shpUID = ""
        _timeStamp = Date.Now
        _projectType = ptPRPFType.project
        _actualDataUntil = Date.MinValue
        _kundenNummer = ""

        _variantName = ""   ' ur:25.6.2014: hinzugefügt, da sonst in der DB variantName mal "" und mal Nothing istshow 
        _variantDescription = ""

        '_ampelErlaeuterung = ""
        '_ampelStatus = 0

        _description = ""
        _businessUnit = ""
        _complexity = 0.0
        _volume = 0.0

        keyMetrics = New clsKeyMetrics

        _updatedAt = ""

    End Sub

    ''' <summary>
    ''' legt ein initiales Projekt mit rootphase con startDate bis EndeDate an, ggf als Union PRojekt ..
    ''' </summary>
    ''' <param name="unionizedP"></param>
    ''' <param name="startDatum"></param>
    ''' <param name="endeDatum"></param>
    Public Sub New(ByVal pName As String,
                   ByVal unionizedP As Boolean,
                   startDatum As Date, endeDatum As Date)

        _Start = CInt(DateDiff(DateInterval.Month, StartofCalendar, startDatum) + 1)
        _startDate = startDatum

        AllPhases = New List(Of clsPhase)
        Dim cphase As New clsPhase(parent:=Me)
        Dim projektdauer As Integer = calcDauerIndays(startDatum, endeDatum)

        With cphase
            .nameID = rootPhaseName
            .changeStartandDauerPhase1(0, projektdauer)
        End With

        Me.AddPhase(cphase)
        If unionizedP Then
            _projectType = ptPRPFType.portfolio
        Else
            _projectType = ptPRPFType.project
        End If

        _actualDataUntil = Date.MinValue
        _kundenNummer = ""

        _extendedView = False
        '_relStart = 1

        _StartOffset = 0

        _earliestStart = 0
        _latestStart = 0

        _earliestStartDate = _startDate.AddMonths(_earliestStart)
        _latestStartDate = _startDate.AddMonths(_latestStart)

        _Status = ProjektStatus(PTProjektStati.geplant)
        _shpUID = ""
        _timeStamp = Date.Now

        _name = pName
        _variantName = ""
        _variantDescription = ""

        If unionizedP Then
            _description = "Summary Project of Program / Portfolio"
        Else
            _description = ""
        End If

        _businessUnit = ""
        _complexity = 0.0
        _volume = 0.0

        keyMetrics = New clsKeyMetrics

        _updatedAt = ""

    End Sub

    Public Sub New(ByVal projektStart As Integer, ByVal earliestValue As Integer, ByVal latestValue As Integer)

        AllPhases = New List(Of clsPhase)
        _extendedView = False
        '_relStart = 1
        _leadPerson = ""

        _StartOffset = 0

        _Start = projektStart
        _earliestStart = earliestValue
        _latestStart = latestValue

        _startDate = StartofCalendar.AddMonths(projektStart)
        _earliestStartDate = _startDate.AddMonths(_earliestStart)
        _latestStartDate = _startDate.AddMonths(_latestStart)

        _Status = ProjektStatus(PTProjektStati.geplant)
        _shpUID = ""
        _timeStamp = Date.Now
        _projectType = ptPRPFType.project
        _actualDataUntil = Date.MinValue
        _kundenNummer = ""

        _variantName = ""
        _variantDescription = ""


        _description = ""
        _businessUnit = ""
        _complexity = 0.0
        _volume = 0.0

        keyMetrics = New clsKeyMetrics

        _updatedAt = ""

    End Sub

    Public Sub New(ByVal startDate As Date, ByVal earliestStartdate As Date, ByVal latestStartdate As Date)

        AllPhases = New List(Of clsPhase)
        extendedView = False
        '_relStart = 1
        _leadPerson = ""

        _StartOffset = 0

        _startDate = startDate
        _earliestStartDate = earliestStartdate
        _latestStartDate = latestStartdate

        _Start = CInt(DateDiff(DateInterval.Month, StartofCalendar, startDate) + 1)
        _earliestStart = CInt(DateDiff(DateInterval.Month, startDate, earliestStartdate))
        _latestStart = CInt(DateDiff(DateInterval.Month, startDate, latestStartdate))

        _Status = ProjektStatus(PTProjektStati.geplant)
        _timeStamp = Date.Now

        _projectType = ptPRPFType.project
        _actualDataUntil = Date.MinValue
        _kundenNummer = ""

        _variantName = ""
        _variantDescription = ""


        _description = ""
        _businessUnit = ""
        _complexity = 0.0
        _volume = 0.0

        keyMetrics = New clsKeyMetrics

        _updatedAt = ""

    End Sub



End Class
