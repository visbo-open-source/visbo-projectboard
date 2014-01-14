Public Class clsProjektvorlage

   

    Public AllPhases As List(Of clsPhase)
    Private relStart As Integer
    Private uuid As Long
    ' als Friend deklariert, damit sie aus der Klasse clsProjekt, die von clsProjektvorlage erbt , erreichbar ist
    Friend _Dauer As Integer
    Private _earliestStart As Integer
    Private _latestStart As Integer


    ''' <summary>
    ''' Bezugsdatum ist hier der StartofCalendar
    ''' während in der addphase der abgeleiteten ProjektKlasse das Projektstartdatum das maßgebliche Datum ist 
    ''' </summary>
    ''' <param name="phase"></param>
    ''' <remarks></remarks>
    Public Overridable Sub AddPhase(ByVal phase As clsPhase)

        Dim phaseEnde As Double
        Dim maxM As Integer

        With phase

            phaseEnde = .startOffsetinDays + .dauerInDays - 1

        End With

        If phaseEnde > 0 Then

            maxM = DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(phaseEnde)) + 1
            If maxM <> _Dauer And maxM > 0 Then
                _Dauer = maxM
                ' hier muss jetzt die Dauer der Allgemeinen Phase angepasst werden ... 
            End If
        End If


        AllPhases.Add(phase)


    End Sub

    Public Property farbe() As Object

    Public Property Schrift() As Integer

    Public Property Schriftfarbe() As Object

    Public Property VorlagenName() As String

    'Public RessourcenDefinitionsBereich As String

    'Public KostenDefinitionsBereich As String


    Public Overridable Sub CopyTo(ByRef newproject As clsProjekt)
        Dim p As Integer
        Dim newphase As clsPhase


        With newproject
            .farbe = farbe
            .Schrift = Schrift
            .Schriftfarbe = Schriftfarbe
            .name = ""
            .VorlagenName = VorlagenName
            .earliestStart = _earliestStart
            .latestStart = _latestStart




            For p = 0 To Me.CountPhases - 1
                newphase = New clsPhase(newproject)
                AllPhases.Item(p).CopyTo(newphase)
                .AddPhase(newphase)
            Next p


        End With

    End Sub

    Public ReadOnly Property Liste() As List(Of clsPhase)

        Get
            Liste = AllPhases
        End Get

    End Property

    Public ReadOnly Property dauerInDays As Integer

        Get
            Dim i As Integer
            Dim max As Double = 0

            ' Bestimmung der Dauer 

            For i = 1 To Me.CountPhases

                With Me.getPhase(i)

                    If max < .startOffsetinDays + .dauerInDays Then
                        max = .startOffsetinDays + .dauerInDays
                    End If

                    For m = 1 To .CountResults
                        If max < .startOffsetinDays + .getResult(m).offset Then
                            max = .startOffsetinDays + .getResult(m).offset
                        End If
                    Next

                End With

            Next i


            dauerInDays = CInt(max)

        End Get
    End Property


    Public Overridable ReadOnly Property Dauer() As Integer


        Get
            Dim i As Integer
            Dim max As Double = 0
            Dim maxM As Integer

            ' neue Bestimmung der Dauer 

            For i = 1 To Me.CountPhases

                With Me.getPhase(i)

                    If max < .startOffsetinDays + .dauerInDays Then
                        max = .startOffsetinDays + .dauerInDays
                    End If

                    For m = 1 To .CountResults
                        If max < .startOffsetinDays + .getResult(m).offset Then
                            max = .startOffsetinDays + .getResult(m).offset
                        End If
                    Next

                End With

            Next i

            maxM = DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(max - 1)) + 1


            If maxM <> _Dauer Then
                _Dauer = maxM
                ' hier muss jetzt die Dauer der Allgemeinen Phase angepasst werden ... 
            End If


            Dauer = _Dauer


        End Get

    End Property

    Public Property UID() As Long

        Get
            UID = uuid
        End Get

        Set(value As Long)
            uuid = value
        End Set

    End Property

    Public ReadOnly Property CountPhases() As Integer

        Get
            CountPhases = AllPhases.Count
        End Get

    End Property

    Public Property Phase(index As Integer) As clsPhase

        Get
            Phase = AllPhases.Item(index - 1)
        End Get

        Set(value As clsPhase)
            AllPhases.Item(index - 1) = value
        End Set

    End Property

    Public ReadOnly Property getPhase(index As Integer) As clsPhase

        Get
            getPhase = AllPhases.Item(index - 1)
        End Get

    End Property

    Public ReadOnly Property getPhase(name As String) As clsPhase

        Get
            Dim index As Integer
            Dim i As Integer
            Dim found As Boolean
            found = False
            i = 1
            While i <= AllPhases.Count And Not found
                If name = AllPhases.Item(i - 1).name Then
                    found = True
                    index = i
                Else
                    i = i + 1
                End If

            End While

            If found Then
                getPhase = AllPhases.Item(index - 1)
            Else
                getPhase = Nothing
            End If

        End Get

    End Property

    '
    ' übergibt in getPhasenBedarf die Werte der Phase <phaseid>
    '
    Public Overridable ReadOnly Property getPhasenBedarf(phaseName As String) As Double()

        Get
            Dim phaseValues() As Double
            Dim anzPhasen As Integer
            Dim i As Integer, p As Integer
            Dim phase As clsPhase


            If _Dauer > 0 Then

                ReDim phaseValues(_Dauer - 1)

                anzPhasen = AllPhases.Count
                If anzPhasen > 0 Then

                    For p = 0 To anzPhasen - 1
                        phase = AllPhases.Item(p)

                        If phase.name = phaseName Then
                            With phase
                                For i = .relStart To .relEnde
                                    phaseValues(i - 1) = phaseValues(i - 1) + 1
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

    '
    ' übergibt in getRessourcenBedarf die Werte der Rolle <roleid>
    '
    Public ReadOnly Property getRessourcenBedarf(roleID As Object) As Double()

        Get
            Dim roleValues() As Double
            Dim anzRollen As Integer
            Dim anzPhasen As Integer
            Dim found As Boolean
            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim role As clsRolle
            Dim lookforIndex As Boolean
            Dim phasenStart As Integer, phasenEnde As Integer
            Dim tempArray() As Double



            If _Dauer > 0 Then

                lookforIndex = IsNumeric(roleID)

                ReDim roleValues(_Dauer - 1)

                anzPhasen = AllPhases.Count

                For p = 0 To anzPhasen - 1
                    phase = AllPhases.Item(p)
                    With phase
                        ' Off1
                        anzRollen = .CountRoles
                        phasenStart = .relStart - 1
                        phasenEnde = .relEnde - 1
                        ReDim tempArray(phasenEnde - phasenStart)

                        For r = 1 To anzRollen
                            role = .getRole(r)
                            found = False

                            With role
                                If lookforIndex Then
                                    If .RollenTyp = roleID Then
                                        found = True
                                    End If
                                Else
                                    If .name = roleID Then
                                        found = True
                                    End If
                                End If

                                If found Then
                                    tempArray = .Xwerte
                                    For i = phasenStart To phasenEnde
                                        roleValues(i) = roleValues(i) + tempArray(i - phasenStart)
                                    Next i
                                End If
                            End With ' role

                        Next r

                    End With ' phase


                Next p ' Loop über alle Phasen

                getRessourcenBedarf = roleValues

            Else
                ReDim roleValues(0)
                getRessourcenBedarf = roleValues
            End If
        End Get

    End Property

    '
    ' übergibt in getUsedRollen eine Collection von Rollen Definitionen, das sind alle Rollen, die in den Phasen vorkommen und einen Bedarf von größer Null haben
    '
    Public ReadOnly Property getUsedRollen() As Collection

        Get
            Dim phase As clsPhase
            Dim aufbauRollen As New Collection
            Dim roleName As String
            Dim hrole As clsRolle
            Dim p As Integer, r As Integer

            'Dim ende As Integer


            If Me._Dauer > 0 Then

                For p = 0 To AllPhases.Count - 1
                    phase = AllPhases.Item(p)
                    With phase
                        For r = 1 To .CountRoles
                            hrole = .getRole(r)
                            If hrole.summe > 0 Then
                                roleName = hrole.name
                                '
                                ' Error-Handling ignorieren - in eine Collection können Werte nur einmal aufgenommen werden
                                '
                                Try
                                    aufbauRollen.Add(roleName, roleName)
                                Catch ex As Exception

                                End Try

                            End If
                        Next r
                    End With
                Next p

            End If


            getUsedRollen = aufbauRollen

        End Get

    End Property


    '
    ''' <summary>
    ''' gibt für Phase 1 ... n die Werte startoffset, dauer zurück 
    ''' Array hat die Dimension 2*anzPhasen -1 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseInfos() As Double()

        Get
            Dim anzPhasen As Integer
            Dim cphase As clsPhase
            Dim tmpvalues() As Double

            anzPhasen = AllPhases.Count
            ReDim tmpvalues(2 * anzPhasen - 1)

            For p = 0 To anzPhasen - 1

                cphase = AllPhases.Item(p)
                tmpvalues(p * 2) = cphase.startOffsetinDays
                tmpvalues(p * 2 + 1) = cphase.dauerInDays

            Next

            getPhaseInfos = tmpvalues

        End Get

    End Property
    '
    ' übergibt in getPersonalKosten die Personal Kosten der Rolle <roleid> über den Projektzeitraum
    '
    Public ReadOnly Property getPersonalKosten(roleID As Object) As Double()
        Get
            Dim costValues() As Double
            Dim anzRollen As Integer
            Dim anzPhasen As Integer
            Dim found As Boolean
            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim role As clsRolle
            Dim lookforIndex As Boolean
            Dim phasenStart As Integer, phasenEnde As Integer
            Dim tempArray() As Double
            Dim tagessatz As Double
            Dim faktor As Double = nrOfDaysMonth

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Then
                faktor = 1
            Else
                faktor = 1
            End If


            If _Dauer > 0 Then
                lookforIndex = IsNumeric(roleID)

                ReDim costValues(_Dauer - 1)

                anzPhasen = AllPhases.Count

                For p = 0 To anzPhasen - 1
                    phase = AllPhases.Item(p)
                    With phase
                        ' Off1
                        anzRollen = .CountRoles
                        phasenStart = .relStart - 1
                        phasenEnde = .relEnde - 1
                        ReDim tempArray(phasenEnde - phasenStart)

                        For r = 1 To anzRollen
                            role = .getRole(r)
                            found = False

                            With role
                                If lookforIndex Then
                                    If .RollenTyp = roleID Then
                                        found = True
                                    End If
                                Else
                                    If .name = roleID Then
                                        found = True
                                    End If
                                End If
                                If found Then
                                    tagessatz = .tagessatzIntern
                                    tempArray = .Xwerte
                                    For i = phasenStart To phasenEnde
                                        costValues(i) = costValues(i) + tempArray(i - phasenStart) * tagessatz * faktor / 1000
                                    Next i
                                End If
                            End With ' role

                        Next r

                    End With ' phase

                Next p ' Loop über alle Phasen

            Else
                ReDim costValues(0)
                costValues(0) = 0
            End If

            getPersonalKosten = costValues

        End Get
    End Property

    '
    ' übergibt in KostenBedarf die Werte der Kostenart <costId>
    '
    Public ReadOnly Property getKostenBedarf(CostID As Object) As Double()

        Get
            Dim costValues() As Double
            Dim anzKostenarten As Integer
            Dim anzPhasen As Integer
            Dim found As Boolean
            Dim i As Integer, p As Integer, k As Integer
            Dim phase As clsPhase
            Dim cost As clsKostenart
            Dim lookforIndex As Boolean, isPersCost As Boolean
            Dim phasenStart As Integer, phasenEnde As Integer
            Dim tempArray() As Double


            If _Dauer > 0 Then

                ReDim costValues(_Dauer - 1)

                lookforIndex = IsNumeric(CostID)
                isPersCost = False

                If lookforIndex Then
                    If CostID = CostDefinitions.Count Then
                        isPersCost = True
                    End If
                Else
                    If CostID = "Personalkosten" Then
                        isPersCost = True
                    End If
                End If

                If isPersCost Then
                    ' costvalues = AllPersonalKosten
                    costValues = Me.getAllPersonalKosten
                Else

                    anzPhasen = AllPhases.Count

                    For p = 0 To anzPhasen - 1
                        phase = AllPhases.Item(p)
                        With phase
                            ' Off1
                            anzKostenarten = .CountCosts
                            phasenStart = .relStart - 1
                            phasenEnde = .relEnde - 1
                            ReDim tempArray(phasenEnde - phasenStart)

                            For k = 1 To anzKostenarten
                                cost = .getCost(k)
                                found = False

                                With cost
                                    If lookforIndex Then
                                        If .KostenTyp = CostID Then
                                            found = True
                                        End If
                                    Else
                                        If .name = CostID Then
                                            found = True
                                        End If
                                    End If
                                    If found Then
                                        tempArray = .Xwerte
                                        For i = phasenStart To phasenEnde
                                            costValues(i) = costValues(i) + tempArray(i - phasenStart)
                                        Next i
                                    End If
                                End With ' cost

                            Next k

                        End With ' phase

                    Next p ' Loop über alle Phasen
                End If
            Else
                ReDim costValues(0)
                costValues(0) = 0
            End If

            getKostenBedarf = costValues


        End Get

    End Property

    '
    ' übergibt in getUsedKosten eine Collection von Kostenarten Definitionen, 
    ' das sind alle Kostenarten, die in den Phasen vorkommen und einen Bedarf von größer Null haben
    '
    Public ReadOnly Property getUsedKosten() As Collection

        Get
            Dim phase As clsPhase
            Dim aufbauKosten As New Collection
            Dim costname As String
            Dim hcost As clsKostenart
            Dim p As Integer, k As Integer

            'Dim ende As Integer

            If _Dauer > 0 Then
                For p = 0 To AllPhases.Count - 1
                    phase = AllPhases.Item(p)
                    With phase
                        For k = 1 To .CountCosts
                            hcost = .getCost(k)
                            If hcost.summe > 0 Then
                                costname = hcost.name
                                '
                                ' Error-Handling ignorieren - in eine Collection können Werte nur einmal aufgenommen werden
                                '
                                Try
                                    aufbauKosten.Add(costname, costname)
                                Catch ex As Exception

                                End Try
                                
                            End If
                        Next k
                    End With
                Next p

            End If


            getUsedKosten = aufbauKosten

        End Get

    End Property

    '
    ' übergibt in getsummekosten die Summe aller Kosten: Personalkosten plus alle anderen Kostenarten
    '
    Public ReadOnly Property getSummeKosten() As Double

        Get
            Dim costValues() As Double
            Dim ErgebnisListe As New Collection
            Dim costSum As Double
            Dim anzKostenarten As Integer
            Dim i As Integer, r As Integer
            Dim costname As String

            If _Dauer > 0 Then

                ReDim costValues(_Dauer - 1)
                costValues = Me.getAllPersonalKosten

                costSum = 0
                For i = 0 To _Dauer - 1
                    costSum = costSum + costValues(i)
                    costValues(i) = 0
                Next i
                '
                ' jetzt sind in der Summe die Personalkosten drin ....
                '

                ' Jetzt werden die einzelnen Kostenarten auf die gleiche Art und Weise geholt
                ErgebnisListe = Me.getUsedKosten

                anzKostenarten = ErgebnisListe.Count
                For r = 1 To anzKostenarten
                    costname = ErgebnisListe.Item(r)
                    costValues = Me.getKostenBedarf(costname)
                    For i = 0 To _Dauer - 1
                        costSum = costSum + costValues(i)
                        costValues(i) = 0
                    Next i
                Next r

                getSummeKosten = costSum

            Else
                getSummeKosten = 0
            End If

        End Get

    End Property

    '
    ' übergibt in getsummekosten die Summe aller Kosten: Personalkosten plus alle anderen Kostenarten
    '
    Public ReadOnly Property getGesamtKostenBedarf() As Double()

        Get
            Dim costValues() As Double, tmpValues() As Double
            Dim ErgebnisListe As New Collection
            Dim anzKostenarten As Integer
            Dim i As Integer, r As Integer
            Dim costname As String


            ReDim costValues(_Dauer - 1)
            ReDim tmpValues(_Dauer - 1)

            If _Dauer > 0 Then

                costValues = Me.getAllPersonalKosten
                '
                ' jetzt sind in costValues die Personalkosten drin ....
                '

                ' Jetzt werden die einzelnen Kostenarten auf die gleiche Art und Weise geholt
                ErgebnisListe = Me.getUsedKosten

                anzKostenarten = ErgebnisListe.Count
                For r = 1 To anzKostenarten
                    costname = ErgebnisListe.Item(r)
                    tmpValues = Me.getKostenBedarf(costname)
                    For i = 0 To _Dauer - 1
                        costValues(i) = costValues(i) + tmpValues(i)
                        tmpValues(i) = 0
                    Next i
                Next r

            End If

            getGesamtKostenBedarf = costValues

        End Get

    End Property

    '
    ' übergibt in getsummekosten die Summe aller Kosten: Personalkosten plus alle anderen Kostenarten
    '
    Public ReadOnly Property getGesamtAndereKosten() As Double()

        Get
            Dim costValues() As Double, tmpValues() As Double
            Dim ErgebnisListe As New Collection
            Dim anzKostenarten As Integer
            Dim i As Integer, r As Integer
            Dim costname As String


            ReDim costValues(_Dauer - 1)
            ReDim tmpValues(_Dauer - 1)

            If _Dauer > 0 Then

                ' Jetzt werden die einzelnen Kostenarten geholt
                ErgebnisListe = Me.getUsedKosten

                anzKostenarten = ErgebnisListe.Count
                For r = 1 To anzKostenarten
                    costname = ErgebnisListe.Item(r)
                    tmpValues = Me.getKostenBedarf(costname)
                    For i = 0 To _Dauer - 1
                        costValues(i) = costValues(i) + tmpValues(i)
                        tmpValues(i) = 0
                    Next i
                Next r

            End If

            getGesamtAndereKosten = costValues

        End Get

    End Property

    '
    ' übergibt in getSummeRessourcen den Ressourcen Bedarf in Mann-Monaten  die Werte der Kostenart <roleId>
    '
    Public ReadOnly Property getSummeRessourcen() As Double

        Get
            Dim roleValues() As Double
            Dim ErgebnisListe As New Collection
            Dim roleSum As Double
            Dim anzRollen As Integer
            Dim i As Integer, r As Integer
            Dim roleName As String


            If _Dauer > 0 Then

                ReDim roleValues(_Dauer - 1)

                roleSum = 0

                ' Jetzt werden die einzelnen Rollen aufsummiert
                ErgebnisListe = Me.getUsedRollen
                anzRollen = ErgebnisListe.Count

                For r = 1 To anzRollen
                    roleName = ErgebnisListe.Item(r)
                    roleValues = Me.getRessourcenBedarf(roleName)
                    For i = 0 To _Dauer - 1
                        roleSum = roleSum + roleValues(i)
                        roleValues(i) = 0
                    Next i
                Next r

                getSummeRessourcen = roleSum

            Else
                getSummeRessourcen = 0
            End If

        End Get

    End Property

    '
    ' übergibt in getSummeRessourcen den Ressourcen Bedarf in Mann-Monaten  die Werte der Kostenart <roleId>
    '
    Public ReadOnly Property getAlleRessourcen() As Double()

        Get
            Dim roleValues() As Double
            Dim alleValues() As Double
            Dim ErgebnisListe As New Collection
            Dim anzRollen As Integer
            Dim i As Integer, r As Integer
            Dim roleName As String

            
            If _Dauer > 0 Then

                ReDim roleValues(_Dauer - 1)
                ReDim alleValues(_Dauer - 1)


                ' Jetzt werden die einzelnen Rollen aufsummiert
                ErgebnisListe = Me.getUsedRollen
                anzRollen = ErgebnisListe.Count

                For r = 1 To anzRollen
                    roleName = ErgebnisListe.Item(r)
                    roleValues = Me.getRessourcenBedarf(roleName)
                    For i = 0 To _Dauer - 1
                        alleValues(i) = alleValues(i) + roleValues(i)
                        roleValues(i) = 0
                    Next i
                Next r

                getAlleRessourcen = alleValues

            Else
                ReDim alleValues(0)
                getAlleRessourcen = alleValues
            End If

        End Get

    End Property


    
    ''' <summary>
    ''' gibt die Personalkosten des betreffenden Projektes zurück ; zugrundgelegt wird der interne Tagessatz 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllPersonalKosten() As Double()

        Get
            Dim costValues() As Double
            Dim anzRollen As Integer
            Dim anzPhasen As Integer
            Dim i As Integer, p As Integer, r As Integer
            Dim phase As clsPhase
            Dim role As clsRolle
            Dim phasenStart As Integer, phasenEnde As Integer
            Dim tempArray() As Double
            Dim tagessatz As Double
            Dim faktor As Double = nrOfDaysMonth

            If awinSettings.kapaEinheit = "PM" Then
                faktor = nrOfDaysMonth
            ElseIf awinSettings.kapaEinheit = "PW" Then
                faktor = 5
            ElseIf awinSettings.kapaEinheit = "PT" Then
                faktor = 1
            Else
                faktor = 1
            End If


            If _Dauer > 0 Then

                ReDim costValues(_Dauer - 1)


                anzPhasen = AllPhases.Count

                For p = 0 To anzPhasen - 1
                    phase = AllPhases.Item(p)
                    With phase
                        ' Off1
                        anzRollen = .CountRoles
                        phasenStart = .relStart - 1
                        phasenEnde = .relEnde - 1
                        ReDim tempArray(phasenEnde - phasenStart) ' xxxxx prüfen !!!

                        For r = 1 To anzRollen
                            role = .getRole(r)

                            With role
                                tagessatz = .tagessatzIntern
                                tempArray = .Xwerte
                                For i = phasenStart To phasenEnde
                                    costValues(i) = costValues(i) + tempArray(i - phasenStart) * tagessatz * faktor / 1000
                                Next i

                            End With ' role

                        Next r

                    End With ' phase

                Next p ' Loop über alle Phasen



            Else

                ReDim costValues(0)
                costValues(0) = 0

            End If

            getAllPersonalKosten = costValues

        End Get

    End Property

    Public Overridable Property earliestStart() As Integer

        Get
            earliestStart = _earliestStart
        End Get
        Set(value As Integer)
            If value > 0 Then
                Throw New ArgumentException("Earliest Start kann nicht nach dem Startdatum liegen")
            Else
                _earliestStart = value
            End If

        End Set

    End Property


    Public Overridable Property latestStart() As Integer

        Get
            latestStart = _latestStart
        End Get
        Set(value As Integer)

            If value < 0 Then
                Throw New ArgumentException("latest Start kann nicht vor dem Startdatum liegen")
            Else
                _latestStart = value
            End If

        End Set

    End Property

    'Public Property Start() As Integer

    '    Get
    '        Start = _Start
    '    End Get

    '    Set(value As Integer)
    '        If _Status = ProjektStatus(1) Or _Status = ProjektStatus(2) Or _
    '                                         _Status = ProjektStatus(2) Then
    '            Call MsgBox("der Startzeitpunkt kann nicht mehr verändert werden ... ")

    '        ElseIf value < _Start + _earliestStart Then
    '            Call MsgBox("der neue Startzeitpunkt liegt vor dem bisher zugelassenen frühestmöglichen Startzeitpunkt ...")

    '        ElseIf value > _Start + _latestStart Then
    '            Call MsgBox("der neue Startzeitpunkt liegt nach dem bisher zugelassenen spätestmöglichen Startzeitpunkt ...")

    '        Else
    '            Dim absEarliest As Integer
    '            Dim absLatest As Integer
    '            Dim earliestDate As Date
    '            Dim newDate As Date = StartofCalendar.AddMonths(value - 1)
    '            Dim Heute As Date = Now

    '            If DateDiff(DateInterval.Month, newDate, Heute) < 0 Then
    '                Call MsgBox("der neue Startzeitpunkt liegt in der Vergangenheit ...")
    '            Else
    '                absEarliest = _Start + _earliestStart
    '                absLatest = _Start + _latestStart

    '                earliestDate = StartofCalendar.AddMonths(absEarliest - 1)
    '                Dim DifferenceInMonths As Long = DateDiff(DateInterval.Month, earliestDate, Heute)

    '                If DifferenceInMonths < 0 Then
    '                    ' frühestmöglicher Startzeitpunkt ist der aktuelle Monat
    '                    absEarliest = DateDiff(DateInterval.Month, Heute, StartofCalendar) + 1
    '                End If

    '                _Start = value
    '                _earliestStart = absEarliest - _Start
    '                _latestStart = absLatest - _Start
    '            End If


    '        End If
    '    End Set
    'End Property

    'Public Property Status() As String
    '    Get
    '        Status = _Status
    '    End Get
    '    Set(value As String)
    '        If value = ProjektStatus(0) Then
    '            _Status = value
    '        ElseIf value = ProjektStatus(1) Or value = ProjektStatus(2) Or _
    '                                           value = ProjektStatus(3) Then
    '            _Status = value
    '            _earliestStart = 0
    '            _latestStart = 0
    '        Else
    '            Call MsgBox("unzulässiger Wert für Status")
    '        End If
    '    End Set
    'End Property

    'Public Property StartOffset As Integer
    '    Get
    '        StartOffset = _StartOffset
    '    End Get

    '    Set(value As Integer)
    '        If value >= _earliestStart And value <= _latestStart Then
    '            _StartOffset = value
    '        Else
    '            Call MsgBox("unzulässiger Wert für StartOffset ...")
    '        End If
    '    End Set

    'End Property



    Public Sub New()

        AllPhases = New List(Of clsPhase)
        relStart = 1
        _Dauer = 0
        '_StartOffset = 0
        '_Start = 1
        _earliestStart = 0
        _latestStart = 0
        '_Status = ProjektStatus(0)

    End Sub

    'Public Sub New(ByVal projektStart As Integer, ByVal earliestValue As Integer, ByVal latestValue As Integer)

    '    AllPhases = New List(Of clsPhase)
    '    relStart = 1
    '    iDauer = 0
    '    _StartOffset = 0
    '    _Start = projektStart
    '    _earliestStart = earliestValue
    '    _latestStart = latestValue
    '    _Status = ProjektStatus(0)

    'End Sub

End Class
