Public Class clsProjektvorlage

   

    Public AllPhases As List(Of clsPhase)

    ' Änderung tk 31.3.15 Hierachie Klasse ergänzt 
    Public hierarchy As clsHierarchy

    Private relStart As Integer
    Private uuid As Long
    ' als Friend deklariert, damit sie aus der Klasse clsProjekt, die von clsProjektvorlage erbt , erreichbar ist
    Friend _Dauer As Integer
    Private _earliestStart As Integer
    Private _latestStart As Integer
    Private _budgetWerte() As Double


    ''' <summary>
    ''' gibt die Budgetwerte des Projekts zurück
    ''' die werden 
    ''' beim Laden aus der Datenbank bestimmt oder 
    ''' beim Ändern des Erlös Werts 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property budgetWerte As Double()
        Get
            budgetWerte = _budgetWerte
        End Get
        Set(value As Double())
            'ReDim _budgetWerte(value.Length - 1)
            If value.Sum > 0 Then
                _budgetWerte = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' Bezugsdatum ist hier der StartofCalendar
    ''' während in der addphase der abgeleiteten ProjektKlasse das Projektstartdatum das maßgebliche Datum ist 
    ''' </summary>
    ''' <param name="phase"></param>
    ''' <remarks></remarks>
    Public Overridable Sub AddPhase(ByVal phase As clsPhase, _
                                    Optional ByVal origName As String = "", _
                                    Optional ByVal parentID As String = "")

        Dim phaseEnde As Double
        Dim maxM As Integer

        With phase

            phaseEnde = .startOffsetinDays + .dauerInDays - 1

        End With

        If phaseEnde > 0 Then

            maxM = CInt(DateDiff(DateInterval.Month, StartofCalendar, StartofCalendar.AddDays(phaseEnde)) + 1)
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

            If origName = "" Then
                .origName = .elemName
            Else
                .origName = origName
            End If

            .indexOfElem = Me.CountPhases

            If parentID = "" Then
                If .indexOfElem = 1 Then
                    .parentNodeKey = ""
                Else
                    .parentNodeKey = calcHryElemKey(".", False)
                End If
            Else
                .parentNodeKey = parentID
            End If

        End With

        With Me.hierarchy
            .addNode(currentElementNode, phase.nameID)
        End With


    End Sub

    ''' <summary>
    ''' gibt den Meilenstein mit Element-ID elemID zurück 
    ''' Nothing, wenn sie nicht existiert 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestoneByID(elemID As String) As clsMeilenstein
        Get
            Dim currentNode As clsHierarchyNode = Me.hierarchy.nodeItem(elemID)
            Dim phaseID As String
            Dim phIndex As Integer, msIndex As Integer

            If Not IsNothing(currentNode) Then

                If elemIDIstMeilenstein(elemID) Then
                    phaseID = currentNode.parentNodeKey
                    phIndex = Me.hierarchy.nodeItem(phaseID).indexOfElem
                    msIndex = currentNode.indexOfElem

                    Dim cphase As clsPhase = Me.getPhase(phIndex)
                    If Not IsNothing(cphase) Then
                        getMilestoneByID = cphase.getMilestone(msIndex)
                    Else
                        getMilestoneByID = Nothing
                    End If
                Else
                    getMilestoneByID = Nothing
                End If

            Else
                getMilestoneByID = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt die Parent-Phase zu der angegebenen Elem-ID zurück; 
    ''' wenn es keine Parent-Phase gibt oder 
    ''' wenn es das Element gar nicht gibt, wird Nothing zurückgegeben 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getParentPhaseByID(ByVal elemID As String) As clsPhase
        Get

            Dim currentNode As clsHierarchyNode
            Dim phaseID As String
            Dim phIndex As Integer


            currentNode = Me.hierarchy.nodeItem(elemID)

            If Not IsNothing(currentNode) Then

                phaseID = currentNode.parentNodeKey
                phIndex = Me.hierarchy.nodeItem(phaseID).indexOfElem
                getParentPhaseByID = Me.getPhase(phIndex)

            Else
                getParentPhaseByID = Nothing
            End If

        End Get
    End Property

    ''' <summary>
    ''' gibt zu einem gegebenen Meilenstein-Namen das clsResult Objekt zurück, sofern es existiert
    ''' Nothing sonst
    ''' </summary>
    ''' <param name="msName">Name des Meilensteins</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestone(ByVal msName As String, Optional ByVal breadcrumb As String = "", Optional ByVal lfdNr As Integer = 1) As clsMeilenstein
        Get
            Dim tmpMilestone As clsMeilenstein = Nothing
            Dim found As Boolean = False

            Dim milestoneIndices(,) As Integer
            milestoneIndices = Me.hierarchy.getMilestoneIndices(msName, breadcrumb)

            If lfdNr > CInt(milestoneIndices.Length / 2) Or lfdNr < 1 Then
                ' kein gültiger Meilenstein 
            Else
                If milestoneIndices(0, lfdNr - 1) > 0 And milestoneIndices(1, lfdNr - 1) > 0 Then
                    ' nur dann existiert dieser Meilenstein
                    tmpMilestone = Me.getMilestone(milestoneIndices(0, lfdNr - 1), milestoneIndices(1, lfdNr - 1))
                End If

            End If

            getMilestone = tmpMilestone


        End Get
    End Property

    ''' <summary>
    ''' gibt den Meilenstein zurück, der in der Phase mit Index PhaseIndex, 
    ''' und dort in der Meilenstein Liste mit Index milestoneIndex vorkommt  
    ''' </summary>
    ''' <param name="phaseIndex"></param>
    ''' <param name="milestoneIndex"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestone(ByVal phaseIndex As Integer, ByVal milestoneIndex As Integer) As clsMeilenstein
        Get

            Dim tmpResult As clsMeilenstein = Nothing
            Dim found As Boolean = False

            If phaseIndex >= 1 And phaseIndex <= AllPhases.Count Then
                Dim cphase As clsPhase = AllPhases.Item(phaseIndex - 1)
                If milestoneIndex >= 1 And milestoneIndex <= cphase.countMilestones Then
                    tmpResult = cphase.getMilestone(milestoneIndex)
                End If
            End If


            getMilestone = tmpResult


        End Get
    End Property


    Public Property farbe() As Object

    Public Property Schrift() As Integer

    Public Property Schriftfarbe() As Object

    Public Property VorlagenName() As String

    'Public RessourcenDefinitionsBereich As String

    'Public KostenDefinitionsBereich As String

    ''' <summary>
    ''' kopiert die Attribute einer Projektvorlage in newproject;  bei der Quelle handelt es sich um eine 
    ''' Vorlage  
    ''' </summary>
    ''' <param name="newproject"></param>
    ''' <remarks></remarks>
    ''' 
    Public Overridable Sub copyAttrTo(ByRef newproject As clsProjekt)

        With newproject
            .farbe = Me.farbe
            .Schrift = Me.Schrift
            .Schriftfarbe = Me.Schriftfarbe
            .VorlagenName = Me.VorlagenName
            .earliestStart = _earliestStart
            .latestStart = _latestStart
            .name = ""
        End With

        ' jetzt wird die Hierarchie kopiert 
        Call copyHryTo(newproject)


    End Sub

    Public Overridable Sub copyTo(ByRef newproject As clsProjekt)
        Dim p As Integer
        Dim newphase As clsPhase
        Dim oldPhase As clsPhase
        'Dim parentID As String
        Dim origName As String = ""

        Call copyAttrTo(newproject)

        For p = 0 To Me.CountPhases - 1
            oldPhase = AllPhases.Item(p)
            newphase = New clsPhase(newproject)

            'parentID = Me.hierarchy.getParentIDOfID(oldPhase.nameID)

            oldPhase.CopyTo(newphase)
            newproject.AddPhase(newphase)
            'newproject.AddPhase(newphase, origName:="", parentID:=parentID)
        Next p


    End Sub


    Public Overridable Sub korrCopyTo(ByRef newproject As clsProjekt, ByVal startdate As Date, ByVal endedate As Date)
        Dim p As Integer
        Dim newphase As clsPhase
        Dim ProjectDauerInDays As Integer
        Dim CorrectFactor As Double

        Call copyAttrTo(newproject)

        newproject.startDate = startdate

        ProjectDauerInDays = calcDauerIndays(startdate, endedate)
        CorrectFactor = ProjectDauerInDays / Me.dauerInDays

        For p = 0 To Me.CountPhases - 1
            newphase = New clsPhase(newproject)
            AllPhases.Item(p).korrCopyTo(newphase, CorrectFactor)

            newproject.AddPhase(newphase)
        Next p


    End Sub

    ''' <summary>
    ''' kopiert die Hierarchie des aktuellen Me Projektes 
    ''' </summary>
    ''' <param name="newproject"></param>
    ''' <remarks></remarks>
    Friend Sub copyHryTo(ByRef newproject As clsProjekt)
        Dim ix As Integer
        Dim curNode As clsHierarchyNode
        Dim copiedNode As clsHierarchyNode
        Dim key As String
        Dim childKey As String

        newproject.hierarchy = New clsHierarchy

        For ix = 1 To Me.hierarchy.count
            curNode = Me.hierarchy.nodeItem(ix)
            key = Me.hierarchy.getIDAtIndex(ix)
            copiedNode = New clsHierarchyNode
            With copiedNode
                .elemName = curNode.elemName
                .indexOfElem = curNode.indexOfElem
                .origName = curNode.origName
                .parentNodeKey = curNode.parentNodeKey
                ' jetzt die Kinder kopieren 
                For cx As Integer = 1 To curNode.childCount
                    childKey = curNode.getChild(cx)
                    .addChild(childKey)
                Next
            End With

            newproject.hierarchy.copyNode(copiedNode, key)

        Next

    End Sub


    Public ReadOnly Property Liste() As List(Of clsPhase)

        Get
            Liste = AllPhases
        End Get

    End Property

    Public Overridable ReadOnly Property dauerInDays As Integer

        Get
            Dim i As Integer
            Dim max As Double = 0

            ' Bestimmung der Dauer 

            For i = 1 To Me.CountPhases

                With Me.getPhase(i)

                    If max < .startOffsetinDays + .dauerInDays Then
                        max = .startOffsetinDays + .dauerInDays
                    End If

                End With

            Next i


            dauerInDays = CInt(max)
            _Dauer = getColumnOfDate(StartofCalendar.AddDays(max - 1))

        End Get
    End Property


    Public ReadOnly Property anzahlRasterElemente() As Integer


        Get

            Dim tmpValue As Integer = 0

            If Me.CountPhases > 0 Then
                With Me.getPhase(1)
                    tmpValue = .relEnde - .relStart + 1
                End With
            End If

            anzahlRasterElemente = tmpValue


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

    ''' <summary>
    ''' gibt die Phase mit Index zurück, wenn Index kleiner bzw. gleich oder größer Anzahl Phasen, 
    ''' dann Nothing 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhase(ByVal index As Integer) As clsPhase

        Get
            If index < 1 Or index > AllPhases.Count Then
                getPhase = Nothing
            Else
                getPhase = AllPhases.Item(index - 1)
            End If
        End Get
    End Property

    Public ReadOnly Property getPhaseCount(ByVal name As String, Optional ByVal breadcrumb As String = "") As Integer
        Get

            Dim phaseIndices() As Integer = Me.hierarchy.getPhaseIndices(name, breadcrumb)
            If phaseIndices.Length = 1 And phaseIndices(0) = 0 Then
                getPhaseCount = 0
            Else
                getPhaseCount = phaseIndices.Length
            End If

        End Get
    End Property

    ''' <summary>
    ''' in der namenListe können Elem-Namen oder Elem-IDs sein; wenn ein Elem-NAme gefunden wird, 
    ''' so wird er ersetzt durch alle Elem-IDs, diesen Namen tragen 
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
            Dim itemName As String = ""
            Dim itemBreadcrumb As String = ""
            Dim iDItem As String
            Dim phaseIndices() As Integer
            Dim milestoneIndices(,) As Integer

            For i As Integer = 1 To namenListe.Count

                itemName = CStr(namenListe.Item(i))

                If istElemID(itemName) Then

                    If Not iDCollection.Contains(itemName) Then
                        iDCollection.Add(itemName, itemName)
                    End If

                Else
                    Call splitHryFullnameTo2(CStr(namenListe.Item(i)), itemName, itemBreadcrumb)

                    If namesAreMilestones Then
                        milestoneIndices = Me.hierarchy.getMilestoneIndices(itemName, itemBreadcrumb)

                        For mx As Integer = 0 To CInt(milestoneIndices.Length / 2) - 1
                            ' wenn der Wert Null ist , so existiert der Wert nicht 
                            If milestoneIndices(0, mx) > 0 And milestoneIndices(1, mx) > 0 Then

                                Try
                                    iDItem = Me.getMilestone(milestoneIndices(0, mx), milestoneIndices(1, mx)).nameID
                                    If Not iDCollection.Contains(iDItem) Then
                                        iDCollection.Add(iDItem, iDItem)
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
                                If Not iDCollection.Contains(iDItem) Then
                                    iDCollection.Add(iDItem, iDItem)
                                End If
                            End If

                        Next
                    End If
                End If

            Next

            getElemIdsOf = iDCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt alle Phasen bzw. Milestone ElemIDs in einer Collection zurück 
    ''' die Milestones gehen alle mit 1§ los, die Phasen alle mit 0§; 
    ''' deshalb markiert das erste Element mit "1§" das Ende der Phasen bzw. 
    ''' den Start der Meilensteine
    ''' </summary>
    ''' <param name="lookingForMS"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAllElemIDs(ByVal lookingForMS As Boolean) As Collection
        Get
            Dim iDCollection As New Collection
            Dim firstIX As Integer, lastIX As Integer
            Dim elemID As String

            If lookingForMS Then
                lastIX = Me.hierarchy.count
                firstIX = Me.hierarchy.getIndexOf1stMilestone
                If firstIX < 0 Then
                    ' es gibt keine Meilensteine 
                Else
                    For mx = firstIX To lastIX
                        elemID = Me.hierarchy.getIDAtIndex(mx)
                        If Not iDCollection.Contains(elemID) Then
                            iDCollection.Add(elemID, elemID)
                        End If
                    Next
                End If
            Else
                ' Phasen holen
                firstIX = 1
                lastIX = Me.hierarchy.getIndexOf1stMilestone - 1
                
                If lastIX < 0 Then
                    ' es gibt keine Meilensteine, sondern nur Phasen 
                    lastIX = Me.hierarchy.count
                End If

                For mx = firstIX To lastIX
                    elemID = Me.hierarchy.getIDAtIndex(mx)
                    If Not iDCollection.Contains(elemID) Then
                        iDCollection.Add(elemID, elemID)
                    End If
                Next

            End If

            getAllElemIDs = iDCollection

        End Get
    End Property

    ''' <summary>
    ''' gibt die Phase zurück, die die folgenden Eigenschaften erfüllt
    ''' hat name als elemName Bestandteil
    ''' hat den optional angegebenen Breadcrumb, wenn der nicht angegeben wird oder "" ist, dann ist es egal, unter welcher Hierarchie Stufe die Phase liegen soll 
    ''' der breadcrum kann die gesamte Hierarchie umfassen oder auch nur die erste Parent-Stufe; Parent-Stufen werden per # voneinander getrennt
    ''' hat die optional angegebene lfdNr, ist also das lfdNr-vielte Vorkommen von name / breadcrumb 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="breadcrumb"></param>
    ''' <param name="lfdNr"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhase(ByVal name As String, Optional ByVal breadcrumb As String = "", Optional ByVal lfdNr As Integer = 1) As clsPhase

        Get

            Dim found As Boolean = False


            Dim phaseIndices() As Integer
            phaseIndices = Me.hierarchy.getPhaseIndices(name, breadcrumb)

            If lfdNr > phaseIndices.Length Or lfdNr < 1 Then
                getPhase = Nothing
            Else
                If phaseIndices(lfdNr - 1) > 0 And phaseIndices(lfdNr - 1) <= AllPhases.Count Then
                    ' wenn phaseIndices(x) = 0 dann gibt es die Phase nicht ..
                    getPhase = AllPhases.Item(phaseIndices(lfdNr - 1) - 1)
                Else
                    getPhase = Nothing
                End If

            End If



            ' alter Code
            'found = False
            'i = 1
            'While i <= AllPhases.Count And Not found
            '    If name = AllPhases.Item(i - 1).name Then
            '        found = True
            '        index = i
            '    Else
            '        i = i + 1
            '    End If

            'End While

            'If found Then
            '    getPhase = AllPhases.Item(index - 1)
            'Else
            '    getPhase = Nothing
            'End If

        End Get

    End Property

    ''' <summary>
    ''' gibt die der ElemID entsprechende Phase zurück 
    ''' </summary>
    ''' <param name="elemID"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getPhaseByID(ByVal elemID As String) As clsPhase

        Get

            Dim phIndex As Integer = Me.hierarchy.getPMIndexOfID(elemID)
            If phIndex >= 0 Or phIndex < AllPhases.Count Then
                getPhaseByID = AllPhases.Item(phIndex - 1)
            Else
                getPhaseByID = Nothing
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
            Dim phasenStart As Integer
            Dim tempArray As Double()


            If _Dauer > 0 Then

                lookforIndex = IsNumeric(roleID)

                ReDim roleValues(_Dauer - 1)

                anzPhasen = AllPhases.Count

                For p = 0 To anzPhasen - 1
                    phase = AllPhases.Item(p)
                    With phase
                        ' Off1
                        anzRollen = .countRoles
                        phasenStart = .relStart - 1

                        ' Änderung: relende, relstart bezeichnet nicht mehr notwendigerweise die tatsächliche Länge des Arrays
                        ' es können Unschärfen auftreten 
                        'phasenEnde = .relEnde - 1


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

                                Dim dimension As Integer
                                If found Then
                                    dimension = .getDimension
                                    ReDim tempArray(dimension)
                                    tempArray = .Xwerte
                                    For i = phasenStart To phasenStart + dimension
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
                        For r = 1 To .countRoles
                            hrole = .getRole(r)
                            If hrole.summe > 0 Then
                                roleName = hrole.name
                                '
                                ' das ist performanter als der Weg über try .. catch 
                                '
                                If Not aufbauRollen.Contains(roleName) Then
                                    aufbauRollen.Add(roleName, roleName)
                                End If

                                'Try
                                '    aufbauRollen.Add(roleName, roleName)
                                'Catch ex As Exception

                                'End Try

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

    Public ReadOnly Property getMilestoneColors() As Double()
        Get
            Dim cphase As clsPhase
            Dim cresult As clsMeilenstein
            Dim tmpvalues() As Double
            Dim colorIndex As Integer
            Dim anzahlMilestones As Integer = 0

            For p = 1 To Me.CountPhases
                anzahlMilestones = anzahlMilestones + Me.getPhase(p).countMilestones
            Next

            If anzahlMilestones > 0 Then

                ReDim tmpvalues(anzahlMilestones - 1)

                Dim index As Integer = 0
                For p = 1 To Me.CountPhases
                    cphase = Me.getPhase(p)

                    For r = 1 To cphase.countMilestones
                        cresult = cphase.getMilestone(r)
                        colorIndex = cresult.getBewertung(1).colorIndex
                        tmpvalues(index) = colorIndex
                        index = index + 1
                    Next r

                Next p

            Else
                Throw New Exception("es gibt keine Meilensteine")
            End If

            getMilestoneColors = tmpvalues

        End Get
    End Property

    ''' <summary>
    ''' gibt zum betreffenden Projekt eine nach dem Datum aufsteigend sortierte Liste der Meilensteine zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns>nach Datum sortierte Liste der MEilensteine im Projekt </returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getMilestones As SortedList(Of Date, String)
        Get
            Dim tmpValues As New SortedList(Of Date, String)
            Dim tmpDate As Date
            Dim cphase As clsPhase
            Dim cresult As clsMeilenstein

            For p = 1 To Me.CountPhases
                cphase = Me.getPhase(p)

                For r = 1 To cphase.countMilestones
                    cresult = cphase.getMilestone(r)
                    tmpDate = cresult.getDate

                    Dim ok As Boolean = False
                    Do Until ok
                        Try
                            tmpValues.Add(tmpDate, cresult.nameID)
                            ok = True
                        Catch ex As Exception
                            tmpDate = tmpDate.AddSeconds(1)
                        End Try
                    Loop

                Next r

            Next p

            getMilestones = tmpValues

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
            Dim phasenStart As Integer
            Dim tempArray() As Double
            Dim tagessatz As Double
            Dim faktor As Double = nrOfDaysMonth
            Dim dimension As Integer

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
                        anzRollen = .countRoles
                        phasenStart = .relStart - 1
                        'phasenEnde = .relEnde - 1


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
                                    dimension = .getDimension
                                    ReDim tempArray(dimension)
                                    tempArray = .Xwerte
                                    For i = phasenStart To phasenStart + dimension
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
            Dim phasenStart As Integer
            Dim tempArray() As Double
            Dim dimension As Integer


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
                            anzKostenarten = .countCosts
                            phasenStart = .relStart - 1
                            'phasenEnde = .relEnde - 1


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
                                        dimension = .getDimension
                                        ReDim tempArray(dimension)
                                        tempArray = .Xwerte
                                        For i = phasenStart To phasenStart + dimension

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
                        For k = 1 To .countCosts
                            hcost = .getCost(k)
                            If hcost.summe > 0 Then
                                costname = hcost.name
                                '
                                ' das ist performanter als über try .. catch 
                                '
                                If Not aufbauKosten.Contains(costname) Then
                                    aufbauKosten.Add(costname, costname)
                                End If
                                'Try
                                '    aufbauKosten.Add(costname, costname)
                                'Catch ex As Exception

                                'End Try

                            End If
                        Next k
                    End With
                Next p

            End If


            getUsedKosten = aufbauKosten

        End Get

    End Property


    ''' <summary>
    ''' übergibt in getsummekosten die Summe aller Kosten: Personalkosten plus alle anderen Kostenarten
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
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
                    costname = CStr(ErgebnisListe.Item(r))
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


    ''' <summary>
    ''' berechnet die Summe nur bis zum index.ten Monaten 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSummeKosten(ByVal index As Integer) As Double

        Get
            Dim costValues() As Double
            Dim ErgebnisListe As New Collection
            Dim costSum As Double
            Dim anzKostenarten As Integer
            Dim i As Integer, r As Integer
            Dim costname As String

            If _Dauer > 0 Then

                If index > _Dauer - 1 Then
                    index = _Dauer - 1
                End If

                ReDim costValues(_Dauer - 1)
                costValues = Me.getAllPersonalKosten

                costSum = 0
                For i = 0 To index

                    costSum = costSum + costValues(i)

                Next i
                '
                ' jetzt sind in der Summe die Personalkosten drin ....
                '

                ' Jetzt werden die einzelnen Kostenarten auf die gleiche Art und Weise geholt
                ErgebnisListe = Me.getUsedKosten

                anzKostenarten = ErgebnisListe.Count
                For r = 1 To anzKostenarten
                    costname = ErgebnisListe.Item(r).ToString

                    ReDim costValues(_Dauer - 1)
                    costValues = Me.getKostenBedarf(costname)
                    For i = 0 To index

                        costSum = costSum + costValues(i)

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
                    costname = CStr(ErgebnisListe.Item(r))
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
                    costname = CStr(ErgebnisListe.Item(r))
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
                    roleName = CStr(ErgebnisListe.Item(r))
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
                    roleName = CStr(ErgebnisListe.Item(r))
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
            Dim phasenStart As Integer
            Dim tempArray() As Double
            Dim tagessatz As Double
            Dim faktor As Double = nrOfDaysMonth
            Dim dimension As Integer

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
                        anzRollen = .countRoles
                        phasenStart = .relStart - 1
                        'phasenEnde = .relEnde - 1


                        For r = 1 To anzRollen
                            role = .getRole(r)

                            With role
                                tagessatz = .tagessatzIntern
                                dimension = .getDimension
                                ReDim tempArray(dimension)
                                tempArray = .Xwerte
                                For i = phasenStart To phasenStart + dimension
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
        ' Änderung tk 31.3.15
        hierarchy = New clsHierarchy

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
