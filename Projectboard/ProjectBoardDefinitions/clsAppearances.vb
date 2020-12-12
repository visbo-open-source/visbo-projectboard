Public Class clsAppearances

    Private _appearanceListe As SortedList(Of String, clsAppearance)


    Private ReadOnly Property Item(ByVal appearanceName As String, ByVal isMilestone As Boolean) As clsAppearance
        Get
            Dim tmpResult As clsAppearance = Nothing
            If _appearanceListe.ContainsKey(appearanceName) Then
                tmpResult = _appearanceListe.Item(appearanceName)
            Else
                If isMilestone Then
                    appearanceName = awinSettings.defaultMilestoneClass
                Else
                    appearanceName = awinSettings.defaultPhaseClass
                End If

                ' does this one exist? 
                If _appearanceListe.ContainsKey(appearanceName) Then
                    tmpResult = _appearanceListe.Item(appearanceName)
                Else
                    ' search for the first existent milestone /phase class 
                    If _appearanceListe.Count > 0 Then
                        For Each kvp As KeyValuePair(Of String, clsAppearance) In _appearanceListe
                            If kvp.Value.isMilestone = isMilestone Then
                                tmpResult = kvp.Value
                                Exit For
                            End If
                        Next
                    End If
                End If
            End If

            Item = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' returns the 
    ''' </summary>
    ''' <returns></returns>
    Public Property liste As SortedList(Of String, clsAppearance)
        Get
            liste = _appearanceListe
        End Get
        Set(value As SortedList(Of String, clsAppearance))
            If Not IsNothing(value) Then
                _appearanceListe = value
            Else
                _appearanceListe = New SortedList(Of String, clsAppearance)
            End If
        End Set
    End Property

    ''' <summary>
    ''' returns true if name exists , either for phase- or milestone appearances 
    ''' </summary>
    ''' <param name="appearanceName"></param>
    ''' <returns></returns>
    Public ReadOnly Property contains(ByVal appearanceName As String) As Boolean
        Get
            contains = _appearanceListe.ContainsKey(appearanceName)
        End Get
    End Property

    ''' <summary>
    ''' returns the appropriate appearance Class
    ''' First Choice: MilestoneDefinitions
    ''' Second Choice: allocation of appearance in Milestone
    ''' Third Choice: awinsettings.defaultMilestoneClass
    ''' Fourth Choice: first milestone appearance clas sin List 
    ''' Fifth Choice: Nothing 
    ''' </summary>
    ''' <param name="ms"></param>
    ''' <returns></returns>
    Public ReadOnly Property getMileStoneAppearance(ByVal ms As clsMeilenstein) As clsAppearance
        Get
            Dim tmpResult As clsAppearance = Nothing
            ' als erstes: gibt es den Namen in der Milestone-Definition?
            If MilestoneDefinitions.Contains(ms.name) Then
                tmpResult = Item(MilestoneDefinitions.getMilestoneDef(ms.name).darstellungsKlasse, True)
            Else
                tmpResult = Item(ms.appearanceName, True)
            End If
            getMileStoneAppearance = tmpResult
        End Get
    End Property

    Public ReadOnly Property getMileStoneAppearance(ByVal msName As String, ByVal appearanceName As String) As clsAppearance
        Get
            Dim tmpResult As clsAppearance = Nothing
            ' als erstes: gibt es den Namen in der Milestone-Definition?
            If MilestoneDefinitions.Contains(msName) Then
                tmpResult = Item(MilestoneDefinitions.getMilestoneDef(msName).darstellungsKlasse, True)
            Else
                tmpResult = Item(appearanceName, True)
            End If

            getMileStoneAppearance = tmpResult
        End Get
    End Property

    ''' <summary>
    ''' returns the appropriate appearance Class
    ''' First Choice: PhaseDefinitions
    ''' Second Choice: allocation of appearance in Milestone
    ''' Third Choice: awinsetitngs.defaultMilestoneClass
    ''' Fourth Choice: first milestone appearance clas sin List 
    ''' Fifth Choice: Nothing 
    ''' </summary>
    ''' <param name="cphase"></param>
    ''' <returns></returns>
    Public ReadOnly Property getPhaseAppearance(ByVal cphase As clsPhase) As clsAppearance
        Get
            Dim tmpResult As clsAppearance = Nothing
            ' als erstes: gibt es den Namen in der Milestone-Definition?
            If PhaseDefinitions.Contains(cphase.name) Then
                tmpResult = Item(PhaseDefinitions.getPhaseDef(cphase.name).darstellungsKlasse, False)
            Else
                tmpResult = Item(cphase.appearanceName, False)
            End If
            getPhaseAppearance = tmpResult
        End Get
    End Property

    Public ReadOnly Property getPhaseAppearance(ByVal phName As String, appearanceName As String) As clsAppearance
        Get
            Dim tmpResult As clsAppearance = Nothing
            ' als erstes: gibt es den Namen in der Milestone-Definition?
            If PhaseDefinitions.Contains(phName) Then
                tmpResult = Item(PhaseDefinitions.getPhaseDef(phName).darstellungsKlasse, False)
            Else
                tmpResult = Item(appearanceName, False)
            End If
            getPhaseAppearance = tmpResult
        End Get
    End Property
    Public Sub New()
        _appearanceListe = New SortedList(Of String, clsAppearance)
    End Sub
End Class
