Imports ProjectBoardDefinitions
Public Class clsOrganisationWeb

    Private _allRoles As List(Of clsRollenDefinitionWeb)
    Private _allCosts As List(Of clsKostenartDefinitionWeb)
    Private _validFrom As Date

    ' tk ergänzt am 17.5 , um die Orga effizienter speichern zu können 
    'Private _OrgaStartOfCalendar As Date

    Public Property allRoles As List(Of clsRollenDefinitionWeb)
        Get
            allRoles = _allRoles
        End Get
        Set(value As List(Of clsRollenDefinitionWeb))
            If Not IsNothing(value) Then
                _allRoles = value
            End If
        End Set
    End Property

    Public Property allCosts As List(Of clsKostenartDefinitionWeb)
        Get
            allCosts = _allCosts
        End Get
        Set(value As List(Of clsKostenartDefinitionWeb))
            If Not IsNothing(value) Then
                _allCosts = value
            End If
        End Set
    End Property

    Public Property validFrom As Date
        Get
            validFrom = _validFrom
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                _validFrom = value
            End If
        End Set
    End Property

    'Public Property OrgaStartOfCalendar As Date
    '    Get
    '        OrgaStartOfCalendar = _OrgaStartOfCalendar
    '    End Get
    '    Set(value As Date)
    '        If Not IsNothing(value) Then
    '            _OrgaStartOfCalendar = value
    '        End If
    '    End Set
    'End Property

    Public ReadOnly Property count As Integer
        Get
            count = _allRoles.Count + _allCosts.Count
        End Get
    End Property


    'Public Sub copyFrom(ByVal orgaDef As clsOrganisation)

    '    With orgaDef

    '        Me.validFrom = .validFrom.ToUniversalTime
    '        'Me.OrgaStartOfCalendar = StartofCalendar

    '        If .allRoles.Count >= 1 Then
    '            For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In .allRoles.liste

    '                Dim rd As New clsRollenDefinition
    '                Dim rdweb As New clsRollenDefinitionWeb
    '                rd = kvp.Value
    '                rdweb.copyFrom(rd)
    '                Me.allRoles.Add(rdweb)
    '            Next
    '        End If

    '        If .allCosts.Count >= 1 Then
    '            For Each kvp As KeyValuePair(Of Integer, clsKostenartDefinition) In .allCosts.liste

    '                Dim kad As New clsKostenartDefinition
    '                Dim kadweb As New clsKostenartDefinitionWeb
    '                kad = kvp.Value
    '                kadweb.copyFrom(kad)
    '                Me.allCosts.Add(kadweb)
    '            Next
    '        End If



    '    End With
    'End Sub

    Public Sub copyTo(ByRef orgaDef As clsOrganisation)

        With orgaDef

            .validFrom = Me.validFrom.ToLocalTime


            If Me.allRoles.Count >= 1 Then
                For Each rdweb As clsRollenDefinitionWeb In Me.allRoles
                    Dim rd As New clsRollenDefinition
                    'rdweb.copyTo(rd, OrgaStartOfCalendar)
                    rdweb.copyTo(rd)
                    .allRoles.Add(rd)
                Next
            End If

            If Me.allCosts.Count >= 1 Then
                For Each kadWeb As clsKostenartDefinitionWeb In Me.allCosts
                    Dim kad As New clsKostenartDefinition
                    kadWeb.copyTo(kad)
                    .allCosts.Add(kad)
                Next
            End If



        End With
    End Sub

    Public Sub New()
        _allRoles = New List(Of clsRollenDefinitionWeb)
        _allCosts = New List(Of clsKostenartDefinitionWeb)
        _validFrom = Date.Now.Date


    End Sub

End Class
