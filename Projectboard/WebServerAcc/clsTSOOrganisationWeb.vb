Imports ProjectBoardDefinitions
Public Class clsTSOOrganisationWeb
    Public Property _id As String
    Public Property vcid As String
    Public Property name As String
    Public Property timestamp As Date
    Private _allRoles As List(Of clsTSORoleDefinitionWeb)
    Private _allCosts As List(Of clsKostenartDefinitionWeb)
    'Private _allUnits As List(Of clsAllUnitsDefinitionWeb)

    ' tk ergänzt am 17.5 , um die Orga effizienter speichern zu können 
    'Private _OrgaStartOfCalendar As Date

    Public Property allRoles As List(Of clsTSORoleDefinitionWeb)
        Get
            allRoles = _allRoles
        End Get
        Set(value As List(Of clsTSORoleDefinitionWeb))
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

    'Public Property allUnits As List(Of clsAllUnitsDefinitionWeb)
    '    Get
    '        allUnits = _allUnits
    '    End Get
    '    Set(value As List(Of clsAllUnitsDefinitionWeb))
    '        If Not IsNothing(value) Then
    '            _allUnits = value
    '        End If
    '    End Set
    'End Property

    Public ReadOnly Property count As Integer
        Get
            count = _allCosts.Count + _allRoles.Count
        End Get
    End Property


    Public Sub copyFrom(ByVal orgaDef As clsOrganisation)

        With orgaDef

            Me.timestamp = .validFrom.ToUniversalTime
            'Me.OrgaStartOfCalendar = StartofCalendar

            If .allRoles.Count >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In .allRoles.liste

                    Dim rd As New clsRollenDefinition
                    Dim rdweb As New clsTSORoleDefinitionWeb
                    rd = kvp.Value
                    rdweb.copyFrom(rd)
                    Me.allRoles.Add(rdweb)
                Next
            End If

            If .allCosts.Count >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, clsKostenartDefinition) In .allCosts.liste

                    Dim kad As New clsKostenartDefinition
                    Dim kadweb As New clsKostenartDefinitionWeb
                    kad = kvp.Value
                    kadweb.copyFrom(kad)
                    Me.allCosts.Add(kadweb)
                Next
            End If

        End With
    End Sub

    Public Sub copyTo(ByRef orgaDef As clsOrganisation)

        With orgaDef

            .validFrom = Me.timestamp.ToLocalTime


            If Me.allRoles.Count >= 1 Then
                For Each rdweb As clsTSORoleDefinitionWeb In Me.allRoles
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
        _id = ""
        _vcid = ""
        _name = "organisation"
        _allRoles = New List(Of clsTSORoleDefinitionWeb)
        _allCosts = New List(Of clsKostenartDefinitionWeb)
        '_allUnits = New List(Of clsAllUnitsDefinitionWeb)
        _timestamp = Date.Now.Date
    End Sub
End Class
