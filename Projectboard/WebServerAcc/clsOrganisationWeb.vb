Imports ProjectBoardDefinitions
Public Class clsOrganisationWeb

    Private _allRoles As List(Of clsRollenDefinitionWeb)
    Private _allCosts As List(Of clsKostenartDefinitionWeb)
    Private _validFrom As Date

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

    Public ReadOnly Property count As Integer
        Get
            count = _allRoles.Count + _allCosts.Count
        End Get
    End Property


    Public Sub copyFrom(ByVal orgaDef As clsOrganisation)

        With orgaDef

            If .allRoles.Count >= 1 Then
                For Each kvp As KeyValuePair(Of Integer, clsRollenDefinition) In .allRoles.liste

                    Dim rd As New clsRollenDefinition
                    Dim rdweb As New clsRollenDefinitionWeb
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

            Me.validFrom = .validFrom.ToUniversalTime

        End With
    End Sub

    Public Sub copyTo(ByRef orgaDef As clsOrganisation)

        With orgaDef

            If Me.allRoles.Count >= 1 Then
                For Each rdweb As clsRollenDefinitionWeb In Me.allRoles
                    Dim rd As New clsRollenDefinition
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

            .validFrom = Me.validFrom.ToLocalTime

        End With
    End Sub
    '''' <summary>
    '''' prüft, ob die neue Organisation gültig ist; wenn ja, werden auch die Kapa-Werte aus der alten auf die neue übertragen
    '''' sie ist nur dann gültig, wenn jede Element-ID aus der alten Organisation auch im neuen vorkommt 
    '''' eine neue Element-ID darf keinen Namen haben, der bereits in der alten vorkommt 
    '''' sollte eine Element-ID im Neuen einen anderen Namen haben, dann ist das nur gültig, wenn dieser Name die alte Organisation 
    '''' </summary>
    '''' <param name="oldOrga"></param>
    'Public Function validityCheckWith(ByVal oldOrga As clsOrganisation, ByRef outputCollection As Collection) As Boolean

    '    Dim Listeneintraege As Integer = outputCollection.Count
    '    Dim errmsg As String = ""

    '    If IsNothing(oldOrga) Then
    '        ' nichts tun , alles i.O
    '    Else
    '        If oldOrga.count = 0 Then
    '            ' nichts tun , alles i.O 
    '        Else
    '            ' jetzt werden die Bedingungen geprüft ...

    '            Dim stillOK = True

    '            Dim oldRoles As clsRollen = oldOrga.allRoles
    '            Dim anzRoles As Integer = oldRoles.Count
    '            Dim moveKapas As Boolean = False
    '            ' ist jede Rollen-ID im alten auch im Neuen ? 
    '            For ixr As Integer = 1 To anzRoles
    '                moveKapas = False
    '                Dim oldRoleDefinition As clsRollenDefinition = oldRoles.getRoledef(ixr)
    '                Dim newRoleDefinition As clsRollenDefinition = _allRoles.getRoleDefByID(oldRoleDefinition.UID)
    '                If Not IsNothing(newRoleDefinition) Then
    '                    ' schon mal ok 
    '                    If newRoleDefinition.name = oldRoleDefinition.name Then
    '                        ' ok 
    '                        moveKapas = True
    '                    Else
    '                        ' nur ok, wenn der neue Name nicht im alten vorkommt und gleichzeitig der alte nicht woanders im neuen 
    '                        stillOK = Not oldRoles.containsName(newRoleDefinition.name) And
    '                                              Not _allRoles.containsName(oldRoleDefinition.name)

    '                        If Not stillOK Then

    '                            If oldRoles.containsName(newRoleDefinition.name) Then
    '                                errmsg = "Konflikt:" & newRoleDefinition.name & " mit anderem Schlüssel bereits in bisheriger Orga-Definition enthalten .."
    '                            Else
    '                                errmsg = "Konflikt:" & oldRoleDefinition.name & " mit anderem Schlüssel in neuer Orga-Definition enthalten .."
    '                            End If
    '                            outputCollection.Add(errmsg)

    '                        Else
    '                            moveKapas = True
    '                        End If
    '                    End If
    '                Else
    '                    ' nicht ok 
    '                    errmsg = "ID: " & oldRoleDefinition.UID.ToString & " : " & oldRoleDefinition.name & " ist nicht in neuer Orga-Definition vorhanden ..."
    '                    outputCollection.Add(errmsg)
    '                End If

    '                ' jetzt werden die Kapas der alten Rollendefinition übernommen ..
    '                If moveKapas Then
    '                    newRoleDefinition.kapazitaet = oldRoleDefinition.kapazitaet
    '                End If
    '            Next


    '            '' jetzt die Kosten ..
    '            stillOK = True

    '            Dim oldCosts As clsKostenarten = oldOrga.allCosts
    '            Dim anzCosts As Integer = oldCosts.Count

    '            ' ist jede Kosten-ID im alten auch im Neuen ? 
    '            For ixc As Integer = 1 To anzCosts
    '                Dim oldCostDefinition As clsKostenartDefinition = oldCosts.getCostdef(ixc)
    '                Dim newCostDefinition As clsKostenartDefinition = _allCosts.getCostDefByID(oldCostDefinition.UID)
    '                If Not IsNothing(newCostDefinition) Then
    '                    ' schon mal ok 
    '                    If newCostDefinition.name = oldCostDefinition.name Then
    '                        ' schon mal ok 
    '                    Else
    '                        ' nur ok, wenn der neue Name nicht im alten vorkommt und gleichzeitig der alte nicht woanders im neuen 
    '                        stillOK = Not oldCosts.containsName(newCostDefinition.name) And
    '                                              Not _allCosts.containsName(oldCostDefinition.name)

    '                        If Not stillOK Then

    '                            If oldCosts.containsName(newCostDefinition.name) Then
    '                                errmsg = "Konflikt:" & newCostDefinition.name & " mit anderer Kosten-ID bereits in bisheriger Orga-Definition enthalten .."
    '                            Else
    '                                errmsg = "Konflikt:" & oldCostDefinition.name & " mit anderem Kosten-ID in neuer Orga-Definition enthalten .."
    '                            End If
    '                            outputCollection.Add(errmsg)

    '                        End If
    '                    End If
    '                Else
    '                    ' nicht ok 
    '                    errmsg = "ID: " & oldCostDefinition.UID.ToString & " : " & oldCostDefinition.name & " ist nicht in neuer Kosten Orga-Definition vorhanden ..."
    '                    outputCollection.Add(errmsg)
    '                End If
    '            Next

    '        End If
    '    End If

    '    ' wenn es keine Einträge gegeben hat, dann ist alles o.k
    '    validityCheckWith = (Listeneintraege = outputCollection.Count)

    'End Function
    Public Sub New()
        _allRoles = New List(Of clsRollenDefinitionWeb)
        _allCosts = New List(Of clsKostenartDefinitionWeb)
        _validFrom = Date.Now.Date
    End Sub
    '' besser? as list(of clsRollendefinitionWeb)
    'Public Property roles As List(Of clsRollenDefinitionWeb)
    'Public Property costs As List(Of clsKostenDefinitionWeb)
    'Public Property validFrom As Date
    'Sub New()
    '    _roles = New List(Of clsRollenDefinitionWeb)
    '    _vccosts = New List(Of clsKostenDefinitionWeb)
    '    _validFrom = Date.Now.Date
    'End Sub
End Class
