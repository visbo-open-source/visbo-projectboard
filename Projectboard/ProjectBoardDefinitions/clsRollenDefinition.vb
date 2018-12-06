Public Class clsRollenDefinition

    ' am 21.11.18 dazu gekommen 
    ' _isExternRole, _isTeam, _teamIDs, _defaultDayKapa (errechnen sich wechselseitig auseinander: defaultDayKapa und defaultKapa errechnen sich über nrdayspMonth) 
    ' weggefallen ist:
    ' tagessatzExtern, externeKapazität
    ' in der ..DB Definition bleiben die alten Definitionen erhalten, sie werden nur nicht mehr hin und herkopiert
    ' in der WebDB Definition sollten sie besser ganz rausfliegen. Wir können jetzt noch auf grüner Wiese anfangern.

    ' wenn es sich um ein Team handelt, dann gibt der Double-Wert an, wieviel Prozent der Kapa der SubRoleID in das Team einfliesst 
    Private _subRoleIDs As SortedList(Of Integer, Double)

    ' 
    ' tk Allianz 21.11.18 Teams abbilden 
    ' gibt die Liste der Teams an, in dem die PErson ist 
    ' der Double Wert sagt, wieviel Prozent der Kapa der Person in das Team einfliesst ; Summe sollte 100% nicht überschreiten;
    ' keine harte Grenze, verursacht nur Warnung 
    Private _teamIDs As SortedList(Of Integer, Double)

    ' gibt an, ob es sich um eine interne oder externe Rolle handelt, nur von Bedeutung wenn es sich um ein Blatt handelt ... 
    ' bei externen Rollen werden die Kapa-Values über die Monate automatisch angepasst ; Beauftragt 100 MT bis Juni, abgerufen bis Mrz 30, dann verbleiben 70 in den Monaten Apr - Jun   
    Private _isExternRole As Boolean
    Public Property isExternRole As Boolean
        Get
            isExternRole = _isExternRole
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _isExternRole = value
            Else
                _isExternRole = False
            End If
        End Set
    End Property

    ' gibt an, ob es sich um eine Team Definition handelt 
    Private _isTeam As Boolean
    Public Property isTeam As Boolean
        Get
            isTeam = _isTeam
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _isTeam = value
            Else
                _isTeam = False
            End If
        End Set
    End Property

    ''' <summary>
    ''' ist quasi ein Test-Check zur isTeam 
    ''' getTeamProperty gibt dann und nur dann true, wenn die Rolle Kinder enthält, die alle Team-Member in der Rolle selber sind ...  
    ''' </summary>
    ''' <returns></returns>
    Public Function getTeamProperty() As Boolean

        Dim tmpResult As Boolean = False
        Dim myUID As Integer = _uuid

        If _subRoleIDs.Count > 0 Then
            tmpResult = True
            Dim i As Integer = 0
            Do While tmpResult = True And i <= _subRoleIDs.Count - 1

                Try
                    Dim childRoleID As Integer = _subRoleIDs.ElementAt(i).Key
                    Dim childRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(childRoleID)

                    tmpResult = childRole.getTeamIDs.ContainsKey(myUID)
                    i = i + 1

                Catch ex As Exception
                    tmpResult = False
                End Try

            Loop


        End If

        getTeamProperty = tmpResult

    End Function

    Public ReadOnly Property defaultDayKapa As Double
        Get
            If nrOfDaysMonth > 0 Then
                defaultDayKapa = _defaultKapa / nrOfDaysMonth
            Else
                defaultDayKapa = 0
            End If

        End Get

    End Property

    Private _defaultKapa As Double
    Public Property defaultKapa As Double
        Get

            defaultKapa = _defaultKapa

        End Get
        Set(value As Double)

            If Not IsNothing(value) Then
                If value >= 0 Then
                    _defaultKapa = value

                Else
                    _defaultKapa = 0
                End If
            Else
                _defaultKapa = 0
            End If

        End Set
    End Property

    ' Ende Ergänzungen tk Allianz 21.11.18

    Private _uuid As Integer
    'Private Kapa() As Double


    Public Property name As String
    Public Property farbe As Object

    Public Property tagessatzIntern As Double
    Public Property kapazitaet As Double()

    ' tk Allianz 21.11.18 nicht mehr gültig ..
    'Public Property tagessatzExtern As Double

    'Public Property externeKapazitaet As Double()

    ''' <summary>
    ''' bestimmt, ob die aktuelle Instanz irgendein Kind oder Kindeskind hat, das in tmpCollection aufgeführt ist
    ''' wird nur aufgerufen, wenn Instanz eine Sammelrolle ist
    ''' </summary>
    ''' <param name="tmpCollection"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property hasAnyOfThemAsChild(ByVal tmpCollection As Collection) As Boolean
        Get
            Dim tmpCheck As Boolean = False
            Dim myRoleName As String = Me.name

            For Each kvp As KeyValuePair(Of Integer, Double) In Me.getSubRoleIDs
                Dim tmpName As String = RoleDefinitions.getRoleDefByID(kvp.Key).name
                If tmpCollection.Contains(tmpName) Then
                    tmpCheck = True
                Else
                    ' 
                    If RoleDefinitions.containsUid(kvp.Key) Then
                        Dim tmpRoleDef As clsRollenDefinition = RoleDefinitions.getRoleDefByID(kvp.Key)
                        If tmpRoleDef.isCombinedRole Then
                            tmpCheck = tmpRoleDef.hasAnyOfThemAsChild(tmpCollection)
                        End If
                    End If

                End If

                If tmpCheck = True Then
                    Exit For
                End If

            Next

            hasAnyOfThemAsChild = tmpCheck
        End Get
    End Property

    ''' <summary>
    ''' gibt die Liste an SubRole IDs als sortierte Liste zurück; 
    ''' Nothing wenn es keine gibt 
    ''' oder Dim = 1 , erstes Element = 0 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSubRoleIDs As SortedList(Of Integer, Double)
        Get
            getSubRoleIDs = _subRoleIDs
        End Get
    End Property

    Public ReadOnly Property getTeamIDs As SortedList(Of Integer, Double)
        Get
            getTeamIDs = _teamIDs
        End Get
    End Property

    ''' <summary>
    ''' gibt zurück, ob es sich um eine Combined Role handelt ... 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isCombinedRole As Boolean
        Get
            Dim tmpValue As Boolean = False
            If IsNothing(_subRoleIDs) Then
                tmpValue = False
            ElseIf _subRoleIDs.Count >= 1 Then
                tmpValue = True
            Else
                tmpValue = False
            End If

            isCombinedRole = tmpValue

        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl SubRoles zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSubRoleCount As Integer
        Get
            Dim tmpValue As Integer = 0
            If Not IsNothing(_subRoleIDs) Then
                tmpValue = _subRoleIDs.Count
            Else
                tmpValue = 0
            End If

            getSubRoleCount = tmpValue
        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl Teams zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTeamCount As Integer
        Get
            Dim tmpValue As Integer = 0
            If Not IsNothing(_teamIDs) Then
                tmpValue = _teamIDs.Count
            Else
                tmpValue = 0
            End If

            getTeamCount = tmpValue
        End Get
    End Property

    ''' <summary>
    ''' fügt die entsprechende uid als SubRole hinzu  .... 
    ''' aber es dürfen keine Team-Memberships existieren ... 
    ''' </summary>
    ''' <param name="subRoleUid"></param>
    ''' <param name="subRolePrz">enthält den Prozentsatz, den die Subrolle zur Kapa der Rolel beiträgt</param>
    ''' <remarks></remarks>
    Public Sub addSubRole(ByVal subRoleUid As Integer, ByVal subRolePrz As Double)

        If Not _subRoleIDs.ContainsKey(subRoleUid) And _teamIDs.Count = 0 Then
            If _teamIDs.Count = 0 Then
                _subRoleIDs.Add(subRoleUid, subRolePrz)
            Else
                Throw New ArgumentException("unzulässig für Parentship: hat Team-Zugehörigkeit " & _teamIDs.Count.ToString)
            End If
        End If

    End Sub

    ''' <summary>
    ''' fügt die entsprechende uid als Team hinzu 
    ''' dann dürfen keine Kinder existieren ! 
    ''' </summary>
    ''' <param name="teamUid"></param>
    ''' <param name="teamPrz"></param>
    Public Sub addTeam(ByVal teamUid As Integer, ByVal teamPrz As Double)

        If Not _teamIDs.ContainsKey(teamUid) Then
            If _subRoleIDs.Count = 0 Then
                _teamIDs.Add(teamUid, teamPrz)
            Else
                Throw New ArgumentException("unzulässig für Team-Membership: hat Kinder " & _subRoleIDs.Count.ToString)
            End If
        End If

    End Sub



    Public Property UID() As Integer

        Get

            UID = _uuid

        End Get

        Set(value As Integer)

            _uuid = value

        End Set

    End Property

    ''' <summary>
    ''' true, if both Roledefinitions are identical , except timestamp 
    ''' </summary>
    ''' <param name="vglRole"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vglRole As clsRollenDefinition) As Boolean
        Get
            Dim stillok As Boolean = True

            If Me._subRoleIDs.Count = vglRole.getSubRoleIDs.Count Then
                If Me._subRoleIDs.Count = 0 Then
                    stillok = True
                Else
                    Dim i As Integer = 0
                    Do While i < Me._subRoleIDs.Count And stillok
                        stillok = (Me._subRoleIDs.ElementAt(i).Key = vglRole.getSubRoleIDs.ElementAt(i).Key And
                                   Me._subRoleIDs.ElementAt(i).Value = vglRole.getSubRoleIDs.ElementAt(i).Value)
                        i = i + 1
                    Loop

                    ' jetzt die TeamIDs
                    i = 0
                    Do While i < Me._teamIDs.Count And stillok
                        stillok = (Me._teamIDs.ElementAt(i).Key = vglRole.getTeamIDs.ElementAt(i).Key And
                                   Me._teamIDs.ElementAt(i).Value = vglRole.getTeamIDs.ElementAt(i).Value)
                        i = i + 1
                    Loop

                End If
            Else
                stillok = False
            End If


            ' jetzt alle anderen Attribute überprüfen ...
            If stillok Then

                stillok = (Me.UID = vglRole.UID) And
                            (Me.name = vglRole.name) And
                            (CLng(Me.farbe) = CLng(vglRole.farbe)) And
                            (Me.defaultKapa = vglRole.defaultKapa) And
                            (Me.isExternRole = vglRole.isExternRole) And
                            (Me.isTeam = vglRole.isTeam) And
                            (Me.tagessatzIntern = vglRole.tagessatzIntern)
                'And _
                '            (Me.tagessatzExtern = vglRole.tagessatzExtern)

            End If

            ' jetzt die Kapa-Arrays vergleichen 
            If stillok Then
                stillok = Not arraysAreDifferent(Me.kapazitaet, vglRole.kapazitaet)
                'And _
                '            Not arraysAreDifferent(Me.externeKapazitaet, vglRole.externeKapazitaet)
            End If

            isIdenticalTo = stillok

        End Get
    End Property

    Public Sub New()

        ' Änderung 29.5.14 damit man zwanzig Jahre vom Start der Projekt-Tafel betrachten kann 
        ' Kapazität: die Null Position hat keine Bedeutung; kapazität(1) = der Wert für StartofCalendar
        ReDim _kapazitaet(240)

        _isExternRole = False
        _isTeam = False

        'ReDim _externeKapazitaet(240)

        _subRoleIDs = New SortedList(Of Integer, Double)
        _teamIDs = New SortedList(Of Integer, Double)

    End Sub

End Class
