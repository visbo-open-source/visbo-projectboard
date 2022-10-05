Public Class clsRollenDefinition

    ' Änderungen
    ' tk 8.1.2020
    ' Neu: _defaultDayKapa as Double: gibt die Anzahl Stunden wieder , die der Mitarbeiter am Tag macht. Und zwar , wenn Urlaub, Krankheit nicht eingerechnet wird. 
    ' in der DefaultKapa ist in der monatlichen KApazität bereits der typische Anteil Urlaub, Krankheit berücksichtigt; in der defaultDayKapa nicht 
    ' Neu: _startDate: Date gibt an, ab wann der Mitarbeiter zur Verfügung steht. Wird nur bei der Ressourcen Zuordnung bzw. beim Import berücksichtigt ... 
    ' Neu: _endDate: Date gibt an, wann der Mitarbeiter das Unternehmen verlassen hat bzw verlassen wird. 
    ' Neu: _employeeNr : String: ist die Personal-Nummer des Mitarbeiters im Unternehmen; dient nur der Namens- und schreibweisen-toleranten Erkennung des Mitarbeiters 



    ' am 21.11.18 dazu gekommen 
    ' _isExternRole, _isTeam, _teamIDs, _defaultDayKapa (errechnen sich wechselseitig auseinander: defaultDayKapa und defaultKapa errechnen sich über nrdayspMonth) 
    ' weggefallen ist:
    ' tagessatzExtern, externeKapazität
    ' in der ..DB Definition bleiben die alten Definitionen erhalten, sie werden nur nicht mehr hin und herkopiert
    ' in der WebDB Definition sollten sie besser ganz rausfliegen. Wir können jetzt noch auf grüner Wiese anfangern.

    ' wenn es sich um ein Team handelt, dann gibt der Double-Wert an, wieviel Prozent der Kapa der SubRoleID in das Team einfliesst 
    Private _subRoleIDs As SortedList(Of Integer, Double)


    ' neue Properties seit 8.1.20 
    Private _aliases As String()
    Public Property aliases As String()
        Get
            aliases = _aliases
        End Get
        Set(value As String())
            _aliases = value
        End Set
    End Property

    Private _employeeNr As String
    Public Property employeeNr As String
        Get
            employeeNr = _employeeNr
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _employeeNr = value
            End If
        End Set
    End Property

    ' ist das Datum, ab wann der Mitarbeiter im Unternehmen bereit steht - wird immer auf den 1.Tag des Monats gesetzt 
    Private _entryDate As Date
    Public Property entryDate As Date
        Get
            entryDate = _entryDate
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                If value > Date.MinValue Then
                    ' immer der erste Tag des Monats
                    _entryDate = value.Date.AddDays(-1 * value.Date.Date.Day + 1)
                Else
                    _entryDate = Date.MinValue
                End If

            End If
        End Set
    End Property

    ' ist das Datum, ab wann der Mitarbeiter nicht mehr im Unternehmen ist. wird immer auf den 1.Tag des Monats gesetzt 
    Private _exitDate As Date
    Public Property exitDate As Date
        Get
            exitDate = _exitDate
        End Get
        Set(value As Date)
            If Not IsNothing(value) Then
                If value > _entryDate Then
                    If value > Date.MinValue Then
                        ' immer der erste Tag des Monats 
                        _exitDate = value.Date.AddDays(-1 * value.Date.Date.Day + 1)
                    Else
                        _exitDate = Date.MinValue
                    End If

                End If

            End If
        End Set
    End Property

    ' tk 8.1.20
    Private _defaultDayCapa As Double
    Public Property defaultDayCapa As Double
        Get
            If _defaultDayCapa < 0 Then
                ' das ist die Vorbesetzung, stellt sicher, dass auch in alten Umgebungen das Ganze noch funktioniert 
                If nrOfDaysMonth > 0 Then
                    defaultDayCapa = _defaultKapa / nrOfDaysMonth
                Else
                    defaultDayCapa = 0
                End If
            Else
                defaultDayCapa = _defaultDayCapa
            End If


        End Get
        Set(value As Double)
            If Not IsNothing(value) Then
                If value > 0 And value <= 20 Then
                    _defaultDayCapa = value
                End If
            End If
        End Set

    End Property

    ' 
    ' tk Allianz 21.11.18 Teams abbilden 
    ' gibt die Liste der Teams an, in dem die PErson ist 
    ' der Double Wert hat keine Wirkung mehr !  
    Private _skillIDs As SortedList(Of Integer, Double)

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
    Private _isSkill As Boolean
    Public Property isSkill As Boolean
        Get
            isSkill = _isSkill Or _isSkillParent
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _isSkill = value
            Else
                _isSkill = False
            End If
        End Set
    End Property

    ''' <summary>
    ''' returns true, when role is active during timeframe given by showrangeleft and showrangeRight
    ''' </summary>
    ''' <returns></returns>
    Public Function isActiveRole() As Boolean
        Dim result As Boolean = True

        If showRangeLeft > 0 And showRangeRight > showRangeLeft Then
            Dim columnOfEntryDate As Integer = getColumnOfDate(entryDate)
            Dim columnOfExitDate As Integer = getColumnOfDate(exitDate)

            result = columnOfEntryDate <= showRangeRight And columnOfExitDate > showRangeLeft
        End If

        isActiveRole = result
    End Function

    ''' <summary>
    ''' returns true, when role is active during timeframe given by fromDateCol and toDateCol 
    ''' </summary>
    ''' <param name="fromDateCol"></param>
    ''' <param name="toDateCol"></param>
    ''' <returns></returns>
    Public Function isActiveRole(ByVal fromDateCol As Integer, ByVal toDateCol As Integer) As Boolean
        Dim result As Boolean = True

        If fromDateCol > 0 And toDateCol >= fromDateCol Then
            Dim columnOfEntryDate As Integer = getColumnOfDate(entryDate)
            Dim columnOfExitDate As Integer = getColumnOfDate(exitDate)

            result = columnOfEntryDate <= toDateCol And columnOfExitDate > fromDateCol

        End If

        isActiveRole = result

    End Function

    ''' <summary>
    ''' ist quasi ein Test-Check zur isTeam 
    ''' getTeamProperty gibt dann und nur dann true, wenn die Rolle Kinder enthält, die alle Team-Member in der Rolle selber sind ...  
    ''' </summary>
    ''' <returns></returns>
    Public Function isSkillLeaf() As Boolean

        Dim tmpResult As Boolean = False
        Dim myUID As Integer = _uuid

        If _subRoleIDs.Count > 0 Then
            tmpResult = True
            Dim i As Integer = 0
            Do While tmpResult = True And i <= _subRoleIDs.Count - 1

                Try
                    Dim childRoleID As Integer = _subRoleIDs.ElementAt(i).Key
                    Dim childRole As clsRollenDefinition = RoleDefinitions.getRoleDefByID(childRoleID)

                    tmpResult = childRole.getSkillIDs.ContainsKey(myUID)
                    i = i + 1

                Catch ex As Exception
                    tmpResult = False
                End Try

            Loop


        End If

        isSkillLeaf = tmpResult

    End Function


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

    Private _isSkillParent As Boolean
    Public Property isSkillParent As Boolean
        Get
            isSkillParent = _isSkillParent
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _isSkillParent = value
            Else
                _isSkillParent = False
            End If

        End Set
    End Property

    Public Property name As String

    Public ReadOnly Property farbe As Integer
        Get
            farbe = visboFarbeBlau
        End Get
    End Property


    Public Property tagessatzIntern As Double

    ' gibt an, ob es sich um eine Rolle handelt, auf die die darunterliegenden Rollenbedarfe aggregiert werden sollen
    Private _isAggregationRole As Boolean
    Public Property isAggregationRole As Boolean
        Get
            isAggregationRole = _isAggregationRole
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _isAggregationRole = value
            Else
                _isAggregationRole = False
            End If
        End Set
    End Property

    ' gibt an, ob es sich um eine Rolle handelt, die keine Person ist
    Private _isSummaryRole As Boolean
    Public Property isSummaryRole As Boolean
        Get
            isSummaryRole = Me.isCombinedRole Or _isSummaryRole
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                ' ur:20210728 restructure of Organisation
                If (value) Then
                    _isSummaryRole = value
                Else
                    _isSummaryRole = Me.isCombinedRole
                End If

            Else
                _isSummaryRole = False
            End If
        End Set
    End Property

    ' gibt an, ob es sich um eine Rolle handelt, die/oder deren Kinder IstDaten erhalten => vor Eintragung der Istdaten werden deren Bedarfe genullt
    Private _isActDataRelevant As Boolean
    Public Property isActDataRelevant As Boolean
        Get
            isActDataRelevant = _isActDataRelevant
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _isActDataRelevant = value
            Else
                _isActDataRelevant = False
            End If
        End Set
    End Property

    Public Property kapazitaet As Double()



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
    ''' wenn es sich um einen Ressourcen-Manager handelt, werden nur die Personen, Untereinheiten angezeigt, die mindestens eine Person aus ihrem eigenen Ressort enthalten  
    ''' leere Liste, wenn es keine gibt 
    ''' oder Dim = 1 , erstes Element = 0 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSubRoleIDs As SortedList(Of Integer, Double)
        Get
            ' tk 15.1.20 wenn es sich bei der Rolle um ein Team handelt und es sich um einen Ressourcen-Manager handelt , dann werden nur die 
            ' Skillgruppen gezeigt, die mindestens ein Mitglied in dem Team haben 
            If myCustomUserRole.customUserRole = ptCustomUserRoles.RessourceManager And Me.isSkill = True Then
                ' alle Kinder der Skill bringen, aber nur die, die Teil der Organisations-Unit von Ressourcen Manager sind .. 

                Dim restrictedToOrgaID As Integer = CInt(myCustomUserRole.specifics)
                Dim restrictedSubRoleIDs As New SortedList(Of Integer, Double)

                For Each kvp As KeyValuePair(Of Integer, Double) In _subRoleIDs

                    ' wenn das Kind der Skill mindestens eine gemeinsame Ressourcen hat ... 
                    If RoleDefinitions.getCommonChildsOfParents(restrictedToOrgaID, kvp.Key).Count > 0 Then
                        restrictedSubRoleIDs.Add(kvp.Key, 1.0)
                    End If

                    Dim roleName As String = RoleDefinitions.getRoleDefByID(kvp.Key).name

                Next

                getSubRoleIDs = restrictedSubRoleIDs

            Else
                getSubRoleIDs = _subRoleIDs
            End If

        End Get
    End Property

    Public ReadOnly Property getSkillIDs As SortedList(Of Integer, Double)
        Get
            getSkillIDs = _skillIDs
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
    ''' dabei muss unterschieden werden, ob es sich um Ressourcen-Manager handelt und die subRoles eines Knoten berechnet werden müssen, der Vater von Teams ist .. 
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
    Public ReadOnly Property getSkillCount As Integer
        Get
            Dim tmpValue As Integer = 0
            If Not IsNothing(_skillIDs) Then
                tmpValue = _skillIDs.Count
            Else
                tmpValue = 0
            End If

            getSkillCount = tmpValue
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

        If Not _subRoleIDs.ContainsKey(subRoleUid) And _skillIDs.Count = 0 Then
            If _skillIDs.Count = 0 Then
                _subRoleIDs.Add(subRoleUid, subRolePrz)
            Else
                Throw New ArgumentException("unzulässig für Parentship: hat Team-Zugehörigkeit " & _skillIDs.Count.ToString)
            End If
        End If

    End Sub

    ''' <summary>
    ''' fügt die entsprechende uid als Team hinzu 
    ''' dann dürfen keine Kinder existieren ! 
    ''' </summary>
    ''' <param name="skillUid"></param>
    ''' <param name="teamPrz"></param>
    Public Sub addSkill(ByVal skillUid As Integer, ByVal teamPrz As Double)

        If Not _skillIDs.ContainsKey(skillUid) Then
            If _subRoleIDs.Count = 0 Then
                _skillIDs.Add(skillUid, teamPrz)
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
    Public ReadOnly Property isIdenticalTo(ByVal vglRole As clsRollenDefinition, Optional ByVal mitKapa As Boolean = True) As Boolean
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

                End If
            Else
                stillok = False
            End If

            ' jetzt die TeamIDs prüfen 
            If Me._skillIDs.Count = vglRole.getSkillIDs.Count Then
                If Me._skillIDs.Count = 0 Then
                    stillok = True
                Else

                    Dim i As Integer = 0
                    Do While i < Me._skillIDs.Count And stillok
                        stillok = (Me._skillIDs.ElementAt(i).Key = vglRole.getSkillIDs.ElementAt(i).Key And
                                   Me._skillIDs.ElementAt(i).Value = vglRole.getSkillIDs.ElementAt(i).Value)
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
                            (Me.defaultKapa = vglRole.defaultKapa) And
                            (Me.isExternRole = vglRole.isExternRole) And
                            (Me.isSkill = vglRole.isSkill) And
                            (Me.tagessatzIntern = vglRole.tagessatzIntern) And
                            (Me.employeeNr = vglRole.employeeNr) And
                            (Me.entryDate.Date = vglRole.entryDate.Date) And
                            (Me.exitDate.Date = vglRole.exitDate.Date) And
                            (Me.defaultDayCapa = vglRole.defaultDayCapa)
                '(CLng(Me.farbe) = CLng(vglRole.farbe)) And

            End If
            ' jetzt die aliases vergleichen
            If stillok Then
                If Not IsNothing(Me.aliases) Then
                    For Each aliasName As String In Me.aliases
                        If Not IsNothing(vglRole.aliases) Then
                            stillok = stillok And vglRole.aliases.Contains(aliasName)
                        Else
                            stillok = False
                        End If
                    Next
                End If
            End If

            ' kapaArray nur vergleichen, wenn mitKapa = true ist
            If mitKapa Then
                ' jetzt die Kapa-Arrays vergleichen 
                If stillok Then
                    stillok = Not arraysAreDifferent(Me.kapazitaet, vglRole.kapazitaet)

                End If
            End If


            isIdenticalTo = stillok

        End Get
    End Property

    Public Sub New()

        ' Änderung 29.5.14 damit man zwanzig Jahre vom Start der Projekt-Tafel betrachten kann 
        ' Kapazität: die Null Position hat keine Bedeutung; kapazität(1) = der Wert für StartofCalendar
        ReDim _kapazitaet(240)

        _isExternRole = False
        _isSkill = False

        ' tk wird aktuell noch nicht in der DB gespeichert, wird beim buildOrgaTeams gesetzt 
        _isSkillParent = False


        _subRoleIDs = New SortedList(Of Integer, Double)
        _skillIDs = New SortedList(Of Integer, Double)

        _employeeNr = ""
        _entryDate = Date.MinValue
        ' _exitDate = CDate("31.12.2200")
        _exitDate = DateAndTime.DateSerial(2200, 12, 31)
        _defaultDayCapa = -1
        _aliases = Nothing
        _isAggregationRole = False
        _isSummaryRole = False
        _isActDataRelevant = False

    End Sub

End Class
