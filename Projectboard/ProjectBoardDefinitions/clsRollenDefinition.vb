Public Class clsRollenDefinition

    Private _subRoleIDs As SortedList(Of Integer, String)

    Private uuid As Integer
    'Private Kapa() As Double


    Public Property name As String
    Public Property farbe As Object
    Public Property defaultKapa As Double
    Public Property tagessatzIntern As Double
    Public Property tagessatzExtern As Double
    Public Property kapazitaet As Double()
    Public Property externeKapazitaet As Double()

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

            For Each kvp As KeyValuePair(Of Integer, String) In Me.getSubRoleIDs
                If tmpCollection.Contains(kvp.Value) Then
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
    Public ReadOnly Property getSubRoleIDs As SortedList(Of Integer, String)
        Get
            getSubRoleIDs = _subRoleIDs
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
    ''' fügt die entsprechende uid und subrolenamen hinzu .... 
    ''' </summary>
    ''' <param name="subRoleUid"></param>
    ''' <param name="subRoleName"></param>
    ''' <remarks></remarks>
    Public Sub addSubRole(ByVal subRoleUid As Integer, ByVal subRoleName As String, ByVal maxNr As Integer)

        If Not _subRoleIDs.ContainsKey(subRoleUid) Then
            If subRoleUid <= maxNr Then
                _subRoleIDs.Add(subRoleUid, subRoleName)
            Else
                Throw New ArgumentException("unzulässige uid für Subrolle:" & subRoleUid.ToString & ", " & subRoleName)
            End If

        End If

    End Sub


    
    Public Property UID() As Integer

        Get

            UID = uuid

        End Get

        Set(value As Integer)

            uuid = value

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
                        stillok = (Me._subRoleIDs.ElementAt(i).Key = vglRole.getSubRoleIDs.ElementAt(i).Key And _
                                   Me._subRoleIDs.ElementAt(i).Value = vglRole.getSubRoleIDs.ElementAt(i).Value)
                        i = i + 1
                    Loop

                End If
            Else
                stillok = False
            End If


            ' jetzt alle anderen Attribute überprüfen ...
            If stillok Then

                stillok = (Me.UID = vglRole.UID) And _
                            (Me.name = vglRole.name) And _
                            (CLng(Me.farbe) = CLng(vglRole.farbe)) And _
                            (Me.defaultKapa = vglRole.defaultKapa) And _
                            (Me.tagessatzIntern = vglRole.tagessatzIntern) And _
                            (Me.tagessatzExtern = vglRole.tagessatzExtern)

            End If

            ' jetzt die Kapa-Arrays vergleichen 
            If stillok Then
                stillok = Not arraysAreDifferent(Me.kapazitaet, vglRole.kapazitaet) And _
                            Not arraysAreDifferent(Me.externeKapazitaet, vglRole.externeKapazitaet)
            End If

            isIdenticalTo = stillok

        End Get
    End Property

    Public Sub New()

        ' Änderung 29.5.14 damit man zwanzig Jahre vom Start der Projekt-Tafel betrachten kann 
        ' Kapazität: die Null Position hat keine Bedeutung; kapazität(1) = der Wert für StartofCalendar
        ReDim _kapazitaet(240)
        ReDim _externeKapazitaet(240)

        _subRoleIDs = New SortedList(Of Integer, String)

    End Sub

End Class
