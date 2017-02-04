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

    Public Sub New()

        ' Änderung 29.5.14 damit man zwanzig Jahre vom Start der Projekt-Tafel betrachten kann 
        ' Kapazität: die Null Position hat keine Bedeutung; kapazität(1) = der Wert für StartofCalendar
        ReDim _kapazitaet(240)
        ReDim _externeKapazitaet(240)

        _subRoleIDs = New SortedList(Of Integer, String)

    End Sub

End Class
