''' <summary>
''' siehe https://visbogmbh.atlassian.net/wiki/spaces/VS/pages/231735299/Erweiterungen+in+Datenmodell+getriggert+durch+Allianz
''' 
''' </summary>
Public Class clsCustomUserRole

    Private _userName As String
    Private _userID As String
    Private _customUserRole As ptCustomUserRoles
    ' gibt im Falle resource Mgr an, welche Orga-Einhairt er nur sehen darf 
    Private _specifics As String
    ' gibt im Falle Portfolio Manager an, welche Rollen ggf aggregiert werden sollen 
    Private _portfolioAggregationRoles() As String
    Private _portfolioAggregationRoleIDs() As Integer
    ' wird benötigt, um bestimmen zu können, welche projectboard Funktionalität erlaubt / nicht erlaubt ist 
    Private _nonAllowance() As String

    Public Sub New()
        _userName = ""
        _userID = ""
        _customUserRole = ptCustomUserRoles.Alles
        _specifics = Nothing
        _portfolioAggregationRoles = {""}
        _portfolioAggregationRoleIDs = {1}
        _nonAllowance = {""}
    End Sub

    ''' <summary>
    ''' setzt , in Abhängigkeit von _customUserRole die Menu-Punkt Allowance
    ''' muss aufgerufen werden, sobald eine customUSerRole gewählt wurde 
    ''' </summary>
    Public Sub setNonAllowances()

        Select Case _customUserRole

            Case ptCustomUserRoles.Alles
                _nonAllowance = {""}

            Case ptCustomUserRoles.OrgaAdmin
                _nonAllowance = {"Pt5G2B1", "Pt5G2B3", "PT5G3M",
                                 "PT4G1M1-2", "PT4G1M1-3", "PT4G2M",
                                 "PTneu", "PTedit", "PTview",
                                 "PTfilter", "PTsort", "PT0G1s9",
                                 "PTOPTB1", "PTreport",
                                 "PTeinst", "PThelp", "PTWebServer"}

            Case ptCustomUserRoles.PortfolioManager
                _nonAllowance = {"Pt5G2B4", "PT4G1M1-1", "PT4G1M1-2",
                                 "PTview", "PTfilter", "PTWebServer"}

            Case ptCustomUserRoles.ProjektLeitung
                _nonAllowance = {"Pt5G2B4", "Pt5G3B1", "PT4G1M1-1",
                                 "PT2G1B1", "PT2G1B3", "PTfilter", "PTsort", "PTeinst", "PThelp",
                                 "PTWebServer"}

            Case ptCustomUserRoles.RessourceManager
                _nonAllowance = {"Pt5G2B4", "PT5G3M",
                                 "PT4G1B8", "PT4G1B12", "PT4G1B11",
                                 "PT4G1M1-2", "PT4G1M1-3",
                                 "PT2G1M2B3", "PT2G1M2B8",
                                 "PT4G1M1B2", "PT2G1B1", "PT2G1B3",
                                 "PTfreezeB1", "PTfreezeB2", "PT2G1M1B4",
                                 "PTview", "PTfilter", "PTsort", "PTeinst", "PThelp",
                                 "PTWebServer"}

            Case Else
                _nonAllowance = {""}
        End Select

    End Sub

    ''' <summary>
    ''' gibt an, ob die userRole für die angegebene MenuID berechtigt ist, dass heisst nicht in der nonAllowance aufgeführt ist 
    ''' true
    ''' </summary>
    ''' <param name="menuID"></param>
    ''' <returns></returns>
    Public Function isEntitledForMenu(ByVal menuID As String) As Boolean
        isEntitledForMenu = Not _nonAllowance.Contains(menuID)
    End Function

    Public Property userName As String
        Get
            userName = _userName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _userName = value
            End If
        End Set
    End Property

    Public Property userID As String
        Get
            userID = _userID
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _userID = value
            End If
        End Set
    End Property

    Public Property customUserRole As ptCustomUserRoles
        Get
            customUserRole = _customUserRole
        End Get
        Set(value As ptCustomUserRoles)

            If Not IsNothing(value) Then
                If [Enum].IsDefined(GetType(ptCustomUserRoles), value) Then
                    _customUserRole = value

                    ' Sonderbehandlung , wenn Portfolio Manager
                    If Not IsNothing(_specifics) And _customUserRole = ptCustomUserRoles.PortfolioManager Then
                        Call setAggregationRoles(_specifics)
                    End If

                End If
            End If

        End Set
    End Property

    Public Property specifics As String
        Get
            specifics = _specifics
        End Get
        Set(value As String)
            If Not IsNothing(value) Then

                _specifics = value

                ' Sonderbehandlung: wenn es sich um einen Portfolio Manager handelt 
                If _customUserRole = ptCustomUserRoles.PortfolioManager Then
                    Call setAggregationRoles(_specifics)
                End If

            Else
                _specifics = ""
            End If

        End Set
    End Property

    ''' <summary>
    ''' gibt die Namen der AggregationRoleNames in einem String-Array zurück
    ''' hat nur Bedeutung wenn userRole = portfolioMgr
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getAggregationRoleNames As String()
        Get
            getAggregationRoleNames = _portfolioAggregationRoles
        End Get
    End Property

    ''' <summary>
    ''' gibt die Namen der AggregationRoleIDs in einem Integer-Array zurück
    ''' hat nur Bedeutung wenn userRole = portfolioMgr
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property getAggregationRoleIDs As Integer()
        Get
            getAggregationRoleIDs = _portfolioAggregationRoleIDs
        End Get
    End Property

    Private Sub setAggregationRoles(ByVal aggregationRoleStr As String)

        Dim tmpStr() As String = aggregationRoleStr.Split(New Char() {CChar(";")})
        Dim tmpCollection As New Collection

        For Each tmpName As String In tmpStr
            If RoleDefinitions.containsName(tmpName.Trim) Then
                tmpCollection.Add(tmpName.Trim)
            End If
        Next

        If tmpCollection.Count > 0 Then

            ReDim _portfolioAggregationRoles(tmpCollection.Count - 1)
            ReDim _portfolioAggregationRoleIDs(tmpCollection.Count - 1)

            Dim i As Integer = 0
            For Each tmpName As String In tmpCollection
                Dim tmpRoleDef As clsRollenDefinition = RoleDefinitions.getRoledef(tmpName)
                _portfolioAggregationRoles(i) = tmpRoleDef.name
                _portfolioAggregationRoleIDs(i) = tmpRoleDef.UID
                i = i + 1
            Next
        Else
            _portfolioAggregationRoles = {""}
            _portfolioAggregationRoleIDs = {1}
        End If

    End Sub



End Class
