''' <summary>
''' siehe https://visbogmbh.atlassian.net/wiki/spaces/VS/pages/231735299/Erweiterungen+in+Datenmodell+getriggert+durch+Allianz
''' 
''' </summary>
Public Class clsCustomUserRole

    Private _userName As String
    Private _userID As String
    Private _customUserRole As ptCustomUserRoles
    Private _specifics As String

    ' wird benötigt, um später bestimmen zu können, welche projectboard Funktionalität erlaubt / nicht erlaubt ist 
    Private _nonAllowance() As String

    Public Sub New()
        _userName = ""
        _userID = ""
        _customUserRole = ptCustomUserRoles.Alles
        _specifics = Nothing
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
                End If
            End If

        End Set
    End Property

    Public Property specifics As String
        Get
            specifics = _specifics
        End Get
        Set(value As String)
            _specifics = value
        End Set
    End Property



End Class
