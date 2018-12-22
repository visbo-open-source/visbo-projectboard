''' <summary>
''' siehe https://visbogmbh.atlassian.net/wiki/spaces/VS/pages/231735299/Erweiterungen+in+Datenmodell+getriggert+durch+Allianz
''' 
''' </summary>
Public Class clsCustomUserRole

    Private _userName As String
    Private _userID As String
    Private _customUserRole As Integer
    Private _specifics As String

    ' wird benötigt, um später bestimmen zu können, welche projectboard Funktionalität erlaubt / nicht erlaubt ist 
    Private _nonAllowance() As String

    Public Sub New()
        _userName = ""
        _userID = ""
        _customUserRole = ptCustomUserRoles.all
        _specifics = Nothing
        _nonAllowance = {""}
    End Sub

    ''' <summary>
    ''' setzt , in Abhängigkeit von _customUserRole die Menu-Punkt Allowance
    ''' muss aufgerufen werden, sobald eine customUSerRole gewählt wurde 
    ''' </summary>
    Public Sub setNonAllowances()

        Select Case _customUserRole

            Case ptCustomUserRoles.all
                _nonAllowance = {""}

            Case ptCustomUserRoles.orgaadmin
                _nonAllowance = {""}

            Case ptCustomUserRoles.portfoliomgr
                _nonAllowance = {""}

            Case ptCustomUserRoles.projectlead
                _nonAllowance = {""}

            Case ptCustomUserRoles.resourcemgr
                _nonAllowance = {""}

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

    Public Property customUserRole As Integer
        Get
            customUserRole = _customUserRole
        End Get
        Set(value As Integer)

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
