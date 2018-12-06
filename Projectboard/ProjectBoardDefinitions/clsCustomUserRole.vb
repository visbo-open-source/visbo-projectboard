''' <summary>
''' siehe https://visbogmbh.atlassian.net/wiki/spaces/VS/pages/231735299/Erweiterungen+in+Datenmodell+getriggert+durch+Allianz
''' 
''' </summary>
Public Class clsCustomUserRole

    Private _userName As String
    Private _userID As String
    Private _customUserRole As Integer
    Private _specifics As Object

    Public Sub New()
        _userName = ""
        _userID = ""
        _customUserRole = ptCustomUserProfils.projectlead
        _specifics = Nothing
    End Sub

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
                If [Enum].IsDefined(GetType(ptCustomUserProfils), value) Then
                    _customUserRole = value
                End If
            End If

        End Set
    End Property

    Public Property specifics As Object
        Get
            specifics = _specifics
        End Get
        Set(value As Object)
            _specifics = value
        End Set
    End Property



End Class
