
Imports ProjectBoardDefinitions
Public Class clsCustomUserRoleWeb
    Public Property userName As String
    Public Property userID As String
    Public Property customUserRole As Integer
    Public Property specifics As String

    ''' <summary>
    ''' der eindeutige Key ist userID+customerUserRole+specifics; diese Kombination kann nur einmal vorkommen
    ''' ein User kann damit viele customUserRoles wahrnehmen
    ''' </summary>
    Public Sub New()
        _userName = ""
        _userID = ""
        _customUserRole = 0
        _specifics = ""
    End Sub


    ''' <summary>
    ''' kopiert den Inhalt der Struktur für den Server in die CustomUserRole (client) 
    ''' </summary>
    ''' <param name="clientRoledef"></param>
    Public Sub copyTo(ByRef clientRoledef As clsCustomUserRole)

        With clientRoledef
            .userName = Me.userName
            .userID = Me.userID
            .customUserRole = Me.customUserRole
            .specifics = Me.specifics
        End With

    End Sub


    ''' <summary>
    ''' kopiert den Inhalt der CustomUserRole (client) in die Struktur für den Server
    ''' </summary>
    ''' <param name="serverRoledef"></param>
    Public Sub copyFrom(ByVal serverRoledef As clsCustomUserRole)
        With serverRoledef
            Me.userName = .userName
            Me.userID = .userID
            Me.customUserRole = .customUserRole
            Me.specifics = .specifics
        End With

    End Sub
End Class
