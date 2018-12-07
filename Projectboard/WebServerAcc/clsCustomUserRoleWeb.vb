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
End Class
