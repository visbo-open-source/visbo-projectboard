Public Class clsCustomUserRolesWeb
    ' der Schlüssel für das Kopieren in die interne Client datenstruktur setzt sich zusammen aus userID, customerUserRole, specifics in clsCustomUserRoleWeb
    Public Property customUserRoles As List(Of clsCustomUserRoleWeb)

    Public Sub New()
        _customUserRoles = New List(Of clsCustomUserRoleWeb)
    End Sub
End Class
