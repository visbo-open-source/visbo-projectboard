Imports ProjectBoardDefinitions
Public Class clsCustomSettingsRPA
    ' der Schlüssel für das Kopieren in die interne Client datenstruktur setzt sich zusammen aus userID, customerUserRole, specifics in clsCustomUserRoleWeb
    Public Property customUserSettingsRPA As List(Of clsCustomSettingRPA)

    Public Sub New()
        _customUserSettingsRPA = New List(Of clsCustomSettingRPA)
    End Sub

End Class
