Imports ProjectBoardDefinitions

''' <summary>
''' Liste der beim Kunden definierten String, Bool, Double CustomFields
''' </summary>
Public Class clsVCSettingCustomfields

    Inherits clsVCSetting
    Public Property value As List(Of clsCustomFieldDefinitions)
    Sub New()
        _value = New List(Of clsCustomFieldDefinitions)
    End Sub
End Class



''' <summary>
''' CustomUerRoles: diese Festlegung triggert, welche Funktionalitäten im Visual Board freigeschaltet sind
''' userID, userName, CustomRolle-Bezeichung, Qualifier Lösungs-Ansatz Allianz-Rollen (et al für Portfolio Planung)
''' </summary>
''' 
Public Class clsVCSettingCustomroles

    Inherits clsVCSetting
    Public Property value As clsCustomUserRolesWeb

    Sub New()
        _value = New clsCustomUserRolesWeb
    End Sub
End Class



''' <summary>
''' Liste aller Rollen-Definitionen
''' Liste aller Kosten-Definitionen
''' </summary>
Public Class clsVCSettingOrganisation

    Inherits clsVCSetting
    Public Property value As clsOrganisationWeb

    Sub New()
        _value = New clsOrganisationWeb
    End Sub
    End Class




