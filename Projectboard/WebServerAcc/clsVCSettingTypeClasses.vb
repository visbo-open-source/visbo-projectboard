Imports ProjectBoardDefinitions

''' <summary>
''' Liste der beim Kunden definierten String, Bool, Double CustomFields
''' </summary>
Public Class clsVCSettingCustomfields

        Inherits clsWebVCSetting
        Public Property value As clsCustomFieldDefinitions

        Sub New()
            _value = New clsCustomFieldDefinitions
        End Sub
    End Class


    ''' <summary>
    ''' CustomUerRoles: diese Festlegung triggert, welche Funktionalitäten im Visual Board freigeschaltet sind
    ''' userID, userName, CustomRolle-Bezeichung, Qualifier Lösungs-Ansatz Allianz-Rollen (et al für Portfolio Planung)
    ''' </summary>
    ''' 
    Public Class clsVCSettingCustomroles

        Inherits clsWebVCSetting
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

        Inherits clsWebVCSetting
        Public Property value As clsVCOrganisationWeb

        Sub New()
            _value = New clsVCOrganisationWeb
        End Sub
    End Class




