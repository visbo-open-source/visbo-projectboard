Imports ProjectBoardDefinitions

''' <summary>
''' Liste der beim Kunden definierten String, Bool, Double CustomFields
''' </summary>
Public Class clsVCSettingCustomfields

    Inherits clsVCSetting
    Public Property value As clsCustomFieldDefinitionsWeb
    Sub New()
        _value = New clsCustomFieldDefinitionsWeb
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
''' <summary>
''' 
''' </summary>
Public Class clsVCSettingEverything

    Inherits clsVCSetting
    Public Property value As Object

    Sub New()
        _value = New Object
    End Sub
End Class

''' <summary>
''' Definitionen, die ursprünglich im Customization-File (Einstellungen) gespeichert waren
''' z.B. StartOfCalendar, Farben für rote, gelbe, grüne Ampel, Farbe für vieles mehr
''' Arbeitstage pro Monat, RollenKostenDefinition aus DB lesen, Zeiteinheit, Kapzitätseinheit, Zeilenhöhe, Spaltenbreite, unbekannte Phase/Meilensteine Definitionen auto aufnehmen,
''' Duplikate eliminieren, Sprachen englisch, ...
''' </summary>
Public Class clsVCSettingCustomization

    Inherits clsVCSetting
    Public Property value As clsCustomizationWeb

    Sub New()
        _value = New clsCustomizationWeb
    End Sub

End Class
''' <summary>
''' Definitionen, die ursprünglich im Customization-File (Darstellungsklassen) gespeichert waren
''' </summary>
Public Class clsVCSettingAppearance

    Inherits clsVCSetting
    Public Property value As clsAppearanceWeb

    Sub New()
        _value = New clsAppearanceWeb
    End Sub

End Class
Public Class clsVCSettingConfiguration
    Inherits clsVCSetting
    Public Property value As clsConfigurationWeb

    Sub New()
        _value = New clsConfigurationWeb
    End Sub

End Class

Public Class clsVCSettingCustomSettingsRPA
    Inherits clsVCSetting
    Public Property value As clsCustomSettingsRPA

    Sub New()
        _value = New clsCustomSettingsRPA
    End Sub

End Class




