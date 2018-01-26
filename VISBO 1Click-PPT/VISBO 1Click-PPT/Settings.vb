'Diese Klasse ermöglicht die Behandlung bestimmter Ereignisse der Einstellungsklasse:
' Das SettingChanging-Ereignis wird ausgelöst, bevor der Wert einer Einstellung geändert wird.
' Das PropertyChanged-Ereignis wird ausgelöst, nachdem der Wert einer Einstellung geändert wurde.
' Das SettingsLoaded-Ereignis wird ausgelöst, nachdem die Einstellungswerte geladen wurden.
' Das SettingsSaving-Ereignis wird ausgelöst, bevor die Einstellungswerte gespeichert werden.
Imports System.ComponentModel
Imports System.Configuration

Partial Friend NotInheritable Class MySettings


    Private Sub MySettings_SettingsLoaded(sender As Object, e As SettingsLoadedEventArgs) Handles Me.SettingsLoaded
        'Call MsgBox("hier loaded")
        'My.Settings.userNamePWD = ""

    End Sub

    Private Sub MySettings_SettingsSaving(sender As Object, e As CancelEventArgs) Handles Me.SettingsSaving
        'Call MsgBox("hier saving")
    End Sub
End Class
