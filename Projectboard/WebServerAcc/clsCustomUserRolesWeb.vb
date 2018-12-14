Imports ProjectBoardDefinitions
Public Class clsCustomUserRolesWeb
    ' der Schlüssel für das Kopieren in die interne Client datenstruktur setzt sich zusammen aus userID, customerUserRole, specifics in clsCustomUserRoleWeb
    Public Property customUserRoles As List(Of clsCustomUserRoleWeb)

    Public Sub New()
        _customUserRoles = New List(Of clsCustomUserRoleWeb)
    End Sub



    ''' <summary>
    ''' kopiert den Inhalt der  Struktur für den Server in die interne Struktur (client)
    ''' </summary>
    ''' <param name="curoles"></param>
    Public Sub copyTo(ByRef curoles As clsCustomUserRoles)

        For Each curole As clsCustomUserRoleWeb In Me.customUserRoles

            Dim internCurole As New clsCustomUserRole
            curole.copyTo(internCurole)
            curoles.addCustomUserRole(internCurole.userName,
                                      internCurole.userID,
                                      internCurole.customUserRole,
                                      internCurole.specifics)

        Next

    End Sub

    ''' <summary>
    ''' kopiert den Inhalt der CustomUserRole (client) in die Struktur für den Server
    ''' </summary>
    ''' <param name="curoles"></param>
    Public Sub copyFrom(ByVal curoles As clsCustomUserRoles)

        For Each kvp As KeyValuePair(Of String, clsCustomUserRole) In curoles.liste

            Dim webCurole As New clsCustomUserRoleWeb
            webCurole.copyFrom(kvp.Value)
            Me.customUserRoles.Add(webCurole)

        Next

    End Sub
End Class
