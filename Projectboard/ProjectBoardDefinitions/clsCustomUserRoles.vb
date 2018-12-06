Public Class clsCustomUserRoles

    ' der Schlüssel setzt sich zusammen aus userID und customerUserRole in clsCustomUserRole
    Private _customUserRoles As SortedList(Of String, clsCustomUserRole)

    Public Sub New()
        _customUserRoles = New SortedList(Of String, clsCustomUserRole)
    End Sub

    ''' <summary>
    ''' gibt eine customerUserRole zurück 
    ''' </summary>
    ''' <param name="userID"></param>
    ''' <param name="customType"></param>
    ''' <returns></returns>
    Public ReadOnly Property getCustomUserRole(ByVal userID As String, ByVal customType As ptCustomUserProfils) As clsCustomUserRole
        Get

            Dim key As String = userID & customType.ToString.Trim
            Dim tmpResult As clsCustomUserRole = Nothing

            If _customUserRoles.ContainsKey(key) Then
                tmpResult = _customUserRoles.Item(key)
            End If

            getCustomUserRole = tmpResult

        End Get

    End Property

    ''' <summary>
    ''' Voraussetzung: hier wurde bereits gecheckt, ob der userName existiert und welche ID er hat ...
    ''' </summary>
    ''' <param name="userName"></param>
    ''' <param name="userID"></param>
    ''' <param name="customRoleType"></param>
    ''' <param name="specifics"></param>
    Public Sub addCustomUserRole(ByVal userName As String, userID As String, ByVal customRoleType As ptCustomUserProfils, ByVal specifics As Object)

        Dim key As String = userName.Trim & customRoleType.ToString.Trim
        If _customUserRoles.ContainsKey(key) Then
            ' Löschen ...
            _customUserRoles.Remove(key)
        End If

        ' jetzt ist sichergestellt, dass der key nicht mehr existiert ..
        Dim newCustomUserRole As New clsCustomUserRole
        With newCustomUserRole
            .userName = userName
            .userID = userID
            .customUserRole = customRoleType
            .specifics = specifics
        End With

        _customUserRoles.Add(key, newCustomUserRole)

    End Sub
End Class
