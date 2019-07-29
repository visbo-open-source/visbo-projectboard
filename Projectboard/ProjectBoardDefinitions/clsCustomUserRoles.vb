Public Class clsCustomUserRoles

    ' der Schlüssel setzt sich zusammen aus name, customerUserRole, specifics in clsCustomUserRole
    Private _customUserRoles As SortedList(Of String, clsCustomUserRole)

    Public Sub New()
        _customUserRoles = New SortedList(Of String, clsCustomUserRole)
    End Sub

    ''' <summary>
    ''' gibt Zugriff auf die sortierte Liste 
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property liste() As SortedList(Of String, clsCustomUserRole)
        Get
            liste = _customUserRoles
        End Get
    End Property


    ''' <summary>
    ''' liefert das Element an der Stelle index. Index kann von 9 bis count-1 gehen
    ''' </summary>
    ''' <param name="index"></param>
    ''' <returns></returns>
    Public ReadOnly Property elementAt(ByVal index As Integer) As clsCustomUserRole
        Get
            If index >= 0 And index < _customUserRoles.Count Then
                elementAt = _customUserRoles.ElementAt(index).Value
            Else
                elementAt = Nothing
            End If
        End Get
    End Property

    Public ReadOnly Property count() As Integer
        Get
            count = _customUserRoles.Count
        End Get
    End Property
    ''' <summary>
    ''' gibt eine Collection of clsCustomUserRole zurück, die zu dem User mit Name userNAme gehören  
    ''' </summary>
    ''' <param name="userName"></param>
    ''' <returns></returns>
    Public ReadOnly Property getCustomUserRoles(ByVal userName As String) As Collection
        Get

            Dim tmpCollection As New Collection

            For Each kvp As KeyValuePair(Of String, clsCustomUserRole) In _customUserRoles

                ' diese String-Vergleich soll immer für Security-Themes hergenommen werden. 
                ' hier ist ignoreCase = true

                If String.Compare(kvp.Value.userName, userName, True) = 0 Then
                    tmpCollection.Add(kvp.Value)
                End If

            Next

            getCustomUserRoles = tmpCollection

        End Get

    End Property

    ''' <summary>
    ''' Voraussetzung: hier wurde bereits gecheckt, ob der userName existiert und welche ID er hat ...
    ''' ein Add funktioniert nur, wenn es nicht konkurrierende Rollen gibt - z.B ist das bei Ressourcen-MAnagern und Team-Managern mit der gleichen Specifics der Fall 
    ''' </summary>
    ''' <param name="userName"></param>
    ''' <param name="userID"></param>
    ''' <param name="customRoleType"></param>
    ''' <param name="specifics"></param>
    Public Sub addCustomUserRole(ByVal userName As String, userID As String, ByVal customRoleType As ptCustomUserRoles, ByVal specifics As String)


        Dim key As String = calcCurKey(userName, customRoleType, specifics)
        Dim chckkey As String = ""

        ' bestimme den sich ausschließenden Key ... 
        If customRoleType = ptCustomUserRoles.RessourceManager Then
            chckkey = calcCurKey(userName, ptCustomUserRoles.TeamManager, specifics)

        ElseIf customRoleType = ptCustomUserRoles.TeamManager Then
            chckkey = calcCurKey(userName, ptCustomUserRoles.RessourceManager, specifics)

        End If

        If _customUserRoles.ContainsKey(key) Then
            ' nichts tun, ist ja schon drin ... 
        Else
            If chckkey <> "" Then
                If _customUserRoles.ContainsKey(chckkey) Then
                    ' nicht erlaubt .. 
                    Throw New ArgumentException("Team-Manager and Ressourcen Manager for same department are not allowed ..")
                End If
            End If
            ' jetzt ist sichergestellt, dass der key noch nicht  existiert ..
            Dim newCustomUserRole As New clsCustomUserRole
            With newCustomUserRole
                .userName = userName
                .userID = userID
                .customUserRole = customRoleType
                .specifics = specifics
            End With

            _customUserRoles.Add(key, newCustomUserRole)
        End If

    End Sub

    Public Sub addCustomUserRole(ByVal curole As clsCustomUserRole)

        Dim key As String = calcCurKey(curole.userName, curole.customUserRole, curole.specifics)
        If _customUserRoles.ContainsKey(key) Then
            ' nichts tun, ist ja schon drin ... 
        Else
            _customUserRoles.Add(key, curole)
        End If

    End Sub
End Class
