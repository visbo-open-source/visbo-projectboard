''' <summary>
''' Die Rollen müssen immer in der customization file in der ursprünglichen Reihenfolge aufgeführt sein; 
''' ein Name kann umbenannt werden , aber er darf auf keinen Fall an eine andere Psoiton verchoben werden 
''' neue Rolle müssen immer ans Ende gestellt werden - alte Rollen müssen immer mitgeschrieben werden ... 
''' </summary>
''' <remarks></remarks>
Public Class clsRollen


    Private _allRollen As SortedList(Of Integer, clsRollenDefinition)



    Public Sub Add(roledef As clsRollenDefinition)

        ' Änderung tk: umgestellt auf 
        If Not _allRollen.ContainsKey(roledef.UID) Then
            _allRollen.Add(roledef.UID, roledef)
        Else
            Throw New ArgumentException(roledef.UID.ToString & " existiert bereits")
        End If



    End Sub

    ''' <summary>
    ''' gibt in einer eindeutigen Liste die Namen aller vorkommenden SubRoles in einer Collection zurück 
    ''' </summary>
    ''' <param name="roleName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getSubRoleNamesOf(ByVal roleName As String) As Collection

        Get
            Dim tmpCollection As New Collection
            Dim tmpRole As clsRollenDefinition = Me.getRoledef(roleName)
            If Not IsNothing(tmpRole) Then

                Dim listOfSubRoles As SortedList(Of Integer, String) = tmpRole.getSubRoleIDs

                If Not IsNothing(listOfSubRoles) Then
                    Dim anzSubroles As Integer = listOfSubRoles.Count

                    If anzSubroles > 0 Then
                        For i As Integer = 1 To anzSubroles
                            Dim subRoleName As String = listOfSubRoles.ElementAt(i - 1).Value
                            If subRoleName <> roleName And Not tmpCollection.Contains(subRoleName) Then
                                tmpCollection.Add(subRoleName, subRoleName)
                            End If
                        Next
                    End If
                Else
                    ' nichts tun
                End If

            End If

            getSubRoleNamesOf = tmpCollection

        End Get
    End Property

    '
    '
    '
    Public ReadOnly Property Count() As Integer

        Get

            Count = _allRollen.Count

        End Get

    End Property

    ''' <summary>
    ''' prüft ob name in der Collection enthalten ist
    ''' </summary>
    ''' <param name="name">Typ String</param>
    ''' <value></value>
    ''' <returns>wahr, wenn name enthalten ist; falsch, sonst</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property containsName(name As String) As Boolean
        Get
            Dim found As Boolean = False
            Dim ix As Integer = 0
            Do While ix <= _allRollen.Count - 1 And Not found
                If _allRollen.ElementAt(ix).Value.name = name Then
                    found = True
                Else
                    ix = ix + 1
                End If
            Loop
            containsName = found
        End Get
    End Property


    ''' <summary>
    ''' gibt die Rollen-Definition mit angegebenem Namen zurück 
    ''' </summary>
    ''' <param name="myitem"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoledef(ByVal myitem As String) As clsRollenDefinition

        Get
            Dim tmpValue As clsRollenDefinition = Nothing

            Dim found As Boolean = False
            Dim ix As Integer = 0

            Do While ix <= _allRollen.Count - 1 And Not found
                If _allRollen.ElementAt(ix).Value.name = myitem Then
                    found = True
                    tmpValue = _allRollen.ElementAt(ix).Value
                Else
                    ix = ix + 1
                End If
            Loop

            getRoledef = tmpValue


        End Get

    End Property

    ''' <summary>
    ''' 1 gibt das erste Element zurück, AnzahlItems das letzte 
    ''' </summary>
    ''' <param name="myitem"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRoledef(ByVal myitem As Integer) As clsRollenDefinition

        Get


            If myitem > 0 And myitem <= _allRollen.Count Then
                getRoledef = _allRollen.ElementAt(myitem - 1).Value
            Else
                getRoledef = Nothing
            End If


        End Get

    End Property

    Public Sub New()

        _allRollen = New SortedList(Of Integer, clsRollenDefinition)

    End Sub

End Class
