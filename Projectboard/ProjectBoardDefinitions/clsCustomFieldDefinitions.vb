Public Class clsCustomFieldDefinitions

    Private listOfDefinitions As SortedList(Of Integer, clsCustomFieldDefinition)


    ''' <summary>
    ''' fügt der Liste an Definitionen ein neues Element hinzu 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="type"></param>
    ''' <remarks></remarks>
    Public Sub add(ByVal name As String, ByVal type As Integer, ByVal uid As Integer)

        Dim tmpCF As clsCustomFieldDefinition

        Try
            tmpCF = New clsCustomFieldDefinition(name, type, uid)

            If Not listOfDefinitions.ContainsKey(uid) Then
                listOfDefinitions.Add(uid, tmpCF)
            Else
                With listOfDefinitions.Item(uid)
                    If name = .name And type = .type Then
                        ' nichts tun, Definition existiert in der exakten Form  ja schon  
                    Else
                        Throw New ArgumentException("Custom Field Definition existiert schon, aber mit anderen Werten:" & _
                                                     .name & ", " & .type.ToString)
                    End If
                End With

            End If

        Catch ex As Exception
            Throw New ArgumentException(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' gibt die sortierte Liste der Custom Fields zurück
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property liste As SortedList(Of Integer, clsCustomFieldDefinition)
        Get
            liste = listOfDefinitions
        End Get
    End Property


    ''' <summary>
    ''' prüft, ob der es ein Custom Field mit der angegebenen uid gibt
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property contains(ByVal uid As Integer) As Boolean
        Get
            contains = listOfDefinitions.ContainsKey(uid)
        End Get
    End Property

    ''' <summary>
    ''' returns true if there is a customField defined with cfName
    ''' comparison is all on lowercase
    ''' </summary>
    ''' <param name="cfName"></param>
    ''' <returns></returns>
    Public ReadOnly Property containsName(ByVal cfName As String) As Boolean
        Get
            Dim found As Boolean = False
            For Each kvp As KeyValuePair(Of Integer, clsCustomFieldDefinition) In listOfDefinitions
                If kvp.Value.name.ToLower = cfName.ToLower Then
                    found = True
                    Exit For
                End If
            Next

            containsName = found
        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl an Custom Fields zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property count As Integer
        Get
            count = listOfDefinitions.Count
        End Get
    End Property

    ''' <summary>
    ''' gibt den NAmen einer UID zurück, Nothing wenn uid nicht existiert
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getName(ByVal uid As Integer) As String
        Get
            If listOfDefinitions.ContainsKey(uid) Then
                getName = listOfDefinitions.Item(uid).name
            Else
                getName = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt den Type einer Uid zurück, Nothing, wenn uid nicht existiert 
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTyp(ByVal uid As Integer) As Integer
        Get
            If listOfDefinitions.ContainsKey(uid) Then
                getTyp = listOfDefinitions.Item(uid).type
            Else
                getTyp = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt den Type eines Customfields Namen zurück, Nothing, wenn NAme nicht existiert
    ''' wenn zwei mit gleichen Namen und unterschiedlichem Typ existieren, wird immer der zuerst 
    ''' gefundene zurückgegegen  
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTyp(ByVal name As String) As Integer
        Get
            Dim tmpValue As Integer = Nothing
            ' suche das erste Auftreten, egal welchen Typs ...
            For Each kvp As KeyValuePair(Of Integer, clsCustomFieldDefinition) In listOfDefinitions
                If name = kvp.Value.name Then
                    tmpValue = kvp.Value.type
                    Exit For
                End If
            Next

            getTyp = tmpValue
        End Get
    End Property

    ''' <summary>
    ''' gibt die uid des CustomFields zurück; -1 wenn nicht existent 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getUid(ByVal name As String, Optional type As Integer = -99) As Integer
        Get

            Dim tmpValue As Integer = -1 ' bei diese Wert wurde es nicht gefunden
            If type = -99 Then
                ' suche das erste Auftreten, egal welchen Typs ...
                For Each kvp As KeyValuePair(Of Integer, clsCustomFieldDefinition) In listOfDefinitions
                    If name = kvp.Value.name Then
                        tmpValue = kvp.Key
                        Exit For
                    End If
                Next
            Else
                ' suche die uid, wo name und type übereinstimmen 
                For Each kvp As KeyValuePair(Of Integer, clsCustomFieldDefinition) In listOfDefinitions
                    If name = kvp.Value.name And type = kvp.Value.type Then
                        tmpValue = kvp.Key
                        Exit For
                    End If
                Next
            End If

            getUid = tmpValue

        End Get
    End Property

    ''' <summary>
    ''' gibt das Element an Stelle item zurück; item kann Werte von 1 .. count annehmen
    ''' </summary>
    ''' <param name="item"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getDef(ByVal item As Integer) As clsCustomFieldDefinition
        Get
            If item >= 1 And item <= listOfDefinitions.Count Then
                getDef = listOfDefinitions.ElementAt(item - 1).Value
            Else
                getDef = Nothing
            End If
        End Get
    End Property


    Sub New()
        listOfDefinitions = New SortedList(Of Integer, clsCustomFieldDefinition)
    End Sub

End Class
