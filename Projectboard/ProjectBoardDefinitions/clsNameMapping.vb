''' <summary>
''' Klasse, über die das Namens-Mapping der Phasen und Meilensteine gemacht wird
''' </summary>
''' <remarks></remarks>
Public Class clsNameMapping

    Private synonyms As SortedList(Of String, String)
    Private coreNames As SortedList(Of String, String)
    Private namesToComplement As SortedList(Of String, String)
    Private hrchy(10) As String
    Private hrchyIndex As Integer

    ''' <summary>
    ''' ergänzt ein Synonym / Standard-Wert Paar  
    ''' </summary>
    ''' <param name="synonym">synonym</param>
    ''' <param name="stdName">Standard-Name, auf den das Synonym abgebildet wird</param>
    ''' <remarks></remarks>
    Public Sub addSynonym(ByVal synonym As String, ByVal stdName As String)


        If Not IsNothing(synonym) And Not IsNothing(stdName) Then

            Dim key As String = synonym.Trim
            Dim value As String = stdName.Trim

            If key.Length > 0 And value.Length > 0 Then

                If synonyms.ContainsKey(key) Then
                    If synonyms(key) = value Then
                        ' alles ok - nichts tun
                    Else
                        Throw New ArgumentException("ein Synonym kann nicht zwei Std-Namen haben!" & vbLf & _
                                                     key & "-> " & synonyms(key) & " / " & value)
                    End If
                Else
                    synonyms.Add(key, value)
                End If
            Else
                Throw New ArgumentException("Synonym und Wert müssen beide ein nicht-leerer String sein")
            End If


        Else
            Throw New ArgumentException("weder Synonym noch Wert dürfen NULL sein")
        End If



    End Sub

    ''' <summary>
    ''' ergänzt Name in der Liste der Core-Names  
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub addcoreName(ByVal name As String)


        If Not IsNothing(name) Then

            Dim key As String = name.Trim

            If key.Length > 0 Then

                If coreNames.ContainsKey(key) Then

                    ' alles ok - nichts tun

                Else
                    coreNames.Add(key, key)
                End If

            Else
                Throw New ArgumentException("Std-Name darf kein leerer String sein")
            End If


        Else
            Throw New ArgumentException("Std-Name darf nicht NULL sein! ")
        End If



    End Sub

    ''' <summary>
    ''' ergänzt Name zu der Liste der zu komplementierenden Namen 
    ''' Die Namen werden durch die letzte Phase komplementiert/ergänzt
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub addNameToComplement(ByVal name As String)


        If Not IsNothing(name) Then

            Dim key As String = name.Trim

            If key.Length > 0 Then

                If namesToComplement.ContainsKey(key) Then

                    ' alles ok - nichts tun

                Else
                    namesToComplement.Add(key, key)
                End If

            Else
                Throw New ArgumentException("zu komplementierender Name darf kein leerer String sein")
            End If


        Else
            Throw New ArgumentException("zu komplementierender Name darf nicht NULL sein! ")
        End If



    End Sub

    Public ReadOnly Property mapToRealName(ByVal parentPhaseName As String, ByVal itemName As String) As String
        Get

            ' erster Check: kommt itemName in Synonym Liste vor ? 
            If Me.synonyms.ContainsKey(itemName.Trim) Then
                itemName = Me.synonyms(itemName).Trim
            Else
                itemName = itemName.Trim
                ' check auf reduction Names
                For Each kvp As KeyValuePair(Of String, String) In coreNames
                    If itemName.Contains(kvp.Key) Then
                        itemName = kvp.Key.Trim
                        Exit For
                    End If
                Next

            End If

            ' check jetzt auf Hierarchie Names
            If Me.namesToComplement.ContainsKey(itemName) Then
                itemName = parentPhaseName & "#" & itemName
            End If

            mapToRealName = itemName

        End Get
    End Property

    Public Sub New()

        synonyms = New SortedList(Of String, String)
        coreNames = New SortedList(Of String, String)
        namesToComplement = New SortedList(Of String, String)
        hrchyIndex = -1

    End Sub
End Class
