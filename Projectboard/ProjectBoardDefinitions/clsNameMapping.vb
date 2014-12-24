''' <summary>
''' Klasse, über die das Namens-Mapping der Phasen und Meilensteine gemacht wird
''' </summary>
''' <remarks></remarks>
Public Class clsNameMapping

    Private synonyms As SortedList(Of String, String)
    Private regExpressionNames As SortedList(Of String, String)
    Private ignoreNames As SortedList(Of String, String)
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
    ''' ergänzt Eintrag in der Liste der Regular-Expressions und ihrer Mappings 
    ''' </summary>
    ''' <param name="regExpress"></param>
    ''' <param name="stdName"></param>
    ''' <remarks></remarks>
    Public Sub addRegExpressName(ByVal regExpress As String, ByVal stdName As String)

        If Not IsNothing(regExpress) And Not IsNothing(stdName) Then

            Dim key As String = regExpress.Trim
            Dim value As String = stdName.Trim

            If key.Length > 0 And value.Length > 0 Then

                If regExpressionNames.ContainsKey(key) Then
                    If regExpressionNames(key) = value Then
                        ' alles ok - nichts tun
                    Else
                        Throw New ArgumentException("eine Regular Expression kann nicht zwei Std-Namen haben!" & vbLf & _
                                                     key & "-> " & regExpressionNames(key) & " / " & value)
                    End If
                Else
                    regExpressionNames.Add(key, value)
                End If
            Else
                Throw New ArgumentException("Regular Expression und Wert müssen beide ein nicht-leerer String sein")
            End If


        Else
            Throw New ArgumentException("weder die Regular Expression noch der Standard-Name dürfen NULL sein")
        End If


    End Sub

    ' ''' <summary>
    ' ''' ergänzt Name in der Liste der Core-Names  
    ' ''' </summary>
    ' ''' <param name="name"></param>
    ' ''' <remarks></remarks>
    'Public Sub addcoreName(ByVal name As String)


    '    If Not IsNothing(name) Then

    '        Dim key As String = name.Trim

    '        If key.Length > 0 Then

    '            If regExpressionNames.ContainsKey(key) Then

    '                ' alles ok - nichts tun

    '            Else
    '                regExpressionNames.Add(key, key)
    '            End If

    '        Else
    '            Throw New ArgumentException("Std-Name darf kein leerer String sein")
    '        End If


    '    Else
    '        Throw New ArgumentException("Std-Name darf nicht NULL sein! ")
    '    End If



    'End Sub

    ''' <summary>
    ''' ergänzt Name in der Ignore Liste
    ''' </summary>
    ''' <param name="name"></param>
    ''' <remarks></remarks>
    Public Sub addIgnoreName(ByVal name As String)


        If Not IsNothing(name) Then

            Dim key As String = name.Trim

            If key.Length > 0 Then

                If ignoreNames.ContainsKey(key) Then

                    ' alles ok - nichts tun

                Else
                    ignoreNames.Add(key, key)
                End If

            Else
                Throw New ArgumentException("Ignore Name darf kein leerer String sein")
            End If


        Else
            Throw New ArgumentException("Ignore Name darf kein leerer String sein! ")
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

    ''' <summary>
    ''' prüft, ob der angegebene Name im Ignore Verzeichnis steht 
    ''' der leere Name liefert true zurück
    ''' </summary>
    ''' <param name="itemName"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property tobeIgnored(ByVal itemName As String) As Boolean
        Get

            Dim ergebnis As Boolean
            itemName = itemName.Trim
            If itemName.Length > 0 Then
                If ignoreNames.ContainsKey(itemName) Then
                    ergebnis = True
                Else
                    ergebnis = False
                End If
            Else
                ergebnis = True
            End If

            tobeIgnored = ergebnis

        End Get
    End Property

    Public ReadOnly Property mapToRealName(ByVal parentPhaseName As String, ByVal itemName As String) As String
        Get


            itemName = itemName.Trim

            ' Test Bedingung
            If itemName.Contains("A-TS") Then
                Dim blabla = "jetzt"
            End If

            Dim realName As String = itemName

            ' erster Check: kommt itemName in Synonym Liste vor ? 
            If Me.synonyms.ContainsKey(itemName) Then
                realName = Me.synonyms(itemName).Trim
            Else

                ' check auf regular Expressions
                For Each kvp As KeyValuePair(Of String, String) In regExpressionNames

                    If regExpressionMatch(kvp.Key, itemName) Then
                        realName = kvp.Value.Trim
                        Exit For
                    End If

                Next

            End If

            ' check jetzt auf Hierarchie Names
            If Me.namesToComplement.ContainsKey(realName) Then
                realName = parentPhaseName & "+" & realName
            End If

            mapToRealName = realName

        End Get
    End Property

    ''' <summary>
    ''' prüft , ob itemName eine Ausprägung der regulären Expression ist
    ''' </summary>
    ''' <param name="regExpression"></param>
    ''' <param name="itemName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function regExpressionMatch(ByVal regExpression As String, ByVal itemName As String) As Boolean
        Dim tmpStr() As String
        Dim matches As Boolean = False
        Dim ix As Integer = 0

        regExpression = regExpression.Trim
        itemName = itemName.Trim

        tmpStr = regExpression.Split(New Char() {CChar("*")}, 30)

        If tmpStr.Length = 1 Then
            ' es kommt kein * vor
            If regExpression.Trim = itemName.Trim Then
                matches = True
            Else
                matches = False
            End If

        ElseIf tmpStr.Length = 2 Then
            ' es kommt ein * vor

            If regExpression.EndsWith("*") Then
                matches = itemName.StartsWith(tmpStr(0))
            ElseIf regExpression.StartsWith("*") Then
                matches = itemName.EndsWith(tmpStr(1))
            Else
                ' der * ist in der Mitte 
                matches = itemName.StartsWith(tmpStr(0)) And _
                    itemName.EndsWith(tmpStr(1))
            End If

        ElseIf tmpStr.Length > 2 Then

            Dim found As Boolean = False
            Dim lastPosition As Integer = -1
            Dim curPosition As Integer = -1
            matches = True

            If tmpStr(0).Length > 0 Then
                ' der itemname muss mit dem ersten Element starten 
                matches = itemName.StartsWith(tmpStr(0))
            End If

            If tmpStr(tmpStr.Length - 1).Length > 0 Then
                ' der itemname muss mit dem letzten Element enden 
                matches = matches And itemName.EndsWith(tmpStr(tmpStr.Length - 1))
            End If

            ix = 1

            Do While ix <= tmpStr.Length - 2 And matches

                If tmpStr(ix).Trim.Length = 0 Then
                    ix = ix + 1
                Else
                    curPosition = itemName.IndexOf(tmpStr(ix))

                    If curPosition >= 0 Then
                        Dim restName As String = ""
                        For k As Integer = curPosition + tmpStr(ix).Length To itemName.Length - 1
                            restName = restName & itemName.Chars(k)
                        Next
                        itemName = restName
                        ix = ix + 1
                    Else
                        matches = False
                    End If

                End If
            Loop



        End If

        regExpressionMatch = matches

    End Function

    Public Sub New()

        synonyms = New SortedList(Of String, String)
        regExpressionNames = New SortedList(Of String, String)
        namesToComplement = New SortedList(Of String, String)
        ignoreNames = New SortedList(Of String, String)
        hrchyIndex = -1

    End Sub
End Class
