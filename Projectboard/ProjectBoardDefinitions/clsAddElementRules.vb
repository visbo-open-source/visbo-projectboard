Public Class clsAddElementRules

    Private _name As String
    Private _regelliste As SortedList(Of String, clsAddElementRule)

    Public Property name As String
        Get
            name = _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    ''' <summary>
    ''' fügt der Regel-Liste eine neue Regel hinzu
    ''' Dabei wird der Schlüssel folgendermaßen bestimmt: Name des neuen Elements + lfdNr, 
    ''' lfdNr beginnt mit 1
    ''' </summary>
    ''' <param name="newRule"></param>
    ''' <remarks></remarks>
    Public Sub addRule(ByVal newRule As clsAddElementRule)
        Dim index As Integer = 1
        Dim key As String = newRule.newElemName & index.ToString("00#")

        While _regelliste.ContainsKey(key)
            index = index + 1
            key = newRule.newElemName & index.ToString("00#")
        End While

        ' jetzt ist key nicht mehr enthalten ...
        _regelliste.Add(key, newRule)

    End Sub

    ''' <summary>
    ''' gibt die Regel zurück, die in der sortierten Liste an der Position index steht
    ''' index kann Werte zwischen 1 und Anzahl_Elemente annehmen  
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRule(ByVal index As Integer) As clsAddElementRule
        Get
            If index >= 1 And index <= _regelliste.Count Then
                getRule = _regelliste.ElementAt(index - 1).Value
            Else
                getRule = Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt die x.te Regel für das Element name zurück
    ''' gibt Nothing zurück, wenn es keine x.te Regel für Element name gibt 
    ''' Nutzt die Sortierung aus  
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="lfdNr"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getRule(ByVal name As String, ByVal lfdNr As Integer) As clsAddElementRule
        Get
            Dim key As String = name & "001"
            Dim position As Integer

            If _regelliste.ContainsKey(key) Then
                position = _regelliste.IndexOfKey(key) + lfdNr - 1
                If position >= 0 And position <= _regelliste.Count - 1 Then
                    If _regelliste.ElementAt(position).Value.newElemName = name Then
                        getRule = _regelliste.ElementAt(position).Value
                    Else
                        getRule = Nothing
                    End If
                Else
                    getRule = Nothing
                End If
            Else
                getRule = Nothing
            End If


        End Get
    End Property


    ''' <summary>
    ''' gibt die Anzahl Regeln zurück, die es für das Element name gibt 
    ''' 0 , wenn es keine einzige Regel gibt 
    ''' </summary>
    ''' <param name="name"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getAnzahlRulesForElem(ByVal name As String) As Integer
        Get
            Dim anzahl As Integer = 0
            Dim key As String = name & "001"
            Dim position As Integer

            If _regelliste.ContainsKey(key) Then
                position = _regelliste.IndexOfKey(key) + 1
                anzahl = 1
                Do While position <= _regelliste.Count - 1
                    If _regelliste.ElementAt(position).Value.newElemName = name Then
                        anzahl = anzahl + 1
                        position = position + 1
                    Else
                        ' jetzt abbrechen 
                        position = _regelliste.Count
                    End If
                Loop
            End If

            getAnzahlRulesForElem = anzahl

        End Get
    End Property

    ''' <summary>
    ''' gibt die Anzahl Elemente in der Regelliste zuürck
    ''' kann Mehrfach-Nennungen enthalten , deshalb ist die Anzahl der newElem oft kleiner 
    ''' als die Gesamtzahl  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property count() As Integer
        Get
            count = _regelliste.Count
        End Get
    End Property

    Public Sub New()
        _regelliste = New SortedList(Of String, clsAddElementRule)
    End Sub
End Class
