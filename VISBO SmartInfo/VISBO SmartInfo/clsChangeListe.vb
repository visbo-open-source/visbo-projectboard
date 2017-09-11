Public Class clsChangeListe
    ' enthält die Informationen zu den Änderungen der Elemente einer Seite, die sich geändert haben 
    Private _changeList As SortedList(Of String, clsChangeItem)

    ''' <summary>
    ''' gibt die Anzahl Einträge in der ChangeList zurück 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getChangeListCount() As Integer
        Get
            Dim tmpValue As Integer = 0
            If Not IsNothing(_changeList) Then
                tmpValue = _changeList.Count
            Else
                tmpValue = 0
            End If
            getChangeListCount = tmpValue
        End Get
    End Property

    ''' <summary>
    ''' setzt die ChangeList mit den Änderungs-Informationen zu TimeStamp- bzw. Variante zurück ...
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub clearChangeList()
        _changeList.Clear()
    End Sub

    ''' <summary>
    ''' fügt die Erläuterung zu dem smart Element mit shapeName hinzu
    ''' shapeName kann ein smart-Milestone, -Phase, -Chart, -Tabelle oder -Komponente sein   
    ''' </summary>
    ''' <param name="shapeName"></param>
    ''' <param name="explanation"></param>
    ''' <remarks></remarks>
    Public Sub addToChangeList(ByVal shapeName As String, ByVal explanation As clsChangeItem)
        ' in der aktuellen Demo Version beschränkt sich das auf das Anzeigen von Unterschieden in Meilensteinen , Phasen 
        If Not _changeList.ContainsKey(shapeName) Then
            _changeList.Add(shapeName, explanation)
        Else
            ' erstmal nichts weiter tun 
        End If
    End Sub

    ''' <summary>
    ''' gibt die index-te Erläuterung zurück; wird insbesondere in der Anzeige der Liste benötigt 
    ''' index läuft von 1 .. count 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getExplanationFromChangeList(ByVal index As Integer) As clsChangeItem
        Get
            If index >= 1 And index <= _changeList.Count Then
                getExplanationFromChangeList = _changeList.ElementAt(index - 1).Value
            Else
                getExplanationFromChangeList = New clsChangeItem
            End If
        End Get
    End Property

    ''' <summary>
    ''' gibt den Namen des Shapes zurück, das an der index-ten Stelle der changeList steht 
    ''' index darf von 1 .. count Werte annehmen
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getShapeNameFromChangeList(ByVal index As Integer) As String
        Get
            If index >= 1 And index <= _changeList.Count Then
                getShapeNameFromChangeList = _changeList.ElementAt(index - 1).Key
            Else
                getShapeNameFromChangeList = ""
            End If
        End Get
    End Property

    Public Sub New()
        _changeList = New SortedList(Of String, clsChangeItem)
    End Sub
End Class
