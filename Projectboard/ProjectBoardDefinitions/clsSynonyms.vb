''' <summary>
''' nimmt die Synonyme für Meilensteine bzw Phasen auf 
''' </summary>
''' <remarks></remarks>
Public Class clsSynonyms

    Private _synkwPair As SortedList(Of String, kwInfo)
    Private _kwsynList As SortedList(Of String, SortedList(Of String, kwInfo))

    Private Class kwInfo
        Friend Property keyword As String
        Friend Property offset As Integer
    End Class

    ''' <summary>
    ''' baut beide Listen auf: 
    ''' 1. die Liste Synonym -> keyword und 
    ''' 2. Keyword -> Liste von Synonymen
    ''' </summary>
    ''' <param name="synonym">ist der Name des Meilensteins</param>
    ''' <param name="keyword">ist der Name des Kompaktplan Meilensteins </param>
    ''' <param name="offset">ist der Start-Offset in Tagen: offset = date(synonym) - date(keyword)  </param>
    ''' <remarks></remarks>
    Public Sub add(ByVal synonym As String, ByVal keyword As String, ByVal offset As Integer)

        ' jetzt wird die Liste Synonym -> keyword, offset aufgebaut 
        If _synkwPair.ContainsKey(synonym) Then
            _synkwPair.Item(synonym).keyword = keyword
            _synkwPair.Item(synonym).offset = offset
        Else
            Dim tmpKWInfo As New kwInfo
            With tmpKWInfo
                .keyword = keyword
                .offset = offset
            End With

            _synkwPair.Add(synonym, tmpKWInfo)

        End If

        ' jetzt wird die Liste keyword -> Liste von Synonymen aufgebaut 
        If _kwsynList.ContainsKey(keyword) Then
            ' jetzt muss das Synonym in der schon vorhandenen Liste aufgebaut werden 
        Else
            Dim tmpSynInfo As New kwInfo

            With tmpSynInfo
                .keyword = synonym
                .offset = offset
            End With

        End If

    End Sub

End Class
