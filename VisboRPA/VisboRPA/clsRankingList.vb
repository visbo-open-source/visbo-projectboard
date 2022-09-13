Imports ProjectBoardDefinitions
Public Class clsRankingList

    Private _rankingList As SortedList(Of Integer, clsRankingParameters)

    Public Property liste() As SortedList(Of Integer, clsRankingParameters)
        Get
            liste = _rankingList
        End Get
        Set(value As SortedList(Of Integer, clsRankingParameters))
            If Not IsNothing(value) Then
                _rankingList = value
            End If
        End Set
    End Property

    Public ReadOnly Property containsPName(ByVal pname As String) As Boolean

        Get
            Dim result As Boolean = False
            For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In _rankingList
                result = (kvp.Value.projectName = pname)

                If result = True Then
                    Exit For
                End If

            Next

            containsPName = result
        End Get

    End Property

    Public ReadOnly Property containsPVName(ByVal pvname As String) As Boolean

        Get
            Dim result As Boolean = False
            For Each kvp As KeyValuePair(Of Integer, clsRankingParameters) In _rankingList

                Dim myPvName As String = calcProjektKey(kvp.Value.projectName, kvp.Value.projectVariantName)
                result = (myPvName = pvname)

                If result = True Then
                    Exit For
                End If

            Next

            containsPVName = result
        End Get

    End Property

    Sub New()
        _rankingList = New SortedList(Of Integer, clsRankingParameters)
    End Sub


End Class
