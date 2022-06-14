Public Class clsPortfolioDefinitions
    Private _listofPortfolioDefinitions As SortedList(Of String, List(Of String))

    Public Function contains(ByVal key As String) As Boolean
        contains = _listofPortfolioDefinitions.ContainsKey(key)
    End Function

    Public Sub addPortfolio(ByVal key As String, ByVal portfolioDefinition As List(Of String))

        Try
            If Not IsNothing(portfolioDefinition) Then

                If _listofPortfolioDefinitions.ContainsKey(key) Then
                    _listofPortfolioDefinitions.Remove(key)
                End If

                _listofPortfolioDefinitions.Add(key, portfolioDefinition)
            End If
        Catch ex As Exception

        End Try


    End Sub

    Public Sub removePortfolio(ByVal key As String)

        Try
            If _listofPortfolioDefinitions.ContainsKey(key) Then
                _listofPortfolioDefinitions.Remove(key)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public ReadOnly Property portfolioListe As SortedList(Of String, List(Of String))
        Get
            portfolioListe = _listofPortfolioDefinitions
        End Get
    End Property

    Public Sub New()

        _listofPortfolioDefinitions = New SortedList(Of String, List(Of String))

    End Sub

End Class
