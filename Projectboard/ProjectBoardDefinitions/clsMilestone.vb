Public Class clsMilestone

    Private listOfErgebnisse As List(Of clsResult)

    Friend Property datum As Date
    Friend Property name As String
    Friend Property offset As Long
    Friend Property offsetType As String

    Friend Sub addErgebnis(ByVal ergebnis As clsResult)

        Try
            listOfErgebnisse.Add(ergebnis)
        Catch ex As Exception

        End Try

    End Sub

    Friend Sub removeErgebnis(ByVal ergebnis As clsResult)

        Try
            listOfErgebnisse.Remove(ergebnis)
        Catch ex As Exception

        End Try
    End Sub

    Friend ReadOnly Property containsErgebnis(ByVal ergebnis As clsResult) As Boolean

        Get
            containsErgebnis = listOfErgebnisse.Contains(ergebnis)
        End Get

    End Property

    Friend ReadOnly Property getErgebnis(ByVal index As Integer) As clsResult
        Get

            Try
                getErgebnis = listOfErgebnisse.ElementAt(index)
            Catch ex As Exception
                getErgebnis = Nothing
            End Try

        End Get
    End Property


    Sub New()
        listOfErgebnisse = New List(Of clsResult)
    End Sub

End Class
