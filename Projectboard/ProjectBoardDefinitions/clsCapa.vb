''' <summary>
''' represents the capa per month of one role in one year
''' </summary>
Public Class clsCapa
    Public Property _id As String
    Public Property vcid As String
    Public Property roleID As String
    Public Property startOfYear As Date
    Public Property capaPerMonth As List(Of Double)

    Public Function isIdenticalTo(ByVal vglCapa As clsCapa) As Boolean

        Dim tmpResult As Boolean = False

        If Not IsNothing(vglCapa) Then
            tmpResult = (roleID = vglCapa.roleID) And
                    (DateDiff(DateInterval.Month, startOfYear, vglCapa.startOfYear) = 0) And
                    (capaPerMonth.Count = vglCapa.capaPerMonth.Count)

            If tmpResult Then
                ' it is now sure that both lists are having the same size ...
                Try
                    For i As Integer = 0 To capaPerMonth.Count - 1
                        tmpResult = tmpResult And (capaPerMonth.Item(i) = vglCapa.capaPerMonth.Item(i))
                    Next
                Catch ex As Exception
                    tmpResult = False
                End Try

            End If
        End If

        isIdenticalTo = tmpResult
    End Function

    Public Sub New()
        _id = ""
        _vcid = ""
        _roleID = ""
        _startOfYear = Date.Now
        _capaPerMonth = New List(Of Double)
    End Sub
End Class
