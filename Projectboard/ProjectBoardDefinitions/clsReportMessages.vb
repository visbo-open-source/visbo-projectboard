Imports System.Xml
Imports System.Xml.Schema


<Serializable()>
Public Class clsReportMessages

    Private _allReportMsg As SortedList(Of Integer, String)



    Public ReadOnly Property Liste As SortedList(Of Integer, String)

        Get
            Liste = _allReportMsg
        End Get

    End Property

    Public ReadOnly Property getmsg(ByVal nr As Integer) As String

        Get

            If nr > 0 And _allReportMsg.Count > nr Then
                getmsg = _allReportMsg.Item(nr)
            Else
                getmsg = ""
            End If

        End Get

    End Property
    Public Sub clear()
        _allReportMsg.Clear()
    End Sub

    Public Sub New()
        _allReportMsg = New SortedList(Of Integer, String)

    End Sub


End Class
