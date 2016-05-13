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
            Dim hstr() As String
            Dim hmsg As String
            Dim ergmsg As String = ""
            Dim i As Integer = 0
            If nr > 0 And _allReportMsg.Count >= nr Then

                hmsg = _allReportMsg.Item(nr)
                hstr = Split(hmsg, "& vblf &", -1)
                While i < hstr.Length
                    If i = 0 Then
                        ergmsg = hstr(i)
                    Else
                        ergmsg = ergmsg & vbLf & hstr(i)
                    End If
                    i = i + 1
                End While

                getmsg = ergmsg
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
