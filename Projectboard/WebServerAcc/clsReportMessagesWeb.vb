Imports ProjectBoardDefinitions
Public Class clsReportMessagesWeb

    Public Property listofRepMessages As List(Of clsReportMessage)

    Public Sub copyFrom(ByVal messageDef As SortedList(Of Integer, clsReportMessage))

        For Each kvp As KeyValuePair(Of Integer, clsReportMessage) In messageDef
            Me.listofRepMessages.Add(kvp.Value)
        Next

    End Sub

    Public Sub copyto(ByRef messageDef As SortedList(Of Integer, clsReportMessage))
        Dim i As Integer = 1
        For Each msgDef In Me.listofRepMessages
            messageDef.Add(i, msgDef)
            i += 1
        Next
    End Sub


    Sub New()
        _listofRepMessages = New List(Of clsReportMessage)
    End Sub


End Class
