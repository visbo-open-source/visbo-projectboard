Imports ProjectBoardDefinitions
Public Class clsAppearanceWeb

    Public Property listofAppearances As List(Of clsAppearance)

    Public Sub copyFrom(ByVal appearanceDef As SortedList(Of String, clsAppearance))

        For Each kvp As KeyValuePair(Of String, clsAppearance) In appearanceDef
            Me.listofAppearances.Add(kvp.Value)
        Next

    End Sub

    Public Sub copyto(ByRef appearanceDef As SortedList(Of String, clsAppearance))
        For Each appDef In Me.listofAppearances
            appearanceDef.Add(appDef.name, appDef)
        Next
    End Sub



    Public Sub New()
        _listofAppearances = New List(Of clsAppearance)
    End Sub



End Class
