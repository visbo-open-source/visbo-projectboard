
Imports ProjectBoardDefinitions
Public Class clsVPf
    Public Property _id As String
    Public Property name As String
    Public Property vpid As String
    Public Property variantName As String
    Public Property timestamp As String
    Public Property updatedAt As String
    Public Property createdAt As String
    Public Property sortType As Integer
    Public Property sortList As List(Of String)

    Public Property allItems As List(Of clsVPfItem)


    Sub New()
        _id = ""
        _name = "not named"
        _vpid = "not yet defined"
        _variantName = ""
        _timestamp = Date.MinValue.ToString
        _updatedAt = Date.MinValue.ToString
        _createdAt = Date.MinValue.ToString
        _sortType = 1
        _sortList = New List(Of String)
        _allItems = New List(Of clsVPfItem)
    End Sub

End Class
