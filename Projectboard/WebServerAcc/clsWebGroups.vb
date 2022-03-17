Public Class clsWebGroups
    Inherits clsWebOutput

    Public Property groups As List(Of clsGroup)
    Public Property count As Integer


    Sub New()
        _groups = New List(Of clsGroup)
        _count = 0
    End Sub
End Class
Public Class clsGroup
    Public Property _id As String
    Public Property name As String
    Public Property vcid As String
    'Public Property Global As String
    Public Property permission As Object
    Public Property users As List(Of clsUser)

End Class
