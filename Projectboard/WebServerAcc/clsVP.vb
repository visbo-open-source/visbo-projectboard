Public Class clsVP
    Public Property _id As String
    Public Property name As String
    Public Property kundennummer As String
    Public Property vcid As String
    Public Property vpvCount As Integer
    Public Property vpType As Integer
    Public Property vpPublic As Boolean

    ' ur: 20210422 properties like businessUnit, strategikFit, risk now moved to vp
    Public Property customFieldString As List(Of clsCustomFieldStr)
    Public Property customFieldDouble As List(Of clsCustomFieldDbl)


    'Public Property users As List(Of clsUser)
    Public Property updatedAt As String
    Public Property createdAt As String
    Public Property lock As List(Of clsVPLock)
    Public Property [Variant] As List(Of clsVPvariant)

    Sub New()
        _id = ""
        _name = "not named"
        _kundennummer = ""
        _vcid = "not yet defined"
        _vpvCount = 0
        _vpType = 0
        _vpPublic = False
        _customFieldString = New List(Of clsCustomFieldStr)
        _customFieldDouble = New List(Of clsCustomFieldDbl)
        '_users = New List(Of clsUser)
        _updatedAt = Date.MinValue.ToString
        _createdAt = Date.MinValue.ToString
        _lock = New List(Of clsVPLock)
        _Variant = New List(Of clsVPvariant)
    End Sub
End Class

' VP properties: _businessUnit', '_risk', '_strategicFit', '_customerID'

' ur: 20210422 vp-Properties businessUnit definitions
Public Class clsCustomFieldStr
    Public Property name As String
    Public Property localName As String
    Public Property value As String
    Public Property type As String
    Sub New()
        _name = ""
        _localName = ""
        _value = ""
        _type = "System"
    End Sub
End Class
' ur: 20210422 vp-Properties strategikFit, risk definitions
Public Class clsCustomFieldDbl
    Public Property name As String
    Public Property localName As String
    Public Property value As Double
    Public Property type As String
    Sub New()
        _name = ""
        _localName = ""
        _value = 0
        _type = "System"
    End Sub
End Class
