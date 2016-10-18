Public Class clsCustomField

    Private _uid As Integer
    Private _value As Object

    Public Property uid As Integer
        Get
            uid = _uid
        End Get
        Set(value As Integer)
            _uid = value
        End Set
    End Property

    Public Property wert As Object
        Get
            wert = _value
        End Get
        Set(value As Object)
            _value = value
        End Set
    End Property

    Public Sub New()
        _uid = -1
        _value = Nothing
    End Sub

    ''' <summary>
    ''' value wird als object übergeben, weil es String, Double oder boolean sein kann 
    ''' über uid kann aus der customfielddefinitions herausgefunden werden, was für ein Typ es ist
    ''' </summary>
    ''' <param name="uid"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal uid As Integer, ByVal value As Object)

        _uid = uid
        _value = value

    End Sub

End Class
