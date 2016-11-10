Public Class clsCustomField

    Private _uid As Integer
    Private _value As Object

    ''' <summary>
    ''' gibt zurück, ob zwei CustomFields identisch sind oder nicht 
    ''' </summary>
    ''' <param name="vCustomField"></param>
    ''' <param name="type"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vCustomField As clsCustomField, ByVal type As Integer) As Boolean
        Get
            Dim tmpResult As Boolean

            With vCustomField
                If type = ptCustomFields.Str Then
                    If _uid = .uid And CStr(_value) = CStr(.wert) Then
                        tmpResult = True
                    Else
                        tmpResult = False
                    End If
                ElseIf type = ptCustomFields.Dbl Then
                    If _uid = .uid And CDbl(_value) = CDbl(.wert) Then
                        tmpResult = True
                    Else
                        tmpResult = False
                    End If
                ElseIf type = ptCustomFields.bool Then
                    If _uid = .uid And CBool(_value) = CBool(.wert) Then
                        tmpResult = True
                    Else
                        tmpResult = False
                    End If
                Else
                    tmpResult = False
                End If
            End With

            isIdenticalTo = tmpResult

        End Get
    End Property
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
