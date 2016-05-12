Public Class clsCustomFieldDefinition

    Private _name As String
    Private _type As Integer
    Private _uid As Integer

    ''' <summary>
    ''' eindeutige ID für das Custom Field
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property uid As Integer
        Get
            uid = _uid
        End Get
        Set(value As Integer)
            _uid = value
        End Set
    End Property

    ''' <summary>
    ''' der Name des Custom Fields
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property name As String
        Get
            name = _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property

    ''' <summary>
    ''' welche Art von Custom-Field ist es ? String-0, Double-1 oder Boolean-2
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property type As Integer
        Get
            type = _type
        End Get
        Set(value As Integer)
            If value = ptCustomFields.Str Or _
                value = ptCustomFields.Dbl Or _
                value = ptCustomFields.bool Then
                _type = value
            Else
                Throw New ArgumentException("unzulässiger Wert für Custom Field Typ: " & value)
            End If
        End Set
    End Property

    Sub New()
        _name = "test"
        _type = ptCustomFields.Str
    End Sub

    Sub New(ByVal name As String, ByVal type As Integer, ByVal uid As Integer)

        If Not IsNothing(name) And Not IsNothing(type) Then
            _name = name
            If type = ptCustomFields.Str Or _
                type = ptCustomFields.Dbl Or _
                type = ptCustomFields.bool Then
                _type = type
            Else
                _type = ptCustomFields.Str
            End If
        Else

            If IsNothing(name) Then
                _name = "test"
                If type = ptCustomFields.Str Or _
                type = ptCustomFields.Dbl Or _
                type = ptCustomFields.bool Then
                    _type = type
                Else
                    _type = ptCustomFields.Str
                End If
                If IsNothing(type) Then
                    _name = name
                    _type = ptCustomFields.Str
                End If
            End If
        End If

        _uid = uid

    End Sub

End Class
