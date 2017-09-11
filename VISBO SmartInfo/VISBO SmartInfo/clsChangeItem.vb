Public Class clsChangeItem
    Private _pName As String
    Private _vName As String
    Private _bestElemName As String
    Private _oldValue As String
    Private _newValue As String
    Private _diffInDays As Integer

    Public Property pName As String
        Get
            pName = _pName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _pName = value
            Else
                _pName = ""
            End If
        End Set
    End Property

    Public Property vName As String
        Get
            vName = _vName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _vName = value
            Else
                _vName = ""
            End If
        End Set
    End Property

    Public Property bestElemName As String
        Get
            bestElemName = _bestElemName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _bestElemName = value
            Else
                _bestElemName = ""
            End If
        End Set
    End Property

    Public Property oldValue As String
        Get
            oldValue = _oldValue
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _oldValue = value
            Else
                _oldValue = ""
            End If
        End Set
    End Property

    Public Property newValue As String
        Get
            newValue = _newValue
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _newValue = value
            Else
                _newValue = ""
            End If
        End Set
    End Property

    Public Property diffInDays As Integer
        Get
            diffInDays = _diffInDays
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                If IsNumeric(value) Then
                    _diffInDays = value
                Else
                    _diffInDays = 0
                End If
            Else
                _diffInDays = 0
            End If
        End Set
    End Property

    Public Sub New()
        _pName = ""
        _vName = ""
        _bestElemName = ""
        _oldValue = ""
        _newValue = ""
        _diffInDays = 0
    End Sub
End Class
