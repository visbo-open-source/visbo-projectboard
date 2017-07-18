Public Class clsroleNode
    Private _roleId As Integer
    Private _level As Integer
    Private _roleParent As Integer
    Private _childs As List(Of Integer)
    'Private _roleNodes As List(Of Integer)

    Public Property roleId As Integer
        Get
            roleId = _roleId
        End Get
        Set(value As Integer)
            _roleId = value
        End Set
    End Property

    'Public Property roleName As String
    '    Get
    '        roleName = _roleName
    '    End Get
    '    Set(value As String)
    '        _roleName = value
    '    End Set
    'End Property

    Public Property level As Integer
        Get
            level = _level
        End Get
        Set(value As Integer)
            _level = value
        End Set
    End Property

    Public Property roleParent As Integer
        Get
            roleParent = _roleParent
        End Get
        Set(value As Integer)
            _roleParent = value
        End Set
    End Property

    Public Property childs As List(Of Integer)
        Get
            childs = _childs
        End Get
        Set(value As List(Of Integer))
            _childs = value
        End Set
    End Property

    'Public Property roleNodes As List(Of Integer)
    '    Get
    '        roleNodes = _roleNodes
    '    End Get
    '    Set(value As List(Of Integer))
    '        _roleNodes = value
    '    End Set
    'End Property

    Public Sub New()

        _roleId = 0
        _level = 0
        _roleParent = 0
        _childs = New List(Of Integer)
     
    End Sub

    Public Sub New(Optional id As Integer = 0, _
                   Optional lv As Integer = 0, Optional rParent As Integer = 0, _
                   Optional chds As List(Of Integer) = Nothing)

        _roleId = id
        _level = lv
        _roleParent = rParent
        _childs = chds

    End Sub

    Public ReadOnly Property getlevelNodes(xlevel As Integer) As List(Of Integer)
        Get


        End Get

    End Property
    Dim listNodes As New List(Of Integer)





End Class
