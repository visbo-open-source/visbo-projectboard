''' <summary>
''' Regel, um ein Element aus einem anderen per Abstands-Regel zu bestimmen 
''' </summary>
''' <remarks></remarks>
Public Class clsAddElementRule

    Private _newElemName As String
    Private _referenceName As String
    Private _offset As Integer
    Private _deliverables As String


    ''' <summary>
    ''' liest / schreibt den Namen des neuen Elements
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property newElemName As String
        Get
            newElemName = _newElemName
        End Get
        Set(value As String)
            _newElemName = value
        End Set
    End Property
    ''' <summary>
    ''' liest/schreibt den Referenz-Namen , der benutzt wird, um das Element zu bestimmen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property referenceName As String
        Get
            referenceName = _referenceName
        End Get
        Set(value As String)
            _referenceName = value
        End Set
    End Property
    ''' <summary>
    ''' liest/schreibt den zeitlichen Abstand in Tagen, den das Element haben soll  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property offset As Integer
        Get
            offset = _offset
        End Get
        Set(value As Integer)
            _offset = value
        End Set
    End Property

    Public Property deliverables As String
        Get
            deliverables = _deliverables
        End Get
        Set(value As String)
            _deliverables = value
        End Set
    End Property

    Public Sub New()
        _newElemName = ""
        _referenceName = ""
        _offset = 0
        _deliverables = ""
    End Sub


End Class
