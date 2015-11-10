''' <summary>
''' Regel, um ein Element aus einem anderen per Abstands-Regel zu bestimmen 
''' </summary>
''' <remarks></remarks>
Public Class clsAddElementRuleItem

    Private _newElemName As String
    Private _referenceName As String

    ' wenn das angegebene Referenz-Objekt eine Phase ist 
    Private _referenceIsPhase As Boolean

    ' wenn das angegebene Referenz-Objekt eine Phase ist: von welchem Datum aus soll der Offset berechnet werden 
    Private _referenceDateIsStart As Boolean

    ' Abstand in Tagen
    Private _offset As Integer



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
    ''' liest schreibt, ob es sich bei dem Referenz-Element um eine Phase handelt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property referenceIsPhase() As Boolean
        Get
            referenceIsPhase = _referenceIsPhase
        End Get
        Set(value As Boolean)
            _referenceIsPhase = value
        End Set
    End Property

    ''' <summary>
    ''' wenn es sich beim Referenz-Element um eine Phase handelt:
    ''' soll der Start als Referenzdatum verwendet werden; true: ja
    ''' false: das Ende soll als Referenzdatum verwendet werden 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property referenceDateIsStart() As Boolean
        Get
            referenceDateIsStart = _referenceDateIsStart
        End Get
        Set(value As Boolean)
            _referenceDateIsStart = value
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


   
    Public Sub New()
        _newElemName = ""
        _referenceName = ""
        _offset = 0
        _referenceIsPhase = False
        _referenceDateIsStart = True
    End Sub


End Class
