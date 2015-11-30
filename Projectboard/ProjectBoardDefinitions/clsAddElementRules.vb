''' <summary>
''' enthält für ein neu zu generierendes Element die Liste alle alternativen Regeln, wie dieses Element erzeugt werden kann 
''' ''' </summary>
''' <remarks></remarks>
Public Class clsAddElementRules

    Private _regelListe As List(Of clsAddElementRuleItem)
    Private _name As String
    Private _elemToCreateIsPhase As Boolean
    Private _duration As Integer
    Private _deliverables As String

    ''' <summary>
    ''' liest/schreibt den Namen des Objekts (das zu erzeugende Element) 
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
    ''' gibt die x.te-Regel für das aktuelle Element zurück 
    ''' index muss zwischen 1 und Count liegen
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getItem(ByVal index As Integer) As clsAddElementRuleItem
        Get
            If index >= 1 And index <= _regelListe.Count Then
                getItem = _regelListe.ElementAt(index - 1)
            Else
                getItem = Nothing
            End If
        End Get
    End Property


    ''' <summary>
    ''' liest, ob das Element, das erzeugt werden soll, eine Phase ist
    ''' dann muss die Duration > 0 sein 
    ''' wird dementsprechend in duration gesetzt 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property elemToCreateIsPhase As Boolean

        Get
            elemToCreateIsPhase = _elemToCreateIsPhase
        End Get

    End Property

    ''' <summary>
    ''' liest/schreibt die Dauer der zu erzeugenden Phase
    ''' wenn Dauer kleiner oder gleich 0 angegeben wird, handelt es sich um einen Meilenstein;
    ''' bei Meilenstein ist Dauer = 0 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property duration As Integer
        Get
            duration = _duration
        End Get
        Set(value As Integer)
            If value > 0 Then
                _elemToCreateIsPhase = True
                _duration = value
            Else
                _elemToCreateIsPhase = False
                _duration = 0
            End If

        End Set
    End Property

    ''' <summary>
    ''' liest/schreibt die Lieferumfänge des neu zu erstellenden Elements
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property deliverables As String
        Get
            deliverables = _deliverables
        End Get
        Set(value As String)
            _deliverables = value
        End Set
    End Property


    ''' <summary>
    ''' gibt die Anzahl der alternativen Regeln zur Erzeugung des Elements zurück   
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property count() As Integer
        Get
            count = _regelListe.Count
        End Get
    End Property
    ''' <summary>
    ''' fügt der Liste für das zu erzeugende Element eine neue  Möglichkeit hinzu, wie es erzeugt werden kann 
    ''' 
    ''' </summary>
    ''' <param name="item"></param>
    ''' <remarks></remarks>
    Public Sub add(ByVal item As clsAddElementRuleItem)
        _regelListe.Add(item)
    End Sub

    Public Sub New()
        _regelListe = New List(Of clsAddElementRuleItem)
        _name = ""
        _elemToCreateIsPhase = False
        _duration = 0
        _deliverables = ""
    End Sub

    Public Sub New(ByVal name As String, ByVal isPhase As Boolean, ByVal duration As Integer, _
                   ByVal deliverables As String)
        _regelListe = New List(Of clsAddElementRuleItem)
        _name = name
        _elemToCreateIsPhase = isPhase
        _duration = duration
        _deliverables = deliverables
    End Sub

End Class
