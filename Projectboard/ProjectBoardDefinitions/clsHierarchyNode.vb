Public Class clsHierarchyNode

    
    Private _elemName As String
    Private _origName As String
    Private _indexOfElem As Integer
    Private _isMilestone As Boolean
    Private _parentNodeKey As String
    Private _childNodeKeys As List(Of String)


    ''' <summary>
    ''' legt einen neuen Hierarchie Knoten mit ID an
    ''' die ID muss einen Wert von >= 0 haben 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        _elemName = ""
        _origName = ""
        _indexOfElem = -1
        _isMilestone = False
        _parentNodeKey = ""
        _childNodeKeys = New List(Of String)

    End Sub

    ''' <summary>
    ''' legt einen neuen Hierarchie Knoten mit allen erforderlichen Angaben an 
    ''' nur die Kind-Knoten müssen dann noch ergänzt werden 
    ''' </summary>
    ''' <param name="elemName"></param>
    ''' <param name="origName"></param>
    ''' <param name="indexOfElem"></param>
    ''' <param name="isMilestone"></param>
    ''' <param name="parentNodeKey"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal elemName As String, ByVal origName As String, _
                       ByVal indexOfElem As Integer, ByVal isMilestone As Boolean, ByRef parentNodeKey As String)




        If Not IsNothing(elemName) Then
            If elemName.Trim.Length > 0 Then
                _elemName = elemName
            Else
                Throw New ArgumentException("Element Name darf nicht Nothing oder leer sein")
            End If
        Else
            Throw New ArgumentException("Element Name darf nicht Nothing oder leer sein")
        End If

        If Not IsNothing(origName) Then
            _origName = origName
        Else
            _origName = ""
        End If


        If indexOfElem >= 1 Then
            _indexOfElem = indexOfElem
        Else
            Throw New ArgumentException(indexOfElem & " ist kein gültiger Phasen-Index")
        End If

        _isMilestone = isMilestone

        If Not IsNothing(parentNodeKey) Then
            _parentNodeKey = parentNodeKey
        Else
            Throw New ArgumentException(parentNodeKey & " ist keine gültige Parent Element-ID")
        End If

        _childNodeKeys = New List(Of String)

    End Sub

    ''' <summary>
    ''' liest die ParentNode-ID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property parentNodeKey As String
        Get
            parentNodeKey = _parentNodeKey
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _parentNodeKey = value
            Else
                Throw New ArgumentException("Parent-Key darf nicht Null sein")
            End If
        End Set
    End Property


    ''' <summary>
    ''' fügt dem Hierarchie Knoten ein neues Kind mit Key = childKey hinzu 
    ''' wenn childKey schon existiert, so wird nichts gemacht, aber auch kein Fehler geworfen
    ''' wenn ChildKey Nothing oder "" ist, dann wird ein Fehler geworfen
    ''' </summary>
    ''' <param name="childKey"></param>
    ''' <remarks></remarks>
    Public Sub addChild(ByVal childKey As String)

        If Not IsNothing(childKey) Then
            If childKey.Trim.Length > 0 Then
                If Not _childNodeKeys.Contains(childKey) Then
                    _childNodeKeys.Add(childKey)
                Else
                    ' nichts tun - dann ist das bereits in der childs-Liste aufgenommen
                End If
            Else
                Throw New ArgumentException("Key für Child darf nicht leer sein")
            End If

        Else
            Throw New ArgumentException("Key für Child darf nicht Null sein")
        End If

    End Sub

    ''' <summary>
    ''' entfernt aus dem Hierarchie-Knoten das Kind mit ID = childID 
    ''' </summary>
    ''' <param name="childKey"></param>
    ''' <remarks></remarks>
    Public Sub removeChild(ByVal childKey As String)
        If Not IsNothing(childKey) Then
            If _childNodeKeys.Contains(childKey) Then
                _childNodeKeys.Remove(childKey)
            Else
                ' nichts tun - dann ist das bereits nicht mehr in der Child Liste 
            End If
        Else
            Throw New ArgumentException(childKey & " ist keine gültige Element-ID")
        End If
    End Sub


    ''' <summary>
    ''' liest bzw. schreibt den Element-Name des Knotens , nicht zu verwechseln mit dem Unique Key in der sortedlist
    ''' unique-Name = elemName#0/1(#laufende Nummer)
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property elemName As String
        Get
            elemName = _elemName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _elemName = value
            Else
                _elemName = ""
            End If
        End Set
    End Property

    ''' <summary>
    ''' liest bzw. schreibt den Original Namen des Objektes, so wie er aus dem Projektplan ausgelesen wurde
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property origName As String
        Get
            origName = _origName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _origName = value
            Else
                _origName = ""
            End If
        End Set
    End Property


    ''' <summary>
    ''' liest oder schreibt den bool'schen Wert, ob es sich um einen Meilenstein handelt 
    ''' oder nicht
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property isMilestone As Boolean
        Get
            isMilestone = _isMilestone
        End Get
        Set(value As Boolean)
            _isMilestone = value
        End Set
    End Property

    ''' <summary>
    ''' liest / schreibt den IndexOfPhase
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property indexOfElem As Integer
        Get
            indexOfElem = _indexOfElem
        End Get
        Set(value As Integer)
            If value >= 1 Then
                _indexOfElem = value
            Else
                Throw New ArgumentException("nicht zugelassener IndexOfPhase: " & value)
            End If
        End Set
    End Property
End Class
