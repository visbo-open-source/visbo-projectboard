Public Class clsWriteProtectionItem
    ' der Projekt-Varianten-Name: pName#vName
    Private _pvName As String
    ' die Kennung, also handelt es sich um eine Projekt-Variante, ein Szenario, ...
    Private _type As Integer
    ' der user-Name, der in der Datenbank eingeloggt ist 
    Private _userName As String
    ' gibt an, ob die Projekt-Variante geschützt ist oder nicht 
    Private _isProtected As Boolean
    ' wenn true, bleibt die Sperre über die Session hinaus bestehen, solnage bis sie explizit aufgehoben wird 
    Private _permanent As Boolean
    Private _lastDateSet As Date
    Private _lastDateReleased As Date

    ''' <summary>
    ''' liest / setzt den pvName (pName#vName)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property pvName() As String
        Get
            pvName = _pvName
        End Get
        Set(value As String)
            If IsNothing(value) Then
                Throw New ArgumentException("value may not be Null")
            ElseIf value = "" Then
                Throw New ArgumentException("value may not be empty string")
            Else
                _pvName = value
            End If

        End Set
    End Property

    Public Property type() As Integer
        Get
            type = _type
        End Get
        Set(value As Integer)
            If IsNothing(value) Then
                Throw New ArgumentException("value may not be Null")
            ElseIf value < 0 Then
                Throw New ArgumentException("value may not be lt 0")
            Else
                _type = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' liest / setzt den UserName 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property userName As String
        Get
            userName = _userName
        End Get
        Set(value As String)
            If IsNothing(value) Then
                Throw New ArgumentException("value may not be Null")
            ElseIf value = "" Then
                Throw New ArgumentException("value may not be empty string")
            Else
                _userName = value
            End If

        End Set
    End Property

    ''' <summary>
    ''' liest / setzt, ob pvName protected ist 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property isProtected As Boolean
        Get
            isProtected = _isProtected
        End Get
        Set(value As Boolean)
            _isProtected = value
            If _isProtected Then
                ' setzen lastSet 
                _lastDateSet = Date.Now
            Else
                _lastDateReleased = Date.Now
            End If
        End Set
    End Property

    ''' <summary>
    ''' liest / setzt, ob permanent geschützt werden soll 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property permanent As Boolean
        Get
            permanent = _permanent
        End Get
        Set(value As Boolean)
            _permanent = value
        End Set
    End Property

    Public Property lastDateSet As Date
        Get
            lastDateSet = _lastDateSet
        End Get
        Set(value As Date)
            _lastDateSet = value
        End Set
    End Property

    Public Property lastDateReleased As Date
        Get
            lastDateReleased = _lastDateReleased
        End Get
        Set(value As Date)
            _lastDateReleased = value
        End Set
    End Property
    Sub New()
        _pvName = ""
        _type = ptWriteProtectionType.project
        _userName = ""
        _isProtected = True
        _permanent = False
        _lastDateSet = Date.Now
        _lastDateReleased = Date.MinValue
    End Sub

    ''' <summary>
    ''' Konstruktor mit allen notwendigen Angaben, um ein pvName zu schützen  
    ''' </summary>
    ''' <remarks></remarks>
    Sub New(ByVal pvN As String, _
            ByVal type As Integer, _
            ByVal userN As String, _
            ByVal prmnnt As Boolean, _
            ByVal protectIT As Boolean)

        _pvName = pvN
        _type = type
        _userName = userN
        _isProtected = protectIT

        If Not protectIT Then
            _permanent = False
            _lastDateSet = Date.MinValue
            _lastDateReleased = Date.Now
        Else
            _permanent = prmnnt
            _lastDateSet = Date.Now
            _lastDateReleased = Date.MinValue
        End If

    End Sub
End Class
