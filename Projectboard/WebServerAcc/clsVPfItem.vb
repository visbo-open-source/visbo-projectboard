Imports ProjectBoardDefinitions
Public Class clsVPfItem

    'Inherits clsConstellationItem
    Private _reasonToExclude As String = ""
    Private _reasonToInclude As String = ""
    Private _projectName As String = ""
    Private _variantName As String = ""
    Private _start As Date = StartofCalendar
    Private _show As Boolean = False
    Private _zeile As Integer = 0

    Public Property reasonToInclude As String
        Get
            reasonToInclude = _reasonToInclude
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _reasonToInclude = value
            End If
        End Set
    End Property

    Public Property reasonToExclude As String
        Get
            reasonToExclude = _reasonToExclude
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _reasonToExclude = value
            End If
        End Set
    End Property

    Public Property projectName As String
        Get
            projectName = _projectName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _projectName = value
            Else
                _projectName = ""
            End If

        End Set
    End Property
    Public Property variantName As String
        Get
            variantName = _variantName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _variantName = value
            Else
                _variantName = ""
            End If

        End Set
    End Property

    Public Property start As Date
        Get
            start = _start
        End Get
        Set(value As Date)
            _start = value
        End Set
    End Property
    Public Property show As Boolean
        Get
            show = _show
        End Get
        Set(value As Boolean)
            If Not IsNothing(value) Then
                _show = value
            Else
                _show = True
            End If
        End Set
    End Property
    Public Property zeile As Integer
        Get
            zeile = _zeile
        End Get
        Set(value As Integer)
            If Not IsNothing(value) Then
                If value >= 0 Then
                    _zeile = value
                Else
                    _zeile = 0
                End If
            Else
                _zeile = 0
            End If
        End Set
    End Property


    Public Property name As String
    Public Property vpid As String
    Public Property _id As String


    Sub New()
        _vpid = ""
        _id = ""
        _name = ""
        _projectName = ""
        _variantName = ""
        _start = StartofCalendar.AddMonths(-1)
        _show = False
        _zeile = 0
        _reasonToInclude = ""
        _reasonToExclude = ""
    End Sub


    'Overloads Sub copyfrom(ByVal item As clsVPfItem)

    '    With item
    '        Me.projectName = .name
    '        Me.variantName = .variantName
    '        Me.Start = .Start.ToUniversalTime
    '        Me.show = .show
    '        Me.zeile = .zeile
    '        Me.reasonToInclude = .reasonToInclude
    '        Me.reasonToExclude = .reasonToExclude
    '    End With
    'End Sub

    'Overloads Sub copyto(ByRef item As clsConstellationItem)

    '    With item
    '        .projectName = Me.name
    '        .variantName = Me.variantName
    '        .start = Me.Start.ToLocalTime
    '        .show = Me.show
    '        .zeile = Me.zeile
    '        .reasonToInclude = Me.reasonToInclude
    '        .reasonToExclude = Me.reasonToExclude
    '    End With

    'End Sub

End Class
