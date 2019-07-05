
Public Class clsConstellationItem

    Private _projectTyp As String = ""
    Private _reasonToInclude As String = ""
    Private _projectName As String = ""
    Private _variantName As String = ""
    Private _vpid As String = ""
    Private _start As Date = StartofCalendar
    Private _show As Boolean = False
    Private _zeile As Integer = 0

    ''' <summary>
    ''' prüft auf Identität 
    ''' </summary>
    ''' <param name="vglCI"></param>
    ''' <returns></returns>
    Public Function isIdentical(ByVal vglCI As clsConstellationItem) As Boolean
        Dim istGleich As Boolean = False

        If projectTyp = vglCI.projectTyp Then
            If reasonToInclude = vglCI.reasonToInclude Then
                If projectName = vglCI.projectName Then
                    If variantName = vglCI.variantName Then
                        If show = vglCI.show Then
                            If DateDiff(DateInterval.Month, start, vglCI.start) = 0 Then
                                If zeile = vglCI.zeile Then
                                    istGleich = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        isIdentical = istGleich
    End Function

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

    Public Property projectTyp As String
        Get
            projectTyp = _projectTyp
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                If value = ptPRPFType.project.ToString Or
                        value = ptPRPFType.portfolio.ToString Then
                    _projectTyp = value
                Else
                    _projectTyp = ptPRPFType.project.ToString
                End If

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
    Public Property vpid As String
        Get
            vpid = _vpid
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                _vpid = value
            Else
                _vpid = ""
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

    ''' <summary>
    ''' kopiert das constellation Item in eine neue Instanz-Variable 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property copy(ByVal Optional prepareForDB As Boolean = False) As clsConstellationItem
        Get
            Dim copyResult As New clsConstellationItem

            With copyResult
                .projectName = Me.projectName
                ' tk 2.3.19
                If prepareForDB And Me.variantName = ptVariantFixNames.pfv.ToString Then
                    .variantName = ""
                Else
                    .variantName = Me.variantName
                End If
                .vpid = Me.vpid
                .show = Me.show
                .projectTyp = Me.projectTyp
                .reasonToInclude = Me.reasonToInclude
                .start = Me.start
                .zeile = Me.zeile
            End With

            copy = copyResult

        End Get
    End Property

    Sub New()

        _projectName = ""
        _variantName = ""
        _vpid = ""
        _start = StartofCalendar.AddMonths(-1)
        _show = False
        _zeile = 0
        _reasonToInclude = ""
        ' war vorher reasonToExclude - istz umgewidmet worden, weil ohnehin nicht gebraucht 
        _projectTyp = ptPRPFType.project.ToString

    End Sub
End Class
