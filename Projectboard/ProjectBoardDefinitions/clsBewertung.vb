Public Class clsBewertung
    ' Änderung tk 2.11.15
    ' Property deliverables eingeführt 

    Private _color As Integer
    Private _description As String = ""
    Private _deliverables As String = ""

    Public Property bewerterName As String
    Public Property datum As Date

    ''' <summary>
    ''' es muss abgefangen werden, dass in description ein Nothing Wert stehen kann 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property description As String
        Get
            description = _description
        End Get
        Set(value As String)
            If IsNothing(value) Then
                _description = ""
            Else
                _description = value
            End If
        End Set
    End Property
    ''' <summary>
    ''' liest/schreibt deliverables
    ''' es muss abgefangen werden, dass in deliverables ein Nothing Wert stehen kann 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property deliverables As String
        Get
            deliverables = _deliverables
        End Get
        Set(value As String)
            If IsNothing(value) Then
                _deliverables = ""
            Else
                _deliverables = value
            End If
        End Set
    End Property
    


    ''' <summary>
    ''' kann die Werte 0, 1, 2, 3 annehmen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property colorIndex As Integer
        Get
            colorIndex = _color
        End Get

        Set(value As Integer)

            If value >= 0 And value <= 3 Then
                _color = value
            Else
                _color = 0
            End If

        End Set
    End Property



    ''' <summary>
    ''' liest / schreibt die Ampelfarben wie in awinsettings definiert  
    ''' 0:keine Bewertung; 1:grün; 2:gelb; 3: rot
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property color As Long
        Get
            If _color = 0 Then
                color = awinSettings.AmpelNichtBewertet
            ElseIf _color = 1 Then
                color = awinSettings.AmpelGruen
            ElseIf _color = 2 Then
                color = awinSettings.AmpelGelb
            ElseIf _color = 3 Then
                color = awinSettings.AmpelRot
            Else
                color = awinSettings.AmpelNichtBewertet
            End If
        End Get
        Set(value As Long)

            If value = 0 Or value = awinSettings.AmpelNichtBewertet Then
                _color = 0
            ElseIf value = 1 Or value = awinSettings.AmpelGruen Then
                _color = 1
            ElseIf value = 2 Or value = awinSettings.AmpelGelb Then
                _color = 2
            ElseIf value = 3 Or value = awinSettings.AmpelRot Then
                _color = 3
            Else
                _color = 0
            End If

        End Set
    End Property



    Public Sub copyto(ByRef newb As clsBewertung)

        With newb
            .description = Me.description
            .deliverables = Me.deliverables
            .color = Me.color
            .bewerterName = Me.bewerterName
            .datum = Me.datum
        End With

    End Sub

    Public Sub copyfrom(ByVal newb As clsBewertung)

        With newb
            Me.description = .description
            Me.deliverables = .deliverables
            Me.color = .color
            Me.bewerterName = .bewerterName
            Me.datum = .datum
        End With

    End Sub


    Public Sub New()
        bewerterName = ""
        datum = Nothing
        color = awinSettings.AmpelNichtBewertet
        _description = ""
        _deliverables = ""
    End Sub

End Class
