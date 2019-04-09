Public Class clsBewertung
    

    Private _colorIndex As Integer
    Private _description As String = ""


    Public Property bewerterName As String
    Public Property datum As Date

    ''' <summary>
    ''' vergleicht eine Bewertung auf Identität
    ''' </summary>
    ''' <param name="vBewertung"></param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isIdenticalTo(ByVal vBewertung As clsBewertung) As Boolean
        Get
            Dim stillOK As Boolean = False

            If Me.colorIndex = vBewertung.colorIndex And
                    Me.description = vBewertung.description And
                    Me.bewerterName = vBewertung.bewerterName Then
                stillOK = True
            Else
                stillOK = False
            End If


            isIdenticalTo = stillOK
        End Get
    End Property

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
    ''' kann die Werte 0, 1, 2, 3 annehmen 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property colorIndex As Integer
        Get
            colorIndex = _colorIndex
        End Get

        Set(value As Integer)

            If value >= 0 And value <= 3 Then
                _colorIndex = value
            Else
                _colorIndex = 0
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
    Public ReadOnly Property color As Long
        Get
            If _colorIndex = 0 Then
                color = awinSettings.AmpelNichtBewertet
            ElseIf _colorIndex = 1 Then
                color = awinSettings.AmpelGruen
            ElseIf _colorIndex = 2 Then
                color = awinSettings.AmpelGelb
            ElseIf _colorIndex = 3 Then
                color = awinSettings.AmpelRot
            Else
                color = awinSettings.AmpelNichtBewertet
            End If
        End Get
        
    End Property



    Public Sub copyto(ByRef newb As clsBewertung)

        With newb
            .description = Me.description
            .colorIndex = Me.colorIndex
            .bewerterName = Me.bewerterName
            .datum = Me.datum
        End With

    End Sub

    Public Sub copyfrom(ByVal newb As clsBewertung)

        With newb
            Me.description = .description
            Me.colorIndex = .colorIndex
            Me.bewerterName = .bewerterName
            Me.datum = .datum
        End With

    End Sub


    Public Sub New()
        bewerterName = ""
        datum = Nothing
        _colorIndex = 0
        _description = ""
    End Sub

End Class
