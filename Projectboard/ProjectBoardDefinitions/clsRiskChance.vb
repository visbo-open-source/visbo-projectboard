''' <summary>
''' hiermit können Risiken / Chancen aufgenommen werden 
''' </summary>
''' <remarks></remarks>
Public Class clsRiskChance

    Private _rcName As String
    ''' <summary>
    ''' setzt / liest den Namen
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property rcName As String
        Get
            rcName = _rcName
        End Get
        Set(value As String)
            If Not IsNothing(value) Then
                If Not value.Trim.Length = 0 Then
                    _rcName = value
                Else
                    Throw New ArgumentException("invalid name " & value)
                End If
            Else
                Throw New ArgumentException("Name for Risk/Chance must not be NULL")
            End If
        End Set
    End Property

    Public Property description As String

    Public Property category As String

    Private _probability As Double
    Public Property probability As Double
        Get
            probability = _probability
        End Get
        Set(value As Double)
            If value > 0 And value <= 1.0 Then
                _probability = value
            Else
                Throw New ArgumentException("Probability needs to be a value gretar than 0 and less or equal 1.0) ")
            End If
        End Set
    End Property

    Public Property potentialVariationInDuration As Integer

    Public Property damageBenefitValue As Double


    Sub New()

        rcName = "dummy"
        description = ""
        category = ""

        _probability = 0.5
        potentialVariationInDuration = 0
        damageBenefitValue = 0.0


    End Sub

End Class
