Public Class clsKeyMetrics

    Public Property costCurrentActual As Double
    Public Property costCurrentTotal As Double
    Public Property costBaseLastActual As Double
    Public Property costBaseLastTotal As Double

    Public Property timeCompletionCurrentActual As Double
    Public Property timeCompletionBaseLastActual As Double

    Public Property timeCompletionCurrentTotal As Double
    Public Property timeCompletionBaseLastTotal As Double
    Public Property endDateCurrent As Date
    Public Property endDateBaseLast As Date

    Public Property deliverableCompletionCurrentActual As Double
    Public Property deliverableCompletionCurrentTotal As Double
    Public Property deliverableCompletionBaseLastActual As Double
    Public Property deliverableCompletionBaseLastTotal As Double

    Sub New()
        costCurrentActual = 0.0
        costCurrentTotal = 0.0
        costBaseLastActual = 0.0
        costBaseLastTotal = 0.0

        timeCompletionCurrentActual = 0.0
        timeCompletionBaseLastActual = 0.0
        timeCompletionCurrentTotal = 0.0
        timeCompletionBaseLastTotal = 0.0
        endDateCurrent = Date.MinValue
        endDateBaseLast = Date.MinValue

        deliverableCompletionCurrentActual = 0.0
        deliverableCompletionCurrentTotal = 0.0
        deliverableCompletionBaseLastActual = 0.0
        deliverableCompletionBaseLastTotal = 0.0
    End Sub

End Class
