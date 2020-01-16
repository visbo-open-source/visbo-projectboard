Public Class clsStopWatch

    Private mlngStart As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long

    Public Sub StartTimer()
        mlngStart = GetTickCount
    End Sub

    Public Function EndTimer() As Long
        EndTimer = (GetTickCount - mlngStart)
    End Function

End Class
