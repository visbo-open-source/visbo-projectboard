Public Class ThisAddIn

    Private Sub ThisAddIn_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        pptAPP = Application

        trafficLightColors(0) = PowerPoint.XlRgbColor.rgbGrey
        trafficLightColors(1) = PowerPoint.XlRgbColor.rgbGreen
        trafficLightColors(2) = PowerPoint.XlRgbColor.rgbYellow
        trafficLightColors(3) = PowerPoint.XlRgbColor.rgbRed
        trafficLightColors(4) = PowerPoint.XlRgbColor.rgbWhite

        showTrafficLights(0) = False
        showTrafficLights(1) = False
        showTrafficLights(2) = False
        showTrafficLights(3) = False
        showTrafficLights(4) = False


    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If
    End Sub

End Class
