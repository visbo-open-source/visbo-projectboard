Public Class ThisAddIn

    Private pane As ucInfo
    'Private ucSearchView As ucSearch
    Private WithEvents thePane As Microsoft.Office.Tools.CustomTaskPane
    'Private WithEvents searchPane As Microsoft.Office.Tools.CustomTaskPane

    Private Sub ThisAddIn_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        pptAPP = Application

        ucPropertiesView = New ucProperties
        propertiesPane = Me.CustomTaskPanes.Add(ucPropertiesView, "Eigenschaften")
        With propertiesPane
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            .Height = 500
            .Width = 500
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            .Width = 320
            .Visible = False
        End With


        ucSearchView = New ucSearch
        searchPane = Me.CustomTaskPanes.Add(ucSearchView, "SEARCH")
        With searchPane
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            .Height = 500
            .Width = 500
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            .Width = 320
            .Visible = False
        End With

        pane = New ucInfo
        thePane = Me.CustomTaskPanes.Add(pane, "Info")
        With thePane
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            .Height = 500
            .Width = 500
            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            .Width = 320
            .Visible = True
        End With


        trafficLightColors(0) = PowerPoint.XlRgbColor.rgbGray
        trafficLightColors(1) = PowerPoint.XlRgbColor.rgbGreen
        trafficLightColors(2) = PowerPoint.XlRgbColor.rgbYellow
        trafficLightColors(3) = PowerPoint.XlRgbColor.rgbRed
        trafficLightColors(4) = PowerPoint.XlRgbColor.rgbWhite

        showTrafficLights(0) = False
        showTrafficLights(1) = False
        showTrafficLights(2) = False
        showTrafficLights(3) = False
        showTrafficLights(4) = False

        ' muss hier noch auf andere Art und Weise bestimmt werden 
        englishLanguage = True


    End Sub



    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If
    End Sub


    Public Function showPropertiesPane(ByVal custPropPane As Microsoft.Office.Tools.CustomTaskPane) As Boolean



    End Function
End Class
