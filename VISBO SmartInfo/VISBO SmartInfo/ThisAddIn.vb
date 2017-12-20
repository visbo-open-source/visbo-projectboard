Imports Microsoft.Office.Interop.PowerPoint

Public Class ThisAddIn

    'Private pane As ucInfo
    'Private ucSearchView As ucSearch
    'Private WithEvents thePane As Microsoft.Office.Tools.CustomTaskPane
    'Private WithEvents searchPane As Microsoft.Office.Tools.CustomTaskPane

    Private Sub ThisAddIn_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        pptAPP = Application

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

    ' see msdn: https://social.msdn.microsoft.com/Forums/sqlserver/en-US/b1c610bf-82ab-4d9e-b425-de21b45ea3fb/same-taskpane-in-multiple-powerpoint-windows?forum=vsto 
    Private listOfWindows As New List(Of Integer)
    Private Sub Application_OpenPresentation(Pres As Presentation) Handles Application.AfterPresentationOpen, Application.AfterNewPresentation

        'Checks if the current window handle exists in the list of TaskPanes
        If listOfWindows.Contains(Application.ActiveWindow.HWND) Then
            'Do nothing, TaskPane has been added previously
        Else
            'Add TaskPanes

            ' Definition des Search-Pane
            ucSearchView = New ucSearch
            If englishLanguage Then
                searchPane = Me.CustomTaskPanes.Add(ucSearchView, "SEARCH")
            Else
                searchPane = Me.CustomTaskPanes.Add(ucSearchView, "SUCHE")
            End If

            With searchPane
                .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                .Height = 500
                .Width = 500
                .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                .Width = 350
                .Visible = False
            End With

            ' Definition des Info-Pane
            ucPropertiesView = New ucProperties
            If englishLanguage Then
                propertiesPane = Me.CustomTaskPanes.Add(ucPropertiesView, "PROPERTIES")
            Else
                propertiesPane = Me.CustomTaskPanes.Add(ucPropertiesView, "EIGENSCHAFTEN")
            End If

            ' tk, try , whether selection Fehler ist immer noch drin 
            With propertiesPane
                .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                .Height = 500
                .Width = 500
                .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                .Width = 320
                .Visible = False
            End With

            'Pane = New ucInfo
            'thePane = Me.CustomTaskPanes.Add(Pane, "Info")
            'With thePane
            '    .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
            '    .Height = 500
            '    .Width = 500
            '    .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
            '    .Width = 320
            '    .Visible = True
            'End With

            'Add reference to the current DocumentWindow to the list
            listOfWindows.Add(Application.ActiveWindow.HWND)
        End If

    End Sub

End Class
