Imports Microsoft.Office.Interop.PowerPoint
Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ProjectBoardBasic

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
        englishLanguage = False

        ' im Powerpoint soll das pwd immer gemerkt werden ..

        awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        If awinSettings.rememberUserPwd Then
            awinSettings.userNamePWD = My.Settings.userNamePWD
        End If

    End Sub




    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        If VisboProtected Then
            Call makeVisboShapesVisible(False)
        End If
    End Sub

    ' see msdn: https://social.msdn.microsoft.com/Forums/sqlserver/en-US/b1c610bf-82ab-4d9e-b425-de21b45ea3fb/same-taskpane-in-multiple-powerpoint-windows?forum=vsto 
    Private listOfWindows As New List(Of Integer)
    Private listOfucProperties As New SortedList(Of Integer, Microsoft.Office.Tools.CustomTaskPane)
    Private listOfucSearch As New SortedList(Of Integer, Microsoft.Office.Tools.CustomTaskPane)
    Private listOfucPropView As New SortedList(Of Integer, ucProperties)
    Private listOfucSearchView As New SortedList(Of Integer, ucSearch)

    Private Sub Application_OpenPresentation(Pres As Presentation) Handles Application.AfterPresentationOpen, Application.AfterNewPresentation

        Dim hsearchPane As Microsoft.Office.Tools.CustomTaskPane
        Dim hPropPane As Microsoft.Office.Tools.CustomTaskPane
        Dim hucsearchView As ucSearch
        Dim hucPropView As ucProperties

        ' Id des aktiven Windows
        Dim hWinID As Integer = Application.ActiveWindow.HWND

        'Checks if the current window handle exists in the list of TaskPanes
        If listOfWindows.Contains(hWinID) Then
            'Do nothing, TaskPane has been added previously
        Else
            'Add TaskPanes

            ' Definition des Search-Pane
            hucsearchView = New ucSearch
            If englishLanguage Then
                hsearchPane = Me.CustomTaskPanes.Add(hucsearchView, "SEARCH" & hWinID.ToString)
            Else
                hsearchPane = Me.CustomTaskPanes.Add(hucsearchView, "SUCHE" & hWinID.ToString)
            End If

            With hsearchPane
                .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                .Height = 500
                .Width = 500
                .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                .Width = 350
                .Visible = False
            End With
            listOfucSearch.Add(hWinID, hsearchPane)
            searchPane = hsearchPane
            listOfucSearchView.Add(hWinID, hucsearchView)
            ucSearchView = hucsearchView

            ' Definition des Info-Pane
            hucPropView = New ucProperties
            If englishLanguage Then
                hPropPane = Me.CustomTaskPanes.Add(hucPropView, "PROPERTIES" & hWinID.ToString)
            Else
                hPropPane = Me.CustomTaskPanes.Add(hucPropView, "EIGENSCHAFTEN" & hWinID.ToString)
            End If

            ' tk, try , whether selection Fehler ist immer noch drin 
            With hPropPane
                .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
                .Height = 500
                .Width = 500
                .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                .Width = 320
                .Visible = False
            End With
            listOfucProperties.Add(hWinID, hPropPane)
            propertiesPane = hPropPane
            listOfucPropView.Add(hWinID, hucPropView)
            ucPropertiesView = hucPropView

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
            listOfWindows.Add(hWinID)
        End If

    End Sub

    Private Sub Application_PresentationBeforeClose(Pres As Presentation, ByRef Cancel As Boolean) Handles Application.PresentationBeforeClose

        ' Id des aktiven Windows
        Dim hWinID As Integer = Application.ActiveWindow.HWND

        ' globale Variablen für Eigenschaften Pane und das Pane selbst löschen
        If listOfucProperties.ContainsKey(hWinID) Then
            'Me.CustomTaskPanes.Remove(propertiesPane)
            listOfucProperties.Remove(hWinID)
        End If
        If listOfucPropView.ContainsKey(hWinID) Then
            'ucPropertiesView = Nothing
            listOfucPropView.Remove(hWinID)
        End If

        ' Username/Pwd in den Settings merken, falls Remember Me gecheckt
        My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
        If My.Settings.rememberUserPWD Then
            My.Settings.userNamePWD = awinSettings.userNamePWD
        End If
        My.Settings.Save()
    End Sub

    Private Sub Application_WindowSelectionChange(Sel As Selection) Handles Application.WindowSelectionChange

        '    'Checks if the current window handle exists in the list of TaskPanes
        '    If listOfWindows.Contains(Application.ActiveWindow.HWND) Then
        '        'Do nothing, TaskPane has been added previously
        '    Else
        '        'Add TaskPanes

        '        ' Definition des Search-Pane
        '        ucSearchView = New ucSearch
        '        If englishLanguage Then
        '            searchPane = Me.CustomTaskPanes.Add(ucSearchView, "SEARCH")
        '        Else
        '            searchPane = Me.CustomTaskPanes.Add(ucSearchView, "SUCHE")
        '        End If

        '        With searchPane
        '            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
        '            .Height = 500
        '            .Width = 500
        '            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
        '            .Width = 350
        '            .Visible = False
        '        End With

        '        ' Definition des Info-Pane
        '        ucPropertiesView = New ucProperties
        '        If englishLanguage Then
        '            propertiesPane = Me.CustomTaskPanes.Add(ucPropertiesView, "PROPERTIES")
        '        Else
        '            propertiesPane = Me.CustomTaskPanes.Add(ucPropertiesView, "EIGENSCHAFTEN")
        '        End If

        '        ' tk, try , whether selection Fehler ist immer noch drin 
        '        With propertiesPane
        '            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
        '            .Height = 500
        '            .Width = 500
        '            .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
        '            .Width = 320
        '            .Visible = False
        '        End With

        '        'Pane = New ucInfo
        '        'thePane = Me.CustomTaskPanes.Add(Pane, "Info")
        '        'With thePane
        '        '    .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionFloating
        '        '    .Height = 500
        '        '    .Width = 500
        '        '    .DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
        '        '    .Width = 320
        '        '    .Visible = True
        '        'End With

        '        'Add reference to the current DocumentWindow to the list
        '        listOfWindows.Add(Application.ActiveWindow.HWND)
        '    End If
    End Sub

    Private Sub Application_WindowActivate(Pres As Presentation, Wn As DocumentWindow) Handles Application.WindowActivate
        'Call MsgBox("windowActivate")   ' Definition des Search-Pane

        ' Id des aktiven DocumentWindow
        Dim hwinid As Integer = Wn.HWND

        ' globale Variablen für Eigenschaften Pane umsetzen
        If listOfucProperties.ContainsKey(Wn.HWND) Then
            propertiesPane = listOfucProperties.Item(Wn.HWND)
        End If
        If listOfucPropView.ContainsKey(Wn.HWND) Then
            ucPropertiesView = listOfucPropView.Item(Wn.HWND)
        End If

        ' globale Variable für search pane umsetzen
        If listOfucSearch.ContainsKey(Wn.HWND) Then
            searchPane = listOfucSearch.Item(Wn.HWND)
        End If
        If listOfucSearchView.ContainsKey(Wn.HWND) Then
            ucSearchView = listOfucSearchView.Item(Wn.HWND)
        End If

    End Sub

    Private Sub Application_WindowDeactivate(Pres As Presentation, Wn As DocumentWindow) Handles Application.WindowDeactivate

        'Call MsgBox("windowDeActivate")

        'Me.CustomTaskPanes.Remove(propertiesPane)
        'Me.CustomTaskPanes.Remove(searchPane)

    End Sub
End Class
