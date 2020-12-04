Imports Microsoft.Office.Interop.PowerPoint
Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ProjectBoardBasic

Public Class ThisAddIn


    Private Sub ThisAddIn_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        If VisboProtected Then
            Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoFalse)
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

        Try
            If englishLanguage Then
                menuCult = ReportLang(PTSprache.englisch)
                repCult = menuCult
                'awinSettings.kapaEinheit = "PD"
            Else
                menuCult = ReportLang(PTSprache.deutsch)
                repCult = menuCult
                'awinSettings.kapaEinheit = "PT"
            End If
        Catch ex As Exception
            awinSettings.englishLanguage = False
            awinSettings.kapaEinheit = "PT"
            menuCult = ReportLang(PTSprache.deutsch)
            repCult = menuCult
        End Try

        ' im Powerpoint soll das pwd immer gemerkt werden ..

        awinSettings.rememberUserPwd = My.Settings.rememberUserPWD
        If awinSettings.rememberUserPwd Then
            awinSettings.userNamePWD = My.Settings.userNamePWD
        End If

        ' visboClient = "VISBO Smartinfo / "
        visboClient = divClients(client.VisboSmartinfo)
        ' Refresh von Projekte im Cache  in Minuten
        cacheUpdateDelay = 10

        ' für Demo Zwecke only !! 
        ' vor Auslieferung auf false setzen ! 
        'awinSettings.autoSetActualDataDate = True

    End Sub


    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

        Try
            If VisboProtected Then
                Call makeVisboShapesVisible(Microsoft.Office.Core.MsoTriState.msoFalse)
            End If
        Catch ex As Exception

        End Try


        Try
            ' User spezifische Settings für Report Creation gemerkt in my.settings
            Call writeSettings()

            ' Username/Pwd in den Settings merken, falls Remember Me gecheckt
            My.Settings.rememberUserPWD = awinSettings.rememberUserPwd
            If My.Settings.rememberUserPWD Then
                My.Settings.userNamePWD = awinSettings.userNamePWD
            End If

            My.Settings.Save()

            ' Logout des Users am Server
            Dim err As New clsErrorCodeMsg

            Dim logoutErfolgreich As Boolean = CType(databaseAcc, DBAccLayer.Request).logout(err)

            If logoutErfolgreich Then
                If awinSettings.visboDebug Then
                    If awinSettings.englishLanguage Then
                        Call MsgBox(err.errorMsg & vbCrLf & "User don't have access to a VisboCenter any longer!")
                    Else
                        Call MsgBox(err.errorMsg & vbCrLf & "User hat keinen Zugriff mehr zu einem VisboCenter!")
                    End If
                End If

            End If

        Catch ex As Exception

        End Try

        'Try
        '    Call closeExcelAPP()
        'Catch ex As Exception

        'End Try


    End Sub

    Private Sub Application_OpenPresentation(Pres As Presentation) Handles Application.AfterPresentationOpen, Application.AfterNewPresentation, Application.PresentationSave

        ' das muss hier bleiben und kann nicht nach Module1 verlagert werden .. 

        currentPresHasVISBOElements = presentationHasAnySmartSlides(Pres)

        If currentPresHasVISBOElements Then

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
                    'hsearchPane = Me.CustomTaskPanes.Add(hucsearchView, "SEARCH" & hWinID.ToString)
                    hsearchPane = Me.CustomTaskPanes.Add(hucsearchView, "SEARCH")
                Else
                    'hsearchPane = Me.CustomTaskPanes.Add(hucsearchView, "SUCHE" & hWinID.ToString)
                    hsearchPane = Me.CustomTaskPanes.Add(hucsearchView, "SUCHE")
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
                    'hPropPane = Me.CustomTaskPanes.Add(hucPropView, "PROPERTIES" & hWinID.ToString)
                    hPropPane = Me.CustomTaskPanes.Add(hucPropView, "PROPERTIES")
                Else
                    'hPropPane = Me.CustomTaskPanes.Add(hucPropView, "EIGENSCHAFTEN" & hWinID.ToString)
                    hPropPane = Me.CustomTaskPanes.Add(hucPropView, "EIGENSCHAFTEN")
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

        End If


    End Sub

    Private Sub Application_PresentationBeforeClose(Pres As Presentation, ByRef Cancel As Boolean) Handles Application.PresentationBeforeClose
        Dim a As Integer = 0
        ' bitte nicht benutzen siehe Module1.vb
    End Sub

    Private Sub Application_WindowSelectionChange(Sel As Selection) Handles Application.WindowSelectionChange
        ' bitte nicht benutzen siehe Module1.vb
        Dim a As Integer = 0
    End Sub

    Private Sub Application_WindowActivate(Pres As Presentation, Wn As DocumentWindow) Handles Application.WindowActivate
        ' bitte nicht benutzen siehe Module1.vb
        Dim a As Integer = 0
    End Sub

    Private Sub Application_WindowDeactivate(Pres As Presentation, Wn As DocumentWindow) Handles Application.WindowDeactivate
        ' bitte nicht benutzen siehe Module1.vb
        Dim a As Integer = 0
    End Sub

    Private Sub Application_SlideSelectionChanged(SldRange As SlideRange) Handles Application.SlideSelectionChanged
        Dim a As Integer = 0
    End Sub

    Private Sub Application_PresentationPrint(Pres As Presentation) Handles Application.PresentationPrint
        ' Es wird noch pptApp_PresentationPrint aufgerufen
        Dim a As Integer = 0
    End Sub

    Private Sub Application_PresentationBeforeSave(Pres As Presentation, ByRef Cancel As Boolean) Handles Application.PresentationBeforeSave
        ' Es wird noch pptApp_PresentationBeforeSave aufgerufen
        Dim a As Integer = 0
    End Sub

    Private Sub ThisAddIn_BindingContextChanged(sender As Object, e As EventArgs) Handles Me.BindingContextChanged

    End Sub

End Class
