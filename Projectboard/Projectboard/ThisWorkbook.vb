Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports MongoDbAccess




Public Class ThisWorkbook
    ' Copyright Philipp Koytek et al. 
    ' 2012 ff
    ' Nicht authorisierte Verwendung nicht gestattet 


    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

    Private Sub ThisWorkbook_ActivateEvent() Handles Me.ActivateEvent

        Application.DisplayFormulaBar = False
        'Application.Worksheets(arrWsNames(3)).Activate()

    End Sub

    ''' <summary>
    ''' stellt sicher, daß die Excel Settings in anderen Workbooks wieder gelten
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ThisWorkbook_Deactivate() Handles Me.Deactivate

        Application.DisplayFormulaBar = True

        With Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = 0
            .DisplayWorkbookTabs = True
            .GridlineColor = RGB(220, 220, 220)
            Try
                .FreezePanes = False
            Catch ex As Exception

            End Try

            .DisplayHeadings = True
        End With

    End Sub

    Private Sub ThisWorkbook_Startup() Handles Me.Startup

        'Dim cbar As CommandBar


        appInstance = Application

        myProjektTafel = appInstance.ActiveWorkbook.Name


        ' die Short Cut Menues aus Excel werden hier nicht mehr de-aktiviert 
        ' das wird jetzt nur in Tabelle1, also der Projekt-Tafel gemacht ...
        ' in anderen Excel Sheets ist das weiterhin aktiv 
        'For Each cbar In appInstance.CommandBars

        '    If cbar.Type = MsoBarType.msoBarTypePopup Then
        '        cbar.Enabled = False
        '    End If
        'Next

        magicBoardCmdBar.cmdbars = appInstance.CommandBars

      

        Try

            appInstance.ScreenUpdating = False
            Call awinsetTypen()

        Catch ex As Exception

            Call MsgBox(ex.Message)

        Finally
            appInstance.ScreenUpdating = True
            appInstance.ShowChartTipNames = True
            appInstance.ShowChartTipValues = True
        End Try

        anzahlCalls = 0


        'Call awinRightClickinPortfolioAendern()
        ' Änderung 19.1.15 Right Click in den Charts de-aktivieren für BMW 
        'Call awinRightClickinPRCCharts()

    End Sub

    


    Private Sub ThisWorkbook_Open() Handles Me.Open


        Dim plantafel As Excel.Window


        CType(Application.Workbooks(myProjektTafel), Excel.Workbook).Activate()

        CType(Application.Worksheets(arrWsNames(3)), Excel.Worksheet).Activate()

        plantafel = Application.ActiveWindow

        With plantafel
            .Caption = windowNames(5)
            .ScrollRow = 1
            .ScrollColumn = 1
            .Visible = True
            .Zoom = 100
        End With


        If appInstance.Windows.Count < 2 Then
            Try
                With appInstance
                    .Windows.Arrange(Excel.XlArrangeStyle.xlArrangeStyleTiled)
                    .Windows(1).WindowState = XlWindowState.xlMaximized
                End With
            Catch ex As Exception
                ' 
            End Try

        End If


        ' hier wird die Projekt Tafel so dargestellt, daß Zeitraum zu sehen ist ... und ein späteres Diagramm 
        ' Änderung 29.06.14 hier nicht mehr notwendig 
        ' Call awinScrollintoView()


    End Sub

    Private Sub ThisWorkbook_BeforeSave(SaveAsUI As Boolean, ByRef Cancel As Boolean) Handles Me.BeforeSave

        'Dim zeitStempel As Date
        'Call MsgBox(" in BeforeSave")

        Cancel = True


        'If AlleProjekte.Count > 0 Then

        '    Call StoreAllProjectsinDB()

        '    zeitStempel = AlleProjekte.First.Value.timeStamp

        '    Call MsgBox("ok, gespeichert!" & vbLf & zeitStempel.ToShortDateString & ", " & zeitStempel.ToShortTimeString)

        '    ' Änderung 18.6 - wenn gespeichert wird, soll die Projekthistorie zurückgesetzt werden 
        '    Try
        '        If projekthistorie.Count > 0 Then
        '            projekthistorie.clear()
        '        End If
        '    Catch ex As Exception

        '    End Try
        'Else
        '    Call MsgBox("keine Projekte zu speichern ...")
        'End If







    End Sub

    Private Sub ThisWorkbook_BeforeClose(ByRef Cancel As Boolean) Handles Me.BeforeClose


        Dim projektespeichern As New frmProjekteSpeichern
        Dim returnValue As DialogResult
        If loginErfolgreich Then


            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)


        'If roentgenBlick.isOn Then
        '    Call awinNoshowProjectNeeds()
        '    With roentgenBlick
        '        .isOn = False
        '        .name = ""
        '        .type = ""
        '    End With
        'End If


        Call awinKontextReset()

        ' tk: nur Fragen , wenn die Datenbank überhaupt läuft 
        Try

            If Request.pingMongoDb() And AlleProjekte.Count > 0 Then
                returnValue = projektespeichern.ShowDialog


                If returnValue = DialogResult.Yes Then

                    Call StoreAllProjectsinDB()

                End If

            Else

                Call MsgBox("keine Projekte zu speichern ...")


            End If
        Catch ex As Exception


        End Try

        appInstance.ScreenUpdating = False
        appInstance.EnableEvents = False

        ' hier sollen jetzt noch die Phasen weggeschrieben werden 
        Try
            Call awinWritePhaseDefinitions()
        Catch ex As Exception
            Call MsgBox("Fehler bei Schreiben Customization File")
        End Try

        End If

        appInstance.ActiveWorkbook.Saved = True


        ' hier wird festgelegt, dass Projectboard.xlsx beim Schließen nicht gespeichert wird, und auch nicht nachgefragt wird.

        Application.Quit()
        appInstance.EnableEvents = True

    End Sub

    Private Sub ThisWorkbook_Shutdown() Handles Me.Shutdown

        'Dim cbar As CommandBar

        ' die Short Cut Menues aus Excel alle wieder aktivieren ...
        'For Each cbar In appInstance.CommandBars

        '    If cbar.Type = MsoBarType.msoBarTypePopup Then
        '        cbar.Enabled = True
        '    End If
        'Next

        'Call MsgBox(" in shutdown")


        Try
            Application.DisplayFormulaBar = True
        Catch ex As Exception

        End Try


        With Application.ActiveWindow
            Try
                .SplitColumn = 0
                .SplitRow = 0
            Catch ex As Exception

            End Try

            Try
                .DisplayWorkbookTabs = True
            Catch ex As Exception

            End Try

            Try
                .GridlineColor = RGB(220, 220, 220)
            Catch ex As Exception

            End Try

            Try
                .FreezePanes = False
            Catch ex As Exception

            End Try

            Try
                .DisplayHeadings = True
            Catch ex As Exception

            End Try

        End With

        appInstance.ShowChartTipNames = True
        appInstance.ShowChartTipValues = True

        Dim anzWindows As Integer = appInstance.Windows.Count


        appInstance.ScreenUpdating = True
        



    End Sub

End Class
