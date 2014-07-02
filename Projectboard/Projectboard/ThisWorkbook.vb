Imports ProjectBoardDefinitions
Imports ClassLibrary1
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel




Public Class ThisWorkbook
    ' Copyright Philipp Koytek et al. 
    ' 2012 ff
    ' Nicht authorisierte Verwendung nicht gestattet 


    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

    Private Sub ThisWorkbook_Startup() Handles Me.Startup

        'Dim cbar As CommandBar

        appInstance = Application
        

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
        End Try

        anzahlCalls = 0


        
        'Call awinRightClickinPortfolioAendern()
        Call awinRightClickinPRCCharts()

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


        appInstance.EnableEvents = True

        'Application.Quit()

    End Sub


    Private Sub ThisWorkbook_Open() Handles Me.Open


        Dim plantafel As Excel.Window


        'Call MsgBox(" in Open")


        Application.Worksheets(arrWsNames(3)).Activate()
        
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

        If roentgenBlick.isOn Then
            Call awinNoshowProjectNeeds()
            With roentgenBlick
                .isOn = False
                .name = ""
                .type = ""
            End With
        End If

        Call awinKontextReset()

        ' hier sollen jetzt noch die Phasen weggeschrieben werden 
        Call awinWritePhaseDefinitions()


        returnValue = projektespeichern.ShowDialog
        Try

            If returnValue = DialogResult.Yes Then

                If AlleProjekte.Count > 0 Then

                    Call StoreAllProjectsinDB()

                Else
                    Call MsgBox("keine Projekte zu speichern ...")
                End If

            End If

            ' hier wird festgelegt, dass Projectboard.xlsx beim Schließen nicht gespeichert wird, und auch nicht nachgefragt wird.
            Application.ActiveWorkbook.Saved = True
            Application.Quit()

        Catch ex As Exception
            ' Bei Fehler, soll Excel nicht geschlossen werden.
            Call MsgBox(ex.Message)
            Application.ActiveWorkbook.Saved = True
            Cancel = True ' Event Schließen soll nicht ausgeführt werden
        End Try
    End Sub

End Class
