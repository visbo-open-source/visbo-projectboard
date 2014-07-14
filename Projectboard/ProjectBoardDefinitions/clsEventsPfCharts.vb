Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop

Public Class clsEventsPfCharts

    Public WithEvents PfChartEvents As Chart

    Public WithEvents PfChartRightClick As CommandBarButton

    Private Sub PfChartEvents_Activate() Handles PfChartEvents.Activate

        appInstance.ShowChartTipNames = False
        appInstance.ShowChartTipValues = False

    End Sub

    Private Sub PfChartEvents_BeforeDoubleClick(ByVal ElementID As Integer, ByVal Arg1 As Integer, ByVal Arg2 As Integer, ByRef Cancel As Boolean) Handles PfChartEvents.BeforeDoubleClick

        Cancel = True

    End Sub

    Private Sub PfChartEvents_BeforeRightClick(ByRef Cancel As Boolean) Handles PfChartEvents.BeforeRightClick

        Cancel = True
        'appInstance.CommandBars("awinRightClickinPortfolio").ShowPopup()


    End Sub

    Private Sub PfChartEvents_Deactivate() Handles PfChartEvents.Deactivate

        appInstance.ShowChartTipNames = True
        appInstance.ShowChartTipValues = True

    End Sub

    
    Private Sub PfChartRightClick_Click(ByVal Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles PfChartRightClick.Click

       

        CancelDefault = True

       
        'Select Case Ctrl.Tag

        '    Case "Loesche aus Portfolio"
        '        Call awinDeleteChartorProject(vprojektname:=selectedProjects(1), firstCall:=True) ' löscht das gewählte Projekt
        '    Case "Show / Noshow"
        '        Call awinShowNoShowProject(pname:=selectedProjects(1))
        '    Case "Bearbeiten Projekt-Attribute"
        '        Call MsgBox("noch nicht imlementiert")
        '    Case "Beauftragen"
        '        Call awinBeauftragung(pname:=selectedProjects(1))
        'End Select


    End Sub

    Private Sub PfChartEvents_Select(ByVal ElementID As Integer, ByVal Arg1 As Integer, ByVal Arg2 As Integer) Handles PfChartEvents.Select

        If (ElementID = Excel.XlChartItem.xlSeries) And Arg1 = 1 And Arg2 > 0 Then
            'Dim i As Integer
            Dim pt As Point
            Dim pname As String

            Dim formerUpdate As Boolean = appInstance.ScreenUpdating
            appInstance.ScreenUpdating = False

            Try
                pname = PfChartBubbleNames(Arg2 - 1)
            Catch ex As ArgumentException
                Call MsgBox(" Projekt nicht in Liste vorhanden ...")
                Exit Sub
            End Try


            With Me.PfChartEvents.SeriesCollection(1)
                If .ApplyDataLabels = Excel.XlDataLabelsType.xlDataLabelsShowNone Then
                    .ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowNone)
                    pt = CType(.points(Arg2), Excel.Point)
                    pt.ApplyDataLabels(Type:=Excel.XlDataLabelsType.xlDataLabelsShowLabel)
                    pt.DataLabel.Text = pname
                End If
            End With

            Dim abstand As Integer
            Call awinClkReset(abstand)
            Dim calledFromPf As Boolean = True
            Call awinSelectProjectiST(pname, calledFromPf)

            appInstance.ScreenUpdating = formerUpdate

        Else
            ' nichts ...

        End If
    End Sub
End Class
