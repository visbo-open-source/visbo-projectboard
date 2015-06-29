Imports xlNS = Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop

Public Class clsEventsPfCharts

    Public WithEvents PfChartEvents As xlNS.Chart

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

            'Dim abstand As Integer
            'Call awinClkReset(abstand)
            Dim calledFromPf As Boolean = True
            Call awinSelectProjectiST(pname, calledFromPf)

            appInstance.ScreenUpdating = formerUpdate

        Else
            ' nichts ...

        End If
    End Sub
    Private Sub PfChartEvents_Resize() Handles PfChartEvents.Resize


        Dim chtobj As xlNS.ChartObject, chtobj1 As xlNS.ChartObject
        Dim IDKennung As String
        Dim foundDiagram As clsDiagramm
        Dim kFontsize As Double
        Dim achsenFontsize As Double
        Dim axisTitleFontsize As Double
        Try
            chtobj = CType(Me.PfChartEvents.Parent, Microsoft.Office.Interop.Excel.ChartObject)
            Try
                chtobj1 = CType(appInstance.ActiveChart.Parent, Microsoft.Office.Interop.Excel.ChartObject)
            Catch ex As Exception

            End Try

            IDKennung = chtobj.Name
            foundDiagram = DiagramList.getDiagramm(IDKennung)

            kFontsize = (chtobj.Width / foundDiagram.width)
            'kHeight = (chtobj.Height / foundDiagram.height)

            With chtobj.Chart

                ' Schriftgröße der Überschrift anpassen
                If .HasTitle Then
                    .ChartTitle.Format.TextFrame2.TextRange.Font.Size = CType(.ChartTitle.Format.TextFrame2.TextRange.Font.Size * kFontsize, Single)
                End If

                ' Schriftgröße der Legende anpassen
                If .HasLegend Then
                    With .Legend
                        .Format.TextFrame2.TextRange.Font.Size = CType(.Format.TextFrame2.TextRange.Font.Size * kFontsize, Single)
                    End With
                End If

                ' Schriftgröße der x-Achse anpassen
                Try
                    With CType(.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary), Excel.Axis)
                        achsenFontsize = CType(.TickLabels.Font.Size * kFontsize, Double)
                        .TickLabels.Font.Size = .TickLabels.Font.Size * kFontsize

                        If .HasTitle Then
                            With .AxisTitle
                                axisTitleFontsize = CType(.Characters.Font.Size * kFontsize, Double)
                                .Characters.Font.Size = .Characters.Font.Size * kFontsize
                            End With
                        End If

                    End With
                Catch ex As Exception

                End Try


                ' Schriftgröße der y-Achse anpassen
                Try
                    With CType(.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue, Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary), Excel.Axis)
                        '.TickLabels.Font.Size = .Ticklabels.Font.Size * kFontsize
                        .TickLabels.Font.Size = achsenFontsize

                        If .HasTitle Then
                            With .AxisTitle
                                .Characters.Font.Size = axisTitleFontsize
                            End With
                        End If
                    End With
                Catch ex As Exception

                End Try


                ' Schriftgröße der eingezeichenten Daten bestimmen
                If .SeriesCollection.Count > 0 Then

                    For j = 1 To .SeriesCollection.Count

                        With CType(.SeriesCollection(j), Excel.Series)
                            For i = 1 To .Points.count
                                With .Points(i)
                                    If .hasdatalabel = True Then
                                        .DataLabel.Font.Size = .DataLabel.Font.Size * kFontsize
                                    End If
                                End With

                            Next i

                        End With

                    Next j

                End If


            End With


            With foundDiagram
                .top = chtobj.Top
                .left = chtobj.Left
                .width = chtobj.Width
                .height = chtobj.Height
            End With
        Catch ex As Exception
            'Call MsgBox("konnte das Chart nicht in der Diagramm-Liste finden ...")
        End Try


        enableOnUpdate = True

    End Sub
End Class
