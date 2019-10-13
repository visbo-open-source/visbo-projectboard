Imports Microsoft.Office.Interop.PowerPoint

Public Class ThisAddIn

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon()
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_SlideSelectionChanged(SldRange As SlideRange) Handles Application.SlideSelectionChanged
        If SldRange.Count = 1 Then
            curSlide = SldRange.Item(1)
        End If
    End Sub

    Private Sub Application_WindowActivate(Pres As Presentation, Wn As DocumentWindow) Handles Application.WindowActivate
        curPresentation = Pres
    End Sub
End Class
