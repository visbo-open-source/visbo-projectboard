Imports Microsoft.Office.Tools.Ribbon
Imports PPTNS = Microsoft.Office.Interop.PowerPoint

Public Class Ribbon1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub activateTab_Click(sender As Object, e As RibbonControlEventArgs) Handles activateTab.Click

        visboInfoActivated = Not visboInfoActivated
        If visboInfoActivated Then
            Me.activateTab.Description = "De-Aktivieren"
            Me.activateTab.ScreenTip = "Info-Modus de-aktivieren"
            'Call MsgBox("Info-Modus aktiviert")
        Else
            Me.activateTab.Description = "Aktivieren"
            Me.activateTab.ScreenTip = "Info-Modus aktivieren"
            'Call MsgBox("Info-Modus de-aktiviert")
        End If

    End Sub

    Private Sub settingsTab_Click(sender As Object, e As RibbonControlEventArgs) Handles settingsTab.Click
        Dim settingsfrm As New frmSettings

        With settingsfrm
            Dim res As System.Windows.Forms.DialogResult = .ShowDialog()
        End With

    End Sub
End Class
