Imports Microsoft.Office.Tools.Ribbon
Imports PPTNS = Microsoft.Office.Interop.PowerPoint

Public Class Ribbon1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub activateTab_Click(sender As Object, e As RibbonControlEventArgs) Handles activateTab.Click

        Dim alreadyProtected As Boolean = VisboProtected
        visboInfoActivated = Not visboInfoActivated

        If visboInfoActivated Then

            If pptAPP.ActivePresentation.Tags.Item(protectionTag) = "PWD" And _
                Not alreadyProtected Then
                ' Formular zur Password Eingabe aufrufen 
                VisboProtected = True

                ' Formular ... 
                Dim pwdFormular As New frmPassword
                If pwdFormular.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    If pwdFormular.pwdText.Text = pptAPP.ActivePresentation.Tags.Item(protectionValue) Then
                        ' in allen Slides den Sicht Schutz aufheben 

                        Call makeVisboShapesVisible(True)
                    End If
                End If

                ' if richtig 
            ElseIf pptAPP.ActivePresentation.Tags.Item(protectionTag) = "COMPUTER" And _
                Not alreadyProtected Then
                ' überprüfen, ob es die richtige Domain ist 
                VisboProtected = True

                Dim userName As String = My.Computer.Name
                If pptAPP.ActivePresentation.Tags.Item(protectionValue) = userName Then
                    ' in allen Slides den Sicht Schutz aufheben 

                    Call makeVisboShapesVisible(True)

                End If
            End If

            Me.activateTab.Label = "De-Aktivieren"
            Me.activateTab.ScreenTip = "Info-Modus de-aktivieren"
            'Call MsgBox("Info-Modus aktiviert")
        Else
            Me.activateTab.Label = "Aktivieren"
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

    Private Sub timeMachineTab_Click(sender As Object, e As RibbonControlEventArgs) Handles timeMachineTab.Click

        ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
        If smartSlideLists.countProjects > 0 Then
            ' nur dann müssen Historien geholt werden 
            If noDBAccessInPPT Then
                Call MsgBox("kein Datenbank Zugriff ... bitte erst einloggen ...")
            Else
                If Not smartSlideLists.historiesExist Then
                    ' für jedes Projekt die ProjektHistorie holen ...
                    Dim anzahlProjekte As Integer = smartSlideLists.countProjects
                    For i As Integer = 1 To anzahlProjekte
                        Dim pvName As String = smartSlideLists.getPVName(i)



                    Next

                End If
            End If
        End If

    End Sub
End Class
