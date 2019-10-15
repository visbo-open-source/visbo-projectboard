Imports Microsoft.Office.Interop.PowerPoint
Imports ProjectBoardDefinitions

Public Class ThisAddIn

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon()
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
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
