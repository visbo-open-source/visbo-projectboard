Imports Microsoft.Office.Tools.Ribbon
Imports PPTNS = Microsoft.Office.Interop.PowerPoint
Imports MongoDbAccess
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic

Public Class Ribbon1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub activateTab_Click(sender As Object, e As RibbonControlEventArgs) Handles activateTab.Click

        
        If userIsEntitled() Then

            ' wird das Formular aktuell angezeigt ? 
            If IsNothing(infoFrm) And Not formIsShown Then
                infoFrm = New frmInfo
                formIsShown = True
                infoFrm.Show()
            End If


        End If

       
    End Sub

    ''' <summary>
    ''' prüft, ob es sich um eine geschützte Präsentation handelt
    ''' kann über pwd, Computer, oder valid Login geschützt werden ... 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function userIsEntitled() As Boolean

        Dim tmpResult As Boolean = False

        If pptAPP.ActivePresentation.Tags.Item(protectionTag) = "PWD" Or _
            pptAPP.ActivePresentation.Tags.Item(protectionTag) = "COMPUTER" Or _
            pptAPP.ActivePresentation.Tags.Item(protectionTag) = "DATABASE" Then

            VisboProtected = True

            If Not protectionSolved Then
                If pptAPP.ActivePresentation.Tags.Item(protectionTag) = "PWD" Then

                    Dim pwdFormular As New frmPassword
                    If pwdFormular.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        If pwdFormular.pwdText.Text = pptAPP.ActivePresentation.Tags.Item(protectionValue) Then
                            ' in allen Slides den Sicht Schutz aufheben 
                            protectionSolved = True
                            Call makeVisboShapesVisible(True)
                        End If
                    End If

                ElseIf pptAPP.ActivePresentation.Tags.Item(protectionTag) = "COMPUTER" Then
                    Dim userName As String = My.Computer.Name
                    If pptAPP.ActivePresentation.Tags.Item(protectionValue) = userName Then
                        ' in allen Slides den Sicht Schutz aufheben 

                        Call makeVisboShapesVisible(True)

                    End If

                ElseIf pptAPP.ActivePresentation.Tags.Item(protectionTag) = "DATABASE" Then
                    ' die Login Maske aufschalten ... 

                End If
            End If

            If protectionSolved Then
                tmpResult = True
            End If
        Else
            tmpResult = True
        End If

        userIsEntitled = tmpResult

    End Function

    Private Sub settingsTab_Click(sender As Object, e As RibbonControlEventArgs) Handles settingsTab.Click


        If userIsEntitled() Then
            Dim settingsfrm As New frmSettings
            With settingsfrm
                Dim res As System.Windows.Forms.DialogResult = .ShowDialog()
            End With
        End If
        
    End Sub

    Private Sub timeMachineTab_Click(sender As Object, e As RibbonControlEventArgs) Handles timeMachineTab.Click

        If userIsEntitled() Then
            ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
            If smartSlideLists.countProjects > 0 Then

                ' muss noch eingeloggt werden ? 
                If noDBAccessInPPT Then
                    ' jetzt die Login Maske aufrufen ... 

                    If awinSettings.databaseURL <> "" And awinSettings.databaseName <> "" Then

                        ' tk: 17.11.16: Einloggen in Datenbank 
                        noDBAccessInPPT = Not loginProzedur()

                        If noDBAccessInPPT Then
                            Call MsgBox("kein Datenbank Zugriff ... ")
                        End If

                    End If

                End If

                If Not noDBAccessInPPT Then

                    If Not smartSlideLists.historiesExist Then

                        ' dazu erst mal alle TimeStamps eines Projektes holen 
                        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

                        Dim anzahlProjekte As Integer = smartSlideLists.countProjects
                        For i As Integer = 1 To anzahlProjekte
                            Dim tmpName As String = smartSlideLists.getPVName(i)
                            Dim pName As String = getPnameFromKey(tmpName)
                            Dim vName As String = getVariantnameFromKey(tmpName)
                            Dim pvName As String = calcProjektKeyDB(pName, vName)

                            Dim tsCollection As Collection = request.retrieveZeitstempelFromDB(pvName)

                            smartSlideLists.addToListOfTS(tsCollection)
                        Next
                        ' jetzt wird das Formular TimeStamps aufgerufen ...
                        Dim tmFormular As New frmPPTTimeMachine
                        Dim dgRes As Windows.Forms.DialogResult = tmFormular.ShowDialog
                    End If

                End If

            Else
                Call MsgBox("es gibt auf dieser Seite keine Datenbank-relevanten Informationen ...")
            End If
        End If

    End Sub
End Class

