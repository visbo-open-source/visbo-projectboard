Imports Microsoft.Office.Tools.Ribbon
Imports PPTNS = Microsoft.Office.Interop.PowerPoint
Imports MongoDbAccess
Imports ProjectBoardDefinitions
Imports ProjectBoardBasic

Public Class Ribbon1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub activateTab_Click(sender As Object, e As RibbonControlEventArgs) Handles activateTab.Click

        Dim msg As String = ""
        If userIsEntitled(msg) Then

            ' wird das Formular aktuell angezeigt ? 
            If IsNothing(infoFrm) And Not formIsShown Then
                infoFrm = New frmInfo
                formIsShown = True
                infoFrm.Show()
            End If

        Else
            Call MsgBox(msg)
        End If


    End Sub



    Private Sub settingsTab_Click(sender As Object, e As RibbonControlEventArgs) Handles settingsTab.Click

        Dim msg As String = ""

        If userIsEntitled(msg) Then
            Dim settingsfrm As New frmSettings
            With settingsfrm
                Dim res As System.Windows.Forms.DialogResult = .ShowDialog()
            End With
        Else
            Call MsgBox(msg)
        End If

    End Sub

    Private Sub timeMachineTab_Click(sender As Object, e As RibbonControlEventArgs) Handles timeMachineTab.Click
        Dim msg As String = ""

        If userIsEntitled(msg) Then
            ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
            If smartSlideLists.countProjects > 0 Then

                ' muss noch eingeloggt werden ? 
                If noDBAccessInPPT Then

                    Call logInToMongoDB()

                End If

                If Not noDBAccessInPPT Then

                    If Not smartSlideLists.historiesExist Then


                        Dim anzahlProjekte As Integer = smartSlideLists.countProjects
                        ' größter kleinster Wert 
                        Dim gkw As Date = Date.MinValue

                        For i As Integer = 1 To anzahlProjekte
                            Dim tmpName As String = smartSlideLists.getPVName(i)
                            Dim pName As String = getPnameFromKey(tmpName)
                            Dim vName As String = getVariantnameFromKey(tmpName)
                            Dim pvName As String = calcProjektKeyDB(pName, vName)
                            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
                            Dim tsCollection As Collection = request.retrieveZeitstempelFromDB(pvName)
                            ' ermitteln des größten kleinstern Wertes ...
                            ' stellt sicher, dass , wenn mehrere Projekte dargesteltl sind, nur TimeStamps abgerufen werden, die jedes Projekt hat ... 

                            Dim kleinsterWert As Date = Date.Now
                            If Not IsNothing(tsCollection) Then
                                If tsCollection.Count > 0 Then
                                    ' tsCollection ist absteigend sortiert ... 
                                    kleinsterWert = tsCollection.Item(tsCollection.Count)
                                End If
                            End If
                            If kleinsterWert > gkw Then
                                gkw = kleinsterWert
                            End If

                            smartSlideLists.addToListOfTS(tsCollection)
                        Next

                        If anzahlProjekte > 1 Then
                            ' jetzt werden aus der TimeStampListe alle TimeStamps rausgeworfen, die kleiner als der gkw sind ... 
                            smartSlideLists.adjustListOfTS(gkw)
                        End If

                    End If

                    ' jetzt wird das Formular TimeStamps aufgerufen ...
                    Dim tmFormular As New frmPPTTimeMachine
                    Dim dgRes As Windows.Forms.DialogResult = tmFormular.ShowDialog
                    'tmFormular.Show()
                End If

            Else
                Call MsgBox("es gibt auf dieser Seite keine Datenbank-relevanten Informationen ...")
            End If
        Else
            Call MsgBox(msg)
        End If

    End Sub
    
   
    Private Sub variantTab_Click_Click(sender As Object, e As RibbonControlEventArgs) Handles variantTab_Click.Click
        Dim msg As String = ""

        If userIsEntitled(msg) Then
            Dim anzahlProjekte As Integer = smartSlideLists.countProjects
            ' prüfen, ob es eine Smart Slide ist und ob die Projekt-Historien bereits geladen sind ...
            If anzahlProjekte > 0 Then

                ' muss noch eingeloggt werden ? 
                If noDBAccessInPPT Then

                    Call logInToMongoDB()

                End If

                If Not noDBAccessInPPT Then

                    ' aktuell nur für ein Projekt implementiert 
                    If anzahlProjekte = 1 Then
                        Dim tmpName As String = smartSlideLists.getPVName(1)

                        ' jetzt wird das Formular Varianten  aufgerufen ...
                        Dim variantFormular As New frmSelectVariant
                        With variantFormular
                            .pName = getPnameFromKey(tmpName)
                            .vName = getVariantnameFromKey(tmpName)
                        End With

                        Dim dgRes As Windows.Forms.DialogResult = variantFormular.ShowDialog

                    Else
                        Call MsgBox("method not yet implemented ...")

                    End If


                End If

            Else
                Call MsgBox("es gibt auf dieser Seite keine Datenbank-relevanten Informationen ...")
            End If
        Else
            Call MsgBox(msg)
        End If
    End Sub
    
End Class

