Imports ProjectBoardDefinitions
Imports DBAccLayer

Public Class frmSelectVariant
    Friend pName As String = ""
    Friend vName As String = ""
    Private Sub frmSelectVariant_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not noDBAccessInPPT Then

            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            ' existiert der Projekt-Name
            If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, vName, Date.Now) Then
                If vName = "" Then
                    ' zeigen nur an, was nicht bereits aktiv ist 
                    ' also hier nichts tun ...
                Else
                    If CType(databaseAcc, DBAccLayer.Request).projectNameAlreadyExists(pName, "", Date.Now) Then
                        variantNamesListBox.Items.Add("Base-Variant")
                    End If

                End If

                Dim namesCollection As Collection = CType(databaseAcc, DBAccLayer.Request).retrieveVariantNamesFromDB(pName)
                If namesCollection.Count > 0 Then
                    For Each tmpStr As String In namesCollection
                        Try
                            ' zeige nur an, was nicht bereits aktiv ist ...
                            If tmpStr.Trim <> vName Then
                                variantNamesListBox.Items.Add(tmpStr)
                            End If

                        Catch ex As Exception
                        End Try
                    Next
                End If
            End If

        End If

    End Sub

    Private Sub variantNamesListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles variantNamesListBox.SelectedIndexChanged

    End Sub

    Private Sub showButton_Click(sender As Object, e As EventArgs) Handles showButton.Click
        Dim selectedVariantName As String = CStr(variantNamesListBox.SelectedItem)

        ' Übersetzen ...
        If selectedVariantName = "Base-Variant" Then
            selectedVariantName = ""
        End If

        If selectedVariantName <> vName Then
            ' die Aktion durchführen 

            Me.UseWaitCursor = True

            previousTimeStamp = currentTimestamp
            previousVariantName = currentVariantname
            currentVariantname = selectedVariantName

            Dim key As String = CType(currentSlide.Parent, PowerPoint.Presentation).Name

            ' wenn das Projekt noch nicht in der Liste verzeichnet ist ... 
            Dim pvName As String = calcProjektKey(pName, selectedVariantName)
            If pvName <> "" Then
                If smartSlideLists.containsProject(pvName) Then
                    ' nichts tun, ist schon drin ..
                Else
                    smartSlideLists.addProject(pvName)
                End If
            End If

            Call moveAllShapes(True)

            ' das Formular aufschalten 

            If IsNothing(changeFrm) Then
                changeFrm = New frmChanges
                'changeFrm.changeliste = chgeLstListe(currentSlide.SlideID)
                changeFrm.changeliste = chgeLstListe.Item(key).Item(currentSlide.SlideID)
                changeFrm.Show()
            Else
                'changeFrm.changeliste = chgeLstListe(currentSlide.SlideID)
                changeFrm.changeliste = chgeLstListe.Item(key).Item(currentSlide.SlideID)
                changeFrm.neuAufbau()
            End If

            Me.UseWaitCursor = False

        Else
            Call MsgBox("wird bereits angezeigt ...")
        End If
    End Sub
End Class