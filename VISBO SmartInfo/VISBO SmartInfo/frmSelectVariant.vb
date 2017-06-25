Imports ProjectBoardDefinitions
Imports MongoDbAccess
Imports ProjectBoardBasic
Public Class frmSelectVariant
    Friend pName As String = ""
    Friend vName As String = ""
    Private Sub frmSelectVariant_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Not noDBAccessInPPT Then

            Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
            ' existiert der Projekt-Name
            If request.projectNameAlreadyExists(pName, vName, Date.Now) Then
                If vName = "" Then
                    ' zeigen nur an, was nicht bereits aktiv ist 
                    ' also hier nichts tun ...
                Else
                    If request.projectNameAlreadyExists(pName, "", Date.Now) Then
                        variantNamesListBox.Items.Add("Base-Variant")
                    End If

                End If

                Dim namesCollection As Collection = request.retrieveVariantNamesFromDB(pName)
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
            previousTimeStamp = currentTimestamp
            previousVariantName = currentVariantname
            currentVariantname = selectedVariantName

            'Call moveAllShapes()

        Else
            Call MsgBox("wird bereits angezeigt ...")
        End If
    End Sub
End Class