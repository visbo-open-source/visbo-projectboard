Public Class frmChanges


    Private Sub frmChanges_FormClosed(sender As Object, e As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        changeFrm = Nothing

        Call undimAllShapes()

        ' Koordinaten merken
        frmCoord(PTfrm.changes, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.changes, PTpinfo.left) = Me.Left


    End Sub

    Private Sub frmChanges_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If frmCoord(PTfrm.changes, PTpinfo.top) > 0 Then
            Me.Top = frmCoord(PTfrm.changes, PTpinfo.top)
            Me.Left = frmCoord(PTfrm.changes, PTpinfo.left)
        Else
            Me.Top = 922
            Me.Left = 24
        End If

        Call listeAufbauen()

        'If Me.Height > changeListTable.Height + 38 Then
        '    Me.Height = changeListTable.Height + 38
        'End If
        

    End Sub

    ''' <summary>
    ''' setzt im Fall englisch die Formular Texte auf englische Bezeichner 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub languageSettings()

        If englishLanguage Then
            With Me
                .Text = "List of Changes"
            End With
        End If

    End Sub

    Private Sub listeAufbauen()
        Dim tmpPreviousVname As String = previousVariantName, tmpCurrentVname As String = currentVariantname
        Dim showVariantMode As Boolean = False
        If currentVariantname <> previousVariantName And _
            previousTimeStamp = currentTimestamp Then

            showVariantMode = True
            If previousVariantName = "" Then
                tmpPreviousVname = "Base-Variant"
            End If
            If currentVariantname = "" Then
                tmpCurrentVname = "Base-Variant"
            End If
            changeListTable.Columns(2).HeaderText = "Variant " & tmpPreviousVname
            changeListTable.Columns(3).HeaderText = "Variant " & tmpCurrentVname

        Else
            If previousTimeStamp < currentTimestamp Then
                changeListTable.Columns(2).HeaderText = "Version " & previousTimeStamp.ToShortDateString
                changeListTable.Columns(3).HeaderText = "Version " & currentTimestamp.ToShortDateString
            ElseIf previousTimeStamp > currentTimestamp Then
                changeListTable.Columns(2).HeaderText = "Version " & currentTimestamp.ToShortDateString
                changeListTable.Columns(3).HeaderText = "Version " & previousTimeStamp.ToShortDateString
            End If
        End If
        

        Dim anzChangeItems As Integer = changeListe.getChangeListCount


        If anzChangeItems > 0 Then

            changeListTable.Rows.Add(anzChangeItems)

            For i As Integer = 0 To anzChangeItems - 1

                Dim currentItem As clsChangeItem = changeListe.getExplanationFromChangeList(i + 1)

                With currentItem
                    If .vName = "" Then
                        changeListTable.Rows(i).Cells(0).Value = .pName
                    Else
                        changeListTable.Rows(i).Cells(0).Value = .pName & " [" & .vName & "]"
                    End If

                    changeListTable.Rows(i).Cells(1).Value = .bestElemName
                    If showVariantMode Then
                        changeListTable.Rows(i).Cells(2).Value = .oldValue
                        changeListTable.Rows(i).Cells(3).Value = .newValue
                    Else
                        If previousTimeStamp < currentTimestamp Then
                            changeListTable.Rows(i).Cells(2).Value = .oldValue
                            changeListTable.Rows(i).Cells(3).Value = .newValue
                        Else
                            changeListTable.Rows(i).Cells(2).Value = .newValue
                            changeListTable.Rows(i).Cells(3).Value = .oldValue
                        End If
                    End If

                    changeListTable.Rows(i).Cells(4).Value = .diffInDays
                End With

                changeListTable.Rows(i).Tag = changeListe.getShapeNameFromChangeList(i + 1)

            Next
        End If
    End Sub

    Friend Sub neuAufbau()

        changeListTable.Rows.Clear()
        Call listeAufbauen()

    End Sub



    Private Sub changeListTable_SelectionChanged(sender As Object, e As EventArgs) Handles changeListTable.SelectionChanged
        Dim nameArrayO() As String
        Dim tmpCollection As New Collection
        Dim anzSelected As Integer = 0

        For i As Integer = 1 To changeListTable.SelectedRows.Count

            Dim tmpShpName As String = changeListTable.SelectedRows.Item(i - 1).Tag
            If Not IsNothing(tmpShpName) Then
                If tmpShpName <> "" Then
                    If Not tmpCollection.Contains(tmpShpName) Then
                        tmpCollection.Add(tmpShpName, tmpShpName)
                    End If
                End If
            End If
            
        Next

        anzSelected = tmpCollection.Count

        ' vorher alle ggf abgedimmten Shapes wieder voll anzeigen 
        Call undimAllShapes()



        If anzSelected >= 1 Then

            ' wenn das erste Element selektiert wird und die Anzahl Marker > 0 ist, dann müssen hier die MArker gelöscht werden 
            If changeListTable.SelectedRows.Count = 1 And markerShpNames.Count > 0 Then
                Call deleteMarkerShapes()
            End If

            ReDim nameArrayO(anzSelected - 1)

            For i As Integer = 0 To anzSelected - 1
                nameArrayO(i) = CStr(tmpCollection.Item(i + 1))
            Next

            Try
                selectedPlanShapes = currentSlide.Shapes.Range(nameArrayO)
                selectedPlanShapes.Select()

                ' jetzt werden alle anderen - relevanten Visbo Shapes - mit Transparenz 80% dargestellt 
                Call dimAllShapesExceptThese(nameArrayO)

                ' jetzt werden noch die Schatten-Previous-Version Shapes gezeichnet ...  
                Call zeichneShadows(nameArrayO, False)

            Catch ex As Exception

            End Try

        Else
            ' nichts tun ...

        End If

    End Sub

   
End Class