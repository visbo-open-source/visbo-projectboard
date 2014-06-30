Public Class frmGetProjectbackFromNoshow

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Dim pname As String
        Dim selectString As String
        Dim hproj As clsProjekt
        Dim tfz As Integer
        Dim toDoListe As New Collection
        Dim atleastOne As Boolean = False
        Dim i As Integer


        For Each selectString In ListBox1.SelectedItems
            If selectString <> "" Then
                pname = selectString
                If AlleProjekte.ContainsKey(pname) Then

                    Try
                        hproj = AlleProjekte.Item(pname)
                        ShowProjekte.Add(hproj)
                    Catch ex As Exception
                        Call MsgBox(" Fehler - kann nicht in Show übernommen werden " & ex.Message)
                        Exit Sub
                    End Try

                    atleastOne = True

                    With hproj

                        tfz = .tfZeile

                    End With

                    ' Änderung 26.7 wird jetzt in zeichneProjektinPlantafel gemacht 
                    'If roentgenBlick.isOn Then
                    '    With roentgenBlick
                    '        Call awinShowNeedsofProject1(mycollection:=.myCollection, type:=.type, projektname:=pname)
                    '    End With

                    'End If

                   

                    Try
                        
                        toDoListe.Add(pname)
                        Dim shortName As String = hproj.name

                        ' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
                        ' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
                        Dim tmpCollection As New Collection
                        Call ZeichneProjektinPlanTafel(tmpCollection, shortName, tfz)
                    Catch ex As Exception

                    End Try


                Else
                    Call MsgBox("Projekt " & pname & " wurde nicht gefunden")
                End If
            End If
        Next

        For i = 1 To toDoListe.Count
            pname = toDoListe.Item(i)
            ListBox1.Items.Remove(pname)
        Next

        If atleastOne Then
            Call awinNeuZeichnenDiagramme(2)
            MyBase.Close()
        Else
            Call MsgBox(" bitte selektieren Sie mindestens ein Projekt")
        End If
        ' mindestens ein Projekt wurde eingefügt 


    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click
        MyBase.Close()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub

    Private Sub frmGetProjectbackFromNoshow_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For Each kvp As KeyValuePair(Of String, clsProjekt) In AlleProjekte

            If ShowProjekte.Liste.ContainsValue(kvp.Value) Then
                ' nichts tun 
            Else
                ' in die Liste schreiben 
                ListBox1.Items.Add(kvp.Key)
            End If


        Next

    End Sub
End Class