Public Class frmSelectCompareProject
    Private formerindex As Integer


    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        'Dim name1 As String = selectedProjects(1), name2 As String
        Dim value As Integer = formerindex

        If formerindex <> ListBox1.SelectedIndex Then
            ' ein anderer Eintrag wurde selektiert 

            

            
            formerindex = ListBox1.SelectedIndex



        End If


    End Sub

    Private Sub frmSelectCompareProject_Load(sender As Object, e As EventArgs) Handles Me.Load

        For Each kvp As KeyValuePair(Of String, clsProjekt) In ShowProjekte.Liste

            ListBox1.Items.Add(kvp.Key)

        Next
        formerindex = -1

    End Sub

    
    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Try
            If compPhases.Checked = True Then
                Call MsgBox(" deprecated ... - check in frmSelectCompareProject")
                ' jetzt wird das in der Listbox selektierte Projekt mit dem auf 
                ' der Plan-Tafel selektierten Projekt (selectedprojects(1)) 
                ' bezgl der Phasen Charakteristik verglichen
                'Call awinCompareProjectPhases(name1:=selectedProjects(1), _
                '                              name2:=ListBox1.Text, _
                '                              compareType:=3)

            Else
                'Call awinCompareProject(pname1:= name2:=ListBox1.Text, compareType:=3)
            End If
        Catch ex As Exception

        End Try
        
    End Sub
End Class