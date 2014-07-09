Imports ProjectBoardDefinitions
Public Class frmDependencies


    Private Sub frmDependencies_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim i As Integer

        statusMeldung.Visible = False

        For i = 1 To selectedProjekte.Count
            If i = 1 Then
                dependentProjectList.Items.Add(selectedProjekte.getProject(i).name)
            Else
                ProjectList.Items.Add(selectedProjekte.getProject(i).name)
            End If
        Next

        degree.Items.Add("schwach")
        degree.Items.Add("stark")
        degree.SelectedItem = ""


    End Sub

    Private Sub moveFromDependent_Click(sender As Object, e As EventArgs) Handles moveFromDependent.Click
        Dim pName As String
        Dim tmpCollection As New Collection
        Dim i As Integer

        For i = 1 To dependentProjectList.SelectedItems.Count
            pName = CStr(dependentProjectList.SelectedItems.Item(i - 1))
            tmpCollection.Add(pName)
        Next

        For i = 1 To tmpCollection.Count
            pName = CStr(tmpCollection.Item(i))

            Try
                If Not ProjectList.Items.Contains(pName) Then
                    ProjectList.Items.Add(pName)
                End If
                dependentProjectList.Items.Remove(pName)
            Catch ex As Exception

            End Try
        Next



    End Sub

    Private Sub copyFromDependent_Click(sender As Object, e As EventArgs) Handles copyFromDependent.Click

        Dim pName As String
        Dim tmpCollection As New Collection
        Dim i As Integer

        For i = 1 To dependentProjectList.SelectedItems.Count
            pName = CStr(dependentProjectList.SelectedItems.Item(i - 1))
            tmpCollection.Add(pName)
        Next

        For i = 1 To tmpCollection.Count
            pName = CStr(tmpCollection.Item(i))

            Try
                If Not ProjectList.Items.Contains(pName) Then
                    ProjectList.Items.Add(pName)
                End If
            Catch ex As Exception

            End Try
        Next



    End Sub

    Private Sub deleteFromDependent_Click(sender As Object, e As EventArgs) Handles deleteFromDependent.Click

        Dim pName As String
        Dim tmpCollection As New Collection
        Dim i As Integer

        For i = 1 To dependentProjectList.SelectedItems.Count
            pName = CStr(dependentProjectList.SelectedItems.Item(i - 1))
            tmpCollection.Add(pName)
        Next

        For i = 1 To tmpCollection.Count
            pName = CStr(tmpCollection.Item(i))

            Try
                dependentProjectList.Items.Remove(pName)
            Catch ex As Exception

            End Try
        Next



    End Sub

    Private Sub moveFromProjects_Click(sender As Object, e As EventArgs) Handles moveFromProjects.Click

        Dim pName As String
        Dim tmpCollection As New Collection
        Dim i As Integer

        For i = 1 To ProjectList.SelectedItems.Count
            pName = CStr(ProjectList.SelectedItems.Item(i - 1))
            tmpCollection.Add(pName)
        Next

        For i = 1 To tmpCollection.Count
            pName = CStr(tmpCollection.Item(i))

            Try
                If Not dependentProjectList.Items.Contains(pName) Then
                    dependentProjectList.Items.Add(pName)
                End If
                ProjectList.Items.Remove(pName)
            Catch ex As Exception

            End Try
        Next



    End Sub

    Private Sub copyFromProjects_Click(sender As Object, e As EventArgs) Handles copyFromProjects.Click

        Dim pName As String
        Dim tmpCollection As New Collection
        Dim i As Integer

        For i = 1 To ProjectList.SelectedItems.Count
            pName = CStr(ProjectList.SelectedItems.Item(i - 1))
            tmpCollection.Add(pName)
        Next

        For i = 1 To tmpCollection.Count
            pName = CStr(tmpCollection.Item(i))

            Try
                If Not dependentProjectList.Items.Contains(pName) Then
                    dependentProjectList.Items.Add(pName)
                End If
            Catch ex As Exception

            End Try
        Next


    End Sub

    Private Sub deleteFromProjects_Click(sender As Object, e As EventArgs) Handles deleteFromProjects.Click

        Dim pName As String
        Dim tmpCollection As New Collection
        Dim i As Integer

        For i = 1 To ProjectList.SelectedItems.Count
            pName = CStr(ProjectList.SelectedItems.Item(i - 1))
            tmpCollection.Add(pName)
        Next

        For i = 1 To tmpCollection.Count
            pName = CStr(tmpCollection.Item(i))

            Try
                ProjectList.Items.Remove(pName)
            Catch ex As Exception

            End Try
        Next

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub dependentProjectList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dependentProjectList.SelectedIndexChanged

        statusMeldung.Visible = False
        Call updateDescriptionAndDegree()

    End Sub

    Private Sub ProjectList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ProjectList.SelectedIndexChanged

        statusMeldung.Visible = False
        Call updateDescriptionAndDegree()

    End Sub

    Private Sub description_TextChanged(sender As Object, e As EventArgs) Handles description.TextChanged

        statusMeldung.Visible = False

    End Sub

    Private Sub updateDescriptionAndDegree()
        Dim lastDescription As String = ""
        Dim lastDegree As Integer
        Dim pName As String
        Dim dpName As String
        Dim firstTime As Boolean = True
        Dim abbruch As Boolean = False
        Dim currentdpndncy As clsDependency

        ' jetzt muss überprüft werden, ob die Description und Degree für alle durch die Selektionen bestimmten Abhängigkeits-Paare identisch ist 
        If dependentProjectList.SelectedItems.Count = 0 Or ProjectList.SelectedItems.Count = 0 Or allDependencies.projectCount = 0 Then
            description.Text = ""
            degree.Text = ""
        Else

            Dim p As Integer = 1
            While p <= ProjectList.SelectedItems.Count And Not abbruch

                pName = CStr(ProjectList.SelectedItems.Item(p - 1))

                Dim d As Integer = 1
                While d <= dependentProjectList.SelectedItems.Count And Not abbruch
                    dpName = CStr(dependentProjectList.SelectedItems(d - 1))
                    currentdpndncy = allDependencies.getDependency(PTdpndncyType.inhalt, pName, dpName)

                    If IsNothing(currentdpndncy) Then
                        abbruch = True
                        description.Text = ""
                        degree.Text = ""
                    Else
                        If firstTime Then
                            firstTime = False
                            lastDescription = currentdpndncy.description
                            lastDegree = currentdpndncy.degree
                        Else
                            If lastDescription <> currentdpndncy.description Then
                                lastDescription = ""
                            End If
                            If lastDegree <> currentdpndncy.degree Then
                                lastDegree = -999
                            End If
                        End If

                        If lastDegree = -999 And lastDescription = "" Then
                            abbruch = True
                            description.Text = ""
                            degree.Text = ""
                        End If

                    End If

                    d = d + 1
                End While

                p = p + 1
            End While

            If Not abbruch Then
                description.Text = lastDescription
                If lastDegree = PTdpndncy.schwach Then
                    degree.Text = "schwach"
                ElseIf lastDegree = PTdpndncy.stark Then
                    degree.Text = "stark"
                Else
                    degree.Text = ""
                End If
            End If
        End If

    End Sub

   
    ''' <summary>
    ''' hier werden alle Paare aus Abhängigkeiten erstellt: alle selektierten Projekte gepaar mit allen selektierten unabhängigen Projekten 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        Dim pName As String, dpName As String
        Dim type As Integer = PTdpndncyType.inhalt
        Dim degreeValue As Integer

        If degree.Text = "schwach" Then
            degreeValue = PTdpndncy.schwach
        ElseIf degree.Text = "stark" Then
            degreeValue = PTdpndncy.stark
        Else
            ' Status-Meldung: Abhängigkeitsgrad wählen, bitte 
            Call MsgBox("Abhängigkeit wählen, bitte ")
            Exit Sub
        End If

        Dim anzahlDP As Integer = 0
        Dim p As Integer
        For p = 1 To ProjectList.SelectedItems.Count
            pName = CStr(ProjectList.SelectedItems.Item(p - 1))
            Dim d As Integer
            For d = 1 To dependentProjectList.SelectedItems.Count
                dpName = CStr(dependentProjectList.SelectedItems.Item(d - 1))

                If pName <> dpName Then
                    ' jetzt die dependency erstellen
                    anzahlDP = anzahlDP + 1
                    allDependencies.Add(pName, dpName, PTdpndncyType.inhalt, degreeValue, description.Text)
                End If

            Next

        Next


        If anzahlDP = 0 Then
            statusMeldung.Text = "keine neue Abhängigkeit erstellt ! "
        ElseIf anzahlDP = 1 Then
            statusMeldung.Text = "eine neue Abhängigkeit erstellt ! "
        ElseIf anzahlDP > 1 Then
            statusMeldung.Text = anzahlDP & " neue Abhängigkeiten erstellt ! "

        End If

        statusMeldung.Visible = True

    End Sub

    Private Sub degree_SelectedIndexChanged(sender As Object, e As EventArgs) Handles degree.SelectedIndexChanged

        statusMeldung.Visible = False

    End Sub
End Class