Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports System.Math
Imports MongoDbAccess
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Public Class frmSelectImportFiles


    Public menueAswhl As Integer
    Public dateiOrdner As String
    Public selectedDateiName As String = ""
    Public selImportFiles As New Collection

    Private Sub frmSelectImportFiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dateiName As String = ""
        Dim dirname As String = ""

        Call defineFrmLanguagesAndVisibility(dirname)

        ' jetzt werden die Importfiles ausgelesen 
        ' Änderung tk 18.3.16 es muss abgefragt werden, ob das Directory überhaupt existiert ... 
        Try
            Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
            Try
                Dim i As Integer
                For i = 1 To listOfImportfiles.Count
                    dateiName = Dir(listOfImportfiles.Item(i - 1))
                    If Not IsNothing(dateiName) Then

                        If menueAswhl = PTImpExp.rplanrxf Then
                            If dateiName.Contains(".rxf") Then
                                ListImportFiles.Items.Add(dateiName)
                            End If
                        ElseIf menueAswhl = PTImpExp.msproject Then
                            If dateiName.Contains(".mpp") Then
                                ListImportFiles.Items.Add(dateiName)
                            End If

                        Else
                            If dateiName.Contains(".xls") Then
                                ListImportFiles.Items.Add(dateiName)
                            End If
                        End If
                    End If


                Next i

                If ListImportFiles.Items.Count > 0 Then
                    ListImportFiles.SelectedIndex = 0
                End If
            Catch ex As Exception
                Call MsgBox(ex.Message & ": " & dateiName)
            End Try
        Catch ex As Exception
            Call MsgBox("Folder existiert nicht: " & dirname)
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            MyBase.Close()
        End Try

    


    End Sub

    Private Sub defineFrmLanguagesAndVisibility(ByRef dirName As String)

        If awinSettings.englishLanguage Then
            alleButton.Text = "All"
            OKButton.Text = "OK"
            SelectAbbruch.Text = "Cancel"
        End If


        If menueAswhl = PTImpExp.visbo Then
            dirName = importOrdnerNames(PTImpExp.visbo)
            If awinSettings.englishLanguage Then
                Me.Text = "select Visbo lean project briefs"
            Else
                Me.Text = "Visbo-Steckbriefe auswählen"
            End If

            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.alleButton.Visible = True

        ElseIf menueAswhl = PTImpExp.rplan Then
            dirName = importOrdnerNames(PTImpExp.rplan)
            If awinSettings.englishLanguage Then
                Me.Text = "select RPLAN Excel files"
            Else
                Me.Text = "RPLAN Excel Dateien auswählen"
            End If

            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.alleButton.Visible = True
        ElseIf menueAswhl = PTImpExp.msproject Then
            dirName = importOrdnerNames(PTImpExp.msproject)
            If awinSettings.englishLanguage Then
                Me.Text = "select MS Project files"
            Else
                Me.Text = "MS-Project Dateien auswählen"
            End If

            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.alleButton.Visible = True
        ElseIf menueAswhl = PTImpExp.rplanrxf Then
            dirName = importOrdnerNames(PTImpExp.rplanrxf)
            If awinSettings.englishLanguage Then
                Me.Text = "select RPLAN RXF-files"
            Else
                Me.Text = "RPLAN RXF Dateien auswählen"
            End If

            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        ElseIf menueAswhl = PTImpExp.simpleScen Then
            dirName = importOrdnerNames(PTImpExp.simpleScen)
            If awinSettings.englishLanguage Then
                Me.Text = "select a portfolio file"
            Else
                Me.Text = "Portfolio Datei auswählen"
            End If

            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        ElseIf menueAswhl = PTImpExp.modulScen Then

            dirName = importOrdnerNames(PTImpExp.modulScen)
            If awinSettings.englishLanguage Then
                Me.Text = "select a modular portfolio file"
            Else
                Me.Text = "modulare Portfolio Datei auswählen"
            End If

            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        ElseIf menueAswhl = PTImpExp.addElements Then
            dirName = importOrdnerNames(PTImpExp.addElements)
            If awinSettings.englishLanguage Then
                Me.Text = "select a rule file"
            Else
                Me.Text = "Regel-Datei auswählen"
            End If


            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        ElseIf menueAswhl = PTImpExp.massenEdit Then
            dirName = importOrdnerNames(PTImpExp.massenEdit)
            If awinSettings.englishLanguage Then
                Me.Text = "select a mass-edit file"
            Else
                Me.Text = "Massen-Edit Datei auswählen"
            End If

            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        ElseIf menueAswhl = PTImpExp.scenariodefs Then
            dirName = importOrdnerNames(PTImpExp.scenariodefs)
            If awinSettings.englishLanguage Then
                Me.Text = "select a portfolio definition file"
            Else
                Me.Text = "Portfolio-Definitions-Datei auswählen"
            End If


            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        End If

    End Sub

    Private Sub ListImportFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListImportFiles.SelectedIndexChanged

     
    End Sub

    Private Sub alleButton_Click(sender As Object, e As EventArgs) Handles alleButton.Click

        Dim element As String = ""
        Dim dirName As String = ""

        If menueAswhl = PTImpExp.visbo Then
            dirName = importOrdnerNames(PTImpExp.visbo)

        ElseIf menueAswhl = PTImpExp.rplan Then
            dirName = importOrdnerNames(PTImpExp.rplan)

        ElseIf menueAswhl = PTImpExp.msproject Then
            dirName = importOrdnerNames(PTImpExp.msproject)

        ElseIf menueAswhl = PTImpExp.rplanrxf Then
            dirName = importOrdnerNames(PTImpExp.rplanrxf)

        ElseIf menueAswhl = PTImpExp.simpleScen Then
            dirName = importOrdnerNames(PTImpExp.simpleScen)

        ElseIf menueAswhl = PTImpExp.modulScen Then
            dirName = importOrdnerNames(PTImpExp.modulScen)

        ElseIf menueAswhl = PTImpExp.addElements Then
            dirName = importOrdnerNames(PTImpExp.addElements)

        ElseIf menueAswhl = PTImpExp.massenEdit Then
            dirName = importOrdnerNames(PTImpExp.massenEdit)

        ElseIf menueAswhl = PTImpExp.scenariodefs Then
            dirName = importOrdnerNames(PTImpExp.scenariodefs)
        End If

        For i = 1 To Me.ListImportFiles.Items.Count
            element = Me.ListImportFiles.Items.Item(i - 1)
            element = dirName & "\" & element

            If selImportFiles.Contains(element) Then
                ' nichts tun 
            Else
                selImportFiles.Add(element)
            End If

        Next

        If selImportFiles.Count < 1 And Me.ListImportFiles.Items.Count < 1 Then

            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            ''Else
            ''    Call MsgBox("bitte wählen sie eine Datei aus")
        End If

        'MyBase.Close()
    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        ' Datenbank wird hier ohnehin nicht benötigt
        'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

        Dim element As String = ""
        Dim dirName As String = ""

        If menueAswhl = PTImpExp.visbo Then
            dirName = importOrdnerNames(PTImpExp.visbo)

        ElseIf menueAswhl = PTImpExp.rplan Then
            dirName = importOrdnerNames(PTImpExp.rplan)

        ElseIf menueAswhl = PTImpExp.msproject Then
            dirName = importOrdnerNames(PTImpExp.msproject)

        ElseIf menueAswhl = PTImpExp.rplanrxf Then
            dirName = importOrdnerNames(PTImpExp.rplanrxf)
            selectedDateiName = dirName & "\" & ListImportFiles.Text
        ElseIf menueAswhl = PTImpExp.simpleScen Then
            dirName = importOrdnerNames(PTImpExp.simpleScen)
            selectedDateiName = dirName & "\" & ListImportFiles.Text
        ElseIf menueAswhl = PTImpExp.modulScen Then
            dirName = importOrdnerNames(PTImpExp.modulScen)
            selectedDateiName = dirName & "\" & ListImportFiles.Text
        ElseIf menueAswhl = PTImpExp.addElements Then
            dirName = importOrdnerNames(PTImpExp.addElements)
            selectedDateiName = dirName & "\" & ListImportFiles.Text
        ElseIf menueAswhl = PTImpExp.massenEdit Then
            dirName = importOrdnerNames(PTImpExp.massenEdit)
            selectedDateiName = dirName & "\" & ListImportFiles.Text
        ElseIf menueAswhl = PTImpExp.scenariodefs Then
            dirName = importOrdnerNames(PTImpExp.scenariodefs)
            selectedDateiName = dirName & "\" & ListImportFiles.Text
        End If


        For i = 1 To Me.ListImportFiles.SelectedItems.Count
            element = Me.ListImportFiles.SelectedItems.Item(i - 1)
            element = dirName & "\" & element

            If selImportFiles.Contains(element) Then
                ' nichts tun 
            Else
                selImportFiles.Add(element)
            End If
      
        Next
        If selImportFiles.Count < 1 Then
            'Call MsgBox("Es wurde keine Datei ausgewählt")
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Else
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            MyBase.Close()
        End If
    End Sub

    Private Sub SelectAbbruch_Click(sender As Object, e As EventArgs) Handles SelectAbbruch.Click
        MyBase.Close()
    End Sub
End Class