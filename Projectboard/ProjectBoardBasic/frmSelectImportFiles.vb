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

        'dirname = awinPath & rplanimportFilesOrdner
        If menueAswhl = PTImpExp.visbo Then
            dirname = importOrdnerNames(PTImpExp.visbo)
            Me.Text = "Visbo-Steckbriefe auswählen"
            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.alleButton.Visible = True
        ElseIf menueAswhl = PTImpExp.rplan Then
            dirname = importOrdnerNames(PTImpExp.rplan)
            Me.Text = "RPLAN Excel Dateien auswählen"
            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.alleButton.Visible = True
        ElseIf menueAswhl = PTImpExp.msproject Then
            dirname = importOrdnerNames(PTImpExp.msproject)
            Me.Text = "MS-Project Dateien auswählen"
            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
            Me.alleButton.Visible = True
        ElseIf menueAswhl = PTImpExp.rplanrxf Then
            dirname = importOrdnerNames(PTImpExp.rplanrxf)
            Me.Text = "RPLAN RXF Dateien auswählen"
            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        ElseIf menueAswhl = PTImpExp.simpleScen Then
            dirname = importOrdnerNames(PTImpExp.simpleScen)
            Me.Text = "Szenario Dateien auswählen"
            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        ElseIf menueAswhl = PTImpExp.modulScen Then
            dirname = importOrdnerNames(PTImpExp.modulScen)
            Me.Text = "modulare Szenario Dateien auswählen"
            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        ElseIf menueAswhl = PTImpExp.addElements Then
            dirname = importOrdnerNames(PTImpExp.addElements)
            Me.Text = "Regel-Dateien auswählen"
            Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.One
            Me.alleButton.Visible = False
        End If


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

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)

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