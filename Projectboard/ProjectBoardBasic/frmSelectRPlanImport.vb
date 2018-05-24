Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports System.Math
Imports DBAccLayer
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel

Public Class frmSelectRPlanImport

    Public menueAswhl As Integer
    Public dateiOrdner As String
    Public selectedDateiName As String = ""

    Private Sub frmSelectRPlanImport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dateiName As String = ""
        Dim dirname As String = ""

        'dirname = awinPath & rplanimportFilesOrdner
        If menueAswhl = PTImpExp.rplan Then
            dirname = importOrdnerNames(PTImpExp.rplan)
            Me.Text = "RPLAN Excel Dateien auswählen"
        ElseIf menueAswhl = PTImpExp.rplanrxf Then
            dirname = importOrdnerNames(PTImpExp.rplanrxf)
            Me.Text = "RPLAN RXF Dateien auswählen"
        ElseIf menueAswhl = PTImpExp.simpleScen Then
            dirname = importOrdnerNames(PTImpExp.simpleScen)
            Me.Text = "Szenario Dateien auswählen"
        ElseIf menueAswhl = PTImpExp.modulScen Then
            dirname = importOrdnerNames(PTImpExp.modulScen)
            Me.Text = "modulare Szenario Dateien auswählen"
        ElseIf menueAswhl = PTImpExp.addElements Then
            dirname = importOrdnerNames(PTImpExp.addElements)
            Me.Text = "Regel-Dateien auswählen"
        ElseIf menueAswhl = PTImpExp.massenEdit Then
            dirname = importOrdnerNames(PTImpExp.massenEdit)
            Me.Text = "Massen-Edit Datei auswählen"
        End If


        ' jetzt werden die RPLANImportfiles ausgelesen 
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
                                RPLANImportDropbox.Items.Add(dateiName)
                            End If
                        ElseIf menueAswhl = PTImpExp.msproject Then
                            If dateiName.Contains(".mpp") Then
                                RPLANImportDropbox.Items.Add(dateiName)
                            End If

                        Else
                            If dateiName.Contains(".xls") Then
                                RPLANImportDropbox.Items.Add(dateiName)
                            End If
                        End If
                    End If


                Next i
            Catch ex As Exception
                Call MsgBox(ex.Message & ": " & dateiName)
            End Try
        Catch ex As Exception
            Call MsgBox("Folder existiert nicht: " & dirname)
        End Try


        

    End Sub

    Private Sub importRPLAN_Click(sender As Object, e As EventArgs) Handles importRPLAN.Click

        If Not noDB Then
            'Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
        End If

        Dim vglName As String = " "
        Dim myCollection As New Collection

        Dim dirName As String = ""

        'dirName = awinPath & rplanimportFilesOrdner
        If menueAswhl = PTImpExp.rplan Then
            dirName = importOrdnerNames(PTImpExp.rplan)
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
        End If

        selectedDateiName = dirName & "\" & RPLANImportDropbox.Text

        MyBase.Close()
    End Sub

    Private Sub SelectAbbruch_Click(sender As Object, e As EventArgs) Handles SelectAbbruch.Click
        MyBase.Close()
    End Sub


    Private Sub RPLANImportDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RPLANImportDropbox.SelectedIndexChanged


    End Sub

End Class