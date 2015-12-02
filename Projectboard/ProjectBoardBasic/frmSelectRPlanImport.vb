Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports System.Math
Imports MongoDbAccess
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
        ElseIf menueAswhl = PTImpExp.rplanrxf Then
            dirname = importOrdnerNames(PTImpExp.rplanrxf)
        ElseIf menueAswhl = PTImpExp.simpleScen Then
            dirname = importOrdnerNames(PTImpExp.simpleScen)
        ElseIf menueAswhl = PTImpExp.modulScen Then
            dirname = importOrdnerNames(PTImpExp.modulScen)
        ElseIf menueAswhl = PTImpExp.addElements Then
            dirname = importOrdnerNames(PTImpExp.addElements)
       
        End If


        ' jetzt werden die RPLANImportfiles ausgelesen 

        Dim listOfImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
        Try
            Dim i As Integer
            For i = 1 To listOfImportfiles.Count
                dateiName = Dir(listOfImportfiles.Item(i - 1))
                If Not IsNothing(dateiName) Then
                    If dirname = importOrdnerNames(PTImpExp.rplanrxf) Then
                        If dateiName.Contains(".rxf") Then
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

    End Sub

    Private Sub importRPLAN_Click(sender As Object, e As EventArgs) Handles importRPLAN.Click

        Dim request As New Request(awinSettings.databaseURL, awinSettings.databaseName, dbUsername, dbPasswort)
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
        End If

        selectedDateiName = dirName & "\" & RPLANImportDropbox.Text

        MyBase.Close()
    End Sub

    Private Sub SelectAbbruch_Click(sender As Object, e As EventArgs) Handles SelectAbbruch.Click
        MyBase.Close()
    End Sub


    Private Sub RPLANImportDropbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RPLANImportDropbox.SelectedIndexChanged
        ' hier muss die selektierte Vorlage genommen werden, um damit den dann bei OK-Button Click den Report anzustoßen
        Dim newTemplate As String = RPLANImportDropbox.Text
    End Sub

End Class