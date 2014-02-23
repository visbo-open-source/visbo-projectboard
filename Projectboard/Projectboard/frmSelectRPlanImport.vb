Imports ClassLibrary1
Imports ProjectBoardDefinitions
Imports System.Math
Imports MongoDbAccess
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel

Public Class frmSelectRPlanImport


    Public RPLANdateiName As String = ""

    Private Sub frmSelectRPlanImport_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dateiName As String = ""
        Dim dirname As String = ""

        dirname = awinPath & rplanimportFilesOrdner

        ' jetzt werden die RPLANImportfiles ausgelesen 

        Dim listOfRPLANImportfiles As Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Computer.FileSystem.GetFiles(dirname)
        Try
            For i = 1 To listOfRPLANImportfiles.Count
                dateiName = Dir(listOfRPLANImportfiles.Item(i - 1))
                RPLANImportDropbox.Items.Add(dateiName)
            Next i
        Catch ex As Exception
            Call MsgBox(ex.Message & ": " & dateiName)
        End Try

    End Sub

    Private Sub importRPLAN_Click(sender As Object, e As EventArgs) Handles importRPLAN.Click

        Dim request As New Request(awinSettings.databaseName)
        Dim vglName As String = " "
        Dim myCollection As New Collection

        Dim dirName As String

        dirName = awinPath & rplanimportFilesOrdner
        RPLANdateiName = dirName & "\" & RPLANImportDropbox.Text

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