Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports ProjectBoardBasic
Imports ProjectBoardDefinitions


Public Class VisboReportRibbon


    Private Sub VisboReportRibbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        ' ''Call MsgBox("VisboReportLoad")

        ' ''Try



        ' ''    awinSettings.awinPath = My.Settings.awinPath

        ' ''    Call awinsetTypenNEW("BHTC")

        ' ''Catch ex As Exception

        ' ''    Call MsgBox(ex.Message)

        ' ''Finally

        ' ''End Try


    End Sub

    Private Sub EPReport_Click(sender As Object, e As RibbonControlEventArgs) Handles EPReport.Click


        Dim reportAuswahl As New frmReportProfil
        Dim hierarchiefenster As New frmHierarchySelection
        Dim returnvalue As DialogResult
        Dim hproj As New clsProjekt
        Dim aktuellesDatum As Date = Date.Now
        Dim filename As String = ""

        'Call MsgBox("EPReport_Click")

        ' Laden des aktuell geladenen Projektes
        Call awinImportMSProject("BHTC", filename, hproj, aktuellesDatum)

        If hproj.name <> "" And Not IsNothing(hproj.name) Then
            reportAuswahl.hproj = hproj
            returnvalue = reportAuswahl.ShowDialog
        End If
     
    End Sub
End Class
