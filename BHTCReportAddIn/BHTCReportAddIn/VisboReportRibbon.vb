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

        Try

            Dim reportAuswahl As New frmReportProfil
            Dim hierarchiefenster As New frmHierarchySelection
            Dim returnvalue As DialogResult
            Dim hproj As New clsProjekt
            Dim aktuellesDatum = Date.Now
            Dim validDatum As Date = "29.Feb.2016"
            Dim filename As String = ""

            ''If MsgBox("Lizenz prüfen?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            If aktuellesDatum > validDatum Then

                ' Testen, ob der User die passende Lizenz besitzt
                Dim user As String = myWindowsName
                Dim komponente As String = LizenzKomponenten(PTSWKomp.Swimlanes2)     ' Swimlanes2

                ' Lesen des Lizenzen-Files

                Dim lizenzen As clsLicences = XMLImportLicences(awinPath & licFileName)

                ' Prüfen der Lizenzen
                If lizenzen.validLicence(user, komponente) Then


                    'Call MsgBox("EPReport_Click")

                    ' Laden des aktuell geladenen Projektes
                    Call awinImportMSProject("BHTC", filename, hproj, aktuellesDatum)

                    If hproj.name <> "" And Not IsNothing(hproj.name) Then
                        reportAuswahl.hproj = hproj
                        returnvalue = reportAuswahl.ShowDialog
                    End If
                Else
                    Call MsgBox("Aktueller User " & myWindowsName & " hat keine passende Lizenz!" _
                                & vbLf & " Bitte kontaktieren Sie ihren Systemadministrator")

                End If


            Else    ' ohne Lizenzprüfung

                ' Laden des aktuell geladenen Projektes
                Call awinImportMSProject("BHTC", filename, hproj, aktuellesDatum)

                If hproj.name <> "" And Not IsNothing(hproj.name) Then
                    reportAuswahl.hproj = hproj
                    returnvalue = reportAuswahl.ShowDialog
                End If

            End If ' Ende if Lizenzprüfung

        Catch ex As Exception

            Throw New ArgumentException(" Bitte kontaktieren Sie ihren Systemadministrator")
            Call MsgBox(" Bitte kontaktieren Sie ihren Systemadministrator")
        End Try
    End Sub
End Class
