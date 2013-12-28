Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

Public Class frmConfirmEditRess

    Public selectedProject As String

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

        'With appInstance.Worksheets(arrWsNames(3))
        '    .activate()
        'End With

        'Änderung 30.7.13 Screenupdating = false gesetzt , damit das Geflacker aufhört 
        appInstance.ScreenUpdating = False
        enableOnUpdate = True
        MyBase.Close()

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        Dim hproj As clsProjekt, newproj As New clsProjekt
        Dim pname As String
        Dim tryzeile As Integer
        Dim key As String
        Dim shpElement As Shape
        Dim tmpShapes As Shapes



        Try
            pname = selectedProject
            hproj = ShowProjekte.getProject(pname)
            key = hproj.name & "#" & hproj.variantName
            tryzeile = hproj.tfZeile
        Catch ex As Exception
            Call MsgBox(ex.Message & " in frmConfirmEditRess! - Abbruch")
            Exit Sub
        End Try

        Try
            enableOnUpdate = False
            hproj.copyAttrTo(newproj, False)

            With newproj
                .name = hproj.name
                .variantName = hproj.variantName
                If hproj.Status = ProjektStatus(1) Then
                    .Status = ProjektStatus(2)
                End If
            End With
            Dim shpUID As String = hproj.shpUID
            Call awinReadProjFromEditRess(newproj)

            ' jetzt müssen die Ampel Bewertung und Meilenstein Bwertungen übernommen werden ...
            hproj.copyBewertungenTo(newproj)

            'Änderung 30.7.13 Screenupdating = false gesetzt , damit das Geflacker aufhört 
            appInstance.ScreenUpdating = False

            ' jetzt muss die Projektdarstellung gelöscht werden ...
            Call clearProjektinPlantafel(pname)

            ' Änderung 26.7: roentgenblick ison wird jetzt in clearProjektinPlantafel behandelt
            'If roentgenBlick.isOn Then
            '    Call NoshowNeedsofProject(pname)
            'End If

            If pname = newproj.name Then

                Try
                    ShowProjekte.Remove(pname)
                    AlleProjekte.Remove(key)
                Catch ex As Exception

                End Try

            End If

            ' Änderung 18.6 : SHPUID = "" wichtig, weil sonst das Paar Shape-ID, Projekt zwiemal eingetragen wird 
            newproj.shpUID = ""
            newproj.timeStamp = Date.Now
            

            ShowProjekte.Add(newproj)
            AlleProjekte.Add(key, newproj)
            pname = newproj.name

            ' dann muss das Projekt neu gezeichnet werden - muss gemacht werden; es könnte sich ja die Darstellung geändert haben 

            ' Änderung 26.7 roentgenblick ison wird jetzt in zeichneProjektinPlantafel behandelt
            Call ZeichneProjektinPlanTafel(pname, tryzeile, False)
            shpUID = newproj.shpUID

            

            ' dann müssen die Diagramme neu gezeichnet werden 
            Call awinNeuZeichnenDiagramme(3)

            'enableOnUpdate = True

            ' jetzt muss der Select des neuen Shapes gemacht werden 

            Try

                With appInstance.Worksheets(arrWsNames(3))
                    tmpShapes = .shapes
                    shpElement = tmpShapes.Item(newproj.name)
                    shpElement.Select()
                End With

            Catch ex As Exception

            End Try

            
            MyBase.Close()

        Catch ex As Exception
            enableOnUpdate = True
            Call MsgBox(ex.Message)

        End Try
        ' Änderung 8.11 
        'enableOnUpdate = True
    End Sub

    Private Sub frmConfirmEditRess_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        'Änderung 8.11.13 Screenupdating = false gesetzt , damit das Geflacker aufhört 
        appInstance.ScreenUpdating = False

        frmCoord(PTfrm.editRess, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.editRess, PTpinfo.left) = Me.Left
       

        With appInstance.Worksheets(arrWsNames(3))
            .activate()
        End With

        enableOnUpdate = True
    End Sub


    Private Sub frmConfirmEditRess_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.Top = frmCoord(PTfrm.editRess, PTpinfo.top)
        Me.Left = frmCoord(PTfrm.editRess, PTpinfo.left)

        enableOnUpdate = False

    End Sub
End Class