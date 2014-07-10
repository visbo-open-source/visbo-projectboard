Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

Public Class frmConfirmEditRess

    Public selectedProject As String
    Private okButtonClicked As Boolean

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

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
            key = calcProjektKey(hproj)
            tryzeile = hproj.tfZeile
        Catch ex As Exception
            Call MsgBox(ex.Message & vbLf & " in frmConfirmEditRess! - Abbruch")
            Exit Sub
        End Try

        Try

            ' Änderung 3.7.14 tk: jetzt dürfen nur noch die Werte der 
            ' existierenden Phasen/Rollen/Kostenarten geändert werden
            Call awinChangeProjFromEditRess(hproj)
            hproj.timeStamp = Date.Now

            ' ----------------------------------------------------
            ' alte Version 
            'hproj.copyAttrTo(newproj)

            'With newproj
            '    .name = hproj.name
            '    .variantName = hproj.variantName
            '    If hproj.Status = ProjektStatus(1) Then
            '        .Status = ProjektStatus(2)
            '    End If
            'End With
            'Dim shpUID As String = hproj.shpUID
            'Call awinReadProjFromEditRess(newproj)

            '' jetzt müssen die Ampel Bewertung und Meilenstein Bwertungen übernommen werden ...
            'hproj.copyBewertungenTo(newproj)

            

            ' jetzt wird das Formular geschlossen ; das ist notwendig, bevor das Shpae geschrieben wird 
            ' es kann sonst zu Seiteneffekten kommen, dass die Shape Koordinaten geändert werden 
            ' durch mybase.close wird auch gewechselt auf Tabelle1 ...


            ''Änderung 30.7.13 Screenupdating = false gesetzt , damit das Geflacker aufhört 
            appInstance.ScreenUpdating = False

            okButtonClicked = True ' damit bei Form_closed kein Enableonupdate bzw ScreenUpdate gemacht wird 
            MyBase.Close()

            ' -------------------------------------------------------------
            ' alte Version, wo noch alles mögliche geändert werden durfte 

            '' jetzt muss die Projektdarstellung gelöscht werden ...
            'Call clearProjektinPlantafel(pname)


            'If pname = newproj.name Then

            '    Try
            '        ShowProjekte.Remove(pname)
            '        AlleProjekte.Remove(key)
            '    Catch ex As Exception

            '    End Try

            'End If

            '' Änderung 18.6 : SHPUID = "" wichtig, weil sonst das Paar Shape-ID, Projekt zwiemal eingetragen wird 
            'newproj.shpUID = ""
            'newproj.timeStamp = Date.Now

            '
            ' alte Version  
            'ShowProjekte.Add(newproj)
            'AlleProjekte.Add(key, newproj)
            'pname = newproj.name

            '' dann muss das Projekt neu gezeichnet werden - muss gemacht werden; es könnte sich ja die Darstellung geändert haben 

            '' Änderung 26.7 roentgenblick ison wird jetzt in zeichneProjektinPlantafel behandelt

            '' wenn bestimmte Projekte beim Suchen nach einem Platz nicht berücksichtigt werden sollen,
            '' dann müssen sie in einer Collection an ZeichneProjektinPlanTafel übergeben werden 
            'Dim tmpCollection As New Collection
            'Call ZeichneProjektinPlanTafel(tmpCollection, pname, tryzeile)
            'shpUID = newproj.shpUID

            
            ' die Diagramme müssen auf alle Fälle neu gezeichnet werden 
            ' dann müssen die Diagramme neu gezeichnet werden 
            Call awinNeuZeichnenDiagramme(3)

            ' jetzt muss der Select des neuen Shapes gemacht werden 

            Try

                With appInstance.Worksheets(arrWsNames(3))
                    tmpShapes = .shapes
                    shpElement = tmpShapes.Item(hproj.name)
                    shpElement.Select()
                End With

            Catch ex As Exception
                shpElement = Nothing
            End Try



        Catch ex As Exception
            enableOnUpdate = True
            appInstance.ScreenUpdating = True
            Call MsgBox(ex.Message)

        End Try

        ' jetzt wird enableOnupdate wieder auf True gesetzt 
        enableOnUpdate = True
        appInstance.ScreenUpdating = True

    End Sub

    Private Sub frmConfirmEditRess_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed

        
        frmCoord(PTfrm.editRess, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.editRess, PTpinfo.left) = Me.Left

        ' jetzt wird auf Tabelle 1 zurückgewechselt ; es ist notwendig, das zu machen, bevor das Shape geschrieben wird; 
        ' andernfalls werden die Shape Koordinaten leicht verändert 
        ' und es ist notwendig es hier zu machen, weil das von Schliessen, Abbrechen, OK her erreicht werden muss 

        If Not okButtonClicked Then
            appInstance.ScreenUpdating = False
        End If

        With appInstance.Worksheets(arrWsNames(3))
            .activate()
        End With

        If Not okButtonClicked Then
            ' jetzt wird enableOnupdate wieder auf True gesetzt 
            enableOnUpdate = True
            appInstance.ScreenUpdating = True
        End If

    End Sub


    Private Sub frmConfirmEditRess_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.Top = frmCoord(PTfrm.editRess, PTpinfo.top)
        Me.Left = frmCoord(PTfrm.editRess, PTpinfo.left)

        okButtonClicked = False


    End Sub
End Class