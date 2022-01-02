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


            ''Änderung 30.7.13 Screenupdating = false gesetzt , damit das Geflacker aufhört 
            appInstance.ScreenUpdating = False

            okButtonClicked = True ' damit bei Form_closed kein Enableonupdate bzw ScreenUpdate gemacht wird 
            MyBase.Close()

            ' die Diagramme müssen auf alle Fälle neu gezeichnet werden 
            ' dann müssen die Diagramme neu gezeichnet werden 
            Call awinNeuZeichnenDiagramme(3)

            ' jetzt muss der Select des neuen Shapes gemacht werden 

            Try

                With CType(appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT)), Microsoft.Office.Interop.Excel.Worksheet)
                    tmpShapes = .Shapes
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

        With appInstance.Workbooks.Item(myProjektTafel).Worksheets(arrWsNames(ptTables.MPT))
            .activate()
        End With

        If Not okButtonClicked Then
            ' jetzt wird enableOnupdate wieder auf True gesetzt 
            enableOnUpdate = True
            appInstance.ScreenUpdating = True
        End If

    End Sub


    Private Sub frmConfirmEditRess_Load(sender As Object, e As EventArgs) Handles Me.Load

        Call getFrmPosition(PTfrm.editRess, Top, Left)

        okButtonClicked = False


    End Sub
End Class