Imports System.Windows.Forms

Public Class frmProjektAendern


    Private Sub frmProjektAendern_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub projectName_TextChanged(sender As Object, e As EventArgs) Handles projectName.TextChanged

    End Sub

    Private Sub vorlagenName_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Erloes_LostFocus(sender As Object, e As EventArgs) Handles Erloes.LostFocus

        With Me.Erloes
            If Not IsNumeric(.Text) Then
                MsgBox("bitte eine Zahl eingeben ")
                .Text = ""
                .Focus()
            ElseIf CType(.Text, Double) < 0 Then
                Call MsgBox(" der Erlös muss eine positive Dezimal-Zahl sein")
                .Text = ""
                .Focus()
            End If
        End With

    End Sub

    Private Sub sFit_LostFocus(sender As Object, e As EventArgs) Handles sFit.LostFocus
        With Me.sFit

            If Not IsNumeric(.Text) Then
                MsgBox("bitte eine Zahl zwischen 0. und 10 eingeben ")
                .Text = ""
                .Focus()
            ElseIf CType(.Text, Double) < 0 Or CType(.Text, Double) > 10 Then
                Call MsgBox(" der strategische Fit muss eine positive Dezimal-Zahl zwischen 0. und 10 sein")
                .Text = ""
                .Focus()
            Else
                Dim hfit As Double
                hfit = CType(.Text, Double)
                .Text = hfit.ToString("0.0")
            End If

        End With
    End Sub


    Private Sub risiko_LostFocus(sender As Object, e As EventArgs) Handles risiko.LostFocus

        With Me.risiko

            If Not IsNumeric(.Text) Then
                MsgBox("bitte eine Zahl zwischen 0. und 10 eingeben ")
                .Text = ""
                .Focus()
            ElseIf CType(.Text, Double) < 0 Or CType(.Text, Double) > 10 Then
                Call MsgBox(" das Risiko muss eine positive Dezimal-Zahl zwischen 0. und 10 sein")
                .Text = ""
                .Focus()
            Else
                Dim hrisk As Double
                hrisk = CType(.Text, Double)
                .Text = hrisk.ToString("0.0")
            End If

        End With

    End Sub

    Private Sub risiko_TextChanged(sender As Object, e As EventArgs) Handles risiko.TextChanged

    End Sub

    Private Sub ruleEngine_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles ruleEngine.LinkClicked
        Dim randomValue As Double

        With Me
            randomValue = appInstance.WorksheetFunction.RandBetween(1, 100) / 10
            .risiko.Text = randomValue.ToString("0.0")
            randomValue = appInstance.WorksheetFunction.RandBetween(1, 100) / 10
            .sFit.Text = randomValue.ToString("0.0")
        End With

    End Sub

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click

        With projectName
            If Len(.Text) < 1 Then
                'MsgBox("Name muss mindestens 1 Zeichen lang sein")
                .Text = ""
                .Focus()
                Exit Sub
            ElseIf IsNumeric(.Text) Then
                MsgBox("Zahlen sind nicht zugelassen")
                .Text = ""
                .Focus()
                Exit Sub
            End If
        End With

        DialogResult = System.Windows.Forms.DialogResult.OK
        MyBase.Close()


    End Sub

    Private Sub AbbrButton_Click(sender As Object, e As EventArgs) Handles AbbrButton.Click

        DialogResult = System.Windows.Forms.DialogResult.Cancel
        MyBase.Close()

    End Sub

    
    Private Sub businessUnit_SelectedIndexChanged(sender As Object, e As EventArgs) Handles businessUnit.SelectedIndexChanged

    End Sub
End Class