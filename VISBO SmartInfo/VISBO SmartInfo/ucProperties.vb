Public Class ucProperties

    Private Sub ucProperties_Leave(sender As Object, e As EventArgs) Handles Me.Leave

    End Sub


    Private Sub ucProperties_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ' label resize
        eleName.MaximumSize = New Drawing.Size(Me.Width - eleName.Margin.Left - eleName.Margin.Right - eleName.Location.X, eleName.MaximumSize.Height)

    End Sub

    Private Sub ucProperties_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If englishLanguage Then
            With Me
                .labelAmpel.Text = "Traffic Light:"
                .labelDate.Text = "Date:"
                .labelDeliver.Text = "Deliverables:"
                .labelRespons.Text = "Responsible:"
            End With
        Else
            With Me
                .labelAmpel.Text = "Ampel:"
                .labelDate.Text = "Datum:"
                .labelDeliver.Text = "Leistungsumfänge:"
                .labelRespons.Text = "Verantwortlich:"
            End With
        End If
    End Sub

    ''' <summary>
    ''' blendet die benötigten Darstellungs-Elemente ein bzw aus und positioniert dieses 
    ''' </summary>
    ''' <param name="isOn"></param>
    ''' <remarks></remarks>
    Public Sub symbolMode(ByVal isOn As Boolean)
        Dim tmpLocation As New System.Drawing.Point
        Dim tmpSize As New System.Drawing.Size

        If isOn Then
            ' der Symbol Mode
            With labelDate
                .Visible = False
            End With

            With labelRespons
                .Visible = False
            End With

            With labelAmpel
                .Visible = False
            End With

            With eleAmpel
                .Visible = False
            End With

            With eleAmpelText
                .Visible = True
                tmpLocation.X = 5
                tmpLocation.Y = 52
                .Location = tmpLocation
                .BorderStyle = Windows.Forms.BorderStyle.None
                tmpSize.Height = 400
                'tmpSize.Width = 276
                .Size = tmpSize
            End With

            With labelDeliver
                .Visible = False
            End With

            With eleDeliverables
                .Visible = False
            End With

        Else
            ' der Normal-Mode
            With labelDate
                .Visible = True
                tmpLocation.X = 5
                tmpLocation.Y = 52
                .Location = tmpLocation
            End With

            With labelRespons
                .Visible = True
            End With

            With labelAmpel
                .Visible = True
            End With

            With eleAmpel
                .Visible = True
            End With

            With eleAmpelText
                .Visible = True
                tmpLocation.X = 10
                tmpLocation.Y = 144
                .Location = tmpLocation
                .BorderStyle = Windows.Forms.BorderStyle.FixedSingle
                tmpSize.Height = 139
                'tmpSize.Width = 276
                .Size = tmpSize
            End With

            With labelDeliver
                .Visible = True
            End With

            With eleDeliverables
                .Visible = True
            End With
        End If
    End Sub

    Private Sub eleAmpelText_MouseDoubleClick(sender As Object, e As Windows.Forms.MouseEventArgs) Handles eleAmpelText.MouseDoubleClick

    End Sub


    Private Sub eleAmpelText_TextChanged(sender As Object, e As EventArgs) Handles eleAmpelText.TextChanged

    End Sub
End Class
