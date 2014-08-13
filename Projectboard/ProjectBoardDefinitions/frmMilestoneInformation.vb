Public Class frmMilestoneInformation

    Public bewertungsListe As SortedList(Of String, clsBewertung)


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        bewertungsListe = New SortedList(Of String, clsBewertung)
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub frmMilestoneInformation_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.msInfo, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.msInfo, PTpinfo.left) = Me.Left

        Call awinDeleteProjectChildShapes(1)
        Call awinDeSelect()



    End Sub


    Private Sub frmMilestoneInformation_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.Top = CInt(frmCoord(PTfrm.msInfo, PTpinfo.top))
        Me.Left = CInt(frmCoord(PTfrm.msInfo, PTpinfo.left))


        If bewertungsListe.Count > 0 Then

            With bewertungsListe.ElementAt(0).Value
                Dim farbe As System.Drawing.Color = Drawing.Color.FromArgb(CInt(.color))

               
                bewertungsText.Text = .description

            End With

        Else

            Dim farbe As System.Drawing.Color = Drawing.Color.FromArgb(CInt(awinSettings.AmpelNichtBewertet))

            
            bewertungsText.Text = "es existiert noch keine Bewertung ...."



        End If

    End Sub

    'Private Sub sliderBewertungen_Scroll(sender As Object, e As EventArgs)



    '    With bewertungsListe.ElementAt(sliderBewertungen.Value).Value

    '        Dim farbe As System.Drawing.Color = Drawing.Color.FromArgb(.color)

    '        bewertungsDatum.Text = .datum.ToShortDateString
    '        'With ampelKreis

    '        '    .BackColor = farbe
    '        '    .FillColor = farbe
    '        '    .FillStyle = PowerPacks.FillStyle.Solid
    '        '    .BorderColor = farbe

    '        'End With
    '        'bewertungsFarbe.BackColor = farbe
    '        bewertungsText.Text = .description
    '        'bewerterName.Text = .bewerterName
    '    End With

    'End Sub

    'Private Sub sliderBewertungen_ValueChanged(sender As Object, e As EventArgs)

    '    With bewertungsListe.ElementAt(sliderBewertungen.Value).Value
    '        Dim farbe As System.Drawing.Color = Drawing.Color.FromArgb(.color)

    '        bewertungsDatum.Text = .datum.ToShortDateString
    '        'With ampelKreis

    '        '    .BackColor = farbe
    '        '    .FillColor = farbe
    '        '    .FillStyle = PowerPacks.FillStyle.Solid
    '        '    .BorderColor = farbe

    '        'End With
    '        'bewertungsFarbe.BackColor = farbe
    '        bewertungsText.Text = .description
    '        'bewerterName.Text = .bewerterName
    '    End With

    'End Sub

    Protected Overrides Sub Finalize()

        MyBase.Dispose(False)

    End Sub


    Private Sub nextButton_Click(sender As Object, e As EventArgs)
        Call MsgBox("noch nicht implementiert")
    End Sub

    Private Sub prevButton_Click_1(sender As Object, e As EventArgs)
        Call MsgBox("noch nicht implementiert")
    End Sub

    
End Class