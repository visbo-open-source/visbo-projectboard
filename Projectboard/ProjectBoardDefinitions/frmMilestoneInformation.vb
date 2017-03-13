Public Class frmMilestoneInformation

    Public bewertungsListe As SortedList(Of String, clsBewertung)
    Public milestone As clsMeilenstein
    Public curProject As clsProjekt


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        bewertungsListe = New SortedList(Of String, clsBewertung)
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub frmMilestoneInformation_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        frmCoord(PTfrm.msInfo, PTpinfo.top) = Me.Top
        frmCoord(PTfrm.msInfo, PTpinfo.left) = Me.Left

        'Call awinDeleteProjectChildShapes(1)
        Call awinDeSelect()



    End Sub


    Private Sub frmMilestoneInformation_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.Top = CInt(frmCoord(PTfrm.msInfo, PTpinfo.top))
        Me.Left = CInt(frmCoord(PTfrm.msInfo, PTpinfo.left))

        Me.showOrigItem.Checked = awinSettings.showOrigName

        rdbDeliverables.Checked = True

        Dim tmpstr() As String = Me.bewertungsText.Text.Split(New Char() {CChar(vbLf), CChar(vbCr)})
        Me.bewertungsText.Lines = tmpstr



    End Sub

    Private Sub fuelleTextBox()

        If bewertungsListe.Count > 0 Then

            With bewertungsListe.ElementAt(0).Value
                Dim farbe As System.Drawing.Color = Drawing.Color.FromArgb(CInt(.color))



                ' Änderung tk: die Zeilen, die durch CRLF getrennt sind, sollen auch so dargestellt werden 
                Dim tmpstr() As String
                Dim tmpDeliverables As String = milestone.getAllDeliverables
                If rdbDeliverables.Checked Then
                    tmpstr = tmpDeliverables.Split(New Char() {CChar(vbLf), CChar(vbCr)}, 100)
                Else
                    tmpstr = .description.Split(New Char() {CChar(vbLf), CChar(vbCr)}, 100)
                End If

                If tmpstr.Length > 0 Then
                    For i As Integer = 1 To tmpstr.Length
                        bewertungsText.Lines(i - 1) = tmpstr(i - 1)
                    Next
                Else
                    If rdbDeliverables.Checked Then
                        bewertungsText.Text = tmpDeliverables
                    Else
                        bewertungsText.Text = .description
                    End If

                End If


            End With

        Else

            Dim farbe As System.Drawing.Color = Drawing.Color.FromArgb(CInt(awinSettings.AmpelNichtBewertet))

            If rdbDeliverables.Checked Then
                bewertungsText.Text = ""
            Else
                bewertungsText.Text = ""
            End If




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



    
    ''' <summary>
    ''' zeigt den urspünglichen Meilenstein-Namen aus Rplan oder anderem PM-System an 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub showOrigItem_CheckedChanged(sender As Object, e As EventArgs) Handles showOrigItem.CheckedChanged

        awinSettings.showOrigName = showOrigItem.Checked

        If showOrigItem.Checked = True Then
            
            resultName.Text = milestone.originalName
            
        Else
            resultName.Text = milestone.name
        End If
    End Sub

    
    ''' <summary>
    ''' es reicht , wenn die Behandlung für einen Radio-Button gemacht wird. 
    ''' fillTextfenster füllt den Text in Abhängigkeit, ob Ergebnisse oder Erläuterung gecheckt sind 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub rdbDeliverables_CheckedChanged(sender As Object, e As EventArgs) Handles rdbDeliverables.CheckedChanged

        Call fuelleTextBox()

    End Sub

    
End Class